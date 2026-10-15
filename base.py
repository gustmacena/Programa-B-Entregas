# -*- coding: utf-8 -*-
"""Consultas à base criada pelo indexador (busca por caso, NF, remessa, CNPJ, placa, texto)."""
from __future__ import annotations

import base64
import json
import os
import re
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import indexador

FUSO = timedelta(hours=-3)  # horário de Brasília (as datas do Salesforce vêm em UTC)


def abrir(db_path: str | Path) -> sqlite3.Connection:
    uri = Path(db_path).resolve().as_uri() + "?mode=ro&immutable=1"
    conn = sqlite3.connect(uri, uri=True, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    versao = conn.execute("SELECT valor FROM meta WHERE chave='schema'").fetchone()
    if not versao or int(versao[0]) != indexador.VERSAO_SCHEMA:
        conn.close()
        raise RuntimeError("A base foi criada por outra versão do programa. Atualize a base.")
    return conn


def meta(conn, chave: str) -> str:
    r = conn.execute("SELECT valor FROM meta WHERE chave=?", (chave,)).fetchone()
    return r[0] if r else ""


def fmt_data(iso: str) -> str:
    """'2026-07-01T10:14:40.000+0000' (UTC) -> '01/07/2026 07:14'."""
    if not iso:
        return ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?", iso)
    if not m:
        return iso
    if m.group(4) is None:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    dt = datetime(*(int(x) for x in m.groups())) + FUSO
    return dt.strftime("%d/%m/%Y %H:%M")


def eh_data_iso(v: str) -> bool:
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}", v or ""))


# --------------------------------------------------------------------------- #
# Busca
# --------------------------------------------------------------------------- #
_PRIORIDADE = {"Caso": 0, "NF": 1, "NF + série": 1, "Remessa": 2, "CNPJ": 3, "Placa": 4,
               "Texto": 5, "Caso (parcial)": 6, "NF (parcial)": 6}


def buscar(conn, entrada: str, limite: int = 500) -> list[dict]:
    q = (entrada or "").strip()
    if not q:
        return []
    achados: dict[str, list[str]] = {}

    def add(sql: str, params: tuple, motivo: str):
        for (n,) in conn.execute(sql, params).fetchall():
            lst = achados.setdefault(n, [])
            if motivo not in lst:
                lst.append(motivo)

    if re.fullmatch(r"[\d.\-/\s]+", q):
        dig = indexador.so_digitos(q)
        m = re.fullmatch(r"(\d+)\s*[-/]\s*(\d{1,3})", q)
        if len(dig) == 14:
            add("SELECT numero FROM casos WHERE cnpj_num=?", (dig,), "CNPJ")
        elif m:
            add("SELECT DISTINCT numero FROM nfs WHERE nf_chave=? AND serie_chave=?",
                (indexador.chave_num(m.group(1)), indexador.chave_num(m.group(2))), "NF + série")
        elif dig:
            if len(dig) <= 8:
                add("SELECT numero FROM casos WHERE numero=?", (dig.zfill(8),), "Caso")
            ch = indexador.chave_num(dig)
            if ch:
                add("SELECT DISTINCT numero FROM nfs WHERE nf_chave=?", (ch,), "NF")
                add("SELECT DISTINCT numero FROM nfs WHERE remessa_chave=?", (ch,), "Remessa")
            if not achados and len(dig) >= 3:
                add("SELECT numero FROM casos WHERE numero LIKE ? LIMIT 200", (f"%{dig}%",), "Caso (parcial)")
                add("SELECT DISTINCT numero FROM nfs WHERE nf_chave LIKE ? LIMIT 200", (dig.lstrip("0") + "%",),
                    "NF (parcial)")
    else:
        if re.fullmatch(r"[A-Za-z]{3}-?\d[A-Za-z0-9]\d{2}", q):
            add("SELECT numero FROM casos WHERE placa_norm=?", (re.sub(r"[^0-9A-Za-z]", "", q).upper(),), "Placa")
        tokens = re.findall(r"\w+", q)
        if tokens:
            expr = " ".join('"%s"*' % t.replace('"', "") for t in tokens)
            try:
                add("SELECT c.numero FROM busca JOIN casos c ON c.rowid = busca.rowid "
                    "WHERE busca MATCH ? ORDER BY busca.rank LIMIT ?", (expr, limite), "Texto")
            except sqlite3.OperationalError:
                pass

    if not achados:
        return []
    numeros = list(achados)[:limite]
    marcas = ",".join("?" * len(numeros))
    linhas = conn.execute(
        f"""SELECT c.numero, c.criado, c.transportadora, c.status, c.motivo, c.semestre,
               (SELECT count(*) FROM nfs n WHERE n.numero=c.numero) AS n_nf,
               (SELECT count(DISTINCT doc_id) FROM anexos a WHERE a.numero=c.numero) AS n_anx,
               (SELECT count(*) FROM comentarios k WHERE k.numero=c.numero) AS n_com
            FROM casos c WHERE c.numero IN ({marcas})""",
        numeros,
    ).fetchall()
    res = []
    for r in linhas:
        d = dict(r)
        d["achado_por"] = ", ".join(achados[d["numero"]])
        d["_prio"] = min(_PRIORIDADE.get(m, 9) for m in achados[d["numero"]])
        res.append(d)
    res.sort(key=lambda d: d["criado"] or "", reverse=True)   # mais recentes primeiro
    res.sort(key=lambda d: d["_prio"])                        # e, antes, as buscas mais exatas
    return res


# --------------------------------------------------------------------------- #
# Detalhe de um caso
# --------------------------------------------------------------------------- #
def caso(conn, numero: str) -> dict | None:
    r = conn.execute("SELECT * FROM casos WHERE numero=?", (numero,)).fetchone()
    if not r:
        return None
    d = dict(r)
    dados = d.pop("dados")
    zd = base64.b64decode(meta(conn, "zdict") or "")
    d["campos"] = indexador.descomprimir_json(dados, zd) if dados else {}
    return d


def casos_relacionados(conn, c: dict) -> dict:
    pai = None
    if c.get("parent_id"):
        r = conn.execute("SELECT numero FROM casos WHERE case_id=?", (c["parent_id"],)).fetchone()
        pai = r[0] if r else None
    filhos = [x[0] for x in conn.execute(
        "SELECT numero FROM casos WHERE parent_id=? AND parent_id<>'' ORDER BY numero", (c["case_id"],))]
    return {"pai": pai, "filhos": filhos}


def comentarios(conn, numero: str) -> list[dict]:
    return [dict(r) for r in conn.execute(
        "SELECT * FROM comentarios WHERE numero=? ORDER BY criado, rowid", (numero,))]


def nfs(conn, numero: str) -> list[dict]:
    res = []
    for r in conn.execute("SELECT * FROM nfs WHERE numero=? ORDER BY nf_chave, serie_chave", (numero,)):
        d = dict(r)
        outros = [x[0] for x in conn.execute(
            "SELECT DISTINCT numero FROM nfs WHERE nf_chave=? AND serie_chave=? AND numero<>? "
            "ORDER BY numero LIMIT 31",
            (d["nf_chave"], d["serie_chave"], numero))]
        d["outros_mais"] = len(outros) > 30
        d["outros_casos"] = outros[:30]
        res.append(d)
    return res


def anexos(conn, numero: str, raiz: str | Path) -> list[dict]:
    pastas = dict(conn.execute("SELECT nome, pasta FROM semestres").fetchall())
    res = []
    for r in conn.execute(
        "SELECT * FROM anexos WHERE numero=? ORDER BY criado, nome, versao", (numero,)
    ):
        d = dict(r)
        d["caminho"] = str(Path(raiz) / pastas.get(d["semestre"], "") / d["arquivo"])
        d["existe"] = os.path.exists(d["caminho"])
        res.append(d)
    return res


def pacote_caso(conn, numero: str, raiz: str | Path) -> dict | None:
    """Tudo o que a tela precisa de um caso (roda fora da thread da tela)."""
    c = caso(conn, numero)
    if not c:
        return None
    return {
        "caso": c,
        "rel": casos_relacionados(conn, c),
        "comentarios": comentarios(conn, numero),
        "nfs": nfs(conn, numero),
        "anexos": anexos(conn, numero, raiz),
    }
