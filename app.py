# -*- coding: utf-8 -*-
"""
Consulta Legado — busca de casos, NFs e anexos da extração do sistema descomissionado.

Execute este arquivo (ou o .exe gerado) e digite o número do caso ou da NF.
"""
from __future__ import annotations

import html
import os
import subprocess
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import (QBuffer, QByteArray, QIODevice, QObject, QRunnable, QSettings, QSize, Qt,
                            QThread, QThreadPool, QTimer, QUrl, Signal)
from PySide6.QtGui import (QAction, QBrush, QColor, QDesktopServices, QFont, QIcon, QImage, QImageReader,
                           QPainter, QPalette, QPen, QPixmap, QTransform)
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QDialog, QFileDialog, QFrame,
                               QGraphicsPixmapItem, QGraphicsScene, QGraphicsView, QHBoxLayout, QHeaderView,
                               QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
                               QPlainTextEdit, QProgressBar, QPushButton, QSizePolicy, QSplitter,
                               QStackedWidget, QTableWidget, QTableWidgetItem, QTabWidget, QTextBrowser,
                               QToolButton, QVBoxLayout, QWidget)

import base
import indexador
from tema import (BORDA, COR_AMARELO_CLARO, COR_BRANCO, COR_CARTAO, COR_CINZA, COR_FUNDO, COR_MARROM,
                  COR_VERDE, COR_VERMELHO, DOURADO, MARROM_HOVER, TEXTO, TEXTO_SUAVE, VERDE_HOVER)

APP_NOME = "Consulta Legado"
EXT_IMAGEM = {"jpg", "jpeg", "jfif", "png", "gif", "bmp", "webp", "tif", "tiff"}

ESTILO = f"""
QMainWindow, QDialog {{ background: {COR_FUNDO}; }}
QWidget {{ font-family: 'Segoe UI', 'Noto Sans', 'DejaVu Sans'; font-size: 10pt; color: {TEXTO}; }}
QLabel {{ background: transparent; }}
QLabel#titulo_app {{ font-size: 18pt; font-weight: 700; color: {COR_MARROM}; }}
QLabel#sub_app {{ font-size: 9pt; color: {COR_MARROM}; }}
QLabel#cartao {{ background: {COR_BRANCO}; border: 1px solid {DOURADO}; border-radius: 16px; padding: 16px 24px; }}
QLabel#cab_caso {{ background: {COR_CINZA}; border: 1px solid {DOURADO}; border-radius: 14px; padding: 10px 16px; }}
QMenuBar {{ background: {COR_MARROM}; color: white; }}
QMenuBar::item {{ background: transparent; padding: 6px 14px; }}
QMenuBar::item:selected {{ background: {MARROM_HOVER}; }}
QMenu {{ background: {COR_CINZA}; color: {COR_MARROM}; border: 1px solid {DOURADO}; }}
QMenu::item {{ padding: 7px 26px; }}
QMenu::item:selected {{ background: {COR_AMARELO_CLARO}; color: {COR_MARROM}; }}
QStatusBar {{ background: transparent; color: {COR_MARROM}; }}
QStatusBar QLabel {{ color: {COR_MARROM}; background: transparent; }}
QToolTip {{ background: {COR_CINZA}; color: {COR_MARROM}; border: 1px solid {DOURADO}; padding: 4px; }}
QLineEdit {{ background: white; border: 2px solid {BORDA}; border-radius: 10px; padding: 6px 10px; }}
QLineEdit#busca {{ font-size: 15pt; padding: 8px 14px; border: 2px solid {COR_MARROM}; border-radius: 14px;
                   background: white; }}
QLineEdit#busca:focus {{ border-color: {COR_VERDE}; }}
QPushButton, QToolButton {{ background: {COR_MARROM}; color: white; border: none; border-radius: 12px;
                            padding: 7px 14px; font-weight: 600; }}
QPushButton:hover, QToolButton:hover {{ background: {MARROM_HOVER}; }}
QPushButton:disabled, QToolButton:disabled {{ background: #EBCB8C; color: #A98A62; }}
QPushButton#primario {{ background: {COR_VERDE}; color: white; font-size: 12pt; font-weight: 700;
                        padding: 10px 30px; border-radius: 14px; }}
QPushButton#primario:hover {{ background: {VERDE_HOVER}; }}
QTableWidget, QListWidget, QTextBrowser, QPlainTextEdit {{ background: white; border: 1px solid {DOURADO};
        border-radius: 12px; alternate-background-color: {COR_CINZA}; gridline-color: {BORDA}; }}
QTableWidget::item:selected, QListWidget::item:selected {{ background: {COR_AMARELO_CLARO}; color: {COR_MARROM}; }}
QHeaderView::section {{ background: {COR_MARROM}; color: white; border: none;
                        border-right: 1px solid {MARROM_HOVER}; padding: 7px; font-weight: 600; }}
QTabWidget::pane {{ border: 1px solid {DOURADO}; background: {COR_CINZA}; border-radius: 12px; top: -1px; }}
QTabBar::tab {{ background: {COR_CARTAO}; color: {COR_MARROM}; padding: 9px 20px; border: 1px solid {DOURADO};
                border-bottom: none; border-top-left-radius: 12px; border-top-right-radius: 12px;
                margin-right: 3px; font-weight: 600; }}
QTabBar::tab:selected {{ background: {COR_CINZA}; color: {COR_MARROM}; }}
QTabBar::tab:hover:!selected {{ background: #FFD36E; }}
QSplitter::handle {{ background: {DOURADO}; }}
QProgressBar {{ border: 1px solid {DOURADO}; border-radius: 9px; background: white; text-align: center;
                min-height: 20px; color: {COR_MARROM}; }}
QProgressBar::chunk {{ background: {COR_VERDE}; border-radius: 8px; }}
QMessageBox {{ background: {COR_FUNDO}; }}
"""


def recurso(nome: str) -> str:
    """Caminho de um arquivo que acompanha o programa (funciona no .py e no .exe do PyInstaller)."""
    pasta = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(pasta / nome)


def esc(s) -> str:
    return html.escape(str(s or ""))


def esc_texto(s) -> str:
    return esc(s).replace("\n", "<br>")


def humano(n: int) -> str:
    n = float(n or 0)
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024
    return ""


def abrir_externo(caminho: str):
    QDesktopServices.openUrl(QUrl.fromLocalFile(caminho))


def mostrar_na_pasta(caminho: str):
    if sys.platform == "win32":
        subprocess.Popen(["explorer", "/select,", os.path.normpath(caminho)])
    else:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(caminho).parent)))


# --------------------------------------------------------------------------- #
# Leitura de arquivos em segundo plano (os arquivos estão no OneDrive e podem
# precisar ser baixados — isso não pode travar a tela)
# --------------------------------------------------------------------------- #
class _Sinal(QObject):
    pronto = Signal(int, str, object)  # geração, chave, resultado


class TarefaArquivo(QRunnable):
    def __init__(self, sinal: _Sinal, geracao: int, chave: str, caminho: str, modo: str, lado: int = 0):
        super().__init__()
        self.sinal, self.geracao, self.chave = sinal, geracao, chave
        self.caminho, self.modo, self.lado = caminho, modo, lado

    def run(self):
        res = None
        try:
            dados = Path(self.caminho).read_bytes()
            if self.modo in ("img_mini", "img_total"):
                buf = QBuffer()
                buf.setData(QByteArray(dados))
                buf.open(QIODevice.OpenModeFlag.ReadOnly)
                rd = QImageReader(buf)
                rd.setAutoTransform(True)  # respeita a rotação gravada na foto (EXIF)
                if self.modo == "img_mini":
                    sz = rd.size()
                    if sz.isValid():
                        sz.scale(self.lado, self.lado, Qt.AspectRatioMode.KeepAspectRatio)
                        rd.setScaledSize(sz)
                img = rd.read()
                res = None if img.isNull() else img
            else:
                res = dados
        except Exception:
            res = None
        try:
            self.sinal.pronto.emit(self.geracao, self.chave, res)
        except RuntimeError:
            pass  # janela já fechada


def tile(texto: str, cor: str = TEXTO_SUAVE, lado: int = 128, fundo: str = "#F6F2EC") -> QPixmap:
    pm = QPixmap(lado, lado)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setBrush(QColor(fundo))
    p.setPen(QPen(QColor("#E3C48A"), 2))
    p.drawRoundedRect(4, 4, lado - 8, lado - 8, 10, 10)
    p.setPen(QColor(cor))
    f = QFont()
    f.setPointSize(16)
    f.setBold(True)
    p.setFont(f)
    p.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, texto)
    p.end()
    return pm


def quadro(img: QImage, lado: int = 128) -> QPixmap:
    """Miniatura centralizada num quadro fixo."""
    pm = QPixmap(lado, lado)
    pm.fill(QColor("#F6F2EC"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    sc = img.scaled(lado - 8, lado - 8, Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation)
    p.drawImage((lado - sc.width()) // 2, (lado - sc.height()) // 2, sc)
    p.setPen(QPen(QColor("#E3C48A"), 1))
    p.drawRect(0, 0, lado - 1, lado - 1)
    p.end()
    return pm


# --------------------------------------------------------------------------- #
# Visor de imagens (zoom com a rodinha, arrastar, girar)
# --------------------------------------------------------------------------- #
class VisorImagem(QGraphicsView):
    def __init__(self):
        super().__init__()
        self._cena = QGraphicsScene(self)
        self.setScene(self._cena)
        self._item = QGraphicsPixmapItem()
        self._item.setTransformationMode(Qt.TransformationMode.SmoothTransformation)
        self._cena.addItem(self._item)
        self._img: QImage | None = None
        self._ajustado = True
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setBackgroundBrush(QBrush(QColor("#2B1A10")))
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setRenderHints(QPainter.RenderHint.SmoothPixmapTransform | QPainter.RenderHint.Antialiasing)

    def definir(self, img: QImage | None):
        self._img = img
        self._item.setPixmap(QPixmap.fromImage(img) if img else QPixmap())
        self._cena.setSceneRect(self._item.boundingRect())
        self.ajustar()

    def ajustar(self):
        self.resetTransform()
        if self._img:
            self.fitInView(self._item, Qt.AspectRatioMode.KeepAspectRatio)
        self._ajustado = True

    def tamanho_real(self):
        self.resetTransform()
        self._ajustado = False

    def zoom(self, fator: float):
        self._ajustado = False
        self.scale(fator, fator)

    def girar(self, graus: int):
        if not self._img:
            return
        self._img = self._img.transformed(QTransform().rotate(graus), Qt.TransformationMode.SmoothTransformation)
        self._item.setPixmap(QPixmap.fromImage(self._img))
        self._cena.setSceneRect(self._item.boundingRect())
        self.ajustar()

    def wheelEvent(self, ev):
        self.zoom(1.18 if ev.angleDelta().y() > 0 else 1 / 1.18)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        if self._ajustado:
            self.ajustar()

    def mouseDoubleClickEvent(self, ev):
        self.tamanho_real() if self._ajustado else self.ajustar()


# --------------------------------------------------------------------------- #
# Painel que mostra o arquivo selecionado (imagem, PDF ou outro)
# --------------------------------------------------------------------------- #
class VisorArquivo(QWidget):
    def __init__(self):
        super().__init__()
        self._sinal = _Sinal()
        self._sinal.pronto.connect(self._chegou)
        self._geracao = 0
        self.atual: dict | None = None
        self._modo = ""

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        barra = QHBoxLayout()
        self.b_abrir = QPushButton("Abrir no programa padrão")
        self.b_pasta = QPushButton("Mostrar na pasta")
        self.b_esq = QToolButton(text="⟲ Girar")
        self.b_dir = QToolButton(text="Girar ⟳")
        self.b_ajust = QToolButton(text="Ajustar")
        self.b_mais = QToolButton(text="＋")
        self.b_menos = QToolButton(text="－")
        for b in (self.b_abrir, self.b_pasta):
            barra.addWidget(b)
        barra.addStretch()
        for b in (self.b_esq, self.b_dir, self.b_ajust, self.b_menos, self.b_mais):
            barra.addWidget(b)
        lay.addLayout(barra)

        self.info = QLabel()
        self.info.setWordWrap(True)
        self.info.setStyleSheet(f"color:{TEXTO_SUAVE}; padding:2px 0;")
        lay.addWidget(self.info)

        self.pilha = QStackedWidget()
        self.msg = QLabel()
        self.msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.msg.setWordWrap(True)
        self.msg.setStyleSheet("background:#F6F2EC; border:2px dashed #E0A21A; border-radius:14px; padding:20px;")
        self.img = VisorImagem()
        self.pdf_doc = QPdfDocument(self)
        self.pdf = QPdfView()
        self.pdf.setPageMode(QPdfView.PageMode.MultiPage)
        self.pdf.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.pdf.setDocument(self.pdf_doc)
        self._pdf_buf: QBuffer | None = None
        self._pdf_dados: QByteArray | None = None
        for w in (self.msg, self.img, self.pdf):
            self.pilha.addWidget(w)
        lay.addWidget(self.pilha, 1)

        self.b_abrir.clicked.connect(lambda: self.atual and abrir_externo(self.atual["caminho"]))
        self.b_pasta.clicked.connect(lambda: self.atual and mostrar_na_pasta(self.atual["caminho"]))
        self.b_esq.clicked.connect(lambda: self.img.girar(-90))
        self.b_dir.clicked.connect(lambda: self.img.girar(90))
        self.b_ajust.clicked.connect(self._ajustar)
        self.b_mais.clicked.connect(lambda: self._zoom(1.25))
        self.b_menos.clicked.connect(lambda: self._zoom(0.8))
        self.limpar()

    # -- controles ---------------------------------------------------------- #
    def _ajustar(self):
        if self._modo == "pdf":
            self.pdf.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        else:
            self.img.ajustar()

    def _zoom(self, f):
        if self._modo == "pdf":
            self.pdf.setZoomMode(QPdfView.ZoomMode.Custom)
            self.pdf.setZoomFactor(max(0.1, self.pdf.zoomFactor() * f))
        else:
            self.img.zoom(f)

    def _controles(self, imagem: bool, pdf: bool, existe: bool):
        self.b_abrir.setEnabled(existe)
        self.b_pasta.setEnabled(existe)
        for b in (self.b_esq, self.b_dir):
            b.setEnabled(imagem)
        for b in (self.b_ajust, self.b_mais, self.b_menos):
            b.setEnabled(imagem or pdf)

    def _mensagem(self, texto: str):
        self._modo = ""
        self.msg.setText(texto)
        self.pilha.setCurrentWidget(self.msg)

    def limpar(self):
        self._geracao += 1
        self.atual = None
        self.info.setText("")
        self._controles(False, False, False)
        self._mensagem("<span style='color:#8A6A5C;font-size:12pt'>Selecione um anexo na lista ao lado para visualizar.</span>")

    # -- exibição ----------------------------------------------------------- #
    def mostrar(self, a: dict):
        self._geracao += 1
        self.atual = a
        caminho = a["caminho"]
        ext = (a["ext"] or "").lower()
        extra = f" · também vinculado ao(s) caso(s) {a['compartilhado']}" if a.get("compartilhado") else ""
        versao = " · <b>versão anterior</b> do arquivo" if not a["mais_recente"] else ""
        self.info.setText(
            f"<b>{esc(a['nome'])}</b> · {esc(ext.upper())} · {humano(a['tamanho'])} · "
            f"enviado por {esc(a['criado_por'])} em {base.fmt_data(a['criado'])}{extra}{versao}"
        )
        existe = a["existe"]
        self._controles(False, False, existe)
        if not existe:
            self._mensagem(
                "<div style='font-size:13pt;color:#b3261e'><b>Arquivo não encontrado</b></div><br>"
                "<span style='color:#555'>Procurei em:</span><br>"
                f"<span style='font-size:9pt'>{esc(caminho)}</span><br><br>"
                "<span style='color:#555'>Se a pasta está no OneDrive, confirme que ela está sincronizada.</span>"
            )
            return
        if ext in EXT_IMAGEM:
            self._mensagem("<span style='font-size:12pt;color:#8A6A5C'>Carregando imagem…</span>")
            QThreadPool.globalInstance().start(
                TarefaArquivo(self._sinal, self._geracao, "img", caminho, "img_total"))
        elif ext == "pdf":
            self._mensagem("<span style='font-size:12pt;color:#8A6A5C'>Carregando PDF…</span>")
            QThreadPool.globalInstance().start(
                TarefaArquivo(self._sinal, self._geracao, "pdf", caminho, "bytes"))
        else:
            self._mensagem(
                f"<div style='font-size:13pt'><b>Arquivo .{esc(ext)}</b></div><br>"
                "Este tipo de arquivo não pode ser exibido aqui.<br>Use o botão "
                "<b>Abrir no programa padrão</b>."
            )

    def _chegou(self, geracao: int, chave: str, res):
        if geracao != self._geracao:
            return  # o usuário já escolheu outro arquivo
        if res is None:
            self._mensagem("<span style='color:#b3261e'>Não foi possível ler este arquivo.</span>")
            return
        if chave == "img":
            self.img.definir(res)
            self._modo = "img"
            self.pilha.setCurrentWidget(self.img)
            self._controles(True, False, True)
        else:
            self.pdf_doc.close()
            self._pdf_dados = QByteArray(res)
            self._pdf_buf = QBuffer(self._pdf_dados)
            self._pdf_buf.open(QIODevice.OpenModeFlag.ReadOnly)
            self.pdf_doc.load(self._pdf_buf)
            if self.pdf_doc.status() != QPdfDocument.Status.Ready:
                self._mensagem("<span style='color:#b3261e'>Não foi possível abrir este PDF.</span>")
                return
            self.pdf.setZoomMode(QPdfView.ZoomMode.FitToWidth)
            self._modo = "pdf"
            self.pilha.setCurrentWidget(self.pdf)
            self._controles(False, True, True)


# --------------------------------------------------------------------------- #
# Aba de anexos: galeria de miniaturas + visor
# --------------------------------------------------------------------------- #
class AbaAnexos(QWidget):
    LADO = 128

    def __init__(self):
        super().__init__()
        self._sinal = _Sinal()
        self._sinal.pronto.connect(self._miniatura)
        self._geracao = 0
        self._itens: dict[str, QListWidgetItem] = {}
        self._pdf_mini = QPdfDocument(self)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        sp = QSplitter(Qt.Orientation.Horizontal)
        self.lista = QListWidget()
        self.lista.setViewMode(QListWidget.ViewMode.IconMode)
        self.lista.setIconSize(QSize(self.LADO, self.LADO))
        self.lista.setGridSize(QSize(self.LADO + 28, self.LADO + 52))
        self.lista.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.lista.setMovement(QListWidget.Movement.Static)
        self.lista.setWordWrap(True)
        self.lista.setSpacing(2)
        self.visor = VisorArquivo()
        sp.addWidget(self.lista)
        sp.addWidget(self.visor)
        sp.setStretchFactor(1, 1)
        sp.setSizes([350, 900])
        lay.addWidget(sp)
        self.lista.currentItemChanged.connect(self._escolheu)

    def carregar(self, anexos: list[dict]):
        self._geracao += 1
        self.lista.clear()
        self._itens.clear()
        self.visor.limpar()
        pool = QThreadPool.globalInstance()
        for i, a in enumerate(anexos):
            existe = a["existe"]
            ext = (a["ext"] or "").lower()
            nome = a["nome"] if a["mais_recente"] else "↺ " + a["nome"]
            it = QListWidgetItem(nome)
            it.setData(Qt.ItemDataRole.UserRole, a)
            it.setToolTip(f"{a['nome']}\n{humano(a['tamanho'])}" + ("" if existe else "\n(arquivo não encontrado)"))
            it.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
            if not existe:
                it.setIcon(QIcon(tile("SEM\nARQUIVO", "#b3261e", self.LADO, "#fdecea")))
            elif ext in EXT_IMAGEM:
                it.setIcon(QIcon(tile("…", TEXTO_SUAVE, self.LADO)))
            elif ext == "pdf":
                it.setIcon(QIcon(tile("PDF", "#b3261e", self.LADO)))
            else:
                it.setIcon(QIcon(tile(ext.upper()[:5], TEXTO_SUAVE, self.LADO)))
            self.lista.addItem(it)
            chave = str(i)
            self._itens[chave] = it
            if existe and ext in EXT_IMAGEM:
                pool.start(TarefaArquivo(self._sinal, self._geracao, chave, a["caminho"], "img_mini", self.LADO))
            elif existe and ext == "pdf":
                pool.start(TarefaArquivo(self._sinal, self._geracao, chave, a["caminho"], "bytes"))
        if anexos:
            primeiro = next((i for i, a in enumerate(anexos) if a["existe"]), 0)
            self.lista.setCurrentRow(primeiro)

    def _escolheu(self, atual, _antes):
        if atual is not None:
            self.visor.mostrar(atual.data(Qt.ItemDataRole.UserRole))

    def _miniatura(self, geracao: int, chave: str, res):
        if geracao != self._geracao or res is None:
            return
        it = self._itens.get(chave)
        if it is None:
            return
        if isinstance(res, QImage):
            it.setIcon(QIcon(quadro(res, self.LADO)))
            return
        # PDF: renderiza a 1ª página
        try:
            dados = QByteArray(res)
            buf = QBuffer(dados)
            buf.open(QIODevice.OpenModeFlag.ReadOnly)
            doc = self._pdf_mini
            doc.close()
            doc.load(buf)
            if doc.status() == QPdfDocument.Status.Ready and doc.pageCount() > 0:
                pt = doc.pagePointSize(0)
                esc_ = (self.LADO * 2) / max(pt.width(), pt.height(), 1)
                img = doc.render(0, QSize(int(pt.width() * esc_), int(pt.height() * esc_)))
                if not img.isNull():
                    it.setIcon(QIcon(quadro(img, self.LADO)))
            doc.close()
        except Exception:
            pass


# --------------------------------------------------------------------------- #
# Montagem do HTML do resumo e dos comentários
# --------------------------------------------------------------------------- #
GRUPOS = [
    ("Identificação", [
        ("CaseNumber", "Nº do caso"), ("Status", "Status"), ("StatusAprovacao__c", "Aprovação"),
        ("Priority", "Prioridade"), ("Origin", "Origem"), ("Type", "Tipo"),
        ("CreatedBy.Name", "Aberto por"), ("LastModifiedBy.Name", "Última alteração por")]),
    ("Transportadora e veículo", [
        ("Account.Name", "Transportadora"), ("Account.TaxNumber__c", "CNPJ"),
        ("BP_Transportadora__c", "BP transportadora"), ("PlacaVeiculo__c", "Placa"),
        ("NomeMotorista__c", "Motorista"), ("Modelo__c", "Modelo do veículo"),
        ("ContactEmail", "E-mail"), ("ContactPhone", "Telefone"), ("ContactMobile", "Celular")]),
    ("Classificação", [
        ("MotivoView__c", "Motivo"), ("SubmotivoView__c", "Submotivo"),
        ("Classificacao_Interna_Motivo__c", "Classificação interna – motivo"),
        ("Classificacao_Interna_Submotivo__c", "Classificação interna – submotivo"),
        ("Regional__c", "Regional"), ("Estabelecimento__c", "Estabelecimento")]),
    ("Datas", [
        ("CreatedDate", "Abertura"), ("DataDeChegada__c", "Chegada"),
        ("HorarioEntradaCD__c", "Entrada no CD"), ("HorarioSaidaCD__c", "Saída do CD"),
        ("DataReentrega__c", "Reentrega"), ("ClosedDate", "Fechamento"),
        ("LastModifiedDate", "Última modificação")]),
    ("Valores e quantidades", [
        ("ValorDescarga__c", "Valor da descarga"), ("QuantidadePallets__c", "Pallets"),
        ("QuantidadeDiarias__c", "Diárias"), ("QuantidadeReentrega__c", "Reentregas"),
        ("Senha__c", "Senha")]),
]
CAMPOS_MOEDA = {"ValorDescarga__c"}
CAMPOS_ZERO_OCULTO = {"ValorDescarga__c", "QuantidadePallets__c", "QuantidadeDiarias__c",
                      "QuantidadeReentrega__c", "Dias_em_aberto__c"}


def valor_campo(chave: str, v: str) -> str:
    if not v:
        return ""
    if base.eh_data_iso(v):
        return base.fmt_data(v)
    if chave in CAMPOS_ZERO_OCULTO:
        try:
            if float(v) == 0:
                return ""
        except ValueError:
            pass
    if chave in CAMPOS_MOEDA:
        try:
            return "R$ " + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except ValueError:
            pass
    return v


def card(titulo: str, linhas: list[tuple[str, str]]) -> str:
    linhas = [(r, v) for r, v in linhas if v]
    if not linhas:
        return ""
    rows = "".join(
        f"<tr><td style='color:{TEXTO_SUAVE};padding:3px 14px 3px 0' valign='top' width='150'>{esc(r)}</td>"
        f"<td style='padding:3px 0'>{esc_texto(v)}</td></tr>" for r, v in linhas)
    return (
        "<table width='100%' cellspacing='0' cellpadding='12' "
        "style='background-color:#ffffff;border:1px solid #E8D9C3'><tr><td>"
        f"<div style='color:{COR_MARROM};font-size:11.5pt;font-weight:600;margin-bottom:6px'>{esc(titulo)}</div>"
        f"<table cellspacing='0' cellpadding='0' width='100%'>{rows}</table></td></tr></table>")


def html_resumo(c: dict, rel: dict, n_nf: int, n_anx: int, n_com: int) -> str:
    campos = c["campos"]
    cards = []
    for titulo, lista in GRUPOS:
        cards.append(card(titulo, [(rot, valor_campo(k, campos.get(k, ""))) for k, rot in lista]))
    cards = [x for x in cards if x]
    corpo = "<table width='100%' cellspacing='8' cellpadding='0'>"
    for i in range(0, len(cards), 2):
        par = cards[i:i + 2]
        corpo += "<tr>" + "".join(f"<td width='50%' valign='top'>{x}</td>" for x in par)
        if len(par) == 1:
            corpo += "<td width='50%'></td>"
        corpo += "</tr>"
    desc = card("Descrição", [("Assunto", campos.get("Subject", "")), ("Descrição", campos.get("Description", ""))])
    if desc:
        corpo += f"<tr><td colspan='2'>{desc}</td></tr>"
    links = []
    if rel["pai"]:
        links.append(("Caso pai", f"<a href='caso:{rel['pai']}'>{rel['pai']}</a>"))
    if rel["filhos"]:
        links.append(("Casos filhos", ", ".join(f"<a href='caso:{n}'>{n}</a>" for n in rel["filhos"][:30])))
    if links:
        rows = "".join(f"<tr><td style='color:{TEXTO_SUAVE};padding:3px 14px 3px 0' width='150'>{r}</td><td>{v}</td></tr>"
                       for r, v in links)
        corpo += ("<tr><td colspan='2'><table width='100%' cellpadding='12' cellspacing='0' "
                  "style='background-color:#ffffff;border:1px solid #E8D9C3'><tr><td>"
                  f"<div style='color:{COR_MARROM};font-size:11.5pt;font-weight:600'>Casos relacionados</div>"
                  f"<table cellspacing='0'>{rows}</table></td></tr></table></td></tr>")
    corpo += "</table>"
    return f"<body style='background-color:#F6F2EC'>{corpo}</body>"


def html_comentarios(lista: list[dict]) -> str:
    if not lista:
        return "<p style='color:#8A6A5C;padding:20px'>Este caso não tem comentários.</p>"
    partes = [f"<p style='color:{TEXTO_SUAVE}'>{len(lista)} comentário(s), do mais antigo para o mais recente.</p>"]
    for c in lista:
        interno = "" if c["publico"] else " <span style='color:#b3261e'>[interno]</span>"
        partes.append(
            "<table width='100%' cellspacing='0' cellpadding='0'><tr>"
            f"<td width='7' bgcolor='{COR_FUNDO}'>&nbsp;</td>"
            f"<td bgcolor='{COR_CINZA}'><table width='100%' cellspacing='0' cellpadding='10'><tr><td>"
            f"<b style='color:{COR_MARROM}'>{esc(c['autor'] or '—')}</b> "
            f"<span style='color:{TEXTO_SUAVE}'>&nbsp;{base.fmt_data(c['criado'])}</span>{interno}<br>"
            f"{esc_texto(c['corpo'])}</td></tr></table></td></tr></table>"
            "<div style='font-size:5pt'>&nbsp;</div>")
    return "".join(partes)


COR_STATUS = {"resolvido": COR_VERDE, "cancelado": COR_VERMELHO, "em finalização": "#C77700"}


# --------------------------------------------------------------------------- #
# Detalhe do caso (cabeçalho + abas)
# --------------------------------------------------------------------------- #
class DetalheCaso(QWidget):
    pedir_busca = Signal(str)

    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.cab = QLabel()
        self.cab.setTextFormat(Qt.TextFormat.RichText)
        self.cab.setObjectName("cab_caso")
        lay.addWidget(self.cab)

        self.abas = QTabWidget()
        lay.addWidget(self.abas, 1)

        self.resumo = QTextBrowser()
        self.resumo.setOpenLinks(False)
        self.resumo.anchorClicked.connect(self._link)
        self.coment = QTextBrowser()
        self.tab_nf = QTableWidget(0, 8)
        self.tab_nf.setHorizontalHeaderLabels(
            ["NF", "Série", "Remessa", "Cliente", "Centro", "Status NF", "Vendedor", "Mesma NF em outros casos"])
        self.tab_nf.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tab_nf.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tab_nf.setAlternatingRowColors(True)
        self.tab_nf.verticalHeader().setVisible(False)
        self.tab_nf.horizontalHeader().setStretchLastSection(True)
        self.tab_nf.cellDoubleClicked.connect(self._nf_dbl)
        self.anexos = AbaAnexos()

        self.todos = QWidget()
        tl = QVBoxLayout(self.todos)
        self.filtro = QLineEdit()
        self.filtro.setPlaceholderText("Filtrar campos (nome ou valor)…")
        self.filtro.textChanged.connect(self._filtrar)
        self.tab_campos = QTableWidget(0, 2)
        self.tab_campos.setHorizontalHeaderLabels(["Campo", "Valor"])
        self.tab_campos.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tab_campos.setAlternatingRowColors(True)
        self.tab_campos.verticalHeader().setVisible(False)
        self.tab_campos.setWordWrap(True)
        self.tab_campos.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.tab_campos.horizontalHeader().setStretchLastSection(True)
        self.tab_campos.setColumnWidth(0, 280)
        nota = QLabel("Todos os campos preenchidos do caso, como vieram do sistema legado "
                      "(campos vazios não são exibidos; datas em horário de Brasília).")
        nota.setStyleSheet(f"color:{TEXTO_SUAVE}")
        tl.addWidget(nota)
        tl.addWidget(self.filtro)
        tl.addWidget(self.tab_campos, 1)

        for w, nome in ((self.resumo, "Resumo"), (self.coment, "Comentários"), (self.tab_nf, "Notas fiscais"),
                        (self.anexos, "Anexos"), (self.todos, "Todos os campos")):
            self.abas.addTab(w, nome)

    def _link(self, url: QUrl):
        s = url.toString()
        if s.startswith("caso:"):
            self.pedir_busca.emit(s[5:])

    def _nf_dbl(self, linha, _col):
        nf = self.tab_nf.item(linha, 0).text()
        serie = self.tab_nf.item(linha, 1).text()
        self.pedir_busca.emit(f"{nf}-{serie}" if serie else nf)

    def _filtrar(self, texto: str):
        t = texto.lower().strip()
        for r in range(self.tab_campos.rowCount()):
            a = self.tab_campos.item(r, 0).text().lower()
            b = self.tab_campos.item(r, 1).text().lower()
            self.tab_campos.setRowHidden(r, bool(t) and t not in a and t not in b)

    def mostrar(self, pacote: dict):
        c = pacote["caso"]
        numero = c["numero"]
        campos = c["campos"]
        rel, com, nfs, anx = pacote["rel"], pacote["comentarios"], pacote["nfs"], pacote["anexos"]

        status = c["status"] or "—"
        cor = COR_STATUS.get(status.lower(), COR_MARROM)
        sub = " · ".join(x for x in (c["transportadora"], c["cnpj"], f"aberto em {base.fmt_data(c['criado'])}") if x.strip())
        self.cab.setText(
            f"<span style='font-size:19pt;font-weight:700;color:{COR_MARROM}'>Caso {esc(numero)}</span> &nbsp; "
            f"<span style='background-color:{cor};color:white;font-size:10pt'>&nbsp;&nbsp;{esc(status)}&nbsp;&nbsp;</span>"
            f"&nbsp; <span style='color:{TEXTO_SUAVE}'>{esc(c['motivo'])}"
            f"{' › ' + esc(c['submotivo']) if c['submotivo'] else ''}</span><br>"
            f"<span style='color:#444'>{esc(sub)}</span><br>"
            f"<span style='color:{TEXTO_SUAVE};font-size:9pt'>Origem: planilha “{esc(c['semestre'])}”</span>")

        self.resumo.setHtml(html_resumo(c, rel, len(nfs), len(anx), len(com)))
        self.coment.setHtml(html_comentarios(com))

        self.tab_nf.setRowCount(len(nfs))
        for r, n in enumerate(nfs):
            outros = ", ".join(n["outros_casos"]) + (" … (e mais)" if n["outros_mais"] else "")
            vals = [n["nf"], n["serie"], n["remessa"], n["cliente"], n["centro"], n["status_nf"], n["vendedor"], outros]
            for col, v in enumerate(vals):
                self.tab_nf.setItem(r, col, QTableWidgetItem(v))
        for col, w in enumerate((100, 60, 110, 260, 70, 110, 330)):
            self.tab_nf.setColumnWidth(col, w)
        self.tab_nf.setToolTip("Dê um duplo clique numa linha para buscar todos os casos com essa NF.")

        self.anexos.carregar(anx)

        itens = list(campos.items())
        self.tab_campos.setRowCount(len(itens))
        for r, (k, v) in enumerate(itens):
            self.tab_campos.setItem(r, 0, QTableWidgetItem(k))
            self.tab_campos.setItem(r, 1, QTableWidgetItem(base.fmt_data(v) if base.eh_data_iso(v) else v))
        self.tab_campos.resizeRowsToContents()
        self.filtro.clear()

        self.abas.setTabText(1, f"Comentários ({len(com)})")
        self.abas.setTabText(2, f"Notas fiscais ({len(nfs)})")
        self.abas.setTabText(3, f"Anexos ({len({a['doc_id'] for a in anx})})")


# --------------------------------------------------------------------------- #
# Trabalhos demorados (copiar a base, ler os Excel) com janela de progresso
# --------------------------------------------------------------------------- #
class ThreadTrabalho(QThread):
    progresso = Signal(str, float)
    concluido = Signal(object)
    falhou = Signal(str)

    def __init__(self, fn):
        super().__init__()
        self.fn = fn

    def run(self):
        try:
            self.concluido.emit(self.fn(lambda m, f: self.progresso.emit(m, f)))
        except Exception as e:
            self.falhou.emit(f"{e}\n\n{traceback.format_exc()}")


class DialogoTrabalho(QDialog):
    def __init__(self, parent, titulo: str, cabecalho: str, fn):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.resize(760, 440)
        self.ok = False
        self.resultado = None
        lay = QVBoxLayout(self)
        self.titulo = QLabel(cabecalho)
        self.titulo.setWordWrap(True)
        self.barra = QProgressBar()
        self.barra.setRange(0, 1000)
        self.status = QLabel("Iniciando…")
        self.relatorio = QPlainTextEdit()
        self.relatorio.setReadOnly(True)
        self.relatorio.setFont(QFont("Consolas", 9))
        self.relatorio.hide()
        self.fechar = QPushButton("Fechar")
        self.fechar.setEnabled(False)
        self.fechar.clicked.connect(self.accept)
        for w in (self.titulo, self.barra, self.status, self.relatorio):
            lay.addWidget(w)
        lay.addWidget(self.fechar, 0, Qt.AlignmentFlag.AlignRight)
        self.th = ThreadTrabalho(fn)
        self.th.progresso.connect(lambda m, f: (self.status.setText(m), self.barra.setValue(int(f * 1000))))
        self.th.concluido.connect(self._fim)
        self.th.falhou.connect(self._erro)
        self.th.start()

    def _fim(self, res):
        self.ok = True
        self.resultado = res
        self.barra.setValue(1000)
        texto = res[0] if isinstance(res, tuple) else (res if isinstance(res, str) else "")
        self.titulo.setText("<b>Concluído.</b>")
        if isinstance(res, tuple):  # relatório da leitura dos Excel
            self.status.setText("Confira o relatório (anexos sem arquivo / arquivos sem registro).")
            self.relatorio.setPlainText(texto)
            self.relatorio.show()
            self.fechar.setEnabled(True)
        else:
            self.accept()

    def _erro(self, msg: str):
        self.titulo.setText("<b style='color:#b3261e'>Não foi possível concluir.</b>")
        self.status.setText("")
        self.relatorio.setPlainText(msg)
        self.relatorio.show()
        self.fechar.setEnabled(True)

    def reject(self):
        if self.fechar.isEnabled():
            super().reject()


# --------------------------------------------------------------------------- #
# Consultas ao banco em segundo plano (a tela nunca fica esperando o disco)
# --------------------------------------------------------------------------- #
class _SinalDB(QObject):
    busca = Signal(int, str, object)
    caso = Signal(int, str, object)
    erro = Signal(int, str)


class TarefaBusca(QRunnable):
    def __init__(self, sinais: _SinalDB, gen: int, texto: str, db: Path):
        super().__init__()
        self.sinais, self.gen, self.texto, self.db = sinais, gen, texto, db

    def run(self):
        try:
            conn = base.abrir(self.db)
            try:
                res = base.buscar(conn, self.texto)
            finally:
                conn.close()
            self.sinais.busca.emit(self.gen, self.texto, res)
        except RuntimeError:
            pass  # janela já fechada
        except Exception as e:
            try:
                self.sinais.erro.emit(self.gen, str(e))
            except RuntimeError:
                pass


class TarefaCaso(QRunnable):
    def __init__(self, sinais: _SinalDB, gen: int, numero: str, db: Path, raiz: str):
        super().__init__()
        self.sinais, self.gen, self.numero, self.db, self.raiz = sinais, gen, numero, db, raiz

    def run(self):
        try:
            conn = base.abrir(self.db)
            try:
                pacote = base.pacote_caso(conn, self.numero, self.raiz)
            finally:
                conn.close()
            self.sinais.caso.emit(self.gen, self.numero, pacote)
        except RuntimeError:
            pass  # janela já fechada
        except Exception as e:
            try:
                self.sinais.erro.emit(self.gen, str(e))
            except RuntimeError:
                pass


# --------------------------------------------------------------------------- #
# Janela principal
# --------------------------------------------------------------------------- #
BOAS_VINDAS = f"""
<p style='font-size:12pt;color:{COR_MARROM}'><b>Digite acima e tecle Enter:</b></p>
<p style='font-size:11pt;line-height:150%'>
&bull; <b>Nº do caso</b> — ex.: <code>00682313</code> ou apenas <code>682313</code><br>
&bull; <b>Nº da NF</b> — ex.: <code>171341</code> ou com série <code>171341-20</code><br>
&bull; <b>Remessa</b>, <b>CNPJ</b> da transportadora ou <b>placa</b> do veículo<br>
&bull; <b>Qualquer texto</b> — descrição, comentários e nomes de anexos</p>
<p style='color:{TEXTO_SUAVE}'>Para cada caso: resumo, comentários, notas fiscais e anexos (fotos e PDFs).</p>
"""


class TelaInicial(QWidget):
    """Tela de boas-vindas: mascote, nome do programa e instruções."""

    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 10, 20, 10)
        lay.addStretch(1)
        mascote = QLabel()
        pm = QPixmap(recurso("logo/mascote_transparente.png"))
        if not pm.isNull():
            mascote.setPixmap(pm.scaledToHeight(230, Qt.TransformationMode.SmoothTransformation))
        mascote.setAlignment(Qt.AlignmentFlag.AlignCenter)
        titulo = QLabel(APP_NOME)
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        titulo.setStyleSheet(f"font-size:28pt;font-weight:700;color:{COR_MARROM};")
        sub = QLabel("Casos, NFs e anexos do sistema legado")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet(f"font-size:12pt;color:{COR_MARROM};")
        cartao = QLabel(BOAS_VINDAS)
        cartao.setObjectName("cartao")
        cartao.setTextFormat(Qt.TextFormat.RichText)
        cartao.setWordWrap(False)
        for w in (mascote, titulo, sub):
            lay.addWidget(w)
        lay.addSpacing(14)
        lay.addWidget(cartao, 0, Qt.AlignmentFlag.AlignHCenter)
        lay.addStretch(2)


class Janela(QMainWindow):
    def __init__(self, iniciar: bool = True):
        super().__init__()
        self.setWindowTitle(f"{APP_NOME} — casos, NFs e anexos")
        self.resize(1450, 880)
        self.cfg = QSettings("Pandurata", "ConsultaLegado")
        self.raiz = ""
        self.db_path: Path | None = None
        self.info_base: dict = {}
        self._res: list[dict] = []
        self._gen_busca = 0
        self._gen_caso = 0
        self._carregando = False

        self.pool_db = QThreadPool(self)
        self.pool_db.setMaxThreadCount(2)
        self.sinais = _SinalDB()
        self.sinais.busca.connect(self._busca_pronta)
        self.sinais.caso.connect(self._caso_pronto)
        self.sinais.erro.connect(self._erro_db)
        self._timer_carga = QTimer(self)
        self._timer_carga.setSingleShot(True)
        self._timer_carga.timeout.connect(self._mostrar_carregando)

        central = QWidget()
        self.setCentralWidget(central)
        lay = QVBoxLayout(central)
        topo = QHBoxLayout()
        topo.setSpacing(12)
        logo = QLabel()
        pm_logo = QPixmap(recurso("logo/cabeca.png"))
        if not pm_logo.isNull():
            logo.setPixmap(pm_logo.scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                          Qt.TransformationMode.SmoothTransformation))
        bloco = QVBoxLayout()
        bloco.setSpacing(0)
        t1 = QLabel(APP_NOME)
        t1.setObjectName("titulo_app")
        t2 = QLabel("casos · NFs · anexos do legado")
        t2.setObjectName("sub_app")
        bloco.addWidget(t1)
        bloco.addWidget(t2)
        topo.addWidget(logo)
        topo.addLayout(bloco)
        self.busca = QLineEdit()
        self.busca.setObjectName("busca")
        self.busca.setPlaceholderText("Nº do caso, NF (ex.: 171341-20), remessa, CNPJ, placa ou texto…")
        self.busca.setClearButtonEnabled(True)
        self.busca.returnPressed.connect(self.buscar)
        self.btn = QPushButton("Buscar")
        self.btn.setObjectName("primario")
        self.btn.clicked.connect(self.buscar)
        topo.addWidget(self.busca, 1)
        topo.addWidget(self.btn)
        lay.addLayout(topo)

        self.sp = QSplitter(Qt.Orientation.Horizontal)
        self.painel_lista = QWidget()
        pl = QVBoxLayout(self.painel_lista)
        pl.setContentsMargins(0, 0, 0, 0)
        self.titulo_lista = QLabel()
        self.titulo_lista.setStyleSheet(f"color:{COR_MARROM};font-weight:700")
        self.tabela = QTableWidget(0, 6)
        self.tabela.setHorizontalHeaderLabels(["Caso", "Abertura", "Transportadora", "Status", "NFs", "Anexos"])
        self.tabela.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabela.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabela.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tabela.setAlternatingRowColors(True)
        self.tabela.verticalHeader().setVisible(False)
        self.tabela.horizontalHeader().setStretchLastSection(False)
        self.tabela.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.tabela.currentCellChanged.connect(lambda r, c, pr, pc: self._selecionou(r))
        pl.addWidget(self.titulo_lista)
        pl.addWidget(self.tabela, 1)

        self.direita = QStackedWidget()
        self.boas_vindas = TelaInicial()
        self.vazio = QLabel()
        self.vazio.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.vazio.setWordWrap(True)
        self.vazio.setStyleSheet(f"background:{COR_CINZA};border:1px solid {DOURADO};border-radius:14px;padding:30px;")
        self.detalhe = DetalheCaso()
        self.detalhe.pedir_busca.connect(self.buscar_texto)
        for w in (self.boas_vindas, self.vazio, self.detalhe):
            self.direita.addWidget(w)

        self.sp.addWidget(self.painel_lista)
        self.sp.addWidget(self.direita)
        self.sp.setStretchFactor(1, 1)
        self.sp.setSizes([700, 780])
        self.painel_lista.hide()
        lay.addWidget(self.sp, 1)

        self.barra_status = self.statusBar()
        rodape = QLabel("by Melhoria Contínua Logística")
        rodape.setStyleSheet(f"color:{COR_MARROM};font-weight:600;padding-right:8px;")
        self.barra_status.addPermanentWidget(rodape)
        self._menu()
        self.busca.setFocus()
        if iniciar:
            QTimer.singleShot(0, self.iniciar)

    def closeEvent(self, ev):
        self._gen_busca += 1
        self._gen_caso += 1
        self.pool_db.clear()
        QThreadPool.globalInstance().clear()
        self.pool_db.waitForDone(3000)
        QThreadPool.globalInstance().waitForDone(3000)
        super().closeEvent(ev)

    # -- menus ---------------------------------------------------------------- #
    def _menu(self):
        m = self.menuBar().addMenu("&Arquivo")
        for texto, fn in (("Escolher a pasta do legado…", self.escolher_pasta),
                          ("Verificar se há base nova na pasta compartilhada…", self.verificar_base_nova),
                          ("Recriar a base (ler os Excel de novo)…", self.reindexar),
                          ("Sobre a base de dados…", self.sobre_base)):
            a = QAction(texto, self)
            a.triggered.connect(fn)
            m.addAction(a)
        m.addSeparator()
        sair = QAction("Sair", self)
        sair.triggered.connect(self.close)
        m.addAction(sair)

    # -- inicialização -------------------------------------------------------- #
    def iniciar(self):
        raiz = self.cfg.value("pasta_legado", "")
        if not raiz or not self._pasta_valida(raiz):
            if not self.escolher_pasta(obrigatorio=True):
                self.close()
            return
        self._preparar_base(raiz)

    @staticmethod
    def _pasta_valida(p: str) -> bool:
        pp = Path(p)
        return pp.is_dir() and (bool(indexador.listar_excels(pp)) or (pp / "ANEXOS").is_dir())

    def escolher_pasta(self, obrigatorio: bool = False) -> bool:
        if obrigatorio:
            QMessageBox.information(
                self, APP_NOME,
                "Escolha a pasta <b>EXTRAÇÃO LEGADO</b> — a que contém os arquivos Excel "
                "(01 - Extração…, 02 - Extração…) e a pasta <b>ANEXOS</b>.<br><br>"
                "Normalmente ela fica no seu OneDrive, em <i>Projeto B Entregas - B Entregas Hist PRD</i>.")
        p = QFileDialog.getExistingDirectory(self, "Pasta EXTRAÇÃO LEGADO", self.cfg.value("pasta_legado", ""))
        if not p:
            return False
        sub = Path(p) / "EXTRAÇÃO LEGADO"
        if not self._pasta_valida(p) and sub.is_dir() and self._pasta_valida(str(sub)):
            p = str(sub)
        if not self._pasta_valida(p):
            QMessageBox.warning(self, APP_NOME, "Essa pasta não parece ser a da extração do legado "
                                "(não encontrei os Excel nem a pasta ANEXOS).")
            return self.escolher_pasta(obrigatorio=False)
        self.cfg.setValue("pasta_legado", p)
        self._preparar_base(p)
        return True

    def _preparar_base(self, raiz: str, forcar_pergunta: bool = False):
        """Garante uma base LOCAL: copia da pasta compartilhada (1ª vez), ou cria lendo os Excel."""
        self.raiz = raiz
        comp = indexador.db_compartilhada(raiz)
        local = indexador.db_local()
        if comp.exists() and not indexador.compartilhada_compativel(raiz):
            # base de uma versão antiga do programa: não vale copiar (é grande); será recriada
            if not local.exists() or forcar_pergunta:
                QMessageBox.information(
                    self, APP_NOME,
                    "A base que está na pasta compartilhada foi criada por uma versão antiga do programa "
                    "(maior e mais lenta).<br><br>Vou criar uma base nova, leve, lendo os Excel. "
                    "Ela substituirá a antiga na pasta compartilhada.")
                self._criar_base()
        elif comp.exists():
            if not local.exists():
                QMessageBox.information(
                    self, APP_NOME,
                    "Vou copiar a base de consulta para o <b>seu computador</b>. Isso é feito só uma vez "
                    "(ou quando houver base nova) e pode levar alguns minutos.<br><br>"
                    "Depois disso o programa abre e busca rapidamente, sem depender do OneDrive.")
                self._copiar_base(comp)
            elif indexador.copia_local_desatualizada(comp):
                r = QMessageBox.question(
                    self, APP_NOME, "Existe uma base mais nova na pasta compartilhada.<br>Copiar para o seu "
                    "computador agora?")
                if r == QMessageBox.StandardButton.Yes:
                    self._copiar_base(comp)
            elif forcar_pergunta:
                QMessageBox.information(self, APP_NOME, "Sua base já é a mais recente.")
        elif not local.exists():
            r = QMessageBox.question(
                self, APP_NOME,
                "A base de consulta ainda não foi criada.<br><br>"
                "Vou ler os arquivos Excel agora e montar a base. Isso é feito só uma vez e "
                "pode levar alguns minutos. Continuar?")
            if r == QMessageBox.StandardButton.Yes:
                self._criar_base()
        if local.exists():
            self.usar_base(raiz, local)

    def _copiar_base(self, comp: Path):
        self._liberar_base()
        dlg = DialogoTrabalho(self, "Copiando a base", "<b>Copiando a base para o seu computador…</b><br>"
                              "Só na primeira vez (ou quando houver base nova).",
                              lambda cb: indexador.copiar_para_local(comp, cb))
        dlg.exec()

    def _criar_base(self):
        self._liberar_base()
        raiz = self.raiz
        dlg = DialogoTrabalho(self, "Preparando a base de consulta",
                              "<b>Lendo os Excel e montando a base de consulta…</b><br>"
                              "Isso é feito só uma vez (ou quando você recriar a base). Pode levar alguns minutos.",
                              lambda cb: indexador.construir_base(raiz, cb))
        dlg.exec()

    def _liberar_base(self):
        """Solta a base atual (no Windows não dá para substituir um arquivo em uso)."""
        self.db_path = None

    def usar_base(self, raiz: str, db: Path):
        try:
            conn = base.abrir(db)
            try:
                info = {"total": base.meta(conn, "total_casos"), "criado": base.meta(conn, "criado_em"),
                        "relatorio": base.meta(conn, "relatorio")}
            finally:
                conn.close()
        except Exception as e:
            r = QMessageBox.question(self, APP_NOME, f"Não consegui abrir a base:\n{e}\n\nDeseja recriá-la agora "
                                     "lendo os Excel?")
            if r == QMessageBox.StandardButton.Yes:
                self.raiz = raiz
                self._criar_base()
                if indexador.db_local().exists():
                    self.usar_base(raiz, indexador.db_local())
            return
        self.raiz, self.db_path, self.info_base = raiz, Path(db), info
        self.barra_status.showMessage(
            f"Base de {int(info['total'] or 0):,}".replace(",", ".") + f" casos · criada em "
            f"{base.fmt_data(info['criado'])} · pasta: {raiz}")
        self.busca.setFocus()

    def verificar_base_nova(self):
        if not self.raiz:
            self.escolher_pasta(obrigatorio=True)
            return
        self._preparar_base(self.raiz, forcar_pergunta=True)

    def reindexar(self):
        if not self.raiz:
            self.escolher_pasta(obrigatorio=True)
            return
        self._criar_base()
        if indexador.db_local().exists():
            self.usar_base(self.raiz, indexador.db_local())

    def sobre_base(self):
        if not self.db_path:
            QMessageBox.information(self, APP_NOME, "Nenhuma base aberta.")
            return
        d = QDialog(self)
        d.setWindowTitle("Sobre a base de dados")
        d.resize(780, 420)
        l = QVBoxLayout(d)
        l.addWidget(QLabel(f"Arquivo da base (neste computador): {self.db_path}"))
        t = QPlainTextEdit(self.info_base.get("relatorio", ""))
        t.setReadOnly(True)
        t.setFont(QFont("Consolas", 9))
        l.addWidget(t)
        b = QPushButton("Fechar")
        b.clicked.connect(d.accept)
        l.addWidget(b, 0, Qt.AlignmentFlag.AlignRight)
        d.exec()

    # -- busca ---------------------------------------------------------------- #
    def buscar_texto(self, texto: str):
        self.busca.setText(texto)
        self.buscar()

    def buscar(self):
        if not self.db_path:
            self.iniciar()
            return
        q = self.busca.text().strip()
        if not q:
            return
        self._gen_busca += 1
        self._gen_caso += 1  # descarta qualquer carregamento de caso em andamento
        self.barra_status.showMessage("Buscando…")
        self.pool_db.start(TarefaBusca(self.sinais, self._gen_busca, q, self.db_path))

    def _erro_db(self, gen: int, msg: str):
        self._timer_carga.stop()
        self.barra_status.showMessage("Erro ao consultar a base.", 8000)
        QMessageBox.warning(self, APP_NOME, f"Erro ao consultar a base:\n{msg}")

    def _busca_pronta(self, gen: int, q: str, res: list):
        if gen != self._gen_busca:
            return  # o usuário já fez outra busca
        self._res = res
        self.tabela.blockSignals(True)
        self.tabela.setRowCount(len(res))
        for r, d in enumerate(res):
            vals = [d["numero"], base.fmt_data(d["criado"])[:10], d["transportadora"], d["status"],
                    str(d["n_nf"]), str(d["n_anx"])]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setToolTip(f"Encontrado por: {d['achado_por']}\nMotivo: {d['motivo']}")
                if c in (4, 5):
                    it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabela.setItem(r, c, it)
        for c, w in ((0, 82), (1, 88), (3, 112), (4, 42), (5, 62)):
            self.tabela.setColumnWidth(c, w)
        self.tabela.blockSignals(False)

        if not res:
            self.painel_lista.hide()
            self.vazio.setText(
                f"<div style='font-size:15pt;color:{COR_MARROM}'><b>Nada encontrado para “{esc(q)}”.</b></div><br>"
                "<span style='color:#555'>Confira o número digitado. Você pode buscar por nº do caso, NF "
                "(com ou sem série), remessa, CNPJ, placa ou parte de um texto.</span>")
            self.direita.setCurrentWidget(self.vazio)
            self.barra_status.showMessage("Nenhum resultado.", 5000)
            self.busca.selectAll()
            return
        if len(res) == 1:
            self.painel_lista.hide()
        else:
            self.painel_lista.show()
            por = res[0]["achado_por"]
            self.titulo_lista.setText(f"{len(res)} casos encontrados (principal critério: {por})"
                                      + (" — limitado aos 500 primeiros" if len(res) >= 500 else ""))
        self.tabela.selectRow(0)
        self._selecionou(0)
        self.busca.selectAll()

    # -- detalhe do caso ------------------------------------------------------ #
    def _selecionou(self, linha: int):
        if linha is None or linha < 0 or linha >= len(self._res) or not self.db_path:
            return
        d = self._res[linha]
        self._gen_caso += 1
        self._carregando = True
        self._numero_carregando = d["numero"]
        self._timer_carga.start(150)  # só mostra "Carregando…" se demorar
        self.pool_db.start(TarefaCaso(self.sinais, self._gen_caso, d["numero"], self.db_path, self.raiz))

    def _mostrar_carregando(self):
        if self._carregando:
            self.vazio.setText(f"<span style='font-size:14pt;color:{TEXTO_SUAVE}'>Carregando o caso "
                               f"{esc(self._numero_carregando)}…</span>")
            self.direita.setCurrentWidget(self.vazio)

    def _caso_pronto(self, gen: int, numero: str, pacote):
        if gen != self._gen_caso:
            return
        self._carregando = False
        self._timer_carga.stop()
        if pacote is None:
            self.vazio.setText("Caso não encontrado na base.")
            self.direita.setCurrentWidget(self.vazio)
            return
        self.direita.setCurrentWidget(self.detalhe)
        self.detalhe.mostrar(pacote)
        if len(self._res) == 1:
            self.barra_status.showMessage(f"Encontrado por: {self._res[0]['achado_por']}", 8000)
        else:
            self.barra_status.showMessage(
                f"Base de {int(self.info_base.get('total') or 0):,}".replace(",", ".") + " casos")


def main():
    if sys.platform == "win32":
        try:  # faz a barra de tarefas mostrar o ícone do programa (e não o do Python)
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Pandurata.ConsultaLegado")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NOME)
    app.setWindowIcon(QIcon(recurso("icone.ico")))
    app.setStyle("Fusion")
    pal = QPalette()
    for papel, cor in ((QPalette.ColorRole.Window, COR_FUNDO), (QPalette.ColorRole.WindowText, TEXTO),
                       (QPalette.ColorRole.Base, COR_BRANCO), (QPalette.ColorRole.AlternateBase, COR_CINZA),
                       (QPalette.ColorRole.Text, TEXTO), (QPalette.ColorRole.Button, COR_MARROM),
                       (QPalette.ColorRole.ButtonText, COR_BRANCO), (QPalette.ColorRole.Highlight, COR_AMARELO_CLARO),
                       (QPalette.ColorRole.HighlightedText, COR_MARROM)):
        pal.setColor(papel, QColor(cor))
    app.setPalette(pal)
    app.setStyleSheet(ESTILO)
    QThreadPool.globalInstance().setMaxThreadCount(6)
    w = Janela()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
