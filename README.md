# Consulta Legado (Programa B Entregas)

Programa de janela (desktop, **sem servidor**) para consultar a extração do sistema legado descomissionado. Busque por **número do caso**, **NF**, remessa, CNPJ, placa ou texto livre e veja, em um só lugar, o resumo do caso, os comentários, as notas fiscais e os **anexos** (fotos e PDFs).

> Desenvolvido pela equipe **Melhoria Contínua Logística** (Pandurata Alimentos / Bauducco).
> O nome "Consulta Legado" é provisório.

---

## Sumário

1. [O que o programa faz](#1-o-que-o-programa-faz)
2. [Como funciona (visão geral)](#2-como-funciona-visão-geral)
3. [Para quem só usa o programa](#3-para-quem-só-usa-o-programa)
4. [Para quem gera a base e o executável](#4-para-quem-gera-a-base-e-o-executável)
5. [Dados de origem](#5-dados-de-origem)
6. [Regras de ligação entre os dados](#6-regras-de-ligação-entre-os-dados)
7. [Busca](#7-busca)
8. [Interface](#8-interface)
9. [Arquitetura e estrutura do repositório](#9-arquitetura-e-estrutura-do-repositório)
10. [Banco de dados (esquema v2)](#10-banco-de-dados-esquema-v2)
11. [Desenvolvimento](#11-desenvolvimento)
12. [Identidade visual](#12-identidade-visual)
13. [Problemas conhecidos e soluções](#13-problemas-conhecidos-e-soluções)
14. [Status e o que não foi testado](#14-status-e-o-que-não-foi-testado)
15. [Pendências e ideias](#15-pendências-e-ideias)
16. [Glossário](#16-glossário)

---

## 1. O que o programa faz

- Digite um número e o programa descobre sozinho se é **caso**, **NF**, **NF + série**, **remessa** ou **CNPJ**.
- Digite uma placa ou um texto e ele faz busca livre (assunto, descrição, transportadora, motorista, comentários, nomes de anexos).
- Mostra o caso em abas: **Resumo**, **Comentários**, **Notas fiscais**, **Anexos** e **Todos os campos**.
- Visualiza **imagens** (zoom, arrastar, girar) e **PDFs** dentro do próprio programa.
- Cada usuário executa na própria máquina. A distribuição é por `.exe`, e os usuários **não precisam ter Python**.

## 2. Como funciona (visão geral)

```
 Excel 01..09 (OneDrive)           Quem gera a base (1 pessoa)
 EXTRAÇÃO LEGADO\*.xlsx   ───►   construir_base()  ───►  legado.db (local)
                                                            │
                                                            ▼ publica
                                           EXTRAÇÃO LEGADO\_indice\legado.db
                                                            │
                                                            ▼ cópia (1 vez por usuário)
                                       %LOCALAPPDATA%\ConsultaLegado\legado.db
                                                            │
                                                            ▼ consultas (somente leitura)
                                                     Consulta Legado (.exe)
```

Ideias centrais:

- **Só uma pessoa lê os Excel.** A carga é pesada (tempo e memória). Ela gera um banco SQLite pronto.
- **Os colegas só copiam o arquivo pronto** (~0,9 GB) e consultam. Nenhum Excel é lido na máquina deles.
- **O programa sempre abre a base local**, nunca direto do OneDrive. Arquivos "só online" do OneDrive eram a causa de travamentos e do erro `Errno 22`.
- **A tela nunca consulta o banco.** Todas as consultas rodam em segundo plano.

## 3. Para quem só usa o programa

Você **não precisa** de Python, de administrador nem de rodar nenhum `.bat`.

1. Receba o `ConsultaLegado.exe` (ou a pasta `ConsultaLegado` inteira, se a versão for em pasta).
2. Abra o programa. Na **primeira vez**, escolha a pasta `EXTRAÇÃO LEGADO` (no OneDrive).
3. O programa copia a base para o seu computador (`%LOCALAPPDATA%\ConsultaLegado\`). Isso acontece uma vez e pode levar alguns minutos. Se o OneDrive falhar, a cópia tenta de novo e continua de onde parou.
4. Digite o número do caso ou da NF no campo de busca e clique em **Buscar**.

**Quando a base for atualizada:** o programa percebe que a cópia compartilhada mudou e pergunta se você quer copiar a nova. Use o menu **Arquivo > Verificar base nova**.

**Não use** a opção "Recriar a base" do menu **Arquivo**. Ela serve só para quem gera a base.

**Dicas para os anexos:**

- Fotos e PDFs são lidos da pasta `ANEXOS` no OneDrive. Se o arquivo estiver "só online", a primeira abertura pode demorar.
- Para evitar a espera, clique com o botão direito na pasta `ANEXOS` e escolha **Sempre manter neste dispositivo** (se houver espaço em disco).
- No primeiro dia, evite que todos abram o programa no mesmo minuto: a cópia de ~0,9 GB sobrecarrega o OneDrive.

**Possíveis avisos:**

- **SmartScreen:** clique em *Mais informações > Executar assim mesmo*.
- **Antivírus da empresa:** o executável de arquivo único se descompacta em pasta temporária. Se for barrado, use a versão em pasta.
- Se a TI redirecionar o `AppData\Local` para a rede, o ganho de desempenho some. Confira digitando `%LOCALAPPDATA%` na barra do Explorer: o caminho deve começar com `C:\Users\...`.

## 4. Para quem gera a base e o executável

### 4.1 Pré-requisitos

- Windows corporativo, **sem permissão de administrador** (nada aqui exige senha de administrador).
- Python 3.14 instalado por usuário (confira com `py -0p`).
- Projeto em disco local (ex.: `C:\ConsultaLegado`), **fora do OneDrive**.

### 4.2 Gerar a base (carga dos Excel)

A carga lê todos os Excel da pasta `EXTRAÇÃO LEGADO`, monta o banco numa pasta temporária, grava a base local e **tenta publicá-la** em `EXTRAÇÃO LEGADO\_indice\`:

| Arquivo em `_indice\` | Conteúdo |
|---|---|
| `legado.db` | Base SQLite pronta para distribuição |
| `versao_schema.txt` | Versão do esquema do banco |
| `relatorio_carga.txt` | Relatório da carga (avisos, linhas lidas) |

Pelo programa: menu **Arquivo > Recriar a base (ler os Excel de novo)**. Depois da publicação, os colegas só precisam copiar.

### 4.3 Gerar o executável

Os `.bat` criam um ambiente virtual `.venv` com a versão definida em `set PYVER=3.14` (topo do arquivo), sem administrador.

| Script | Resultado |
|---|---|
| `build_exe_unico.bat` | Um único arquivo `dist\ConsultaLegado.exe` |
| `build_exe.bat` | Pasta `dist\ConsultaLegado` (enviar a pasta inteira) |
| `executar.bat` | Abre o programa a partir do código (só para desenvolvimento) |

O `.exe` leva o ícone e a pasta `logo/` (`--add-data`). O código usa `recurso()` para localizar esses arquivos, funcionando tanto no `.py` quanto no `.exe`.

> **Atenção:** nenhum `.bat` roda na máquina do usuário final. Eles são só para quem gera o executável.

### 4.4 Distribuição

Envie o executável pela pasta compartilhada do OneDrive (ou pela seção **Releases** do GitHub). **Nunca** coloque dados (Excel, `legado.db`, anexos) no repositório Git.

## 5. Dados de origem

Pasta no OneDrive: `Projeto B Entregas - B Entregas Hist PRD\EXTRAÇÃO LEGADO`, contendo:

- **9 Excel numerados:** `01 - Extração - 2022.xlsx`, `02 - Extração - Primeiro Semestre 2023.xlsx` … `09 - Extração - Segundo Semestre 2026.xlsx` (48 a 129 MB cada, exceto o 01 com 90 KB e o 09 com 3 MB).
- **`ANEXOS\<NN - nome do semestre>\arquivos\`** com os arquivos físicos (fotos de WhatsApp, PDFs de NF etc.).
- Os arquivos do OneDrive aparecem como "só online" (baixados sob demanda).

**Nome do arquivo físico do anexo:** `<NumeroCaso>_<ContentDocumentId>_<PathOnClient>`

Exemplo: `00682335_069U600000rVK8oIAG_WhatsApp Image 2026-07-01 at 16.31.27.jpeg`

A pasta do semestre é deduzida do nome do Excel: `09 - Extração - Segundo Semestre 2026.xlsx` → `ANEXOS\09 - Segundo Semestre 2026`.

**Estrutura verificada (somente no arquivo 09):** 5 abas.

| Aba | Linhas (arq. 09) | Papel |
|---|---|---|
| CASOS | 3.344 (108 colunas) | Cadastro do caso |
| COMENTÁRIOS DO CASO | 5.935 | Histórico |
| NF-CASO | 2.824 | NFs vinculadas |
| CONTENTDOCUMENTLINK | 2.152 | Liga anexo ao caso |
| CONTENTVERSION | 2.151 | Dados do arquivo |

> **Não verificado:** a estrutura dos arquivos 01 a 08. O leitor localiza abas e colunas pelo nome e avisa no relatório quando algo falta, mas isso nunca foi confirmado com os arquivos reais.

Tamanho real já carregado: bases com 519.963 e, depois, 635.163 casos.

## 6. Regras de ligação entre os dados

- `CASOS.Id` ↔ `CONTENTDOCUMENTLINK.LinkedEntityId`; `CONTENTDOCUMENTLINK.ContentDocumentId` ↔ `CONTENTVERSION.ContentDocumentId`.
- `CONTENTVERSION.CASENUMBER` e a coluna E de `CONTENTDOCUMENTLINK` são fórmulas `VLOOKUP`. O leitor usa o valor em cache. O programa refaz o cruzamento por `Id` e usa o `CASENUMBER` só para montar o nome do arquivo.
- **Comentários:** `Parent.CaseNumber` (ou `ParentId` como reserva).
- **NFs:** `NumeroCaso__c` (ou `Caso__c`). Número em `NumeroNota__c`, série em `SerieNota__c`, remessa em `Remessa__c`. O campo `NumNotaFiscal__c` vem com HTML e é limpo. A comparação ignora zeros à esquerda.
- **Todo caso tem uma NF principal** em `CASOS.NumeroNF__c`. No arquivo 09, 1.608 casos têm a NF só nesse campo (sem linha em NF-CASO), então o programa **une as duas fontes** (2.824 + 1.608 = 4.432 NFs).
- **50 NFs aparecem em mais de um caso**, então a busca por NF pode devolver uma lista.
- **Anexos ligados a mais de um caso** (ex.: casos 00682775/00682776 e 00682957/00682959) têm um só arquivo físico e aparecem nos dois casos. Há anexo com 2 versões (`IsLatest`).
- Se o mesmo caso aparece em dois Excel, vale o de `LastModifiedDate` mais recente.
- **Datas** vêm em UTC. A tela mostra o horário de Brasília (UTC−3 fixo, sem horário de verão).
- `Dias_em_aberto__c` é calculado na data da extração e engana em casos resolvidos: **não aparece no Resumo** (continua na aba "Todos os campos").
- Sequências `_x000D_` e quebras de linha do Excel são limpas.

## 7. Busca

Função `base.buscar`:

- **Só números e pontuação:**
  - 14 dígitos → **CNPJ**
  - formato `NNN-SS` → **NF + série**
  - caso contrário, tenta **caso** (completando com zeros até 8 dígitos), **NF** e **remessa**
  - se nada for achado, faz busca **parcial** (caso contém / NF começa com)
- **Texto:** placa (padrão `AAA-9A99`) e/ou busca livre (FTS5, prefixo, todos os termos).
- **Ordem:** critério mais exato primeiro, depois data de abertura decrescente. Limite de **500 resultados**.
- Cada resultado informa "encontrado por": Caso, NF, NF + série, Remessa, CNPJ, Placa ou Texto.

## 8. Interface

- **Topo:** mascote, título, caixa de busca e botão **Buscar** (verde). **Rodapé:** "by Melhoria Contínua Logística".
- **Tela inicial** com mascote grande e instruções. Com mais de um resultado, a lista aparece à esquerda.
- **Abas do caso:**
  - **Resumo:** cartões de identificação, transportadora e veículo, classificação, datas, valores, descrição e casos pai/filhos
  - **Comentários:** do mais antigo ao mais novo
  - **Notas fiscais:** duplo clique busca todos os casos daquela NF
  - **Anexos:** miniaturas e visor
  - **Todos os campos:** com filtro
- **Visor de arquivos:**
  - Imagens: zoom (roda do mouse), arrastar, girar, duplo clique = tamanho real (rotação EXIF respeitada)
  - PDF: aberto dentro do programa (QtPdf)
  - Outros formatos: abrem no programa padrão do Windows
  - Arquivo ausente: aviso com o caminho esperado
- **Threads:** a tela nunca consulta o banco. `TarefaBusca` e `TarefaCaso` rodam num pool próprio (2 threads), cada uma com sua conexão. Arquivos (miniaturas, imagens, PDFs) rodam no pool global (6 threads), porque os anexos podem estar "só online". Contadores de geração descartam resultados antigos. O aviso "Carregando…" só aparece se passar de 150 ms.
- **Menu Arquivo:** escolher pasta do legado, verificar base nova, recriar a base e ver informações da base (relatório).
- **Status do caso (cores):** Resolvido = verde, Cancelado = vermelho, Em finalização = âmbar, demais = marrom.

## 9. Arquitetura e estrutura do repositório

**Linguagem:** Python. **Interface:** PySide6 (Qt). **Banco:** SQLite com FTS5. **Leitura de Excel:** `python-calamine` (rápido, opcional) com reserva em `openpyxl`.

```
Programa-B-Entregas/
├── app.py                  # Interface (janelas, abas, visor, tarefas em segundo plano)
├── base.py                 # Consultas ao banco
├── indexador.py            # Excel → SQLite, compressão, cópia e publicação da base
├── tema.py                 # Paleta oficial e cores derivadas
├── logo/                   # mascote.svg, icone.svg (fontes), PNGs e icone.ico
├── icone.ico               # Ícone do executável
├── mascote_consulta.ico    # Ícone alternativo
├── build_exe.bat           # Gera dist\ConsultaLegado (pasta)
├── build_exe_unico.bat     # Gera dist\ConsultaLegado.exe (arquivo único)
├── executar.bat            # Abre pelo código (desenvolvimento)
├── requirements.txt        # PySide6>=6.6, openpyxl>=3.1
├── LEIAME.md               # Manual de uso e distribuição
└── README.md               # Este arquivo
```

**Classes principais de `app.py`:**

| Classe | Papel |
|---|---|
| `Janela` | Tela principal |
| `TelaInicial` | Tela de boas-vindas |
| `DetalheCaso` | Cabeçalho e abas do caso |
| `AbaAnexos` | Galeria de anexos |
| `VisorArquivo` / `VisorImagem` | Visor de imagem, PDF e outros; zoom e rotação |
| `TarefaBusca` / `TarefaCaso` | Consultas em segundo plano |
| `TarefaArquivo` | Leitura de arquivos em segundo plano |
| `ThreadTrabalho` / `DialogoTrabalho` | Janela de progresso |

**Funções principais:**

- `base.py`: `abrir`, `buscar`, `caso`, `casos_relacionados`, `comentarios`, `nfs`, `anexos`, `pacote_caso`, `fmt_data`.
- `indexador.py`: `construir_base`, `Planilha`, `montar_zdict`, `comprimir_json`, `copiar_para_local`, `publicar_na_pasta`, `copia_local_desatualizada`, `compartilhada_compativel`. Constantes: `VERSAO_SCHEMA = 2`, pasta `_indice`, arquivo `legado.db`.

**Dependências:**

- `PySide6>=6.6` e `openpyxl>=3.1` (em `requirements.txt`).
- `python-calamine` é instalado à parte e só se houver pacote pronto (`--only-binary=:all:`, nunca compila). Sem ele, o programa usa o `openpyxl` (resultado idêntico, mais lento: 7,4 s contra ~0,5 s no arquivo de 3 MB).

## 10. Banco de dados (esquema v2)

Tabelas: `meta`, `semestres`, `casos`, `comentarios`, `nfs`, `anexos` e a tabela virtual `busca` (FTS5).

| Tabela | Conteúdo |
|---|---|
| `casos` | Colunas de busca/lista (numero, case_id, semestre, criado, transportadora, cnpj, cnpj_num, status, motivo, submotivo, placa, placa_norm, motorista, parent_id) e `dados` (BLOB com **todos** os campos do caso, em JSON comprimido com zlib e dicionário pré-definido, guardado em `meta` com a chave `zdict`, em base64) |
| `comentarios` | numero, criado, autor, corpo, publico (0/1) |
| `nfs` | numero, nf, nf_chave, serie, serie_chave, remessa, remessa_chave, cliente, centro, status_nf, vendedor |
| `anexos` | numero, doc_id, arquivo (só o nome), nome, ext, tamanho, criado, criado_por, mais_recente, versao, compartilhado, semestre, achado. Caminho completo = raiz do legado + `semestres.pasta` + `arquivo` |
| `busca` (FTS5, sem conteúdo, `detail=column`) | Usa o `rowid` do caso. Indexa número, assunto, descrição, transportadora, placa, motorista, motivo, submotivo, comentários e nomes de anexos |

**Índices:** NF (`nf_chave`, `serie_chave`), remessa, CNPJ, placa, `parent_id`, `case_id`, comentários e anexos por caso.

**Ao mudar o esquema, suba `VERSAO_SCHEMA`.** Bases de outra versão são detectadas e recriadas.

### Onde fica a base (regra central)

- O programa **sempre abre a cópia local:** `%LOCALAPPDATA%\ConsultaLegado\legado.db`, em modo somente leitura (`mode=ro&immutable=1`).
- A cópia compartilhada (`EXTRAÇÃO LEGADO\_indice\`) serve só de distribuição.
- `origem.json` (local) guarda tamanho e data da cópia compartilhada já copiada. Se a compartilhada mudar, o programa pergunta se você quer copiar a nova.
- A cópia tolera falhas do OneDrive: tenta de novo com blocos menores (8 MB → 1 MB → 256 KB → 64 KB), retoma de onde parou e, em último caso, usa `shutil.copyfile` (cópia nativa do Windows).

## 11. Desenvolvimento

### Regras de código do projeto

1. **Nunca consultar o banco na thread da tela.**
2. **Subir `VERSAO_SCHEMA`** ao mudar o esquema do banco.
3. **Cores só em `tema.py`.**
4. **Textos da interface em português.**
5. O programa **sempre abre uma base local**, nunca direto do OneDrive.

### Rodar a partir do código

```bat
py -3.14 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install --only-binary=:all: python-calamine
python app.py
```

(O calamine é opcional. Se não houver pacote pronto para a sua versão do Python, pule essa linha.)

### Testar a interface sem tela

```bash
QT_QPA_PLATFORM=offscreen python seu_teste.py
```

Use `widget.grab().save("x.png")` para capturar a tela.

> **Armadilha:** `QTest.qWait` longo segura o interpretador e trava as threads de consulta. Em testes, espere em passos curtos (ex.: 50 ms).

### Teste de escala

Base sintética replicando os 3.344 casos 155 vezes (~520 mil casos). Resultado medido no Linux: ~894 MB; buscas de 0,1 ms a 1 s (a mais lenta é a busca parcial por número); pior pausa da tela inferior a 0,1 s.

### Leitor sem calamine

Simule com `sys.modules['python_calamine'] = None`.

### Fluxo com Git

```
git add .
git status
git commit -m "descrição da mudança"
git push
```

Confira o `git status` antes do commit. O `.gitignore` bloqueia `.venv/`, `build/`, `dist/`, `__pycache__/`, `*.pyc`, `*.spec`, `*.db`, `*.xlsx`, `*.zip`, `ANEXOS/` e `_indice/`.

> **Importante:** os dados do legado contêm informações de transportadoras, CNPJs e placas. **Nunca** suba Excel, `legado.db` ou anexos para o repositório.

## 12. Identidade visual

Paleta oficial, definida em `tema.py`:

| Constante | Cor | Uso |
|---|---|---|
| `COR_FUNDO` | `#FDB834` | Fundo (laranja/amarelo) |
| `COR_MARROM` | `#653024` | Botões e destaques |
| `COR_AMARELO_CLARO` | `#F2F0A1` | Apoio |
| `COR_BRANCO` | `#FFFFFF` | Superfícies |
| `COR_VERDE` | `#2E7D32` | Ação principal, "Resolvido" |
| `COR_VERMELHO` | `#B3261E` | "Cancelado" |
| `COR_CINZA` | `#F6F2EC` | Fundos neutros |
| `COR_CARTAO` | `#FFC84F` | Cartões |

Derivadas: `MARROM_HOVER #7F4334`, `VERDE_HOVER #256628`, `TEXTO #3B1D14`, `TEXTO_SUAVE #8A6A5C`, `BORDA #E8D9C3`, `DOURADO #E0A21A`.

Referência de estilo: tela do "Motor de Diagnóstico — Pagamento de Fretes" (fundo laranja, botões marrons arredondados, ação principal verde).

**Mascote:** gota amarela com fone de ouvido e colete refletivo, desenhada em vetor (`logo/mascote.svg`). O ícone usa só cabeça e lupa para ficar legível em 16 px.

## 13. Problemas conhecidos e soluções

| Problema | Causa | Decisão |
|---|---|---|
| Janela "Não está respondendo" com 520 mil casos | Base de ~3 GB lida direto do OneDrive e consultas na thread da tela | Base sempre local, consultas em segundo plano, esquema v2 (~0,9 GB) |
| `Errno 22` ao copiar a base do OneDrive | Leitura grande de arquivo "só online" | Cópia com nova tentativa, blocos menores e cópia nativa de reserva |
| `pip install python-calamine` falhou | Python sem pacote pronto tenta compilar com Rust e falta `link.exe` | `--only-binary`, calamine opcional, reserva em `openpyxl` |
| `.bat` rodado em máquina sem Python | Os `.bat` precisam de Python | `.bat` só para quem gera; o usuário final só abre o `.exe` |
| Repositório Git com 78 MB | `dist/` e `build/` versionados antes do `.gitignore` valer | `.gitignore` corrigido e histórico recriado |

## 14. Status e o que não foi testado

**Testado (em ambiente Linux/offscreen e base sintética):** interface com Qt offscreen, escala de ~520 mil casos, buscas, leitor com e sem calamine.

**Ainda não validado em produção:**

- Execução do `.exe` em Windows de usuários finais (o executável foi gerado, mas o uso real pelos colegas ainda não foi confirmado).
- Excel 01 a 08: estrutura real e uso de memória no arquivo de 129 MB.
- Leitura de anexos reais do OneDrive e cópia da base real com falhas reais.
- Compatibilidade do PyInstaller com Python 3.14/3.15.

Diante de qualquer erro, anote a **mensagem completa** (ou tire um print) antes de reportar.

## 15. Pendências e ideias

- [ ] Definir o **nome** definitivo do programa e aplicá-lo (`APP_NOME` em `app.py`, `--name` nos `.bat`). Sugestões levantadas: Memória B, B Entregas Histórico, Achei!, Garimpo, Arquivo Vivo.
- [ ] Confirmar a estrutura dos Excel 01–08 e ajustar o leitor se houver colunas diferentes.
- [ ] Publicar a base com **nome versionado** (ex.: `legado_AAAA-MM-DD.db` mais um arquivo apontando para a versão atual), evitando troca de arquivo no meio de uma leitura.
- [ ] Esconder ou proteger a opção "Recriar a base" para quem não gera a base.
- [ ] Trocar a letra "B" do colete do mascote se houver símbolo oficial.
- [ ] Exportar "dossiê do caso" (PDF/ZIP).
- [ ] Filtros por status, período e transportadora.
- [ ] Histórico das últimas buscas.
- [ ] Assinatura do executável para evitar o aviso do SmartScreen.

## 16. Glossário

| Termo | Significado |
|---|---|
| **Caso** | Registro de atendimento do sistema legado (número de 8 dígitos, ex.: `00682335`) |
| **NF / série / remessa** | Nota fiscal, série da nota e remessa de entrega ligadas ao caso |
| **ContentDocumentId / PathOnClient** | Identificador e nome original do anexo |
| **Base local / compartilhada** | `legado.db` no `%LOCALAPPDATA%` (usada) × em `_indice` no OneDrive (distribuição) |
| **Legado** | Sistema descomissionado de onde vêm os Excel |
| **FTS5** | Mecanismo de busca de texto completo do SQLite |

---

*Mantido pela equipe Melhoria Contínua Logística. Mantenha este README, o `LEIAME.md` e o código coerentes entre si a cada mudança.*
