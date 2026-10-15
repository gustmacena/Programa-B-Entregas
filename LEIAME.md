# Consulta Legado

Programa de janela (sem servidor) para consultar a extração do sistema legado:
digite o **número do caso** ou da **NF** e veja resumo, comentários, notas fiscais e anexos (fotos e PDFs).

## Como usar
1. Abra o programa (`executar.bat`, ou o `ConsultaLegado.exe` depois de gerar com `build_exe.bat`).
2. Na primeira vez ele pede a pasta **EXTRAÇÃO LEGADO** (a que tem os Excel `01 - Extração…` a `09 - Extração…`
   e a pasta `ANEXOS`). Cada pessoa escolhe a sua, pois o caminho do OneDrive muda por usuário.
3. O programa prepara a base de consulta **no seu computador** (ver abaixo) e já pode buscar.
4. Digite e tecle Enter. Aceita: nº do caso (`00682313` ou `682313`), NF (`171341` ou `171341-20`),
   remessa, CNPJ, placa ou qualquer texto (descrição, comentários, nomes de anexos).

## Quem gera × quem usa (importante)
- **Quem gera** (tem Python): roda `build_exe.bat` (gera a pasta `dist\ConsultaLegado`) **ou**
  `build_exe_unico.bat` (gera um único `dist\ConsultaLegado.exe`).
- **Quem usa** (não precisa de Python): recebe só o resultado e abre o `ConsultaLegado.exe`.
  **Não** rode `executar.bat` nem `build_*.bat` nessa máquina — eles precisam do Python.
- Com `build_exe.bat`, envie a **pasta inteira** `dist\ConsultaLegado` (o `.exe` sozinho não funciona).

## Visual
As cores seguem a paleta oficial e ficam todas em `tema.py` (fundo, marrom, verde, vermelho etc.).
Para mudar uma cor, edite esse arquivo e gere o executável de novo. O mascote e o ícone ficam na pasta `logo/`
(o `build_exe.bat` já leva essa pasta junto).

## Como a base funciona (e por que o programa não trava)
- O programa **sempre abre uma cópia da base no disco do próprio computador**
  (`%LOCALAPPDATA%\ConsultaLegado\legado.db`). Ele nunca lê o banco direto do OneDrive/SharePoint:
  banco de dados em pasta sincronizada fica lento e pode dar conflito de sincronização.
- **Primeira pessoa:** abre o programa, ele lê os Excel e monta a base (uma vez só, alguns minutos) e a
  publica em `EXTRAÇÃO LEGADO\_indice\` para os colegas.
- **Colegas:** ao abrir, o programa **copia** essa base para o computador (uma vez, cerca de 1 GB para
  ~520 mil casos) e não precisa ler os Excel de novo.
- Quando alguém recriar a base, os outros recebem o aviso "há base nova" e copiam de novo.
- Menu **Arquivo**: *Verificar se há base nova…* e *Recriar a base (ler os Excel de novo)…*.
- As buscas e a abertura dos casos rodam em segundo plano: a janela não congela, mesmo com os anexos
  ainda sendo baixados do OneDrive.

## Dicas
- Anexos no OneDrive marcados como "só online" são baixados na hora de abrir. Para ficar rápido, clique com o
  botão direito na pasta `ANEXOS` > **Sempre manter neste dispositivo**.
- Fotos: roda do mouse = zoom, arrastar = mover, duplo clique = tamanho real, botões para girar.
- Aba **Notas fiscais**: duplo clique numa linha busca todos os casos com aquela NF.
- Aba **Todos os campos**: os campos originais do legado, com filtro.
- Datas são mostradas em horário de Brasília (o legado grava em UTC).
- A busca por texto livre usa descrição, comentários, motorista, transportadora e nomes de anexos.

## Como os dados se ligam
`Caso.Id` -> `ContentDocumentLink.LinkedEntityId` -> `ContentVersion` (por `ContentDocumentId`).
Arquivo físico: `ANEXOS\<semestre>\arquivos\<NumeroCaso>_<ContentDocumentId>_<PathOnClient>`.
Um anexo ligado a mais de um caso aparece em todos eles.
