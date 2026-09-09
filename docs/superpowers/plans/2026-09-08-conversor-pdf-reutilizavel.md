# Conversor PDF Analítico Reutilizável Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Encapsular a extração coordenada usada no PDF de referência em `converter_pdf.py`, gerando automaticamente um `.xlsx` com código e nome do beneficiário em colunas separadas.

**Architecture:** O programa lerá os streams textuais nativos do PDF, reconstruirá linhas por posição e aplicará âncoras de layout centralizadas em uma configuração. Um estado hierárquico carregará o titular entre páginas; lançamentos serão emitidos individualmente e validados por `TOTAL TITULAR`. A planilha será criada com `openpyxl`, mantendo valores como texto e registrando alertas em uma aba separada.

**Tech Stack:** Python 3.9+, `openpyxl`, biblioteca padrão (`argparse`, `logging`, `re`, `zlib`, `zipfile`, `xml.etree`).

**Spec:** Requisitos fornecidos pelo usuário na conversa de 2026-09-08.

## Global Constraints

- O PDF é a única fonte oficial dos dados.
- Uma linha do Excel representa um lançamento do PDF.
- O Excel principal terá exatamente as 10 colunas solicitadas, nesta ordem.
- Valores e códigos serão gravados como texto para não perder sinal, zeros ou casas decimais.
- Nenhuma validação poderá alterar dados extraídos.
- O contexto do titular deverá sobreviver à mudança de página.
- Linhas sem beneficiário repetido serão associadas pelo agrupamento e registradas em `Observações`.

### Task 1: Criar o núcleo de leitura e reconstrução do PDF

**Files:**
- Create: `converter_pdf.py`
- Test: `test_converter_pdf.py`

**Interfaces:**
- Produces `split_code_and_name(text) -> tuple[str, str]`.
- Produces `extract_records(pdf_path) -> ExtractionResult`.
- Produces `validate_records(result) -> list[str]`.

- [x] **Step 1: Write failing tests** para separação por regex, campos de lançamento, continuidade de página e preservação de valores.
- [x] **Step 2: Run tests** e confirmar falha pela ausência de `converter_pdf.py`.
- [x] **Step 3: Implementar tokenizer PDF, leitura de streams, agrupamento de linhas e parser por âncoras configuráveis.**
- [x] **Step 4: Run tests** e confirmar que os testes unitários passam.

### Task 2: Implementar validação, logging e CLI

**Files:**
- Modify: `converter_pdf.py`
- Modify: `test_converter_pdf.py`

**Interfaces:**
- CLI: `python converter_pdf.py INPUT.pdf [-o OUTPUT.xlsx]`.
- Produces logging de arquivo, páginas, lançamentos, titulares, linhas ignoradas e alertas.

- [x] **Step 1: Write failing tests** para `TOTAL TITULAR`, argumentos de saída e erro de arquivo inexistente.
- [x] **Step 2: Run tests** e confirmar falha.
- [x] **Step 3: Implementar validação e CLI com `argparse` e `logging`.**
- [x] **Step 4: Run tests** e confirmar aprovação.

### Task 3: Gerar Excel e requirements

**Files:**
- Modify: `converter_pdf.py`
- Create: `requirements.txt`
- Modify: `test_converter_pdf.py`

**Interfaces:**
- Produces workbook with headers `CÓDIGO BENEFICIÁRIO`, `BENEFICIÁRIO`, `OP. LANÇAMENTO`, `FILIAL`, `CENTROCUSTO`, `QUANT.`, `VALOR UNIT.`, `VALOR BONIF.`, `VALOR TOTAL`, `MENSAGEM`.

- [x] **Step 1: Write failing tests** para cabeçalhos, 185 linhas, filtros, congelamento e valores textuais.
- [x] **Step 2: Run tests** e confirmar falha.
- [x] **Step 3: Implementar geração funcional com `openpyxl` e declarar somente `openpyxl` em `requirements.txt`.**
- [x] **Step 4: Run tests** e confirmar aprovação.

### Task 4: Testar contra o PDF de referência e comparar com o Excel anterior

**Files:**
- Create: `ANALITICO_TAXA_0054041451-1_reproduzido.xlsx`

- [x] **Step 1: Executar** `python converter_pdf.py ANALITICO_TAXA_0054041451-1.PDF -o ANALITICO_TAXA_0054041451-1_reproduzido.xlsx`.
- [x] **Step 2: Comparar** as 185 linhas e os campos do Excel anterior, ignorando somente a mudança de uma coluna concatenada para código/nome.
- [x] **Step 3: Validar** todos os 137 `TOTAL TITULAR`, sinais, zeros, casas decimais e linhas de transição.
- [x] **Step 4: Executar** a suíte completa e a verificação estrutural do `.xlsx`.
