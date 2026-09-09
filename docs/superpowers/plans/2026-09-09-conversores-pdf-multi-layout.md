# Conversores PDF Multi-Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transformar o conversor monolítico atual em uma aplicação Flask multi-layout com `UnimedParser` e `BradescoDentalParser`, ambos retornando `pandas.DataFrame`, com Excel desacoplado e seleção centralizada por registry.

**Architecture:** O leitor estrutural de PDF será extraído para `rateiosrh.core` e ficará livre de regras de fornecedor. Cada layout terá um parser isolado derivado de `BaseParser`; `ConversionService` resolverá o parser pelo registry, validará compatibilidade e o DataFrame, e `ExcelExporter` cuidará apenas da saída. A CLI e o Flask reutilizarão esse mesmo serviço.

**Tech Stack:** Python 3.9+, Flask 3.x, pandas 2.x, openpyxl 3.x, `Decimal` para a futura camada financeira, pytest para testes e biblioteca padrão para leitura estrutural do PDF.

**Spec:** `docs/superpowers/specs/2026-09-09-conversores-pdf-multi-layout-design.md`

## Global Constraints

- O PDF é a única fonte oficial dos dados.
- Uma linha do DataFrame representa um registro/lançamento do PDF.
- O DataFrame principal do Bradesco Dental terá exatamente as 13 colunas aprovadas, nesta ordem.
- Os campos serão preservados como texto durante a extração.
- Linhas de continuidade manterão vazios os campos não repetidos visualmente.
- Nenhum parser poderá corrigir, preencher por inferência, recalcular, deduplicar, consolidar ou descartar registros.
- Um parser só conhecerá regras do seu próprio layout.
- A escolha do parser será feita pelo identificador selecionado pelo usuário e centralizada no registry.
- PDFs incompatíveis serão rejeitados antes da geração do Excel.
- `Decimal` será usado na futura conversão financeira; `float` não será usado indiscriminadamente.
- O rateamento e o CSV final não serão implementados nesta etapa.
- Todos os comandos Python e pip deverão usar `venv/bin/python` ou `venv/bin/python -m pip`.
- O workspace atual não possui um repositório Git válido; não executar `git init` nem alterar metadados Git sem autorização do usuário.

---

## Mapa de arquivos

### Criar

- `rateiosrh/__init__.py`: pacote da aplicação.
- `rateiosrh/core/__init__.py`: exportações do núcleo.
- `rateiosrh/core/pdf_reader.py`: leitor PDF comum, spans e agrupamento de linhas.
- `rateiosrh/core/exceptions.py`: exceções de estrutura, layout, registry e DataFrame.
- `rateiosrh/parsers/__init__.py`: exportações dos parsers.
- `rateiosrh/parsers/base.py`: contrato abstrato dos parsers.
- `rateiosrh/parsers/unimed.py`: migração isolada do layout atual.
- `rateiosrh/parsers/bradesco_dental.py`: parser do novo layout.
- `rateiosrh/registry.py`: registry e metadados dos conversores.
- `rateiosrh/services/__init__.py`: exportações dos serviços.
- `rateiosrh/services/validation.py`: validação comum de schema/DataFrame.
- `rateiosrh/services/excel_exporter.py`: DataFrame para XLSX.
- `rateiosrh/services/conversion.py`: orquestração da conversão.
- `rateiosrh/web/__init__.py`: pacote web.
- `rateiosrh/web/routes.py`: rotas Flask.
- `rateiosrh/web/templates/index.html`: formulário de seleção/upload.
- `app.py`: factory Flask e ponto de entrada web.
- `cli.py`: ponto de entrada de linha de comando.
- `test_architecture.py`: contrato, registry e serviços.
- `test_bradesco_dental.py`: regressão do novo layout.
- `test_web.py`: rotas e respostas Flask.

### Alterar

- `converter_pdf.py`: tornar-se uma fachada de compatibilidade para a CLI e funções públicas do UNIMED.
- `test_converter_pdf.py`: manter os testes existentes e apontá-los para a fachada/parser migrado quando necessário.
- `requirements.txt`: declarar pandas, Flask e pytest junto às dependências existentes.
- `contexto_sessao.md`: registrar a implementação, validações e pendências ao final.

### Manter como fixtures de validação

- `ANALITICO_TAXA_0054041451-1.PDF`: referência UNIMED, 10 páginas e 185 lançamentos.
- `ANALITICO_TAXA_0054041451-1.xlsx`: Excel anterior.
- `ANALITICO_TAXA_0054041451-1_reproduzido.xlsx`: saída validada anterior.
- `FLEXIV_1166046_1_082026.pdf`: referência Bradesco Dental, 7 páginas e 279 registros.
- `FLEXIV_1166046_1_082026_referencia.xlsx`: Excel de referência conferido manualmente.

---

### Task 1: Preparar dependências e o pacote da aplicação

**Files:**
- Modify: `requirements.txt`
- Create: `rateiosrh/__init__.py`, `rateiosrh/core/__init__.py`, `rateiosrh/parsers/__init__.py`, `rateiosrh/services/__init__.py`, `rateiosrh/web/__init__.py`
- Test: `test_architecture.py`

**Interfaces:**
- Produces: imports válidos para `rateiosrh.core`, `rateiosrh.parsers`, `rateiosrh.services` e `rateiosrh.web`.
- Dependency contract: `pandas>=2.2,<3`, `openpyxl>=3.1,<4`, `Flask>=3.0,<4`, `pytest>=8,<9`.

- [ ] **Step 1: Escrever o teste de importação que falha**

Adicionar em `test_architecture.py`:

```python
def test_application_packages_are_importable():
    import rateiosrh.core
    import rateiosrh.parsers
    import rateiosrh.services
    import rateiosrh.web

    assert rateiosrh.core is not None
```

- [ ] **Step 2: Executar o teste na venv e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_architecture.py::test_application_packages_are_importable`

Expected: FAIL com `ModuleNotFoundError: No module named 'rateiosrh'`.

- [ ] **Step 3: Atualizar dependências e criar os pacotes mínimos**

Atualizar `requirements.txt` para:

```text
pandas>=2.2,<3
openpyxl>=3.1,<4
Flask>=3.0,<4
pytest>=8,<9
```

Criar os cinco `__init__.py` vazios ou apenas com exportações de pacote, sem regras de parsing.

- [ ] **Step 4: Instalar e verificar dependências exclusivamente na venv**

Run: `venv/bin/python -m pip install -r requirements.txt`

Run: `venv/bin/python -c "import flask, pandas, openpyxl, pytest; print(flask.__version__, pandas.__version__, openpyxl.__version__)"`

Expected: os quatro imports funcionam e as versões atendem aos limites declarados.

- [ ] **Step 5: Executar o teste e confirmar aprovação**

Run: `venv/bin/python -m pytest -q test_architecture.py::test_application_packages_are_importable`

Expected: PASS.

---

### Task 2: Extrair o leitor PDF comum e preservar ambos os formatos de árvore de páginas

**Files:**
- Create: `rateiosrh/core/pdf_reader.py`, `rateiosrh/core/exceptions.py`
- Modify: `test_architecture.py`, `test_converter_pdf.py`
- Modify: `converter_pdf.py` para importar o núcleo, sem duplicar a implementação.

**Interfaces:**
- `PdfDocument(data: bytes)` indexa objetos e expõe `object_body`, `stream` e `page_streams()`.
- `TextSpan(y: float, x: float, text: str)` representa texto posicionado.
- `group_rows(stream: bytes, row_precision: int = 2) -> list[tuple[float, list[tuple[float, str]]]]` agrupa spans por linha.
- `clean(value: str) -> str` normaliza espaços apenas para reconstrução textual, sem alterar valores extraídos.
- `PdfStructureError` representa PDF sem objetos, páginas, conteúdos ou streams válidos.

- [ ] **Step 1: Escrever testes falhos para os dois PDFs**

Adicionar em `test_architecture.py`:

```python
from pathlib import Path

from rateiosrh.core.pdf_reader import PdfDocument


def test_reader_finds_ten_pages_in_unimed_reference():
    document = PdfDocument(Path("ANALITICO_TAXA_0054041451-1.PDF").read_bytes())
    assert len(document.page_streams()) == 10


def test_reader_finds_seven_pages_in_bradesco_dental_reference():
    document = PdfDocument(Path("FLEXIV_1166046_1_082026.pdf").read_bytes())
    assert len(document.page_streams()) == 7
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_architecture.py -k 'reader_finds'`

Expected: FAIL porque `rateiosrh.core.pdf_reader` ainda não existe.

- [ ] **Step 3: Mover a implementação comum do conversor atual**

Transferir para `pdf_reader.py` as rotinas atualmente comuns de `converter_pdf.py`: indexação de objetos, leitura/descompressão de streams, tokenizer, spans, limpeza e agrupamento.

Alterar a busca do dicionário de páginas para aceitar as duas formas:

```python
re.search(rb"/Type\s*/Pages", body)
```

Usar a mesma tolerância para `/Type /Page` e `/Type/Page`. Não adicionar cabeçalhos, coordenadas de fornecedor ou regex de registro ao núcleo.

- [ ] **Step 4: Atualizar o conversor atual para usar o núcleo**

Substituir as definições duplicadas em `converter_pdf.py` por imports de `rateiosrh.core.pdf_reader`, preservando aliases temporários para `PdfDocument`, `TextSpan`, `LayoutConfig` e funções públicas usadas pelos testes existentes.

- [ ] **Step 5: Executar as regressões do leitor**

Run: `venv/bin/python -m pytest -q test_architecture.py -k 'reader_finds' test_converter_pdf.py::test_reference_pdf_reconstructs_all_records_and_preserves_fields`

Expected: PASS; 10 páginas do UNIMED e 7 do Bradesco Dental são reconhecidas, e os valores do teste UNIMED permanecem iguais.

---

### Task 3: Implementar o contrato `BaseParser` e a validação comum

**Files:**
- Create: `rateiosrh/parsers/base.py`, `rateiosrh/services/validation.py`
- Modify: `rateiosrh/core/exceptions.py`, `test_architecture.py`

**Interfaces:**
- `class BaseParser(ABC)` declara `parser_id`, `display_name`, `columns`.
- `validate_layout(pdf_path: Path) -> bool` verifica compatibilidade.
- `parse(pdf_path: Path) -> pd.DataFrame` é obrigatório.
- `validate_dataframe(df: pd.DataFrame) -> list[str]` retorna mensagens sem alterar o DataFrame.
- `get_columns() -> list[str]` retorna a lista declarada.
- `ValidationService.validate_schema(df, expected_columns) -> list[str]` valida somente presença, ordem e quantidade de colunas.

- [ ] **Step 1: Escrever teste falho do contrato**

Adicionar:

```python
import pandas as pd
import pytest

from rateiosrh.parsers.base import BaseParser


def test_base_parser_requires_parse_implementation():
    with pytest.raises(TypeError):
        BaseParser()


def test_schema_validation_rejects_wrong_column_order():
    from rateiosrh.services.validation import ValidationService

    frame = pd.DataFrame([["x", "y"]], columns=["B", "A"])
    assert ValidationService.validate_schema(frame, ["A", "B"])
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_architecture.py -k 'base_parser or schema_validation'`

Expected: FAIL por ausência dos módulos e classes.

- [ ] **Step 3: Implementar o contrato abstrato e as exceções**

Criar `LayoutMismatchError`, `UnknownParserError` e `DataFrameValidationError`. Implementar `BaseParser` com `ABC`, `ClassVar` e métodos abstratos, sem coordenadas nem regex de fornecedor.

- [ ] **Step 4: Implementar a validação comum**

`ValidationService.validate_schema` deverá retornar mensagens para:

```python
if list(df.columns) != list(expected_columns):
    errors.append("Colunas ausentes ou fora de ordem.")
```

Também deverá detectar DataFrame não vazio somente quando o serviço receber essa regra explicitamente; não aplicar preenchimento automático nem coerção de tipos.

- [ ] **Step 5: Executar os testes do contrato**

Run: `venv/bin/python -m pytest -q test_architecture.py -k 'base_parser or schema_validation'`

Expected: PASS.

---

### Task 4: Migrar o conversor atual para `UnimedParser`

**Files:**
- Create: `rateiosrh/parsers/unimed.py`
- Modify: `converter_pdf.py`, `test_converter_pdf.py`, `test_architecture.py`

**Interfaces:**
- `UnimedParser.parser_id == "unimed"`.
- `UnimedParser.display_name == "Converter UNIMED"`.
- `UnimedParser.columns` contém as 10 colunas atuais.
- `UnimedParser.validate_layout(PDF) -> True` para `ANALITICO_TAXA_0054041451-1.PDF`.
- `UnimedParser.validate_layout(FLEXIV_1166046_1_082026.pdf) -> False`.
- `UnimedParser.parse(pdf_path) -> pd.DataFrame` retorna 185 linhas no PDF de referência.
- A fachada `converter_pdf.py` continuará exportando `extract_records`, `validate_records`, `write_workbook` e `run` para compatibilidade.

- [ ] **Step 1: Escrever teste falho de regressão via DataFrame**

Adicionar:

```python
from pathlib import Path

from rateiosrh.parsers.unimed import UnimedParser


def test_unimed_parser_returns_the_validated_dataframe():
    frame = UnimedParser().parse(Path("ANALITICO_TAXA_0054041451-1.PDF"))
    assert len(frame) == 185
    assert frame.columns.tolist() == [
        "CÓDIGO BENEFICIÁRIO", "BENEFICIÁRIO", "OP. LANÇAMENTO",
        "FILIAL", "CENTROCUSTO", "QUANT.", "VALOR UNIT.",
        "VALOR BONIF.", "VALOR TOTAL", "MENSAGEM",
    ]
    assert frame.iloc[0]["VALOR TOTAL"] == "407,19"
    assert frame.iloc[0]["CÓDIGO BENEFICIÁRIO"] == "0976.0663.000060-00"
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_architecture.py::test_unimed_parser_returns_the_validated_dataframe`

Expected: FAIL porque `UnimedParser` ainda não existe.

- [ ] **Step 3: Encapsular a lógica validada no parser UNIMED**

Mover para `unimed.py` `LayoutConfig`, `Record`, `TitularTotal`, `ExtractionResult`, regras de titular, lançamentos, totais e validações específicas. O parser deverá construir um DataFrame com `pd.DataFrame(rows, columns=columns).astype("string")` sem alterar os textos.

Usar como assinatura de compatibilidade o marcador textual `ANALÍTICO DE TAXA - FATURA NRO` e a estrutura de campos do layout existente. Não usar apenas o valor da filial para aceitar um documento.

- [ ] **Step 4: Preservar a API legada como fachada**

Fazer `converter_pdf.py` importar as classes e funções do parser UNIMED. Manter as assinaturas legadas de `extract_records`, `validate_records`, `write_workbook` e `run`, definir `unimed` implicitamente apenas nessa fachada e não adicionar seleção de outros fornecedores nesse módulo. A troca da implementação de `run` para `ConversionService` será feita explicitamente na Task 7, depois que o serviço existir.

- [ ] **Step 5: Executar regressão completa do UNIMED**

Run: `venv/bin/python -m pytest -q test_converter_pdf.py test_architecture.py -k 'unimed or reference or validation or cli'`

Expected: PASS com os 185 lançamentos, 137 totais, somas e valores previamente validados.

---

### Task 5: Implementar `BradescoDentalParser` usando o Excel conferido

**Files:**
- Create: `rateiosrh/parsers/bradesco_dental.py`
- Create/Modify: `test_bradesco_dental.py`
- Modify: `rateiosrh/core/exceptions.py` somente se uma exceção específica de seção for necessária.

**Interfaces:**
- `BradescoDentalParser.parser_id == "bradesco_dental"`.
- `BradescoDentalParser.display_name == "Converter Bradesco Dental"`.
- `BradescoDentalParser.columns` é exatamente a lista de 13 colunas aprovada.
- `BradescoDentalParser.validate_layout(PDF) -> True` para `FLEXIV_1166046_1_082026.pdf`.
- `BradescoDentalParser.validate_layout(ANALITICO_TAXA_0054041451-1.PDF) -> False`.
- `BradescoDentalParser.parse(pdf_path) -> pd.DataFrame` retorna 279 linhas e 13 colunas.

**Regras de layout a implementar:**

- reconhecer `BRADESCO DENTAL - FATURA TECNICA` e o cabeçalho detalhado `Certif.`, `Nome Beneficiário`, `Mês/Ano`, `Valor`, `Part. Benef.`;
- na referência, processar as páginas 2–6; no parser, selecionar páginas pela presença do cabeçalho da seção detalhada e ignorar páginas de resumo, boleto e mensagens;
- ignorar resumo da página 1, boleto e mensagens da página 7;
- usar coordenadas próprias para Certif. `41.31`, Nome `77.00`, nascimento `264.98`, sexo `299.50/299.67`, estado civil `318.66/320.33`, parentesco `340.33/342.00`, plano `362.83`, início `388.49`, movimento `426.83–428.50`, mês `442.16`, valor `492.99–501.32` e participação `548.32–551.32`;
- manter a coluna `Subfatura Nº N=Nova` vazia quando o PDF não apresentar valor;
- emitir uma linha para cada linha detalhada, inclusive as 77 linhas sem certificado repetido;
- preservar movimentos `CM`, `CR`, `IR`, `IM` e valores como `15,00-` sem normalizá-los;
- criar DataFrame com `dtype="string"` e sem colunas auxiliares.

- [ ] **Step 1: Escrever os testes falhos contra o PDF e a referência**

Adicionar:

```python
from pathlib import Path

from rateiosrh.parsers.bradesco_dental import BradescoDentalParser


PDF = Path("FLEXIV_1166046_1_082026.pdf")


def test_bradesco_dental_parser_has_exact_schema_and_row_count():
    frame = BradescoDentalParser().parse(PDF)
    assert frame.shape == (279, 13)
    assert frame.columns.tolist() == [
        "Certif.", "Nome Beneficiário", "Subfatura Nº N=Nova",
        "Data Nascimento", "Sexo", "Est. Civil", "Paren.", "Plano",
        "Data Início", "Mov.", "Mês/ano", "Valor", "Part. Benef",
    ]


def test_bradesco_dental_preserves_continuations_and_trailing_negative_sign():
    frame = BradescoDentalParser().parse(PDF)
    assert frame["Certif."].eq("").sum() == 77
    continuation = frame[frame["Certif."].eq("")].iloc[0]
    assert continuation["Mov."] == "CM"
    assert continuation["Valor"] == "15,00-"
    assert frame["Subfatura Nº N=Nova"].eq("").all()


def test_bradesco_dental_rejects_unimed_pdf():
    assert not BradescoDentalParser().validate_layout(
        Path("ANALITICO_TAXA_0054041451-1.PDF")
    )
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_bradesco_dental.py`

Expected: FAIL porque o parser ainda não existe.

- [ ] **Step 3: Implementar validação de compatibilidade**

Abrir os streams até encontrar os marcadores do cabeçalho; retornar `False` para PDFs sem a assinatura Bradesco Dental e sem o cabeçalho detalhado. Não usar o nome do arquivo como critério.

- [ ] **Step 4: Implementar a extração por coordenadas e seção**

Reconstruir cada linha por `y`, extrair os campos por faixas de `x`, selecionar linhas entre `y=4.70` e `y=748.70` nas páginas 2–6 e emitir todas as linhas que contenham qualquer campo de registro detalhado. Deixar vazios os campos que não possuem span na linha.

- [ ] **Step 5: Implementar validação específica do DataFrame**

Verificar exatamente 13 colunas, relações estruturais entre linhas completas e continuidades, movimentos esperados e preservação de valores sem conversão numérica. O teste da referência deverá exigir 279 registros, 202 certificados e 77 continuidades; o parser não deverá fixar esses totais para relatórios futuros do mesmo layout. Retornar mensagens de validação sem modificar o DataFrame.

- [ ] **Step 6: Executar os testes do parser**

Run: `venv/bin/python -m pytest -q test_bradesco_dental.py`

Expected: PASS com 279 linhas e os sinais/textos conferidos.

---

### Task 6: Criar registry, validação de DataFrame, exportador Excel e serviço de conversão

**Files:**
- Create: `rateiosrh/registry.py`, `rateiosrh/services/validation.py`, `rateiosrh/services/excel_exporter.py`, `rateiosrh/services/conversion.py`
- Modify: `test_architecture.py`, `test_converter_pdf.py`

**Interfaces:**
- `PARSER_REGISTRY: dict[str, type[BaseParser]]` contém `unimed` e `bradesco_dental`.
- `get_parser(parser_id: str) -> BaseParser` retorna uma instância ou levanta `UnknownParserError`.
- `list_parsers() -> list[dict[str, str]]` retorna `id` e `display_name` sem if/elif por fornecedor.
- `ExcelExporter.export(df: pd.DataFrame, output_path: Path, observations: Sequence[str] = ()) -> Path` grava o Excel.
- `ConversionOutput(dataframe: pd.DataFrame, parser_id: str, observations: list[str])` carrega o resultado da conversão.
- `ConversionService.convert(parser_id: str, pdf_path: Path) -> ConversionOutput` executa registry, layout, parsing e validação.

- [ ] **Step 1: Escrever testes falhos do registry e serviço**

Adicionar:

```python
from pathlib import Path
import pytest

from rateiosrh.core.exceptions import UnknownParserError
from rateiosrh.registry import get_parser, list_parsers
from rateiosrh.services.conversion import ConversionService


def test_registry_lists_only_implemented_converters():
    assert [item["id"] for item in list_parsers()] == ["unimed", "bradesco_dental"]
    assert get_parser("bradesco_dental").parser_id == "bradesco_dental"


def test_registry_rejects_unknown_converter():
    with pytest.raises(UnknownParserError):
        get_parser("bradesco_seguro")


def test_conversion_service_returns_dataframe():
    output = ConversionService().convert(
        "bradesco_dental", Path("FLEXIV_1166046_1_082026.pdf")
    )
    assert output.parser_id == "bradesco_dental"
    assert output.dataframe.shape == (279, 13)
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_architecture.py -k 'registry or conversion_service'`

Expected: FAIL por ausência do registry e dos serviços.

- [ ] **Step 3: Implementar registry e resolução central**

Importar apenas as classes concretas no registry, ordenar `list_parsers()` de forma estável e levantar `UnknownParserError` para identificadores ausentes. Não introduzir condicionais espalhadas em CLI ou Flask.

- [ ] **Step 4: Implementar `ValidationService` e `ConversionService`**

Executar exatamente nesta ordem:

```python
parser = get_parser(parser_id)
if not parser.validate_layout(pdf_path):
    raise LayoutMismatchError(parser_id, pdf_path)
frame = parser.parse(pdf_path)
errors = parser.validate_dataframe(frame)
if errors:
    raise DataFrameValidationError(errors)
```

Retornar o DataFrame sem conversão de valores e observações separadas.

- [ ] **Step 5: Escrever teste falho do ExcelExporter**

Adicionar:

```python
def test_excel_exporter_writes_exact_dataframe_columns(tmp_path):
    output = tmp_path / "bradesco.xlsx"
    result = ConversionService().convert(
        "bradesco_dental", Path("FLEXIV_1166046_1_082026.pdf")
    )
    from rateiosrh.services.excel_exporter import ExcelExporter

    ExcelExporter.export(result.dataframe, output)
    assert output.exists()
```

- [ ] **Step 6: Implementar o exportador sem regras de parsing**

Usar `openpyxl` apenas para formatação: cabeçalho baseado em `df.columns`, filtro, congelamento em `A2`, larguras e formato textual. Se houver observações, gravá-las em aba separada. Nunca alterar o DataFrame nem inferir tipos.

- [ ] **Step 7: Executar testes de registry, serviço e Excel**

Run: `venv/bin/python -m pytest -q test_architecture.py -k 'registry or conversion_service or excel_exporter'`

Expected: PASS; o Excel Bradesco Dental possui 13 colunas, 280 linhas incluindo cabeçalho e filtro `A1:M280`.

---

### Task 7: Reapontar a CLI para o serviço comum e preservar compatibilidade

**Files:**
- Create: `cli.py`
- Modify: `converter_pdf.py`, `test_converter_pdf.py`

**Interfaces:**
- `cli.run(argv: Sequence[str] | None = None) -> int` aceita `converter_id`, PDF e `-o/--output`.
- `converter_pdf.run` continuará aceitando a forma legada `run([pdf, "-o", output])` e usará `unimed` por compatibilidade.
- Nova forma: `venv/bin/python cli.py bradesco_dental FLEXIV_1166046_1_082026.pdf -o resultado.xlsx`.

- [ ] **Step 1: Escrever testes falhos para CLI multi-layout**

Adicionar:

```python
def test_cli_writes_bradesco_dental_workbook(tmp_path):
    output = tmp_path / "resultado.xlsx"
    from cli import run

    assert run([
        "bradesco_dental",
        "FLEXIV_1166046_1_082026.pdf",
        "-o", str(output),
    ]) == 0
    assert output.exists()


def test_cli_rejects_incompatible_selected_parser(tmp_path):
    output = tmp_path / "nao_deve_existir.xlsx"
    from cli import run

    assert run([
        "bradesco_dental",
        "ANALITICO_TAXA_0054041451-1.PDF",
        "-o", str(output),
    ]) != 0
    assert not output.exists()
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_converter_pdf.py -k 'cli'`

Expected: os testes novos falham porque a CLI multi-layout ainda não existe.

- [ ] **Step 3: Implementar `cli.py` usando `ConversionService`**

Resolver o parser somente pelo `parser_id`, registrar logs de páginas/linhas/avisos e retornar códigos não zero sem criar o Excel quando ocorrer erro.

- [ ] **Step 4: Atualizar `converter_pdf.py` como fachada UNIMED**

Manter a assinatura legada, definir `unimed` implicitamente apenas nessa fachada e delegar a exportação ao `ExcelExporter`. Não adicionar seleção de outros fornecedores nesse módulo.

- [ ] **Step 5: Executar testes de CLI e regressão**

Run: `venv/bin/python -m pytest -q test_converter_pdf.py`

Expected: PASS nos testes antigos e nos novos comportamentos de saída/erro.

---

### Task 8: Implementar interface/backend Flask

**Files:**
- Create: `app.py`, `rateiosrh/web/routes.py`, `rateiosrh/web/templates/index.html`
- Create: `test_web.py`

**Interfaces:**
- `create_app() -> Flask` constrói a aplicação.
- `GET /` responde 200 e lista os itens de `list_parsers()`.
- `POST /convert` recebe `converter_id` e campo multipart `pdf`, retornando `.xlsx` em sucesso.
- Erro de conversor/arquivo: 400.
- Layout incompatível: 422.
- Falha interna controlada: 500.

- [ ] **Step 1: Escrever testes falhos usando Flask test client**

Adicionar:

```python
from io import BytesIO

from app import create_app


def test_index_lists_registered_converters():
    response = create_app().test_client().get("/")
    assert response.status_code == 200
    assert b"Converter UNIMED" in response.data
    assert b"Converter Bradesco Dental" in response.data


def test_convert_returns_xlsx_for_selected_parser():
    client = create_app().test_client()
    with open("FLEXIV_1166046_1_082026.pdf", "rb") as source:
        response = client.post(
            "/convert",
            data={
                "converter_id": "bradesco_dental",
                "pdf": (BytesIO(source.read()), "entrada.pdf"),
            },
            content_type="multipart/form-data",
        )
    assert response.status_code == 200
    assert response.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def test_convert_rejects_pdf_for_wrong_parser():
    client = create_app().test_client()
    with open("ANALITICO_TAXA_0054041451-1.PDF", "rb") as source:
        response = client.post(
            "/convert",
            data={
                "converter_id": "bradesco_dental",
                "pdf": (BytesIO(source.read()), "entrada.pdf"),
            },
            content_type="multipart/form-data",
        )
    assert response.status_code == 422
```

- [ ] **Step 2: Executar e confirmar a falha**

Run: `venv/bin/python -m pytest -q test_web.py`

Expected: FAIL porque `app.py` e as rotas ainda não existem.

- [ ] **Step 3: Implementar factory e template**

`create_app()` deverá registrar as rotas e não conter condicionais por fornecedor. O template deverá iterar sobre os metadados do registry e enviar `converter_id` junto do arquivo. `app.py` deverá expor `app = create_app()` e permitir execução direta com `venv/bin/python app.py` usando `app.run(debug=False)`.

- [ ] **Step 4: Implementar upload temporário e resposta**

Usar `tempfile.TemporaryDirectory()` ou arquivos temporários com limpeza garantida. Processar com `ConversionService`, exportar para um buffer/arquivo temporário e retornar com `send_file`. Garantir que o arquivo esteja disponível durante o envio e seja removido após a resposta.

- [ ] **Step 5: Implementar mapeamento de erros HTTP**

Mapear `UnknownParserError`/arquivo ausente para 400, `LayoutMismatchError` para 422 e `DataFrameValidationError` para 422, com mensagens controladas no template/resposta.

- [ ] **Step 6: Executar testes web**

Run: `venv/bin/python -m pytest -q test_web.py`

Expected: PASS nos casos de listagem, conversão e rejeição.

---

### Task 9: Validar integração completa e atualizar documentação de continuidade

**Files:**
- Modify: `requirements.txt`, `contexto_sessao.md`
- Review: implementation files and tests from Tasks 1–8; qualquer correção encontrada nesta etapa deverá vir acompanhada de um teste de regressão específico.

**Interfaces:**
- A mesma instância de serviço deve atender CLI e Flask.
- Os dois parsers devem permanecer independentes e registrados sem condicionais distribuídas.

- [ ] **Step 1: Executar a suíte completa dentro da venv**

Run: `venv/bin/python -m pytest -q`

Expected: todos os testes passam.

- [ ] **Step 2: Verificar compilação dos módulos**

Run: `venv/bin/python -m compileall -q rateiosrh app.py cli.py converter_pdf.py`

Expected: comando termina com código 0.

- [ ] **Step 3: Executar conversões finais pela CLI**

Run: `venv/bin/python cli.py unimed ANALITICO_TAXA_0054041451-1.PDF -o /tmp/unimed-final.xlsx`

Run: `venv/bin/python cli.py bradesco_dental FLEXIV_1166046_1_082026.pdf -o /tmp/bradesco-dental-final.xlsx`

Expected: dois arquivos são gerados; o primeiro mantém 185 linhas e 10 colunas, o segundo 279 linhas e 13 colunas.

- [ ] **Step 4: Comparar as saídas com as referências**

Usar `venv/bin/python` e `pandas.read_excel(..., dtype="string")` apenas para comparação/auditoria, verificando:

```python
assert generated.columns.tolist() == reference.columns.tolist()
assert generated.fillna("").equals(reference.fillna(""))
```

Para o UNIMED, comparar contra o Excel anterior com a regra já validada de código/nome separado. Para o Bradesco Dental, comparar contra `FLEXIV_1166046_1_082026_referencia.xlsx`.

- [ ] **Step 5: Verificar que valores financeiros continuam textuais**

Confirmar no DataFrame e no Excel que `15,00-`, `6,00-`, `0,00` e `08/2026` não foram convertidos durante a extração. Não executar cálculo financeiro nesta tarefa.

- [ ] **Step 6: Atualizar `contexto_sessao.md`**

Registrar arquivos criados, testes, quantidade de registros de cada parser, comandos executados, compatibilidade Flask/CLI e qualquer pendência. Manter o procedimento para adicionar novos layouts apontando para a especificação arquitetural.

- [ ] **Step 7: Revisar a documentação final e o estado do workspace**

Run: `rg -n "if|elif|PARSER_REGISTRY|DataFrame|bradesco_dental|unimed" rateiosrh app.py cli.py converter_pdf.py`

Confirmar que a seleção por fornecedor aparece somente no registry, que nenhum parser importa o outro e que não há rateamento implementado.

---

## Checkpoints de revisão

- Após Task 2: leitor comum reconhece os dois PDFs e a regressão UNIMED continua verde.
- Após Task 5: parser Bradesco Dental reproduz exatamente os 279 registros do Excel conferido.
- Após Task 6: DataFrame, validação, registry e Excel funcionam sem Flask.
- Após Task 8: Flask lista os conversores, processa o parser escolhido e rejeita layout incompatível.
- Após Task 9: suíte completa, compilação e comparações finais concluídas.

## Critério de conclusão

A implementação só será considerada concluída quando os dois parsers retornarem DataFrames validados, a CLI e o Flask usarem o registry e o `ConversionService`, os Excels forem gerados exclusivamente a partir dos DataFrames, o PDF incompatível for rejeitado e a documentação de continuidade permitir adicionar um novo layout sem alterar parsers existentes.
