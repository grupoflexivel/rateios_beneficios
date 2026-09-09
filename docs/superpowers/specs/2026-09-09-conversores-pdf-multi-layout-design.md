# Desenho arquitetural — conversores PDF multi-layout

## Status

Desenho aprovado pelo usuário em 2026-09-09. A implementação ainda não foi iniciada.

## Objetivo

Evoluir o conversor atual para uma aplicação com vários conversores isolados por layout, usando `pandas.DataFrame` como contrato central, Excel como saída de auditoria nesta fase e Flask como interface/backend web.

O layout anterior será identificado como `unimed`. O novo PDF, `FLEXIV_1166046_1_082026.pdf`, será implementado como `bradesco_dental` depois da validação manual já concluída.

## Escopo atual

- Migrar o conversor validado do layout anterior para `UnimedParser`, preservando seus resultados.
- Criar `BradescoDentalParser` para o layout Bradesco Dental.
- Criar contrato base, registry, serviços comuns, exportador Excel e aplicação Flask.
- Manter a CLI reutilizando o mesmo serviço de conversão.
- Atualizar dependências com `pandas`, `openpyxl` e `Flask`.
- Não implementar rateamento, cálculos financeiros de negócio ou CSV final nesta etapa.

## Arquitetura

```text
rateiosrh/
├── app.py
├── cli.py
├── requirements.txt
├── rateiosrh/
│   ├── core/
│   │   ├── pdf_reader.py
│   │   └── exceptions.py
│   ├── parsers/
│   │   ├── base.py
│   │   ├── unimed.py
│   │   └── bradesco_dental.py
│   ├── registry.py
│   ├── services/
│   │   ├── conversion.py
│   │   ├── validation.py
│   │   └── excel_exporter.py
│   └── web/
│       ├── routes.py
│       └── templates/
└── tests/
```

O leitor PDF comum conhece apenas a leitura estrutural, streams, spans, coordenadas e agrupamento de linhas. Coordenadas, regex, cabeçalhos e exceções de um fornecedor pertencem exclusivamente ao parser correspondente.

## Contrato dos parsers

Cada parser deverá declarar:

- `parser_id`;
- `display_name`;
- `columns`, em ordem fixa;
- `validate_layout(pdf_path)`, para compatibilidade básica;
- `parse(pdf_path) -> pandas.DataFrame`;
- `validate_dataframe(df)`;
- `get_columns()`.

`parse` deverá sempre retornar um DataFrame. A validação de layout deverá rejeitar o arquivo antes da extração quando os marcadores esperados não forem encontrados. O serviço não tentará adivinhar outro parser.

O DataFrame deverá conter exatamente as colunas declaradas pelo parser, na ordem declarada. O parser não poderá preencher, corrigir, recalcular, deduplicar, consolidar ou descartar registros.

## Esquemas

`UnimedParser` preservará o esquema atual validado de 10 colunas.

`BradescoDentalParser` usará exatamente:

1. `Certif.`
2. `Nome Beneficiário`
3. `Subfatura Nº N=Nova`
4. `Data Nascimento`
5. `Sexo`
6. `Est. Civil`
7. `Paren.`
8. `Plano`
9. `Data Início`
10. `Mov.`
11. `Mês/ano`
12. `Valor`
13. `Part. Benef`

Todos os campos serão extraídos como texto. Linhas de continuidade do Bradesco Dental manterão vazios os campos que não aparecem visualmente no PDF. Não serão adicionadas colunas técnicas ao DataFrame principal.

## Novo layout Bradesco Dental

O PDF possui 7 páginas:

- página 1: resumo e fatura;
- páginas 2 a 6: detalhe;
- página 7: mensagens.

A conversão validada possui 279 registros, distribuídos por página como `[0, 63, 63, 63, 63, 27, 0]`. Há 202 linhas com certificado e 77 linhas de continuidade. Os movimentos observados são `CM`, `CR`, `IR` e `IM`.

O parser deverá reconhecer a seção detalhada pelo cabeçalho e pelas coordenadas próprias do layout, ignorando resumo, boleto e mensagens. A leitura deverá aceitar a forma estrutural `/Type/Pages` usada neste PDF e continuar compatível com a forma do layout UNIMED.

O Excel de referência validado é `FLEXIV_1166046_1_082026_referencia.xlsx`.

## Registry e serviços

O registry será a única fonte de seleção:

```python
PARSER_REGISTRY = {
    "unimed": UnimedParser,
    "bradesco_dental": BradescoDentalParser,
}
```

`ConversionService` receberá um identificador e um PDF, resolverá o parser, validará o layout, executará a extração e validará o DataFrame.

`ValidationService` verificará schema, ordem, campos e regras específicas sem alterar valores.

`ExcelExporter` receberá um DataFrame e um caminho de saída. A planilha principal terá exatamente as colunas do DataFrame, filtro, congelamento da primeira linha e formatação textual. Observações poderão ser uma aba separada, sem modificar o schema principal.

## Flask

- `GET /`: exibe os conversores registrados.
- `POST /convert`: recebe `converter_id` e PDF multipart.
- O upload será armazenado temporariamente.
- O serviço comum processará o arquivo e o Excel será retornado para download.
- Arquivos temporários serão removidos ao fim do processamento.
- Conversor inexistente ou arquivo ausente: HTTP 400.
- PDF incompatível com o parser escolhido: HTTP 422.
- Falha interna controlada: HTTP 500.
- Nenhum Excel será gerado quando a validação falhar.

A CLI utilizará o mesmo `ConversionService`; não haverá regras duplicadas de seleção de parser.

## Fidelidade e valores financeiros

O texto original será preservado, inclusive zeros, casas decimais, datas, separadores e sinais negativos à direita, como `15,00-`.

Conversões numéricas não ocorrerão durante a extração. Na futura camada de rateamento, o texto será convertido explicitamente para `Decimal`, normalizando separador decimal e sinal à direita. O campo textual original permanecerá disponível; `float` não será usado indiscriminadamente.

## Testes

Os testes deverão cobrir:

- regressão do UNIMED contra os 185 registros já validados;
- Bradesco Dental contra os 279 registros do Excel de referência;
- exatamente 13 colunas e ordem correta;
- linhas de continuidade;
- datas e `Mês/ano` como texto;
- valores positivos, zerados e negativos;
- transições de página;
- validação e rejeição de layout incompatível;
- registry;
- upload, download e erros Flask;
- integridade do Excel.

## Procedimento para adicionar um novo layout

Este procedimento deverá ser seguido em toda nova sessão:

1. Ler `contexto_sessao.md` e este documento antes de alterar código.
2. Identificar o fornecedor, o nome do layout e um `parser_id` estável.
3. Analisar integralmente o PDF: páginas, cabeçalhos, rodapés, coordenadas, agrupamentos, continuidades, totais e exceções.
4. Criar um PDF de referência e obter uma conversão Excel validada manualmente.
5. Definir a lista exata de colunas e seus tipos de preservação textual.
6. Criar uma classe isolada em `rateiosrh/parsers/<novo_layout>.py`, derivada de `BaseParser`.
7. Reutilizar somente infraestrutura comum; não copiar regras específicas para outro parser.
8. Implementar `validate_layout` com assinaturas próprias do documento.
9. Implementar `parse(pdf) -> DataFrame`, mantendo uma linha por registro do PDF.
10. Implementar validações específicas sem alterar os dados extraídos.
11. Criar testes com o PDF de referência e comparar dados, ordem, sinais, datas e quantidade de linhas.
12. Registrar a classe em `PARSER_REGISTRY` com o novo identificador.
13. Confirmar que a interface Flask passa a listar o conversor automaticamente.
14. Validar incompatibilidade com pelo menos um PDF de outro layout.
15. Atualizar `requirements.txt`, testes e `contexto_sessao.md` somente se necessário.

Adicionar um layout não exige alterar parsers existentes, regras de rateamento ou condicionais espalhados pela aplicação.

## Evolução futura

O `RateioService` receberá o DataFrame validado e produzirá um DataFrame processado. Um futuro `CsvExporter` consumirá esse resultado. Nenhuma dessas responsabilidades será incorporada aos parsers.
