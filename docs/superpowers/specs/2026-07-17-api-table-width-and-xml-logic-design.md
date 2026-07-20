# API Test Table Width And XML Logic Design

## Goal

Improve two generated-material details without changing ledger filtering, attachment extraction, template section replacement, output filenames, or material dependency order:

1. Make every N07 API test-report input-parameter table occupy the same layout width as its output-parameter table.
2. Replace generic N08 statistics-analysis XML summaries with deterministic, SQL-grounded descriptions of source tables, field mappings, joins, filters, comments, and supported database concepts.

## N07 Table Layout

The legacy template uses incompatible table properties: the input prototype has a `3806 pct` width and fixed layout, while the output prototype uses `7440 dxa` and an autofit layout. The generator currently preserves both prototypes exactly, so most input tables render visibly narrower.

The fix remains local to `materials/n07/api_test_report.py`. After cloning an input table, the builder copies the output prototype's table width, indentation, and layout mode, scales the input grid to the output prototype's total grid width, and writes matching `dxa` widths into each unmerged cell. Input-specific fonts, shading, borders, paragraph formatting, and column proportions remain unchanged.

## N08 XML Logic Analysis

The current builder reduces SQL to boolean hints such as whether `JOIN`, `CASE`, or `GROUP BY` appears. It also extracts `--` comments but does not use them in the business description. The replacement stays deterministic and does not call an LLM or add a SQL-parser dependency, because the source contains mixed Hive/Oracle syntax, custom functions, and placeholders such as `${taskid}`.

Each SQL node will be analyzed for:

- Node label and ordered natural-language `--` comments.
- Target table, source tables, and `INSERT ... SELECT` target-to-expression mappings.
- Join type and concrete `ON` field pairs.
- Important filters, latest-partition selection, deletion flags, and region constants.
- Aggregation, window functions, set operations, deduplication, and conditional expressions.

Descriptions will use real table and field names. Long direct mappings will be summarized as a count plus representative mappings; transformed or constant mappings take priority. Repeated district nodes will be grouped by mapping shape while retaining district/source-table differences.

Natural-language comments are included as node intent. Commented-out SQL fragments, dialect markers, and empty comments are excluded so inactive conditions are not presented as executed logic.

Database terminology is evidence-based:

- Equi-join or outer join -> relational algebra join terminology.
- Multiple identity fields in a join -> deterministic record linkage.
- `ROW_NUMBER ... PARTITION BY` -> window function and partition ordering.
- `MAX(partition)` latest-batch filters -> temporal latest-snapshot selection.
- `GROUP BY` with aggregates -> grouped aggregation.
- `UNION` -> set union.
- `CASE` -> rule-driven classification.

No concept is named unless the matching SQL construct exists.

## Output Impact

Regenerate these files in dependency order:

1. `N07-API接口开发/04-接口测试报告（含《API接口列表》）.docx`
2. `N08-数据统计分析/04-数据统计分析_结果表及使用说明.xlsx`
3. `N08-数据统计分析/01-数据统计分析_需求文档.docx`
4. `N08-数据统计分析/02-数据统计分析_设计文档.docx`

The statistics test document is unaffected because it does not consume the relation-description text.

## Verification

- Regression test proves input and output parameter tables have equal effective width after generation.
- Unit tests prove concrete field mappings, joins, comments, and evidence-based terminology are present.
- Unit tests prove unsupported theory names are absent from simple SQL.
- Full test suite passes.
- Regenerated Office packages pass ZIP integrity checks.
- N07 generated input/output table widths match in document XML.
- N08 result count, embedded flowchart count, and media count remain equal.
- Downstream N08 Word documents contain the new concrete mappings.
