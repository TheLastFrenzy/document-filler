# Document Filler Behavior-Preserving Optimization Design

## Goal

Reduce duplicated infrastructure and maintenance coupling without changing any generated document content, section structure, image selection, field mapping, output filename, output format, or failure rule.

## Hard Compatibility Boundary

The optimization must preserve:

- all currently supported material type names and aliases;
- all builder function signatures used by `scripts/fill_document.py`;
- generated DOC, DOCX, and XLSX filenames and formats;
- ledger filtering, merged-row handling, template replacement, text generation, image extraction, and screenshot selection;
- current strict/partial generation behavior, including existing error conditions;
- current CLI arguments and return values.

Business-rule corrections identified during the audit, such as changing statistics-test completion percentages or material-specific required columns, are explicitly out of scope because they would change generation behavior.

## Design

### Shared Word Infrastructure

Create `scripts/materials/shared/office_word.py` as the public home for legacy DOC-to-DOCX conversion, DOCX-to-DOC conversion, and document saving. Keep compatibility wrappers in the existing N07 modules so existing imports, tests, and external callers continue to work.

The shared functions accept the TOC updater as an injected callback where needed. This preserves the current call path and allows existing tests that patch module-level TOC functions to remain effective.

### Material Registry Metadata

Keep the current dispatch behavior, but make the registry expose canonical material types, aliases, and supported-type display data from one source. The CLI error message will be built from registered N07 types plus the unchanged N08 legacy list instead of duplicating one long literal string.

No N08 builder will be moved in this phase. Moving the 3,705-line legacy implementation is a later, separately verifiable refactor and would create unnecessary risk in the same change set.

### Dependency Cleanup

Remove `pypdf` from required dependencies only after a failing test demonstrates that production code does not import it. Keep `reportlab` because the current DOCX statistics-test builder dynamically imports the PDF helper module and therefore still requires it.

### Regression Protection

Add focused tests for the new shared Word interface, compatibility wrappers, registry canonical names, alias resolution, output-path resolution, and unsupported-material reporting. Run the complete existing suite after every refactor step.

## Error Handling

Preserve existing exception types and output checks. Conversion continues to raise when the expected converted file is absent or empty. TOC update behavior remains unchanged.

## Success Criteria

- The full existing suite remains green.
- New tests fail before implementation and pass afterward.
- `git diff` contains no business-rule or document-content changes.
- CLI help, compilation, and dependency checks pass.
- The worktree remains free of generated Office artifacts.
