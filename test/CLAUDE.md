# CLAUDE.md — test/

Integration tests that create actual DOCX files and validate the output.

## Running Tests

```bash
go test ./test/                        # All integration tests
go test -v ./test/ -run TestTemplate   # Specific test pattern
go test -short ./test/                 # Skip long-running tests
```

## Conventions

- Tests are in `package test` (external test package, uses public API only)
- Test output files go to `test_output/` — always clean up with `defer os.RemoveAll("test_output")`
- Test data (images, template DOCX files) in `test/testdata/`
- Tests create real DOCX files, save them, then reopen and validate content
- Table-driven test pattern: `tests := []struct{name string; ...}{...}`

## Test Coverage by Area

- `document_test.go` — Document lifecycle (create, save, open, modify)
- `table_*.go` — Table styles, dynamic merging, insert/merge fixes, nested tables
- `template_*.go` — Template rendering, inheritance, style preservation
- `markdown_*.go` — Math formulas, tables, task lists
- `footnotes_test.go` — Footnote config, number formats, positions
- `page_settings_test.go` — Page size, margins, orientation
- `text_formatting_test.go` — Text styling and formatting
- `toc_update_test.go` — Table of contents generation
