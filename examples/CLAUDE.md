# CLAUDE.md — examples/

Standalone example programs demonstrating wordZero features. Each subdirectory is a separate `main` package.

## Running Examples

```bash
go run examples/basic/main.go
go run examples/table/main.go
```

Output files are written to `examples/output/` or the example's directory.

## Conventions

- Each example is self-contained with its own `main.go`
- New features should include a corresponding example
- Examples serve as both documentation and manual integration verification
- Use realistic content (mixed English/Chinese text) to test encoding

## Example Categories

- **basic/, formatting/** — Document creation, text styling
- **table/, table_style/, table_layout/, cell_advanced/** — Table operations
- **template_demo/, template_from_file_demo/, template_inheritance_demo/** — Template engine
- **markdown_demo/, word_to_markdown_demo/** — Markdown conversion
- **floating_images_demo/, image_persistence_demo/** — Image handling
- **page_settings/** — Page layout configuration
- **style_demo/** — Style system usage
- **toc_update_demo/** — Table of contents
- **advanced_features/** — Combined feature demonstrations
