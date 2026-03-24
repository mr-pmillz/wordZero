# CLAUDE.md — benchmark/

Cross-language performance benchmarks comparing Go, JavaScript, and Python implementations.

## Running Benchmarks

```bash
# Go benchmarks only
go test -bench=. ./benchmark/golang/

# All languages (requires Node.js and Python)
task benchmark    # Uses Taskfile.yml

# Compare results
python3 benchmark/compare_results.py
```

## Structure

- `golang/` — Go benchmarks using `testing.B`
- `javascript/` — Node.js benchmarks for comparison
- `python/` — Python benchmarks for comparison
- `compare_results.py` — Cross-language analysis
- `results/` — Output artifacts (JSON, CSV)

## Benchmark Categories

Standard iteration counts: basic (50), complex (30), table (20), largeTable (10), largeDoc (5), memory (10)

Tests: basic document creation, complex formatting, table operations, large tables, large documents, memory usage.
