SHELL := /bin/bash

.PHONY: build test test-verbose test-cover test-short lint fmt vet clean

# Build all packages
build:
	go build ./...

# Run all tests (excludes examples which have known build issues)
test:
	go test ./pkg/... ./test/...

# Run tests with verbose output
test-verbose:
	go test -v ./pkg/... ./test/...

# Run tests with coverage report
test-cover:
	go test -cover ./pkg/... ./test/...

# Run tests with coverage profile and HTML report
test-cover-html:
	mkdir -p coverage
	go test -coverprofile=coverage/coverage.out ./pkg/... ./test/...
	go tool cover -html=coverage/coverage.out -o coverage/coverage.html
	@echo "Coverage report: coverage/coverage.html"

# Run short tests only (skip long-running tests)
test-short:
	go test -short ./pkg/... ./test/...

# Run tests with JSON output (for CI)
test-ci:
	mkdir -p coverage
	go test -json -v ./pkg/... ./test/... 2>&1 | tee coverage/gotest.log

# Run golangci-lint
lint:
	golangci-lint run --config .golangci-lint.yml

# Format code
fmt:
	go fmt ./...

# Run go vet
vet:
	go vet ./pkg/... ./test/...

# Run all checks (format, vet, lint, test)
check: fmt vet lint test

# Clean generated files
clean:
	rm -rf coverage/ test_output/
	find . -name "*.docx" -path "*/test_output/*" -delete
