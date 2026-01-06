.PHONY: help install build clean lint format test test-unit test-integration test-e2e test-browser test-dist test-ts typecheck

# Default target
help:
	@echo "ExcelTS - Available commands:"
	@echo ""
	@echo "  make install         Install dependencies"
	@echo "  make build           Build browser bundles (dist/)"
	@echo "  make clean           Remove dist/"
	@echo ""
	@echo "  make test            Run all tests (build + unit + integration + e2e + browser + dist + ts)"
	@echo "  make test-unit       Run unit tests only"
	@echo "  make test-integration Run integration tests only"
	@echo "  make test-e2e        Run end-to-end tests only"
	@echo "  make test-browser    Run browser tests only"
	@echo "  make test-dist       Run dist tests only"
	@echo "  make test-ts         Run TypeScript tests only"
	@echo ""
	@echo "  make lint            Run oxlint"
	@echo "  make format          Check formatting (oxfmt)"
	@echo "  make format-fix      Auto-fix formatting"
	@echo "  make typecheck       Run TypeScript type checker"
	@echo ""
	@echo "  make check           Run lint + format + typecheck"
	@echo "  make ci              Full CI pipeline (check + test)"

# Setup
install:
	bun install

# Build
build:
	bun run build

clean:
	rm -rf dist

# Testing
test:
	bun test

test-unit:
	bun test tests/unit --timeout 20000

test-integration:
	bun test tests/integration --timeout 20000

test-e2e:
	bun test tests/end-to-end --timeout 20000

test-browser:
	bun test tests/browser --timeout 20000

test-dist:
	bun run build
	bun test tests/dist --timeout 20000

test-ts:
	bun test tests/typescript --timeout 20000

# Code quality
lint:
	bun run lint

format:
	bun run format

format-fix:
	bun run format:write

typecheck:
	bunx tsc --noEmit

# Combined targets
check: lint format typecheck

ci: check test
