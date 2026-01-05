# ExcelTS Update

## Overview
- Fork modernization: Bun-only runtime, Bun build pipeline, and oxc lint/format tooling.
- Repo cleanup: removed legacy tooling, dotfiles, and old build artifacts.
- Code layout: renamed `lib/` to `src/`, converted source files to `.ts`.
- Tests: converted to TypeScript, migrated from chai/mocha to Bun's native expect.
- Dependencies: stripped to minimal set, removed entire chai/mocha ecosystem.
- **ESM migration complete**: Converted from CommonJS to ES Modules.

## Test Results
- **All tests passing**: 1090 pass, 1 skip, 0 fail
- **0 errors** from module system

## Dependencies

### Production (9 packages)
```
archiver, dayjs, fast-csv, jszip, readable-stream, saxes, tmp, unzipper, uuid
```

### Dev (5 packages)
```
@types/bun, express, oxfmt, oxlint, typescript
```

### Removed
- `@types/node`, `@types/chai`, `chai`, `chai-datetime`, `chai-xml`, `dirty-chai`
- `got`, `mocha`, `ts-node`, `eslint`, `grunt`, `babel`, `prettier`

## Key Changes

### Build & Tooling
- Build uses `bun build` for browser bundles (IIFE format)
- Linting: `oxlint`
- Formatting: `oxfmt`
- Test runner: Bun native (`bun test`)

### ESM Migration
- Package.json: `"type": "module"`
- All `require()` → `import`
- All `module.exports` → `export default` / named exports
- Relative imports use `.ts` extension (Bun-native)
- Removed legacy entry points (excel.js, index.ts)
- Main entry: `./src/exceljs.nodejs.ts`

### Codebase
- `lib/` → `src/` (all `.ts`)
- `spec/` → `tests/` (all `.ts`)
- Tests use Bun's native `expect` with custom XML/date helpers
- Preload script (`tests/preload.ts`) sets up global `verquire`
- XML comparison is attribute-order-insensitive

### Removed
- `build/`, `dist/es5/` (old artifacts)
- `gruntfile.js`, `.babelrc`, `.browserslistrc`, `.prettier*`, `.npmrc`
- `TODO.txt`, `UPGRADE-4.0.md`, `MODEL.md`
- Legacy eslint/mocha configs
- Mocha shim (`bun.test.ts`, `bun-setup.js`)
- `excel.js`, `index.ts` (CJS entry points)

## File Structure
```
excelts/
├── src/                    # TypeScript source (ESM)
├── tests/
│   ├── preload.ts         # Global setup (verquire)
│   ├── unit/              # Unit tests
│   ├── integration/       # Integration tests
│   ├── end-to-end/        # E2E tests (express)
│   ├── browser/           # Browser tests
│   ├── typescript/        # TS integration tests
│   └── utils/             # Test utilities
├── docs/
│   └── excel-format-notes.md  # Excel XML format reference
├── dist/                  # Built browser bundles
├── bunfig.toml           # Bun test config (preload)
├── tsconfig.json
├── package.json
├── index.d.ts            # Type declarations (manual)
└── AGENTS.md             # LLM agent guide
```

## Commands
```bash
bun install            # Install deps
bun run build          # Build browser bundles
bun test               # Run all tests
bun test tests/unit    # Run unit tests only
bun run lint           # Run oxlint
bun run format         # Check formatting
bun run format:write   # Auto-format
```

## Import Usage (ESM)
```typescript
// Default import
import ExcelJS from 'exceljs';

// Create workbook
const workbook = new ExcelJS.Workbook();
```

## CI/CD
- GitHub Actions workflow updated for Bun (`tests.yml`)
- Removed old npm/Node workflow (`asset-size.yml`)
- Tests run on ubuntu, macos, windows

## Linting
- Added `oxlint.json` config
- Disabled `no-sparse-arrays` (intentional test data for Excel rows)
- Disabled `no-unassigned-vars` (WIP features)
- All lint warnings fixed: 0 warnings, 0 errors

## Next Steps
- [x] ~~Fix remaining 20 XML element ordering test failures (update test data)~~ **DONE**
- [x] ~~Fix remaining 9 integration test failures~~ **DONE**
- [x] ~~Update GitHub Actions for Bun~~ **DONE**
- [x] ~~Fix all lint warnings~~ **DONE**
- [ ] Generate `index.d.ts` from source types
- [ ] Review remaining dependencies for further reduction
