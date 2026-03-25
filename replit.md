# Workspace

## Overview

pnpm workspace monorepo using TypeScript. Each package manages its own dependencies.

## Stack

- **Monorepo tool**: pnpm workspaces
- **Node.js version**: 24
- **Package manager**: pnpm
- **TypeScript version**: 5.9
- **API framework**: Express 5
- **Database**: PostgreSQL + Drizzle ORM
- **Validation**: Zod (`zod/v4`), `drizzle-zod`
- **API codegen**: Orval (from OpenAPI spec)
- **Build**: esbuild (CJS bundle)

## Structure

```text
artifacts-monorepo/
├── artifacts/              # Deployable applications
│   ├── api-server/         # Express API server with payroll engine
│   └── payroll-app/        # React + Vite frontend for payroll calculator
├── lib/                    # Shared libraries
│   ├── api-spec/           # OpenAPI spec + Orval codegen config
│   ├── api-client-react/   # Generated React Query hooks
│   ├── api-zod/            # Generated Zod schemas from OpenAPI
│   └── db/                 # Drizzle ORM schema + DB connection
├── scripts/                # Utility scripts (single workspace package)
│   └── src/                # Individual .ts scripts, run via `pnpm --filter @workspace/scripts run <script>`
├── pnpm-workspace.yaml     # pnpm workspace (artifacts/*, lib/*, lib/integrations/*, scripts)
├── tsconfig.base.json      # Shared TS options (composite, bundler resolution, es2022)
├── tsconfig.json           # Root TS project references
└── package.json            # Root package with hoisted devDeps
```

## TypeScript & Composite Projects

Every package extends `tsconfig.base.json` which sets `composite: true`. The root `tsconfig.json` lists all packages as project references. This means:

- **Always typecheck from the root** — run `pnpm run typecheck` (which runs `tsc --build --emitDeclarationOnly`). This builds the full dependency graph so that cross-package imports resolve correctly. Running `tsc` inside a single package will fail if its dependencies haven't been built yet.
- **`emitDeclarationOnly`** — we only emit `.d.ts` files during typecheck; actual JS bundling is handled by esbuild/tsx/vite...etc, not `tsc`.
- **Project references** — when package A depends on package B, A's `tsconfig.json` must list B in its `references` array. `tsc --build` uses this to determine build order and skip up-to-date packages.

## Root Scripts

- `pnpm run build` — runs `typecheck` first, then recursively runs `build` in all packages that define it
- `pnpm run typecheck` — runs `tsc --build --emitDeclarationOnly` using project references

## Packages

### `artifacts/api-server` (`@workspace/api-server`)

Express 5 API server. Routes live in `src/routes/` and use `@workspace/api-zod` for request and response validation and `@workspace/db` for persistence.

- Entry: `src/index.ts` — reads `PORT`, starts Express
- App setup: `src/app.ts` — mounts CORS, JSON/urlencoded parsing, routes at `/api`
- Routes: `src/routes/index.ts` mounts sub-routers; `src/routes/health.ts` exposes `GET /health` (full path: `/api/health`)
- **Payroll routes** (`src/routes/payroll.ts`): `POST /api/payroll/upload-master`, `POST /api/payroll/upload-reports`, `POST /api/payroll/process`, `GET /api/payroll/download`, `GET /api/payroll/status`, `POST /api/payroll/reset`
- **Payroll engine** (`src/lib/payroll-engine.ts`): Business logic for salary calculation using ExcelJS. In-memory session state.
- Depends on: `@workspace/db`, `@workspace/api-zod`, `exceljs`, `multer`
- `pnpm --filter @workspace/api-server run dev` — run the dev server
- `pnpm --filter @workspace/api-server run build` — production esbuild bundle (`dist/index.cjs`)

### `artifacts/payroll-app` (`@workspace/payroll-app`)

React + Vite frontend for the "ММ Расчёт Графика" payroll schedule calculator. All UI is in Russian. Dark professional theme with glass-morphism effects.

- Entry: `src/main.tsx` → `src/App.tsx` → `src/pages/Home.tsx`
- Key components: `StepIndicator` (3-step wizard), `Dropzone` (drag-drop file upload), `ResultsTable` (sortable salary results table with existing/new hours breakdown)
- Hooks: `src/hooks/use-payroll.ts` — wraps generated React Query hooks for payroll API, includes file download logic
- Uses: `@workspace/api-client-react` for generated hooks, `react-dropzone`, `framer-motion`, `@tanstack/react-table`, `lucide-react`

### `lib/db` (`@workspace/db`)

Database layer using Drizzle ORM with PostgreSQL.

### `lib/api-spec` (`@workspace/api-spec`)

Owns the OpenAPI 3.1 spec (`openapi.yaml`) and the Orval config (`orval.config.ts`). Running codegen produces output into two sibling packages.
Run codegen: `pnpm --filter @workspace/api-spec run codegen`

### `lib/api-zod` (`@workspace/api-zod`)

Generated Zod schemas from the OpenAPI spec.

### `lib/api-client-react` (`@workspace/api-client-react`)

Generated React Query hooks and fetch client from the OpenAPI spec.

## Payroll Business Rules

- **Per-employee salary (оклад)**: Read from master file column 70 (e.g., 4500 or 5100 RUB). Default 4500 if not found.
- **Salary limit**: 24,500 RUB max per employee (hours scaled proportionally if exceeded)
- **Overtime**: 2x hourly rate for hours above monthly production calendar norm
- **Hourly rate**: employee_salary / production_calendar_norm_hours
- **2026 production calendar norms (40-hour week)**: Jan=120h, Feb=152h, Mar=168h, Apr=175h, May=151h, Jun=167h, Jul=184h, Aug=168h, Sep=176h, Oct=176h, Nov=159h, Dec=176h
- **Hours merging**: Existing hours from master + report hours merged per day (take max if both exist)
- **Fired before 17th**: 0 hours/salary
- **Fired on/after 17th**: only hours up to dismissal date counted
- **ОТПУСК (vacation)**: 0 hours
- **БОЛЬНИЧНЫЙ (sick)**: 0 hours unless end date specified, then hours after sick-end
- **ЗА СВОЙ СЧЁТ (unpaid leave)**: 0 hours/salary
- **Employee filtering**: Skip rows containing "бухгалтер", "руководитель", "главн"
- **Employee section detection**: Stop parsing after 3 consecutive empty rows in column B
- **FIO matching**: Exact normalized match + fuzzy by surname prefix and initials (dots removed)
- **Master file format**: Excel with day numbers in row 3, 2 columns per day (col1=request #1, col2=request #2), employees start from headerRow+2, salary in col 70
- **Output file**: New hours written to col2 of day pair, green fill (FF90EE90) for genuinely new hours, red fill for fired employees
- **Report file format**: Headers row 1 (Дата открытия, Кассир, НАЧИСЛЕНО/ОКРУГЛЕНИЕ), data rows 2+. Date can be Date object, string (DD.MM.YYYY), or Excel serial. Hours can be number or "HH:MM" string.
