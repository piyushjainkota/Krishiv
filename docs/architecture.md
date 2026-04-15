# System Architecture

## 1. Recommended Stack

- Frontend: `Next.js 15` + `TypeScript` + `App Router`
- Backend: `Node.js` + `Fastify` + `TypeScript`
- Database: `PostgreSQL`
- ORM: `Prisma`
- Validation: `Zod`
- Auth: session or JWT-backed RBAC
- Excel import/export: `xlsx` or `exceljs`
- PDF/print: server-side HTML template to PDF plus browser print layout

This stack is chosen for fast delivery, strong typing across layers, and easy report/export support.

## 2. High-Level Components

### Web App

- operator workflows
- master data management
- reporting UI
- dashboard and audit views

### API

- authentication and authorization
- import processing
- registration query services
- intake receipt orchestration
- lot allocation engine
- reporting queries and exports

### Database

- transactional seed intake model
- audit trail
- import history
- report source data

## 3. Core Backend Modules

1. `auth`
2. `users`
3. `seasons`
4. `imports`
5. `farmers`
6. `crop-registrations`
7. `godowns`
8. `stacks`
9. `intake-receipts`
10. `lots`
11. `quality`
12. `reports`
13. `audit`

## 4. Architectural Rules

1. `intake_receipt_lines` are the operational input unit.
2. `lot_allocations` are the traceability and quantity truth source.
3. `crop_registrations.total_received_qtl` is a maintained aggregate, not an editable field.
4. Reports read from normalized transactional tables or materialized views, never from denormalized manual sheets.
5. Critical actions run inside database transactions.

## 5. Transaction Boundary for Intake Save

One intake save transaction should:

1. validate registration status and balance
2. lock the registration row
3. create receipt header
4. create receipt lines
5. allocate quantity across eligible or new lots
6. update lot quantities
7. update registration totals
8. create audit log
9. commit

## 6. Future Extensibility

The model intentionally leaves room for:

- grading batches linked to raw lots
- stock transfer between stacks/godowns
- packing and finished goods lots
- weighbridge integrations
- media attachments and GPS evidence

## 7. Reporting Strategy

### Primary approach

Query normalized tables into report DTOs, then render:

- HTML print layouts
- Excel export worksheets
- PDF export

### Template support

If the official workbook template becomes available, map the report DTOs into exact worksheet coordinates and preserve headings and groupings.
