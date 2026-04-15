# Product Requirements Document

## 1. Product Summary

KRISHIV Seed Intake Platform is a production-grade web application for seed producing companies that need to intake raw seed from registered farmers, enforce certification-agency rules, maintain physical and transactional traceability, and reproduce official raw seed intake reports in required formats.

The product is designed for operational staff, godown teams, quality users, auditors, and administrators. It treats a lot as a certification traceability unit rather than as a simple storage grouping.

## 2. Problem Statement

Current manual Excel-driven processes create risk in five areas:

1. intake may happen against invalid or blocked crop registrations
2. lots may accidentally exceed certification limits of `200 QTL`
3. physical stack segregation may be lost in reporting
4. report totals may not reconcile back to source transactions
5. audit scrutiny is difficult because changes are not fully logged

## 3. Goals

1. Enforce intake only against valid crop registration records imported from Excel.
2. Auto-create and allocate certification lots without ever crossing `200 QTL`.
3. Preserve traceability from intake receipt to lot, stack, godown, and registration.
4. Generate Excel/PDF/print outputs in the structure of the required raw seed intake workbook.
5. Provide an extensible foundation for grading, packing, stock transfer, and audit workflows.

## 4. Non-Goals for Phase 1

1. ERP integration
2. mobile offline sync
3. weighbridge hardware integration
4. grading/packing execution
5. multi-language UI

## 5. Personas

### Admin

- configures seasons, users, tolerances, masters
- manages overrides and audit review

### Intake Operator

- searches approved registration
- enters receipt and stack details
- prints inward slip

### Godown Manager

- monitors stack occupancy, lot status, stock position
- reviews godown-wise reporting

### Quality User

- marks hold/reject status
- ensures incompatible quality statuses do not mix in same lot

### Viewer / Auditor

- reviews traceability, reports, overrides, and historical actions

## 6. Functional Scope

### 6.1 Authentication and Roles

- secure login
- role-based page and action access
- critical action audit logging

### 6.2 Farmer Master Import

- upload Excel file by season
- preview parsed rows
- validate mandatory headers and row content
- identify duplicates by season + crop registration code
- import valid rows and flag invalid rows
- store import batch metadata and source file reference

### 6.3 Registration Master

- searchable registration list
- status filters
- derived fields: received quantity, pending balance, lot count
- detail screen with receipt and lot history

### 6.4 Intake Receipt Entry

- registration search by farmer name, village, farmer code, crop registration code
- operator selects exact registration record
- system shows expected yield, certified area, status, received, balance
- operator enters receipt header and one or more receipt lines
- each line selects one godown and one stack
- system auto-allocates accepted quantity across eligible/open or newly created lots

### 6.5 Lot Engine

- determines whether existing open lot can continue
- creates new lot on stack change or quantity overflow
- maintains sequence per registration
- writes allocation lines for every receipt line

### 6.6 Godown and Stack

- godown master
- stack master under godown
- live stock per stack and per lot

### 6.7 Quality and Acceptance

- moisture capture
- sample number
- accepted / hold / rejected handling
- quality-status-based lot compatibility rules

### 6.8 Reporting

- Godown Wise Detail
- Farmer Wise Detail
- Summary
- Daily Intake Register
- Registration Pending vs Received
- Lot-wise Stock Ledger
- Stack-wise Stock Position

## 7. Mandatory Business Rules

### BR-01 Registration linkage required

Every intake line must reference an existing active crop registration. Manual free-entry intake is forbidden.

### BR-02 Lot quantity cap

Any single lot must remain `<= 200 QTL`.

### BR-03 Stack segregation

Material in different stacks creates separate lot continuity boundaries.

### BR-04 Compatible continuation only

Continuation into an existing lot is allowed only when farmer, registration, crop, variety, class stage, stack, and quality compatibility all match and capacity remains.

### BR-05 Overflow split

If a receipt line quantity would exceed available lot balance, the engine must allocate remaining quantity into one or more new lots.

### BR-06 Registration cap

Total received quantity for a registration cannot exceed allowed intake without an explicit authorized override.

### BR-07 Blocked records

Registrations with zero certified area, zero expected yield, or blocked/rejected/closed status cannot receive intake.

## 8. Key Derived Metrics

- `allowed_intake_qtl`: default from expected yield, overridable by policy
- `total_received_qtl`: sum of accepted allocated quantity across all linked lot allocations
- `balance_qtl`: allowed intake minus total received
- `lot_balance_qtl`: `200 - current_qty_qtl`

## 9. Acceptance Criteria

1. Intake cannot be saved without valid registration and stack selection.
2. Any intake that would breach registration cap is blocked unless override is permitted and captured.
3. No persisted lot quantity exceeds `200 QTL`.
4. Every report row can drill back to receipt line and lot allocation source.
5. Godown-wise, farmer-wise, and summary totals reconcile for the same filters.

## 10. Risks and Controls

- Excel data quality risk: solved with preview, validation, and exception logging.
- Quantity mismatch risk: solved with derived totals from allocation lines, not manual summaries.
- Certification non-compliance risk: solved with service-level rule enforcement and DB constraints.
- Operator error risk: solved with guided search, autofill, stack validation, and print previews.
