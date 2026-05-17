# KRISHIV Seed Intake System Handover Report

Date: 20-Apr-2026  
Project Workspace: `D:\KRISHIV seed DATA`  
Prepared For: system transfer to another computer and future development continuity

## 1. Purpose of This Report

This report explains what has been built in the KRISHIV Seed Intake System from scratch till date, how the system currently works, what has already been deployed, what remains pending, and what the next computer or next developer should know before continuing the work.

This is a practical handover note, not just a technical summary.

It is intended to help the next setup:
- understand the business flow already implemented
- understand the application architecture
- know the important modules and rules
- know the current deployment status
- know what to test first after moving to another computer
- continue development safely without breaking existing work

---

## 2. System Overview

The system is a seed intake, lot allocation, discrepancy handling, slip generation, reporting, and farmer payment control application for KRISHIV.

It is designed around the real operational workflow of a seed procurement / seed intake business where farmer registrations are imported first, then receipt-wise intake is recorded, lots are created stack-wise, discrepancies are tracked, slips are generated, reports are exported, and financial vouchers are created against accepted intake.

The system has gradually moved from a simple intake tool into a more structured operations platform with:
- role-based login
- intake control
- lot traceability
- discrepancy management
- report generation
- farmer financial vouchering
- payment ledger tracking
- database backup/restore
- LAN and online deployment readiness

---

## 3. Current Technology Stack

### Frontend
- Next.js
- React
- TypeScript
- local PDF generation using `jspdf` and `jspdf-autotable`
- Excel export using `xlsx`

Frontend location:
- [apps/web](/D:/KRISHIV%20seed%20DATA/apps/web)

### Backend
- Fastify
- TypeScript
- Mongoose
- MongoDB

Backend location:
- [apps/api](/D:/KRISHIV%20seed%20DATA/apps/api)

### Shared Domain Layer
- shared allocation logic and contracts

Shared package:
- [packages/domain](/D:/KRISHIV%20seed%20DATA/packages/domain)

### Database
- MongoDB Atlas is now connected for online use

### Source Control
- GitHub repository:
  - [piyushjainkota/Krishiv](https://github.com/piyushjainkota/Krishiv)

---

## 4. Repository / Folder Structure

Main working folders:
- [apps/api](/D:/KRISHIV%20seed%20DATA/apps/api) - backend API
- [apps/web](/D:/KRISHIV%20seed%20DATA/apps/web) - frontend application
- [packages/domain](/D:/KRISHIV%20seed%20DATA/packages/domain) - shared logic
- [docs](/D:/KRISHIV%20seed%20DATA/docs) - design notes and UI previews

Important documentation / preview files already created:
- [erp-busy-ui-preview.html](/D:/KRISHIV%20seed%20DATA/docs/erp-busy-ui-preview.html)
- [erp-preview-01-classic-admin.html](/D:/KRISHIV%20seed%20DATA/docs/erp-preview-01-classic-admin.html)
- [erp-preview-02-manufacturing-desk.html](/D:/KRISHIV%20seed%20DATA/docs/erp-preview-02-manufacturing-desk.html)
- [erp-preview-03-tabular-mis.html](/D:/KRISHIV%20seed%20DATA/docs/erp-preview-03-tabular-mis.html)
- [erp-preview-04-form-ledger.html](/D:/KRISHIV%20seed%20DATA/docs/erp-preview-04-form-ledger.html)
- [erp-preview-05-gov-enterprise-hybrid.html](/D:/KRISHIV%20seed%20DATA/docs/erp-preview-05-gov-enterprise-hybrid.html)
- [gov-portal-ui-preview.html](/D:/KRISHIV%20seed%20DATA/docs/gov-portal-ui-preview.html)
- [mobile-user-ui-preview.html](/D:/KRISHIV%20seed%20DATA/docs/mobile-user-ui-preview.html)

---

## 5. Business Flow Implemented So Far

The current system broadly works in this order:

1. Farmer registration master is imported from Excel.
2. Registration master becomes the intake base.
3. Receipt-wise seed intake is recorded against a valid registration.
4. Each receipt line is assigned to a godown and stack.
5. Lot allocation is generated automatically within the 200 QTL cap.
6. If intake exceeds expected yield, discrepancy entry is created.
7. Discrepancy can later be shifted / resolved through a separate flow.
8. Slip Print Center generates:
   - individual slips
   - overall consolidated slips
   - bulk overall slip PDF
9. Reports are generated in Excel and PDF.
10. Financial Voucher module creates farmer-wise payment voucher.
11. Voucher supports payment ledger, part payments, status progression, and ledger PDF.

---

## 6. Core Functional Modules Built

## 6.1 Login and Role Management

Roles currently implemented:
- `ADMIN`
- `MANAGER`
- `USER`

Current default users seeded by backend:
- `admin`
- `manager`
- `user`

These are ensured in backend startup compatibility logic.

Role behavior currently:
- `ADMIN`
  - full control
  - import
  - validate
  - voucher actions
  - maintenance
  - delete/edit
- `MANAGER`
  - entry and edit
  - voucher access
  - no import
  - no maintenance
  - no delete
- `USER`
  - mainly intake entry use
  - no edit/delete/import/voucher/maintenance

Important note:
- login is currently simple credential matching through the database
- passwords are still simple stored values in `passwordHash` field
- this is functional, but should be hardened before serious public production rollout

Relevant files:
- [apps/api/src/modules/seed/seed.routes.ts](/D:/KRISHIV%20seed%20DATA/apps/api/src/modules/seed/seed.routes.ts)
- [apps/api/src/modules/seed/seed.service.ts](/D:/KRISHIV%20seed%20DATA/apps/api/src/modules/seed/seed.service.ts)
- [apps/web/app/page.tsx](/D:/KRISHIV%20seed%20DATA/apps/web/app/page.tsx)

---

## 6.2 Farmer Master Import / Registration Master

Implemented:
- Excel-based import of farmer registration master
- import is restricted to `ADMIN`
- import also requires separate extra admin import password
- imported records include:
  - reg code
  - farmer name
  - father name
  - village
  - block
  - district
  - crop
  - variety
  - class
  - expected yield
  - balance

Registration status logic supports:
- `ACTIVE`
- `BLOCKED`
- `CLOSED`
- `EXHAUSTED`

Registration Master also supports:
- deposit view popup
- receipt-wise detail
- from deposit view, receipt can be directly opened into edit flow

---

## 6.3 Godown and Stack Masters

Implemented:
- Godown creation
- Stack creation under selected godown
- admin-oriented masters module
- master list / register style display

This is the base for physical stock placement and lot traceability.

---

## 6.4 Intake Entry

Implemented:
- receipt-wise intake entry
- registration-based intake only
- multiple receipt lines
- gross weight
- bag count
- weight per bag
- net weight
- vehicle no.
- stack allocation
- remarks
- moisture field exists in intake data structure

Business rules:
- intake cannot be entered without valid registration
- intake line must map to godown and stack
- receipt quantity is calculated and recorded carefully
- registration totals update after save

The intake system was further improved with:
- better user feedback dialogs
- mobile responsiveness
- deposit view linkage

---

## 6.5 Intake Entry Edit

Implemented:
- existing receipt can be loaded and edited
- receipt can also be deleted and re-entered if necessary
- edit flow is accessible from deposit view

This is the main correction path for data entry mistakes.

Current practical rule:
- small mistake -> edit receipt
- fully wrong receipt -> delete and re-enter

---

## 6.6 Lot Allocation and Lot Ledger

Implemented:
- automatic lot creation and filling logic
- lot capacity rule of `200 QTL`
- lot code generation
- stack-wise lot separation
- lot status handling:
  - `OPEN`
  - `FULL`
  - `VOID`
  - `CANCELLED`

Lot summary is visible in slips and related stock views.

Important later refinement already done:
- `VOID` lot entries should not appear in slip print preview or downloaded PDF summaries

---

## 6.7 Discrepancy Register and Discrepancy Shift

Implemented:
- automatic discrepancy generation when intake crosses expected yield
- discrepancy details include:
  - discrepancy no.
  - receipt no.
  - expected quantity
  - total received after receipt
  - excess quantity
  - estimated excess bags
- status handling:
  - `OPEN`
  - `SHIFT_PENDING`
  - `RESOLVED`

Shift handling implemented:
- discrepancy stock can be shifted to another stack
- approval and remarks captured

Data integrity work already performed:
- discrepancy register was checked against live receipt data
- a known incorrect Abdul Majid discrepancy receipt linkage was corrected
- current discrepancy register was revalidated at that point

---

## 6.8 Slip Print Center

Implemented:
- slip generation center
- single slip preview
- overall consolidated intake slip
- PDF download
- bulk overall slip download

Filters added for bulk overall slip:
- search
- district
- village
- seed class
- crop
- only registrations with intake

Important refinements already done:
- `VOID` lot details excluded from overall slip lot summary
- net weights in preview/PDF rounded to 2 decimal places

Slip module is already operationally useful.

---

## 6.9 Reports Module

The reports module has grown substantially.

Currently implemented report types include:
- Godown wise detail
- District wise detail
- Farmer wise detail
- Overall intake
- Summary
- Daily intake register
- Registration pending received
- Lot wise stock ledger
- Stack wise stock position
- Stack card register
- Discrepancy register

Important reporting enhancements already done:

### Daily Intake Register
- receipt-line-based reporting
- Excel export
- PDF export

### District Wise Detail
- district filter added
- Excel export creates separate sheet per district in one workbook
- totals included

### Overall Intake Report
fields:
- sr no.
- farmer name
- father name
- village
- district
- expected yield
- net intake quantity
- balance quantity

### Stack Card Register
implemented in Reports module with Excel export

Fields:
- SR. NO
- FARMER NAME
- FARMER REG CODE
- NUMBER OF BAGS IN STACK
- TOTAL NET WEIGHT IN STACK

Excel output:
- one workbook
- separate sheet for each stack

---

## 6.10 Financial Voucher Module

This module is now one of the most advanced parts of the system.

Main document:
- Farmer Overall Intake Cum Payment Voucher

Key implemented features:
- voucher generation registration-wise
- discrepancy-aware voucher calculation
- hides discrepancy section when discrepancy does not exist
- PDF download only
- paper size changed to portrait as requested
- header field font size reduced
- total bags added
- voucher numbering changed to season serial style
  - example: `RABI25-26/01`
- voucher starts in `DRAFT`

Voucher rules implemented:
- one voucher per registration
- existing voucher can be edited
- duplicate/new second voucher for same reg code is blocked

Voucher register actions implemented:
- Edit
- Delete
- Paid / Mark Paid
- Download PDF
- Ledger View

Paid control logic:
- once voucher is paid, edit/delete requires admin password
- draft voucher can be edited/deleted normally

Later enhancements implemented:
- final amount rounding
- part payment support
- payment ledger popup
- net paid and balance columns in voucher register
- payment status progression:
  - `DRAFT`
  - `PART PAID`
  - `PAID`

Payment entry fields:
- payment date
- amount
- transaction no.
- remarks
- mode treated as RTGS/NEFT workflow

Ledger improvements:
- ledger popup now behaves more like deposit view
- ledger PDF redesigned in account-ledger style
- A5 paper size
- includes:
  - voucher no.
  - farmer name
  - father name
  - village
  - district
  - seed purchase amount
  - seed payment
  - final payable amount
  - ledger rows with debit/credit/balance
  - transaction no.
  - totals and signature block

Relevant files:
- [apps/web/app/page.tsx](/D:/KRISHIV%20seed%20DATA/apps/web/app/page.tsx)
- [apps/web/app/mvp.ts](/D:/KRISHIV%20seed%20DATA/apps/web/app/mvp.ts)
- [apps/api/src/modules/seed/seed.service.ts](/D:/KRISHIV%20seed%20DATA/apps/api/src/modules/seed/seed.service.ts)

---

## 6.11 Database Backup and Restore

Implemented:
- admin-only modules
  - Database Backup
  - Database Restore
- asks for backup directory
- asks for restore folder/directory
- uses MongoDB Database Tools

Practical note:
- it depends on `mongodump.exe` and `mongorestore.exe`
- tools were found locally under the downloaded MongoDB tools folder

Important operational learning:
- not every backup folder generated earlier was valid
- restore reliability must always be checked against actual dump contents

Atlas migration was later completed successfully and validated separately.

---

## 6.12 Dashboard

Implemented dashboard improvements include:
- gross weight
- net weight
- clearer intake summary visibility

The dashboard is functional but still likely to evolve further depending on final UI redesign.

---

## 6.13 Dialog Feedback / User Messaging

System-wide feedback was improved:
- on save / error where page does not change
- browser dialog-style feedback now appears
- intended to make operator actions clearer

This was especially important for phone/LAN use.

---

## 7. Current UI / UX State

The live UI has gone through many refinements, but it is still in transition.

### What has been done
- sidebar redesigned into ERP-style grouped module -> submodule structure
- reports, finance, masters, administration now grouped more cleanly
- mobile responsiveness improved
- login page cleaned up
- admin-only maintenance controls added

### UI direction work done in preview form
Multiple static preview directions exist in `docs/`:
- ERP-style variants
- government portal style variant
- Busy ERP-like variant
- mobile-first USER role variant

### Important UX conclusion already reached
- `USER` role should not see the full desktop ERP clutter on mobile
- a simpler mobile-first intake console is more appropriate

This is why the file below was created:
- [mobile-user-ui-preview.html](/D:/KRISHIV%20seed%20DATA/docs/mobile-user-ui-preview.html)

This preview is specifically meant for:
- `user` login
- phone-only use
- intake-first workflow

---

## 8. Online / Deployment Status

## 8.1 GitHub

Code is pushed to:
- [https://github.com/piyushjainkota/Krishiv](https://github.com/piyushjainkota/Krishiv)

Recent push includes:
- finance ledger and UI preview updates
- `.gitignore` updated to exclude local dump/export folders

---

## 8.2 MongoDB Atlas

Atlas migration has been done.

Important notes:
- direct `mongodb+srv://` was not working from local machine due DNS SRV resolution problem
- non-SRV Atlas connection string was used successfully
- local database was dumped and restored into Atlas
- backend was then pointed to Atlas

At the time of successful migration, live data verification confirmed presence of core collections such as:
- registrations
- receipts
- lots
- discrepancies
- users

Do not expose Atlas credentials in future notes or commits.

If the system is moved to another computer:
- update local backend `.env` with the correct Atlas URI
- confirm backend connects before starting frontend

---

## 8.3 Render Backend

Backend is deployed on Render and is live.

Live URL:
- [https://krishiv-api.onrender.com](https://krishiv-api.onrender.com)

Health URL:
- [https://krishiv-api.onrender.com/health](https://krishiv-api.onrender.com/health)

Confirmed health response:
- `{"ok":true,"service":"krishiv-api"}`

Important Render setup detail:
- because of monorepo/workspace behavior, backend deploy was adjusted carefully
- build command used:
  - `npm install && npm run build --workspace @krishiv/api`
- start command used:
  - `cd apps/api && node --import tsx src/index.ts`

This start command is important because compiled ESM runtime had module resolution issues on Render, while `tsx` source start worked reliably.

---

## 8.4 Vercel Frontend

Frontend deployment process was started on Vercel.

Configuration used:
- project name similar to `krishiv-seed`
- root directory:
  - `apps/web`
- preset:
  - `Next.js`
- env variable:
  - `NEXT_PUBLIC_API_BASE=https://krishiv-api.onrender.com`

At the time of this report:
- backend is confirmed live
- frontend deployment flow was in progress and had intermittent Vercel UI glitches during import

This part should be rechecked on the target machine / final deployment session.

---

## 9. Current Known Operational URLs

### Local
- Frontend:
  - `http://127.0.0.1:3000`
- Backend:
  - `http://127.0.0.1:4000/health`

### LAN
- Frontend:
  - `http://192.168.1.33:3000`
- Backend:
  - `http://192.168.1.33:4000/health`

### Online
- Backend:
  - [https://krishiv-api.onrender.com](https://krishiv-api.onrender.com)
- Health:
  - [https://krishiv-api.onrender.com/health](https://krishiv-api.onrender.com/health)

Frontend online URL should be added here after Vercel is finalized.

---

## 10. Important Credentials / Sensitive Setup Notes

This report should not store actual live secrets.

What the next computer must know:
- backend uses `.env`
- MongoDB Atlas URI must be configured there
- `IMPORT_ADMIN_PASSWORD` must be configured there
- Atlas password was exposed during setup discussion and should be rotated if not already rotated

Current seeded login identities in the system:
- `admin`
- `manager`
- `user`

The next machine should verify the actual current passwords from the database or secure setup note, not from this report.

---

## 11. Important Data / Backup Notes

### Local backup/export folders
These were intentionally excluded from git:
- `atlas-migration-backup/`
- `exports/`

`.gitignore` was updated accordingly.

### Before moving to another computer
Do these steps:
1. confirm GitHub code is latest
2. confirm Atlas data is current
3. copy any important local docs and previews if needed
4. on new machine:
   - clone repo
   - install dependencies
   - configure backend `.env`
   - start backend
   - start frontend
   - test login and data load

---

## 12. Recommended Transfer Procedure to Another Computer

## Step 1: Clone project
- clone from GitHub repo

## Step 2: Install Node dependencies
From repo root:
- `npm install`

## Step 3: Configure backend environment
In backend env:
- `MONGODB_URI`
- `IMPORT_ADMIN_PASSWORD`
- any future production env values

## Step 4: Start backend
Recommended local dev start:
- from root:
  - `npm run dev:api`
or from API folder:
  - `npm run dev`

## Step 5: Start frontend
From root:
- `npm run dev:web`
or from web folder:
  - `npm run dev`

## Step 6: Validate immediately
Check:
- backend `/health`
- login page
- admin login
- registration master data visible
- one receipt open/edit
- one report preview
- one voucher open

---

## 13. Important Development Decisions Already Taken

These are important because a future developer should not accidentally undo them.

1. Voucher is integrated intake + payment voucher, not separate disconnected finance paper.
2. Voucher is one per registration, but editable.
3. Paid voucher requires admin password to edit/delete.
4. Ledger is account-ledger style, not just a payment list.
5. Slip PDFs should exclude `VOID` lot rows.
6. Net weights should be displayed with 2 decimal precision in slips.
7. `USER` role should get a simpler phone-focused UI, not the full ERP clutter.
8. Sidebar should be grouped module -> submodule, not a long flat list.
9. Database backup/restore should be admin-only.
10. Farmer master import requires extra admin import password.

---

## 14. Current Known Risks / Technical Debt

The system is functional, but there are still some areas that should be considered technical debt or follow-up work.

### 14.1 Authentication security
- passwords are still simple direct values in DB field
- should move to hashed passwords before serious public rollout

### 14.2 Frontend deployment finalization
- Vercel flow should be rechecked and completed
- online frontend URL should be documented after success

### 14.3 Render backend runtime approach
- current Render start uses `tsx` source startup
- it works, but later a cleaner production build/runtime path may be desirable

### 14.4 Backup/restore trust
- backup/restore should always be validated with actual test restore
- do not assume every generated backup folder is valid without inspection

### 14.5 UI consistency
- many functional improvements were added quickly
- a deeper role-wise UI cleanup is still a good future step

---

## 15. Suggested Immediate Next Development Steps

If development continues from the next machine, best next sequence is:

1. Finish Vercel frontend deployment and verify online login.
2. Rotate Atlas password if not yet rotated.
3. Move login credentials to better security model.
4. Implement mobile-first `USER` role UI based on:
   - [mobile-user-ui-preview.html](/D:/KRISHIV%20seed%20DATA/docs/mobile-user-ui-preview.html)
5. Continue ERP / government portal style redesign only after role-based workflows are stable.
6. Add final production checklist and smoke-test documentation.

---

## 16. Recommended Smoke Test After Transfer

After setting up on another machine, perform this short test:

1. Login as `admin`
2. Open Registration Master and confirm farmer data loads
3. Open deposit view for one registration
4. Open receipt edit from deposit view
5. Open Reports and preview:
   - Overall Intake
   - District Wise Detail
   - Daily Intake Register
6. Open Financial Voucher register
7. Open Ledger View for a voucher
8. Download one PDF
9. Check backend `/health`
10. If online frontend is live, confirm frontend can call Render backend

---

## 17. Final Summary

This system is no longer a blank or early prototype.

It already contains:
- real data import
- receipt-wise intake
- lot creation and stack traceability
- discrepancy control
- slip generation
- Excel/PDF reporting
- financial vouchers
- payment ledger
- admin maintenance features
- LAN use
- Atlas database migration
- live Render backend deployment

The system is now at a stage where:
- day-to-day operational use is possible
- handover to another machine is practical
- future development can continue in an organized way

The most important thing for the next person is:
- do not treat this as a generic app
- it is built around specific seed intake business logic and many decisions were made intentionally to match that workflow

If this report is followed, the system can be transferred and continued without losing the logic, direction, and operational understanding built so far.
