# KRISHIV Seed Intake Platform

Production-oriented monorepo scaffold for a seed producing company to manage:

- farmer registration imports from Excel
- raw seed intake against valid crop registrations
- certification-safe lot creation and allocation
- godown and stack traceability
- audit logs and role-aware workflows
- official raw seed intake reporting outputs

## Workspace Layout

- `docs/` implementation-ready product, architecture, reporting, and roadmap documents
- `apps/api/` Fastify + Prisma API scaffold
- `apps/web/` Next.js operator dashboard scaffold
- `packages/domain/` shared domain contracts and allocation rules

## Core Design Principles

- intake cannot exist without a valid linked crop registration
- lot quantity cannot exceed `200 QTL`
- stack segregation drives lot separation
- every receipt line remains traceable to registration, stack, lot, and report row
- certification rules are enforced at UI, API, service, and persistence layers
