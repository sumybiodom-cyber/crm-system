# CRM Cloud Migration Plan

## Current storage

The CRM still uses browser `localStorage` as the default storage layer. This keeps the existing UI safe, but every browser has its own isolated data.

Current local keys include:

- `crm_employees`
- `crm_clients`
- `crm_orders`
- `crm_documents`
- `crm_suppliers`
- `crm_warehouses`
- `crm_products`
- `crm_payments`
- `crm_engineer_payments`
- `crm_finance_expenses`
- `crm_tax_invoices`
- `crm_contracts`
- `crm_contract_acts`
- `crm_bank_imports`
- `crm_requirements`
- `crm_purchases`
- `crm_receipts`
- `crm_movements`
- `crm_action_logs`
- `crm_backups`
- `crm_cash_shift`
- `sessionUserId`, `activeUserId`, `page`

## Phase 1 cloud bridge

Phase 1 does not rewrite the domain model. It adds shared Supabase/PostgreSQL tables that store the current frontend records as `jsonb` payloads:

- `crm_users`
- `crm_clients`
- `crm_orders`
- `crm_documents`

This lets the existing CRM move its most critical operational data from isolated browsers to a shared database without breaking orders, clients, documents, printing, QR, AI Command Bar, or role permissions.

Run:

```sql
db/supabase_phase1_cloud_sync.sql
```

Then set:

```env
VITE_SUPABASE_URL=https://your-project.supabase.co
VITE_SUPABASE_ANON_KEY=your-anon-key
VITE_CRM_CLOUD_SYNC=true
VITE_CRM_SYNC_INTERVAL_MS=5000
```

## Sync behavior

When cloud sync is enabled:

- CRM first loads local data so the app remains usable.
- It pulls shared `users`, `clients`, `orders`, and `documents` from Supabase.
- It merges remote data into the current workspace.
- Local edits are pushed back to Supabase with upsert semantics.
- Other workstations pull changes on the configured interval.

The SQL also enables Supabase Realtime publication for these tables. The frontend currently uses polling as the safe first transport; a later step can replace polling with Supabase channel subscriptions.

## Access model

Phase 1 keeps the CRM role/permission checks in the existing frontend logic. The Supabase SQL contains permissive internal test policies so the current frontend can sync with the anon key.

Before production, replace these policies with authenticated user policies:

- Users can read only data allowed by their role.
- Managers can read/write operational clients, orders, and allowed documents.
- Engineers can read assigned repairs and update allowed repair statuses.
- Accounting can access payments and documents, not system settings.
- Admins can manage users and permissions.

## Migration roadmap

### Stage 1: critical operational ERP

- Cloud sync: users, clients, orders, documents.
- Keep localStorage fallback.
- Keep existing AI workflow and print workflow.
- Verify multi-user scenario: admin creates a client, manager sees it without local import/export.

### Stage 2: warehouse and finance

- Move products, warehouses, receipts, movements, purchases, payments, expenses, cash shift.
- Add server-side validation for stock movement and payments.
- Add audit trail in PostgreSQL.

### Stage 3: role-aware backend

- Add real Supabase Auth or existing backend auth.
- Replace permissive test RLS policies with role-aware RLS.
- Move permission decisions that matter financially to the backend.

### Stage 4: AI automation

- Store AI Command Bar history and assistant events in shared tables.
- Allow AI to work with shared ERP data.
- Add guarded server actions for safe operational commands.

### Stage 5: normalized ERP schema

- Gradually migrate from `jsonb` bridge tables into normalized tables from `db/schema.sql`.
- Keep compatibility views or adapters while old UI state is phased out.

## Files touched in Phase 1

- `src/cloudSync.ts`
- `src/WarehouseCRM.tsx`
- `.env.example`
- `db/supabase_phase1_cloud_sync.sql`
- `docs/cloud-migration-plan.md`
