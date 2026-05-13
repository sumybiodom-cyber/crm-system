-- Phase 1 cloud bridge for the current React CRM state shape.
-- Run in Supabase SQL editor before setting VITE_CRM_CLOUD_SYNC=true.
-- These tables intentionally keep jsonb payloads so the existing CRM UI can
-- move from localStorage to shared PostgreSQL without a full domain rewrite.

create extension if not exists pgcrypto;

create table if not exists crm_users (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists crm_clients (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists crm_orders (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists crm_documents (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create or replace function crm_cloud_sync_set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists crm_users_set_updated_at on crm_users;
create trigger crm_users_set_updated_at before update on crm_users
for each row execute function crm_cloud_sync_set_updated_at();

drop trigger if exists crm_clients_set_updated_at on crm_clients;
create trigger crm_clients_set_updated_at before update on crm_clients
for each row execute function crm_cloud_sync_set_updated_at();

drop trigger if exists crm_orders_set_updated_at on crm_orders;
create trigger crm_orders_set_updated_at before update on crm_orders
for each row execute function crm_cloud_sync_set_updated_at();

drop trigger if exists crm_documents_set_updated_at on crm_documents;
create trigger crm_documents_set_updated_at before update on crm_documents
for each row execute function crm_cloud_sync_set_updated_at();

alter table crm_users enable row level security;
alter table crm_clients enable row level security;
alter table crm_orders enable row level security;
alter table crm_documents enable row level security;

-- Internal test-mode policy for the current frontend-only phase.
-- Replace with authenticated role policies before production launch.
drop policy if exists "crm phase1 read users" on crm_users;
create policy "crm phase1 read users" on crm_users for select using (true);
drop policy if exists "crm phase1 write users" on crm_users;
create policy "crm phase1 write users" on crm_users for insert with check (true);
drop policy if exists "crm phase1 update users" on crm_users;
create policy "crm phase1 update users" on crm_users for update using (true) with check (true);

drop policy if exists "crm phase1 read clients" on crm_clients;
create policy "crm phase1 read clients" on crm_clients for select using (true);
drop policy if exists "crm phase1 write clients" on crm_clients;
create policy "crm phase1 write clients" on crm_clients for insert with check (true);
drop policy if exists "crm phase1 update clients" on crm_clients;
create policy "crm phase1 update clients" on crm_clients for update using (true) with check (true);

drop policy if exists "crm phase1 read orders" on crm_orders;
create policy "crm phase1 read orders" on crm_orders for select using (true);
drop policy if exists "crm phase1 write orders" on crm_orders;
create policy "crm phase1 write orders" on crm_orders for insert with check (true);
drop policy if exists "crm phase1 update orders" on crm_orders;
create policy "crm phase1 update orders" on crm_orders for update using (true) with check (true);

drop policy if exists "crm phase1 read documents" on crm_documents;
create policy "crm phase1 read documents" on crm_documents for select using (true);
drop policy if exists "crm phase1 write documents" on crm_documents;
create policy "crm phase1 write documents" on crm_documents for insert with check (true);
drop policy if exists "crm phase1 update documents" on crm_documents;
create policy "crm phase1 update documents" on crm_documents for update using (true) with check (true);

do $$
begin
  alter publication supabase_realtime add table crm_users;
exception when duplicate_object then null;
end $$;

do $$
begin
  alter publication supabase_realtime add table crm_clients;
exception when duplicate_object then null;
end $$;

do $$
begin
  alter publication supabase_realtime add table crm_orders;
exception when duplicate_object then null;
end $$;

do $$
begin
  alter publication supabase_realtime add table crm_documents;
exception when duplicate_object then null;
end $$;
