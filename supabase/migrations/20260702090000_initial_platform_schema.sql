create extension if not exists pgcrypto;

create table if not exists public.organizations (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  slug text not null unique,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.organization_members (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null check (role in ('owner', 'admin', 'operator', 'viewer')),
  created_at timestamptz not null default now(),
  unique (organization_id, user_id)
);

create table if not exists public.conversion_templates (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  name text not null,
  source_company text not null,
  document_type text not null check (document_type in ('sales', 'credit_note', 'mixed')),
  source_columns jsonb not null default '[]'::jsonb,
  dexy_columns jsonb not null default '[]'::jsonb,
  rules jsonb not null default '{}'::jsonb,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.conversion_batches (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  template_id uuid references public.conversion_templates(id) on delete set null,
  source_company text not null,
  document_type text not null check (document_type in ('sales', 'credit_note', 'mixed')),
  status text not null check (status in ('draft', 'processing', 'success', 'failed')),
  source_file_name text,
  output_file_name text,
  currency_mode text check (currency_mode in ('CDF', 'USD', 'EUR', 'MULTI')),
  invoice_count integer not null default 0,
  source_row_count integer not null default 0,
  tax_row_count integer not null default 0,
  total_tax numeric(18, 4) not null default 0,
  error_message text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  completed_at timestamptz
);

create table if not exists public.conversion_files (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  batch_id uuid not null references public.conversion_batches(id) on delete cascade,
  file_role text not null check (file_role in ('source', 'reference', 'output', 'audit')),
  file_name text not null,
  storage_path text,
  mime_type text,
  size_bytes bigint,
  created_at timestamptz not null default now()
);

create table if not exists public.certified_invoices (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  batch_id uuid references public.conversion_batches(id) on delete set null,
  source_company text not null,
  invoice_number text not null,
  normalized_invoice_number text,
  dgi_code text,
  document_type text not null check (document_type in ('FV', 'FA')),
  invoice_date date,
  original_invoice_date date,
  currency text not null default 'CDF',
  exchange_rate numeric(18, 8),
  exchange_rate_date date,
  total_ht numeric(18, 4),
  total_ttc numeric(18, 4),
  total_tax numeric(18, 4),
  created_at timestamptz not null default now(),
  unique (organization_id, source_company, invoice_number, document_type)
);

create table if not exists public.credit_note_links (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  credit_note_invoice_id uuid not null references public.certified_invoices(id) on delete cascade,
  original_invoice_id uuid not null references public.certified_invoices(id) on delete restrict,
  reference_type text not null check (reference_type in ('RAN', 'COR')),
  reference_description text not null check (reference_description in ('ANNULATION', 'CORRECTION')),
  created_at timestamptz not null default now(),
  unique (credit_note_invoice_id)
);

create table if not exists public.conversion_events (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  batch_id uuid references public.conversion_batches(id) on delete cascade,
  level text not null check (level in ('info', 'warning', 'error')),
  message text not null,
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

alter table public.organizations enable row level security;
alter table public.organization_members enable row level security;
alter table public.conversion_templates enable row level security;
alter table public.conversion_batches enable row level security;
alter table public.conversion_files enable row level security;
alter table public.certified_invoices enable row level security;
alter table public.credit_note_links enable row level security;
alter table public.conversion_events enable row level security;

create or replace function public.is_organization_member(target_organization_id uuid)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.organization_members member
    where member.organization_id = target_organization_id
      and member.user_id = auth.uid()
  );
$$;

create policy "members can read organizations"
on public.organizations
for select
using (public.is_organization_member(id));

create policy "members can read organization membership"
on public.organization_members
for select
using (public.is_organization_member(organization_id));

create policy "members can manage conversion templates"
on public.conversion_templates
for all
using (public.is_organization_member(organization_id))
with check (public.is_organization_member(organization_id));

create policy "members can manage conversion batches"
on public.conversion_batches
for all
using (public.is_organization_member(organization_id))
with check (public.is_organization_member(organization_id));

create policy "members can manage conversion files"
on public.conversion_files
for all
using (public.is_organization_member(organization_id))
with check (public.is_organization_member(organization_id));

create policy "members can manage certified invoices"
on public.certified_invoices
for all
using (public.is_organization_member(organization_id))
with check (public.is_organization_member(organization_id));

create policy "members can manage credit note links"
on public.credit_note_links
for all
using (public.is_organization_member(organization_id))
with check (public.is_organization_member(organization_id));

create policy "members can manage conversion events"
on public.conversion_events
for all
using (public.is_organization_member(organization_id))
with check (public.is_organization_member(organization_id));
