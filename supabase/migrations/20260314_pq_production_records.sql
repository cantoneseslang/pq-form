-- pq-form production records history
create table if not exists public.pq_production_records (
  id uuid primary key default gen_random_uuid(),
  record_date date not null,
  page_type text not null check (page_type in ('molding', 'auto')),
  product_types jsonb default '{}'::jsonb,
  machines jsonb default '{}'::jsonb,
  main_data jsonb not null,
  material_data jsonb default '{}'::jsonb,
  sheet_row int,
  sheet_name text default 'pq-form',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  corrected_at timestamptz,
  correction_note text
);

create index if not exists pq_production_records_record_date_idx
  on public.pq_production_records (record_date desc);

create index if not exists pq_production_records_created_at_idx
  on public.pq_production_records (created_at desc);

create index if not exists pq_production_records_page_type_idx
  on public.pq_production_records (page_type);

alter table public.pq_production_records enable row level security;
