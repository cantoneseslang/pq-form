-- soft delete for pq-form production records (keep row for restore)
alter table public.pq_production_records
  add column if not exists deleted_at timestamptz;

create index if not exists pq_production_records_deleted_at_idx
  on public.pq_production_records (deleted_at);
