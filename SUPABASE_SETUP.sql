-- CMF 排班助手 Supabase 初始化/升级脚本（可重复执行）

create table if not exists public.schedules (
    id uuid primary key default gen_random_uuid(),
    file_name text not null,
    storage_path text not null,
    batch_name text,
    sheet_name text,
    start_date date,
    end_date date,
    task_count integer not null default 0,
    uploaded_by text not null,
    status text not null default 'draft',
    file_hash text not null,
    source_type text not null default 'workbook',
    records jsonb,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

alter table public.schedules
    add column if not exists source_type text not null default 'workbook',
    add column if not exists records jsonb;

create unique index if not exists schedules_storage_path_unique
    on public.schedules (storage_path);

create unique index if not exists schedules_file_hash_unique
    on public.schedules (file_hash);

create index if not exists schedules_status_index
    on public.schedules (status);

create index if not exists schedules_date_index
    on public.schedules (start_date, end_date);

create index if not exists schedules_created_at_index
    on public.schedules (created_at desc);

alter table public.schedules enable row level security;
