-- Alice Office · private per-user workspaces and day-to-day Excel continuity
-- Applies to the user's Supabase project; never grant anon data access.
create extension if not exists pgcrypto;

create table if not exists public.alice_office_drafts (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  work_date date not null,
  workflow text not null check (workflow in ('prep', 'daily')),
  payload jsonb not null default '{}'::jsonb,
  revision integer not null default 1 check (revision > 0),
  updated_at timestamptz not null default now(),
  created_at timestamptz not null default now(),
  unique(owner_id, work_date, workflow)
);

create table if not exists public.alice_office_files (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  category text not null check (category in ('template', 'source', 'output')),
  role text not null,
  filename text not null,
  object_path text not null,
  work_date date,
  workflow text check (workflow in ('prep', 'daily') or workflow is null),
  byte_size bigint check (byte_size >= 0),
  created_at timestamptz not null default now(),
  unique(owner_id, object_path)
);
create index if not exists alice_office_files_owner_role_created_idx
  on public.alice_office_files(owner_id,category,role,created_at desc);
create index if not exists alice_office_files_owner_work_idx
  on public.alice_office_files(owner_id,work_date,workflow,created_at desc);

create table if not exists public.alice_office_vehicles (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  work_date date not null,
  flow text not null check (flow in ('STP','PTR')),
  booking text not null,
  car_no integer not null check (car_no > 0),
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  unique(owner_id,work_date,flow,booking,car_no)
);

alter table public.alice_office_drafts enable row level security;
alter table public.alice_office_files enable row level security;
alter table public.alice_office_vehicles enable row level security;

revoke all on public.alice_office_drafts from anon;
revoke all on public.alice_office_files from anon;
revoke all on public.alice_office_vehicles from anon;
grant select,insert,update,delete on public.alice_office_drafts,public.alice_office_files,public.alice_office_vehicles to authenticated;

drop policy if exists "alice_drafts_select_own" on public.alice_office_drafts;
drop policy if exists "alice_drafts_insert_own" on public.alice_office_drafts;
drop policy if exists "alice_drafts_update_own" on public.alice_office_drafts;
drop policy if exists "alice_drafts_delete_own" on public.alice_office_drafts;
create policy "alice_drafts_select_own" on public.alice_office_drafts
 for select to authenticated using ((select auth.uid()) = owner_id);
create policy "alice_drafts_insert_own" on public.alice_office_drafts
 for insert to authenticated with check ((select auth.uid()) = owner_id);
create policy "alice_drafts_update_own" on public.alice_office_drafts
 for update to authenticated using ((select auth.uid()) = owner_id) with check ((select auth.uid()) = owner_id);
create policy "alice_drafts_delete_own" on public.alice_office_drafts
 for delete to authenticated using ((select auth.uid()) = owner_id);

drop policy if exists "alice_files_select_own" on public.alice_office_files;
drop policy if exists "alice_files_insert_own" on public.alice_office_files;
drop policy if exists "alice_files_update_own" on public.alice_office_files;
drop policy if exists "alice_files_delete_own" on public.alice_office_files;
create policy "alice_files_select_own" on public.alice_office_files
 for select to authenticated using ((select auth.uid()) = owner_id);
create policy "alice_files_insert_own" on public.alice_office_files
 for insert to authenticated with check ((select auth.uid()) = owner_id);
create policy "alice_files_update_own" on public.alice_office_files
 for update to authenticated using ((select auth.uid()) = owner_id) with check ((select auth.uid()) = owner_id);
create policy "alice_files_delete_own" on public.alice_office_files
 for delete to authenticated using ((select auth.uid()) = owner_id);

drop policy if exists "alice_vehicles_select_own" on public.alice_office_vehicles;
drop policy if exists "alice_vehicles_insert_own" on public.alice_office_vehicles;
drop policy if exists "alice_vehicles_update_own" on public.alice_office_vehicles;
drop policy if exists "alice_vehicles_delete_own" on public.alice_office_vehicles;
create policy "alice_vehicles_select_own" on public.alice_office_vehicles
 for select to authenticated using ((select auth.uid()) = owner_id);
create policy "alice_vehicles_insert_own" on public.alice_office_vehicles
 for insert to authenticated with check ((select auth.uid()) = owner_id);
create policy "alice_vehicles_update_own" on public.alice_office_vehicles
 for update to authenticated using ((select auth.uid()) = owner_id) with check ((select auth.uid()) = owner_id);
create policy "alice_vehicles_delete_own" on public.alice_office_vehicles
 for delete to authenticated using ((select auth.uid()) = owner_id);

insert into storage.buckets (id, name, public, file_size_limit)
values ('alice-office-private', 'alice-office-private', false, 31457280)
on conflict (id) do update set public=false, file_size_limit=31457280;

drop policy if exists "alice_storage_select_own" on storage.objects;
drop policy if exists "alice_storage_insert_own" on storage.objects;
drop policy if exists "alice_storage_update_own" on storage.objects;
drop policy if exists "alice_storage_delete_own" on storage.objects;
create policy "alice_storage_select_own" on storage.objects
 for select to authenticated
 using (bucket_id='alice-office-private' and (storage.foldername(name))[1] = (select auth.uid())::text);
create policy "alice_storage_insert_own" on storage.objects
 for insert to authenticated
 with check (bucket_id='alice-office-private' and (storage.foldername(name))[1] = (select auth.uid())::text);
create policy "alice_storage_update_own" on storage.objects
 for update to authenticated
 using (bucket_id='alice-office-private' and (storage.foldername(name))[1] = (select auth.uid())::text)
 with check (bucket_id='alice-office-private' and (storage.foldername(name))[1] = (select auth.uid())::text);
create policy "alice_storage_delete_own" on storage.objects
 for delete to authenticated
 using (bucket_id='alice-office-private' and (storage.foldername(name))[1] = (select auth.uid())::text);

do $$
declare tab text;
begin
  foreach tab in array ARRAY['alice_office_drafts','alice_office_files','alice_office_vehicles'] loop
    if not exists (
      select 1 from pg_publication_tables
      where pubname='supabase_realtime' and schemaname='public' and tablename=tab
    ) then
      execute format('alter publication supabase_realtime add table public.%I',tab);
    end if;
  end loop;
end $$;
