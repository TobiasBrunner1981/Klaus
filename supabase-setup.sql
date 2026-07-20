-- Tandem: run this once in the Supabase SQL Editor (new project).
-- Three tables + Realtime. The photos bucket is created in the Storage UI.

create table if not exists public.tasks (
  id uuid primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);
create table if not exists public.nudges (
  id uuid primary key,
  data jsonb not null,
  created_at timestamptz not null default now()
);
create table if not exists public.household (
  id text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

-- Row level security on, with open policies for the anon key
-- (the household is private because the anon key is private to you two).
alter table public.tasks enable row level security;
alter table public.nudges enable row level security;
alter table public.household enable row level security;

create policy "tandem tasks all" on public.tasks
  for all using (true) with check (true);
create policy "tandem nudges all" on public.nudges
  for all using (true) with check (true);
create policy "tandem household all" on public.household
  for all using (true) with check (true);

-- Live sync between the two phones
alter publication supabase_realtime add table public.tasks;
alter publication supabase_realtime add table public.nudges;
alter publication supabase_realtime add table public.household;

-- After running this, in the dashboard: Storage -> New bucket ->
-- name it  photos  and tick "Public bucket". Then add these policies:
create policy "tandem photos read" on storage.objects
  for select using (bucket_id = 'photos');
create policy "tandem photos write" on storage.objects
  for insert with check (bucket_id = 'photos');
create policy "tandem photos update" on storage.objects
  for update using (bucket_id = 'photos');
