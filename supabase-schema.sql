create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  email text unique,
  full_name text not null default '',
  role text not null default 'rep' check (role in ('admin', 'rep')),
  active boolean not null default true,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.app_settings (
  id uuid primary key default gen_random_uuid(),
  singleton_key text not null unique default 'default',
  assumptions jsonb not null default '{}'::jsonb,
  routing_rules jsonb not null default '{}'::jsonb,
  lead_sources jsonb not null default '[]'::jsonb,
  statuses jsonb not null default '[]'::jsonb,
  products jsonb not null default '[]'::jsonb,
  carriers jsonb not null default '[]'::jsonb,
  updated_by uuid references public.profiles (id),
  updated_at timestamptz not null default timezone('utc', now())
);

alter table public.app_settings add column if not exists routing_rules jsonb not null default '{}'::jsonb;

create table if not exists public.opportunities (
  id uuid primary key default gen_random_uuid(),
  lead_number text not null unique,
  assigned_user_id uuid not null references public.profiles (id),
  assigned_rep_name text not null,
  date_received date not null,
  lead_source text not null default '',
  business_name text not null default '',
  target_niche text not null default '',
  product_focus text not null default '',
  contact_name text not null default '',
  contact_email text not null default '',
  contact_phone text not null default '',
  carrier text not null default '',
  incumbent_carrier text not null default '',
  policy_type text not null default 'New',
  policy_term_months integer not null default 12,
  renewal_status text not null default 'Not Started',
  effective_date date,
  expiration_date date,
  lead_cost numeric(12,2) not null default 0,
  premium_quoted numeric(12,2) not null default 0,
  premium_bound numeric(12,2) not null default 0,
  status text not null default 'New Lead',
  first_attempt_date date,
  last_activity_date date,
  next_follow_up_date date,
  next_task text not null default '',
  task_priority text not null default 'Medium',
  notes text not null default '',
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

alter table public.opportunities add column if not exists contact_name text not null default '';
alter table public.opportunities add column if not exists contact_email text not null default '';
alter table public.opportunities add column if not exists contact_phone text not null default '';
alter table public.opportunities add column if not exists incumbent_carrier text not null default '';
alter table public.opportunities add column if not exists effective_date date;
alter table public.opportunities add column if not exists expiration_date date;
alter table public.opportunities add column if not exists policy_term_months integer not null default 12;
alter table public.opportunities add column if not exists renewal_status text not null default 'Not Started';
alter table public.opportunities add column if not exists next_task text not null default '';
alter table public.opportunities add column if not exists task_priority text not null default 'Medium';

create table if not exists public.coaching_notes (
  id uuid primary key default gen_random_uuid(),
  rep_user_id uuid not null references public.profiles (id) on delete cascade,
  week_start date not null,
  biggest_gap text not null default '',
  behavior_to_improve text not null default '',
  action_commitment text not null default '',
  next_review_notes text not null default '',
  created_by uuid references public.profiles (id),
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  unique (rep_user_id, week_start)
);

create table if not exists public.opportunity_activity (
  id uuid primary key default gen_random_uuid(),
  opportunity_id uuid not null references public.opportunities (id) on delete cascade,
  actor_id uuid references public.profiles (id) on delete set null,
  actor_name text not null default '',
  title text not null default '',
  detail text not null default '',
  created_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.opportunity_attachments (
  id uuid primary key default gen_random_uuid(),
  opportunity_id uuid not null references public.opportunities (id) on delete cascade,
  file_name text not null default '',
  file_path text not null unique,
  file_type text not null default 'Other',
  file_size bigint not null default 0,
  mime_type text not null default 'application/octet-stream',
  created_by uuid references public.profiles (id) on delete set null,
  created_by_name text not null default '',
  created_at timestamptz not null default timezone('utc', now())
);

insert into storage.buckets (id, name, public)
values ('opportunity-files', 'opportunity-files', false)
on conflict (id) do nothing;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, full_name, role)
  values (
    new.id,
    new.email,
    coalesce(new.raw_user_meta_data ->> 'full_name', split_part(coalesce(new.email, ''), '@', 1)),
    coalesce(new.raw_user_meta_data ->> 'role', 'rep')
  )
  on conflict (id) do update
    set email = excluded.email,
        full_name = excluded.full_name,
        role = excluded.role,
        updated_at = timezone('utc', now());

  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute procedure public.handle_new_user();

insert into public.app_settings (
  singleton_key,
  assumptions,
  routing_rules,
  lead_sources,
  statuses,
  products,
  carriers
)
values (
  'default',
  '{
    "averageCommissionPct": 0.12,
    "sameDayWorkedTargetPct": 0.9,
    "contactRateTargetPct": 0.5,
    "quoteRateTargetPct": 0.2,
    "quoteToBindTargetPct": 0.25,
    "leadToBindTargetPct": 0.05,
    "crmComplianceTargetPct": 0.95,
    "followUpDueWindowDays": 3,
    "freshLeadWindowDays": 3
  }'::jsonb,
  '{
    "autoAssignEnabled": false,
    "mode": "round_robin",
    "roundRobinCursor": 0,
    "sourceRules": []
  }'::jsonb,
  '["Purchased Leads","Warm Transfer","Referral","Website / Organic","Partner / Network","Recycled Lead","Self-Generated"]'::jsonb,
  '["New Lead","Attempted","Contacted","Qualified","Quoted","Pending Decision","Bound","Lost","Nurture / Recycle"]'::jsonb,
  '["GL / BOP","Workers Comp","Package / Multi-Line"]'::jsonb,
  '[
    {"name":"AmTrust","newPct":0.16,"renewalPct":0.1,"notes":"Typical"},
    {"name":"Berxi","newPct":0.13,"renewalPct":0.13,"notes":"Typical"},
    {"name":"Blitz","newPct":0.125,"renewalPct":0.125,"notes":"Typical"},
    {"name":"Chubb","newPct":0.14,"renewalPct":0.12,"notes":"Typical"},
    {"name":"Coterie","newPct":0.12,"renewalPct":0.1,"notes":"Typical"},
    {"name":"First","newPct":0.16,"renewalPct":0.16,"notes":"Typical"},
    {"name":"Hiscox","newPct":0.14,"renewalPct":0.12,"notes":"Typical"},
    {"name":"Pathpoint","newPct":0.11,"renewalPct":0.11,"notes":"Typical"},
    {"name":"Simply Business","newPct":0.12,"renewalPct":0.12,"notes":"Typical"},
    {"name":"THREE","newPct":0.12,"renewalPct":0.12,"notes":"Typical"}
  ]'::jsonb
)
on conflict (singleton_key) do nothing;

alter table public.profiles enable row level security;
alter table public.app_settings enable row level security;
alter table public.opportunities enable row level security;
alter table public.coaching_notes enable row level security;
alter table public.opportunity_activity enable row level security;
alter table public.opportunity_attachments enable row level security;

create or replace function public.is_admin(user_id uuid)
returns boolean
language sql
stable
as $$
  select exists (
    select 1
    from public.profiles
    where id = user_id
      and role = 'admin'
      and active = true
  );
$$;

drop policy if exists "profiles self or admin select" on public.profiles;
create policy "profiles self or admin select"
on public.profiles
for select
to authenticated
using (id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "profiles self update" on public.profiles;
create policy "profiles self update"
on public.profiles
for update
to authenticated
using (id = auth.uid())
with check (id = auth.uid());

drop policy if exists "profiles admin manage" on public.profiles;
create policy "profiles admin manage"
on public.profiles
for update
to authenticated
using (public.is_admin(auth.uid()))
with check (public.is_admin(auth.uid()));

drop policy if exists "settings read authenticated" on public.app_settings;
create policy "settings read authenticated"
on public.app_settings
for select
to authenticated
using (true);

drop policy if exists "settings admin update" on public.app_settings;
create policy "settings admin update"
on public.app_settings
for all
to authenticated
using (public.is_admin(auth.uid()))
with check (public.is_admin(auth.uid()));

drop policy if exists "opportunities read own or admin" on public.opportunities;
create policy "opportunities read own or admin"
on public.opportunities
for select
to authenticated
using (assigned_user_id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "opportunities insert own or admin" on public.opportunities;
create policy "opportunities insert own or admin"
on public.opportunities
for insert
to authenticated
with check (assigned_user_id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "opportunities update own or admin" on public.opportunities;
create policy "opportunities update own or admin"
on public.opportunities
for update
to authenticated
using (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
with check (assigned_user_id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "opportunities delete own or admin" on public.opportunities;
create policy "opportunities delete own or admin"
on public.opportunities
for delete
to authenticated
using (assigned_user_id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "coaching read own or admin" on public.coaching_notes;
create policy "coaching read own or admin"
on public.coaching_notes
for select
to authenticated
using (rep_user_id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "coaching admin manage" on public.coaching_notes;
create policy "coaching admin manage"
on public.coaching_notes
for all
to authenticated
using (public.is_admin(auth.uid()))
with check (public.is_admin(auth.uid()));

drop policy if exists "activity read own or admin" on public.opportunity_activity;
create policy "activity read own or admin"
on public.opportunity_activity
for select
to authenticated
using (
  exists (
    select 1
    from public.opportunities
    where id = opportunity_id
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "activity insert own or admin" on public.opportunity_activity;
create policy "activity insert own or admin"
on public.opportunity_activity
for insert
to authenticated
with check (
  exists (
    select 1
    from public.opportunities
    where id = opportunity_id
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "attachments read own or admin" on public.opportunity_attachments;
create policy "attachments read own or admin"
on public.opportunity_attachments
for select
to authenticated
using (
  exists (
    select 1
    from public.opportunities
    where id = opportunity_id
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "attachments insert own or admin" on public.opportunity_attachments;
create policy "attachments insert own or admin"
on public.opportunity_attachments
for insert
to authenticated
with check (
  exists (
    select 1
    from public.opportunities
    where id = opportunity_id
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "attachments delete own or admin" on public.opportunity_attachments;
create policy "attachments delete own or admin"
on public.opportunity_attachments
for delete
to authenticated
using (
  exists (
    select 1
    from public.opportunities
    where id = opportunity_id
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "storage opportunity files read own or admin" on storage.objects;
create policy "storage opportunity files read own or admin"
on storage.objects
for select
to authenticated
using (
  bucket_id = 'opportunity-files'
  and exists (
    select 1
    from public.opportunities
    where id::text = (storage.foldername(name))[1]
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "storage opportunity files upload own or admin" on storage.objects;
create policy "storage opportunity files upload own or admin"
on storage.objects
for insert
to authenticated
with check (
  bucket_id = 'opportunity-files'
  and exists (
    select 1
    from public.opportunities
    where id::text = (storage.foldername(name))[1]
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);

drop policy if exists "storage opportunity files delete own or admin" on storage.objects;
create policy "storage opportunity files delete own or admin"
on storage.objects
for delete
to authenticated
using (
  bucket_id = 'opportunity-files'
  and exists (
    select 1
    from public.opportunities
    where id::text = (storage.foldername(name))[1]
      and (assigned_user_id = auth.uid() or public.is_admin(auth.uid()))
  )
);
