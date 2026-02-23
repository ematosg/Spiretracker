-- Spire Online Schema (Supabase / Postgres)
-- Run in Supabase SQL Editor.

create extension if not exists pgcrypto;

-- --------------------------------------------
-- Profiles (1:1 with auth.users)
-- --------------------------------------------
create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  username text unique not null,
  account_type text not null check (account_type in ('gm','player')),
  created_at timestamptz not null default now()
);

-- Keep profile row in sync when new auth user is created.
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, username, account_type)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'username', split_part(new.email, '@', 1), 'user_' || substr(new.id::text, 1, 8)),
    coalesce(new.raw_user_meta_data->>'account_type', 'player')
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute function public.handle_new_user();

-- --------------------------------------------
-- Campaigns + Membership + Invite Codes
-- --------------------------------------------
create table if not exists public.campaigns (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  owner_user_id uuid not null references auth.users(id) on delete restrict,
  data jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.campaign_members (
  campaign_id uuid not null references public.campaigns(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null check (role in ('gm','player')),
  joined_at timestamptz not null default now(),
  primary key (campaign_id, user_id)
);

create table if not exists public.invite_codes (
  id uuid primary key default gen_random_uuid(),
  campaign_id uuid not null references public.campaigns(id) on delete cascade,
  code text not null unique,
  role_to_grant text not null default 'player' check (role_to_grant in ('player','gm')),
  created_by_user_id uuid not null references auth.users(id) on delete cascade,
  max_uses int not null default 1 check (max_uses > 0),
  used_count int not null default 0 check (used_count >= 0),
  expires_at timestamptz,
  revoked boolean not null default false,
  created_at timestamptz not null default now()
);

create index if not exists idx_campaign_members_user on public.campaign_members(user_id);
create index if not exists idx_invite_codes_campaign on public.invite_codes(campaign_id);

create or replace function public.touch_campaign_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_campaign_touch on public.campaigns;
create trigger trg_campaign_touch
before update on public.campaigns
for each row execute function public.touch_campaign_updated_at();

-- --------------------------------------------
-- Helper functions / RPCs
-- --------------------------------------------

create or replace function public.user_campaign_role(c_id uuid)
returns text
language sql
stable
security definer
set search_path = public
as $$
  select cm.role
  from public.campaign_members cm
  where cm.campaign_id = c_id
    and cm.user_id = auth.uid()
  limit 1;
$$;

create or replace function public.create_campaign(p_name text)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_campaign_id uuid;
begin
  if auth.uid() is null then
    raise exception 'Not authenticated';
  end if;

  insert into public.campaigns (name, owner_user_id, data)
  values (coalesce(nullif(trim(p_name), ''), 'New Campaign'), auth.uid(), '{}'::jsonb)
  returning id into v_campaign_id;

  insert into public.campaign_members (campaign_id, user_id, role)
  values (v_campaign_id, auth.uid(), 'gm')
  on conflict do nothing;

  return v_campaign_id;
end;
$$;

create or replace function public.generate_invite_code(
  p_campaign_id uuid,
  p_role_to_grant text default 'player',
  p_max_uses int default 1,
  p_expires_minutes int default 1440
)
returns text
language plpgsql
security definer
set search_path = public
as $$
declare
  v_code text;
  v_role text := coalesce(p_role_to_grant, 'player');
  v_is_gm boolean;
begin
  if auth.uid() is null then
    raise exception 'Not authenticated';
  end if;

  select exists (
    select 1 from public.campaign_members
    where campaign_id = p_campaign_id
      and user_id = auth.uid()
      and role = 'gm'
  ) into v_is_gm;

  if not v_is_gm then
    raise exception 'Only GM can generate invite codes';
  end if;

  if v_role not in ('player','gm') then
    raise exception 'Invalid role_to_grant';
  end if;

  v_code := upper(encode(gen_random_bytes(5), 'hex')); -- 10-char code

  insert into public.invite_codes (
    campaign_id,
    code,
    role_to_grant,
    created_by_user_id,
    max_uses,
    expires_at
  ) values (
    p_campaign_id,
    v_code,
    v_role,
    auth.uid(),
    greatest(p_max_uses, 1),
    now() + make_interval(mins => greatest(p_expires_minutes, 1))
  );

  return v_code;
end;
$$;

create or replace function public.join_campaign_with_code(p_code text)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_invite public.invite_codes%rowtype;
  v_campaign_id uuid;
begin
  if auth.uid() is null then
    raise exception 'Not authenticated';
  end if;

  select * into v_invite
  from public.invite_codes
  where code = upper(trim(p_code))
  limit 1;

  if not found then
    raise exception 'Invalid invite code';
  end if;

  if v_invite.revoked then
    raise exception 'Invite code revoked';
  end if;

  if v_invite.expires_at is not null and now() > v_invite.expires_at then
    raise exception 'Invite code expired';
  end if;

  if v_invite.used_count >= v_invite.max_uses then
    raise exception 'Invite code has no remaining uses';
  end if;

  insert into public.campaign_members (campaign_id, user_id, role)
  values (v_invite.campaign_id, auth.uid(), v_invite.role_to_grant)
  on conflict (campaign_id, user_id)
  do update set role = excluded.role;

  update public.invite_codes
  set used_count = used_count + 1
  where id = v_invite.id;

  v_campaign_id := v_invite.campaign_id;
  return v_campaign_id;
end;
$$;

-- --------------------------------------------
-- RLS
-- --------------------------------------------
alter table public.profiles enable row level security;
alter table public.campaigns enable row level security;
alter table public.campaign_members enable row level security;
alter table public.invite_codes enable row level security;

-- Profiles: users can read all profiles, edit own profile only.
drop policy if exists profiles_select_all on public.profiles;
create policy profiles_select_all on public.profiles
for select using (true);

drop policy if exists profiles_update_own on public.profiles;
create policy profiles_update_own on public.profiles
for update using (auth.uid() = id)
with check (auth.uid() = id);

-- Campaigns: members can read. GM can insert/update/delete.
drop policy if exists campaigns_select_members on public.campaigns;
create policy campaigns_select_members on public.campaigns
for select
using (
  exists (
    select 1 from public.campaign_members cm
    where cm.campaign_id = campaigns.id
      and cm.user_id = auth.uid()
  )
);

drop policy if exists campaigns_insert_gm on public.campaigns;
create policy campaigns_insert_gm on public.campaigns
for insert
with check (
  owner_user_id = auth.uid()
);

drop policy if exists campaigns_update_gm on public.campaigns;
create policy campaigns_update_gm on public.campaigns
for update
using (
  exists (
    select 1 from public.campaign_members cm
    where cm.campaign_id = campaigns.id
      and cm.user_id = auth.uid()
      and cm.role = 'gm'
  )
)
with check (
  exists (
    select 1 from public.campaign_members cm
    where cm.campaign_id = campaigns.id
      and cm.user_id = auth.uid()
      and cm.role = 'gm'
  )
);

drop policy if exists campaigns_delete_owner on public.campaigns;
create policy campaigns_delete_owner on public.campaigns
for delete
using (owner_user_id = auth.uid());

-- Memberships: members can view memberships in campaigns they belong to.
-- Only GM can insert/update/delete membership rows in their campaigns.
drop policy if exists members_select_campaign_members on public.campaign_members;
create policy members_select_campaign_members on public.campaign_members
for select
using (
  exists (
    select 1 from public.campaign_members me
    where me.campaign_id = campaign_members.campaign_id
      and me.user_id = auth.uid()
  )
);

drop policy if exists members_manage_gm on public.campaign_members;
create policy members_manage_gm on public.campaign_members
for all
using (
  exists (
    select 1 from public.campaign_members gm
    where gm.campaign_id = campaign_members.campaign_id
      and gm.user_id = auth.uid()
      and gm.role = 'gm'
  )
)
with check (
  exists (
    select 1 from public.campaign_members gm
    where gm.campaign_id = campaign_members.campaign_id
      and gm.user_id = auth.uid()
      and gm.role = 'gm'
  )
);

-- Invite codes: members can read active invites in their campaign.
-- Only GM can create/update/delete.
drop policy if exists invites_select_members on public.invite_codes;
create policy invites_select_members on public.invite_codes
for select
using (
  exists (
    select 1 from public.campaign_members cm
    where cm.campaign_id = invite_codes.campaign_id
      and cm.user_id = auth.uid()
  )
);

drop policy if exists invites_manage_gm on public.invite_codes;
create policy invites_manage_gm on public.invite_codes
for all
using (
  exists (
    select 1 from public.campaign_members cm
    where cm.campaign_id = invite_codes.campaign_id
      and cm.user_id = auth.uid()
      and cm.role = 'gm'
  )
)
with check (
  exists (
    select 1 from public.campaign_members cm
    where cm.campaign_id = invite_codes.campaign_id
      and cm.user_id = auth.uid()
      and cm.role = 'gm'
  )
);

-- Grant execute for RPCs to authenticated users.
grant execute on function public.create_campaign(text) to authenticated;
grant execute on function public.generate_invite_code(uuid, text, int, int) to authenticated;
grant execute on function public.join_campaign_with_code(text) to authenticated;
grant execute on function public.user_campaign_role(uuid) to authenticated;
