-- 한컷영단어 v4.10.0: 학생 그룹 및 학습세트 배정
-- 기존 카드/학습기록을 삭제하거나 변경하지 않습니다. Supabase SQL Editor에서 1회 실행하세요.
create extension if not exists pgcrypto;
create table if not exists public.onecut_groups (
 id uuid primary key default gen_random_uuid(), name text not null check(length(name) between 1 and 80),
 created_by uuid references auth.users(id) on delete set null, created_at timestamptz not null default now()
);
create table if not exists public.onecut_group_members (
 group_id uuid not null references public.onecut_groups(id) on delete cascade,
 user_id uuid not null references auth.users(id) on delete cascade,
 created_at timestamptz not null default now(), primary key(group_id,user_id)
);
create table if not exists public.onecut_assignments (
 id uuid primary key default gen_random_uuid(), title text not null check(length(title) between 1 and 100),
 card_ids jsonb not null check(jsonb_typeof(card_ids)='array' and jsonb_array_length(card_ids) between 1 and 50),
 target_type text not null check(target_type in ('student','group')), target_id uuid not null,
 created_by uuid references auth.users(id) on delete set null, assigned_at timestamptz not null default now(),
 due_at timestamptz, active boolean not null default true
);
create index if not exists onecut_group_members_user_idx on public.onecut_group_members(user_id);
create index if not exists onecut_assignments_target_idx on public.onecut_assignments(target_type,target_id,active);
alter table public.onecut_groups enable row level security;
alter table public.onecut_group_members enable row level security;
alter table public.onecut_assignments enable row level security;
revoke all on public.onecut_groups,public.onecut_group_members,public.onecut_assignments from anon,authenticated;
-- 브라우저 직접 접근은 막고 Python 서버의 service role만 사용합니다.
