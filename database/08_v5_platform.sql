-- 한컷영단어 v5.0 플랫폼 확장
-- 기존 데이터 삭제 없음. 분류/커리큘럼/시험 기록용 테이블만 추가합니다.

create table if not exists public.onecut_card_meta (
  card_id text primary key,
  difficulty text not null default '미분류' check (difficulty in ('미분류','초급','중급','고급')),
  school_level text not null default '미분류' check (school_level in ('미분류','초등','중등','고등','수능')),
  exam_scope text not null default '',
  topic text not null default '',
  tags jsonb not null default '[]'::jsonb,
  updated_by uuid null,
  updated_at timestamptz not null default now()
);

create table if not exists public.onecut_curricula (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  group_id uuid null,
  settings jsonb not null default '{}'::jsonb,
  active boolean not null default true,
  created_by uuid null,
  created_at timestamptz not null default now()
);

create table if not exists public.onecut_exam_attempts (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  mode text not null check (mode in ('meaning','spelling','listening','mixed')),
  total integer not null check (total between 1 and 100),
  correct integer not null check (correct between 0 and 100),
  score integer not null check (score between 0 and 100),
  card_ids jsonb not null default '[]'::jsonb,
  answers jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now()
);
create index if not exists onecut_exam_attempts_user_created_idx on public.onecut_exam_attempts(user_id, created_at desc);

alter table public.onecut_card_meta enable row level security;
alter table public.onecut_curricula enable row level security;
alter table public.onecut_exam_attempts enable row level security;

-- 브라우저가 위 테이블을 직접 읽고 쓰지 않습니다.
-- 서버는 service-role로 접근하고, 관리자/학생 인증을 서버에서 먼저 검증합니다.
revoke all on table public.onecut_card_meta from anon, authenticated;
revoke all on table public.onecut_curricula from anon, authenticated;
revoke all on table public.onecut_exam_attempts from anon, authenticated;
