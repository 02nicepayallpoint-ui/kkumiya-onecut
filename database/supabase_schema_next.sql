-- 한컷영단어 클라우드 DB 초안. 아직 실제 Supabase에 적용하지 않음.
-- 공개 읽기와 관리자 쓰기 권한은 서비스 서버에서 권한 검증 후 구현해야 합니다.
create extension if not exists pgcrypto;
create table if not exists public.vocab_cards (
 id text primary key,
 word text not null,
 part_of_speech text not null,
 ipa text not null,
 hangul_pronunciation text,
 meaning text not null,
 stress_start integer not null,
 stress_end integer not null,
 mnemonic text,
 scene_description text,
 scene_image_path text,
 example_sentence text,
 example_translation text,
 status text not null default 'draft' check(status in ('draft','review','approved','published','rejected')),
 quality_score integer not null default 0 check(quality_score between 0 and 100),
 pronunciation_checked boolean not null default false,
 image_checked boolean not null default false,
 created_at timestamptz not null default now(),
 updated_at timestamptz not null default now(),
 unique(word,part_of_speech)
);
create table if not exists public.vocab_revisions (
 id uuid primary key default gen_random_uuid(),
 card_id text not null references public.vocab_cards(id) on delete cascade,
 revision_type text not null,
 rejection_reason text,
 snapshot jsonb not null,
 created_at timestamptz not null default now()
);
create table if not exists public.user_vocab_progress (
 user_id uuid not null,
 card_id text not null references public.vocab_cards(id),
 rating text,
 next_review_at timestamptz,
 favorite boolean not null default false,
 attempt_count integer not null default 0,
 updated_at timestamptz not null default now(),
 primary key(user_id,card_id)
);
alter table public.vocab_cards enable row level security;
alter table public.vocab_revisions enable row level security;
alter table public.user_vocab_progress enable row level security;
-- 출판된 콘텐츠만 공개적으로 읽을 수 있도록 최소 정책 설정
create policy "published cards read" on public.vocab_cards for select to anon, authenticated using (status='published');
create policy "own learning progress read" on public.user_vocab_progress for select to authenticated using (auth.uid()=user_id);
create policy "own learning progress insert" on public.user_vocab_progress for insert to authenticated with check (auth.uid()=user_id);
create policy "own learning progress update" on public.user_vocab_progress for update to authenticated using (auth.uid()=user_id) with check (auth.uid()=user_id);
-- 관리자 쓰기는 공개 API 정책 없이 서버 측 서비스 계정 및 추가 권한 검증으로 처리할 것.
