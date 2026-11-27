-- Supabase/PostgreSQL schema for One-Cut Vocabulary, draft v1
-- Execute in a NEW Supabase project. Admin assignment is a manual privileged operation.
create extension if not exists pg_trgm;

create table public.admin_members (
 user_id uuid primary key references auth.users(id) on delete cascade,
 created_at timestamptz not null default now()
);
create or replace function public.is_vocabulary_admin()
returns boolean language sql stable security definer set search_path = '' as $$
 select exists(select 1 from public.admin_members where user_id = (select auth.uid()));
$$;
revoke all on function public.is_vocabulary_admin() from public;
grant execute on function public.is_vocabulary_admin() to authenticated;

create table public.words (
 id bigint generated always as identity primary key,
 lemma text not null unique,
 normalized_lemma text generated always as (lower(trim(lemma))) stored,
 level int not null default 1 check (level between 1 and 5),
 created_at timestamptz not null default now()
);
create index words_search_trgm on public.words using gin (normalized_lemma gin_trgm_ops);

create table public.word_senses (
 id bigint generated always as identity primary key,
 word_id bigint not null references public.words(id) on delete cascade,
 part_of_speech text not null,
 meaning_ko text not null,
 ipa text not null,
 hangul_hint text,
 stress_start int,
 stress_end int,
 status text not null default 'draft' check(status in ('draft','review','approved','published','rejected')),
 check(stress_start is null or (stress_start >= 0 and stress_end > stress_start)),
 unique(word_id, part_of_speech, ipa)
);

create table public.mnemonics (
 id bigint generated always as identity primary key,
 sense_id bigint not null references public.word_senses(id) on delete cascade,
 mnemonic_ko text not null,
 scene_description text,
 image_path text,
 primary_card boolean not null default false,
 status text not null default 'draft' check(status in ('draft','review','approved','published','rejected')),
 created_at timestamptz not null default now()
);
create unique index mnemonics_one_primary_published
on public.mnemonics(sense_id) where primary_card and status='published';

create table public.examples (
 id bigint generated always as identity primary key,
 sense_id bigint not null references public.word_senses(id) on delete cascade,
 sentence_en text not null,
 translation_ko text not null,
 audio_path text,
 status text not null default 'draft' check(status in ('draft','review','approved','published','rejected'))
);

create table public.word_sets (
 id bigint generated always as identity primary key,
 slug text not null unique,
 title text not null,
 level int not null check (level between 1 and 5),
 published boolean not null default false
);
create table public.word_set_items (
 word_set_id bigint not null references public.word_sets(id) on delete cascade,
 sense_id bigint not null references public.word_senses(id) on delete cascade,
 day_no int not null,
 sort_order int not null,
 primary key(word_set_id,sense_id)
);

create table public.learning_progress (
 user_id uuid not null references auth.users(id) on delete cascade,
 sense_id bigint not null references public.word_senses(id) on delete cascade,
 outcome text not null check(outcome in ('again','hard','good','easy')),
 attempts int not null default 1,
 next_review_at timestamptz,
 updated_at timestamptz not null default now(),
 primary key(user_id,sense_id)
);
create table public.search_requests (
 id bigint generated always as identity primary key,
 query_normalized text not null,
 user_id uuid references auth.users(id) on delete set null,
 created_at timestamptz not null default now()
);

-- NOTE: No direct anonymous read of words table. Read through a published-content view.
create view public.published_cards with (security_invoker = true) as
 select w.id word_id, w.lemma, w.level, s.id sense_id, s.part_of_speech,
 s.meaning_ko, s.ipa, s.hangul_hint, s.stress_start, s.stress_end,
 m.mnemonic_ko, m.image_path
 from public.words w
 join public.word_senses s on s.word_id=w.id
 join public.mnemonics m on m.sense_id=s.id
 where s.status='published' and m.status='published' and m.primary_card;

alter table public.admin_members enable row level security;
alter table public.words enable row level security;
alter table public.word_senses enable row level security;
alter table public.mnemonics enable row level security;
alter table public.examples enable row level security;
alter table public.word_sets enable row level security;
alter table public.word_set_items enable row level security;
alter table public.learning_progress enable row level security;
alter table public.search_requests enable row level security;

-- Published read (drafts remain hidden). word metadata is readable, but no unpublished sense/image data.
create policy "word_read" on public.words for select to anon,authenticated using (true);
create policy "sense_read" on public.word_senses for select to anon,authenticated
 using (status='published' or public.is_vocabulary_admin());
create policy "mnemonic_read" on public.mnemonics for select to anon,authenticated
 using (status='published' or public.is_vocabulary_admin());
create policy "example_read" on public.examples for select to anon,authenticated
 using (status='published' or public.is_vocabulary_admin());
create policy "sets_read" on public.word_sets for select to anon,authenticated
 using (published or public.is_vocabulary_admin());
create policy "set_items_read" on public.word_set_items for select to anon,authenticated
 using (exists(select 1 from public.word_sets ws where ws.id=word_set_id and ws.published) or public.is_vocabulary_admin());

-- Admin-only content writes. Register admin by SQL service role (not in student browser).
create policy "words_admin_all" on public.words for all to authenticated
 using (public.is_vocabulary_admin()) with check (public.is_vocabulary_admin());
create policy "senses_admin_all" on public.word_senses for all to authenticated
 using (public.is_vocabulary_admin()) with check (public.is_vocabulary_admin());
create policy "mnemonics_admin_all" on public.mnemonics for all to authenticated
 using (public.is_vocabulary_admin()) with check (public.is_vocabulary_admin());
create policy "examples_admin_all" on public.examples for all to authenticated
 using (public.is_vocabulary_admin()) with check (public.is_vocabulary_admin());
create policy "sets_admin_all" on public.word_sets for all to authenticated
 using (public.is_vocabulary_admin()) with check (public.is_vocabulary_admin());
create policy "set_items_admin_all" on public.word_set_items for all to authenticated
 using (public.is_vocabulary_admin()) with check (public.is_vocabulary_admin());

create policy "progress_owner_read" on public.learning_progress for select to authenticated
 using (user_id=auth.uid());
create policy "progress_owner_insert" on public.learning_progress for insert to authenticated
 with check (user_id=auth.uid());
create policy "progress_owner_update" on public.learning_progress for update to authenticated
 using (user_id=auth.uid()) with check (user_id=auth.uid());

-- Unmatched searches: logged by a rate-limited server function in production, NOT open browser INSERT.
-- No search_requests insert policy is intentionally defined.
-- Storage: create public-read bucket for ONLY approved image assets. Keep drafts private.
