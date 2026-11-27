-- 한컷영단어 v3.4 기존 onecut-vocab Supabase 프로젝트용 추가 SQL
-- 이미 만든 app_admins, vocab_cards, vocab_images, vocab-images를 삭제하지 않습니다.
-- 기존 관리자 검수 데이터와 이미지 버전 구조를 보존하기 위한 JSON 카드 상태 테이블입니다.
-- 실행은 새 SQL Editor 쿼리에서 단 한 번.
-- 한컷영단어 v2.6: 새 Supabase 프로젝트에서 SQL Editor로 실행.
-- 중요: 서비스 역할 키는 Python 서버 환경변수 전용입니다. 브라우저에 노출 금지.
create extension if not exists pgcrypto;

create table if not exists public.onecut_state (
  id integer primary key default 1 check (id=1),
  version integer not null default 0,
  cards jsonb not null default '[]'::jsonb check (jsonb_typeof(cards)='array'),
  updated_at timestamptz not null default now()
);

create table if not exists public.onecut_history (
  id bigint generated always as identity primary key,
  version integer not null,
  actor uuid references auth.users(id) on delete set null,
  cards jsonb not null,
  saved_at timestamptz not null default now()
);

insert into public.onecut_state(id,version,cards) values (1,0,'[]'::jsonb)
on conflict (id) do nothing;

alter table public.app_admins enable row level security;
alter table public.onecut_state enable row level security;
alter table public.onecut_history enable row level security;
-- 기본적으로 anon/authenticated 어느 누구에게도 표 직접 접근 권한을 부여하지 않습니다.
revoke all on public.onecut_state, public.onecut_history from anon,authenticated;
-- Python 서버의 서비스 역할만 RPC 호출. 브라우저가 직접 관리자 DB 접근 금지.

create or replace function public.onecut_get_cards()
returns jsonb language sql security definer set search_path = public, pg_temp as $$
 select jsonb_build_object('version',version,'cards',cards) from public.onecut_state where id=1
$$;

create or replace function public.onecut_save_cards(p_cards jsonb,p_expected_version integer,p_actor uuid)
returns jsonb language plpgsql security definer set search_path = public,pg_temp as $$
declare old_version integer; new_version integer; card jsonb; previous jsonb; prev_card jsonb;
begin
 if not exists(select 1 from public.app_admins where user_id=p_actor) then
   raise exception '관리자 권한이 없습니다.';
 end if;
 if jsonb_typeof(p_cards)<>'array' or jsonb_array_length(p_cards)>3000 then
   raise exception '카드 배열이 올바르지 않습니다.';
 end if;
 select version,cards into old_version,previous from public.onecut_state where id=1 for update;
 if old_version is distinct from p_expected_version then
   raise exception '다른 관리자가 먼저 수정했습니다. 클라우드 새로고침 후 다시 진행하세요.';
 end if;
 for card in select value from jsonb_array_elements(p_cards)
 loop
   if card->>'status' in ('approved','published') then
     if coalesce((card->>'sourceChecked')::boolean,false)=false
        or coalesce((card->>'imageChecked')::boolean,false)=false
        or card->>'sceneImageMode'<>'clean' or coalesce(card->>'sceneImage','')=''
        or coalesce(card->>'mnemonic','')='' then
        raise exception '승인되지 않은 카드가 포함되어 있습니다.';
     end if;
   end if;
   -- v3.3: single-action 승인 및 공개 allowed; enforce the same server-side gates.
   if card->>'status'='published' then
     if coalesce(jsonb_typeof(card->'scores'),'') <> 'array'
        or jsonb_array_length(card->'scores')<>4
        or coalesce((select sum((v.value)::numeric) from jsonb_array_elements_text(card->'scores') as v(value)),0) < 85
        or card->'evaluation'->>'status' not in ('ai','manual')
        or coalesce(jsonb_array_length(card->'flags'),0)>0
        or coalesce((card->>'sourceChecked')::boolean,false)=false
        or coalesce((card->>'imageChecked')::boolean,false)=false
        or (coalesce((card->>'speechRequired')::boolean,false)=true and
           (coalesce((card->>'speechChecked')::boolean,false)=false or nullif(card->>'sceneSpeech','') is null))
        or card->>'sceneImageMode'<>'clean'
        or coalesce(card->>'sceneImage','')='' then
       raise exception '학생 공개 조건을 충족하지 않은 카드입니다.';
     end if;
   end if;
 end loop;
 new_version:=old_version+1;
 update public.onecut_state set version=new_version,cards=p_cards,updated_at=now() where id=1;
 insert into public.onecut_history(version,actor,cards) values(new_version,p_actor,p_cards);
 return jsonb_build_object('version',new_version,'saved',true);
end;
$$;

-- 학생에게는 공개된 카드의 학습용 필드만 전달. 점수·메모·이력·거절 사유 제외.
create or replace function public.onecut_public_cards()
returns jsonb language sql security definer set search_path = public,pg_temp as $$
 select coalesce(jsonb_agg(jsonb_build_object(
   'id',c->>'id','word',c->>'word','pos',c->>'pos','ipa',c->>'ipa',
   'hangul',c->>'hangul','meaning',c->>'meaning',
   'stress',c->'stress','mnemonic',c->>'mnemonic','scene',c->>'scene',
   'sceneImage',c->>'sceneImage','sceneImageMode','clean',
   'example',c->>'example','translation',c->>'translation','status','published'
  )), '[]'::jsonb)
 from public.onecut_state st, lateral jsonb_array_elements(st.cards) c
 where st.id=1 and c->>'status'='published'
   and c->>'sceneImageMode'='clean' and c->>'imageChecked'='true'
   and c->>'sourceChecked'='true'
   and coalesce((select sum(score.value::integer) from jsonb_array_elements_text(c->'scores') as score(value)),0)>=85
   and coalesce(jsonb_array_length(c->'flags'),0)=0
$$;

revoke all on function public.onecut_get_cards() from public,anon,authenticated;
revoke all on function public.onecut_save_cards(jsonb,integer,uuid) from public,anon,authenticated;
revoke all on function public.onecut_public_cards() from public,anon,authenticated;
grant execute on function public.onecut_get_cards() to service_role;
grant execute on function public.onecut_save_cards(jsonb,integer,uuid) to service_role;
grant execute on function public.onecut_public_cards() to service_role;

