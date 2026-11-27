-- v3.4.6: 1인 관리자의 최종 콘텐츠 공개 승인. 기존 데이터 삭제 없음.
-- 품질평가 85점 및 과거 flags는 경고사항이며 공개 차단 조건이 아닙니다.
-- 클라우드 저장은 관리자 전용 RPC에서만 가능하고 앱 서버는 이미지 실재 여부를 검사합니다.
create or replace function public.onecut_save_cards(p_cards jsonb,p_expected_version integer,p_actor uuid)
returns jsonb language plpgsql security definer set search_path=public,pg_temp as $$
declare old_version integer; new_version integer; card jsonb; previous jsonb; previous_card jsonb;
begin
 if not exists(select 1 from public.app_admins where user_id=p_actor) then
   raise exception '관리자 권한이 없습니다.';
 end if;
 if p_cards is null or jsonb_typeof(p_cards)<>'array' or jsonb_array_length(p_cards)>3000 then
   raise exception '카드 배열이 올바르지 않습니다.';
 end if;
 select version,cards into old_version,previous from public.onecut_state where id=1 for update;
 if old_version is distinct from p_expected_version then
   raise exception '다른 저장 내역이 있습니다. 클라우드를 새로고침하세요.';
 end if;
 for card in select value from jsonb_array_elements(p_cards) loop
   if coalesce(card->>'id','')='' or coalesce(card->>'word','')='' or card->>'status' not in ('draft','review','approved','published','rejected') then
     raise exception '카드 ID 또는 상태가 잘못되었습니다.';
   end if;
   select value into previous_card from jsonb_array_elements(previous) where value->>'id'=card->>'id' limit 1;
   -- Existing published cards keep their state; new publication requires essential content.
   if card->>'status'='published' and
      (previous_card is null or previous_card->>'status' is distinct from 'published') then
      if coalesce(card->>'pos','')='' or coalesce(card->>'ipa','')='' or coalesce(card->>'meaning','')=''
         or coalesce(card->>'mnemonic','')='' or coalesce(card->>'scene','')=''
         or coalesce(card->>'example','')='' or coalesce(card->>'translation','')=''
         or coalesce(card->>'sceneImage','')='' or card->>'sceneImageMode'<>'clean'
         or coalesce((card->>'sourceChecked')::boolean,false)=false
         or coalesce((card->>'imageChecked')::boolean,false)=false then
         raise exception '필수 콘텐츠 및 이미지 최종 확인이 필요합니다.';
      end if;
   end if;
 end loop;
 new_version:=old_version+1;
 update public.onecut_state set version=new_version,cards=p_cards,updated_at=now() where id=1;
 insert into public.onecut_history(version,actor,cards) values(new_version,p_actor,p_cards);
 return jsonb_build_object('version',new_version,'saved',true);
end;
$$;
-- Preserve existing server-only execute grants; do not expose the RPC to client roles.
revoke all on function public.onecut_save_cards(jsonb,integer,uuid) from public,anon,authenticated;
grant execute on function public.onecut_save_cards(jsonb,integer,uuid) to service_role;
