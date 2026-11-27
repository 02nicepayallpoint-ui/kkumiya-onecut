-- v3.4.2: 레거시 공개 카드가 변경되지 않았으면 신규 검수 재요구 없이 저장 허용
-- 변경된 공개/승인 카드에는 기존 품질 기준을 그대로 적용
-- 기존 데이터 삭제/초기화 없음
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
   if card->>'status' in ('approved','published') AND NOT EXISTS (
      SELECT 1 FROM jsonb_array_elements(previous) AS existing_card(value)
      WHERE existing_card.value->>'id' = card->>'id' AND existing_card.value = card
    ) then
     if coalesce((card->>'sourceChecked')::boolean,false)=false
        or coalesce((card->>'imageChecked')::boolean,false)=false
        or card->>'sceneImageMode'<>'clean' or coalesce(card->>'sceneImage','')=''
        or coalesce(card->>'mnemonic','')='' then
        raise exception '승인되지 않은 카드가 포함되어 있습니다.';
     end if;
   end if;
   -- v3.3: single-action 승인 및 공개 allowed; enforce the same server-side gates.
   if card->>'status'='published' AND NOT EXISTS (
      SELECT 1 FROM jsonb_array_elements(previous) AS existing_card(value)
      WHERE existing_card.value->>'id' = card->>'id' AND existing_card.value = card
    ) then
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

