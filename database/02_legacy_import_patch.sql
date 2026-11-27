-- 한컷영단어 v3.4.1: 기존 로컬 카드의 최초 이전만 허용.
-- 반드시 이전 01_cloud_bridge_existing_project.sql 실행 후 적용하세요.
-- 기존 데이터 삭제 / 상태 덮어쓰기를 수행하지 않습니다.
-- 공개 카드의 기존 published 상태는 복원합니다. 신규 승인 검증은 onecut_save_cards에 남겨둡니다.
CREATE OR REPLACE FUNCTION public.onecut_migrate_legacy_cards(
  p_cards jsonb, p_expected_version integer, p_actor uuid
) RETURNS jsonb
LANGUAGE plpgsql SECURITY DEFINER SET search_path=public,pg_temp AS $$
DECLARE old_version integer; existing_cards jsonb; new_version integer;
BEGIN
  IF NOT EXISTS (SELECT 1 FROM public.app_admins WHERE user_id=p_actor) THEN
    RAISE EXCEPTION '관리자 권한이 없습니다.';
  END IF;
  IF p_cards IS NULL OR jsonb_typeof(p_cards) <> 'array'
     OR jsonb_array_length(p_cards)=0 OR jsonb_array_length(p_cards)>3000 THEN
    RAISE EXCEPTION '이전할 카드 데이터가 올바르지 않습니다.';
  END IF;
  IF EXISTS (SELECT 1 FROM jsonb_array_elements(p_cards) AS item
        WHERE jsonb_typeof(item) <> 'object'
           OR coalesce(item->>'id','') = '' OR coalesce(item->>'word','') = ''
           OR coalesce(item->>'status','') NOT IN ('draft','review','approved','published','rejected')) THEN
    RAISE EXCEPTION '카드 ID, 단어, 상태 필드가 올바르지 않습니다.';
  END IF;
  IF (SELECT count(DISTINCT item->>'id') FROM jsonb_array_elements(p_cards) AS item)
       <> jsonb_array_length(p_cards) THEN
    RAISE EXCEPTION '카드 ID가 중복되었습니다.';
  END IF;
  SELECT version,cards INTO old_version,existing_cards
    FROM public.onecut_state WHERE id=1 FOR UPDATE;
  IF NOT FOUND THEN RAISE EXCEPTION 'onecut_state 초기화가 필요합니다.'; END IF;
  IF old_version IS DISTINCT FROM p_expected_version THEN
    RAISE EXCEPTION '클라우드 버전이 변경됐습니다. 새로고침 후 확인하세요.';
  END IF;
  IF jsonb_array_length(existing_cards)<>0 THEN
    RAISE EXCEPTION '클라우드에 이미 카드가 있어 중복 이전을 차단했습니다.';
  END IF;
  new_version:=old_version+1;
  UPDATE public.onecut_state SET version=new_version,cards=p_cards,updated_at=now() WHERE id=1;
  INSERT INTO public.onecut_history(version,actor,cards) VALUES(new_version,p_actor,p_cards);
  RETURN jsonb_build_object('version',new_version,'migrated',true,'count',jsonb_array_length(p_cards));
END;
$$;
REVOKE ALL ON FUNCTION public.onecut_migrate_legacy_cards(jsonb,integer,uuid)
  FROM PUBLIC,anon,authenticated;
GRANT EXECUTE ON FUNCTION public.onecut_migrate_legacy_cards(jsonb,integer,uuid) TO service_role;

-- 이미 검수되어 published였던 레거시 카드의 공개 이력을 그대로 유지.
-- 신규 공개 시의 검수 검증은 기존 onecut_save_cards RPC와 Python validate_cards에서 계속 적용.
CREATE OR REPLACE FUNCTION public.onecut_public_cards()
RETURNS jsonb LANGUAGE sql SECURITY DEFINER SET search_path=public,pg_temp AS $$
  SELECT coalesce(jsonb_agg(jsonb_build_object(
    'id',c->>'id','word',c->>'word','pos',c->>'pos','ipa',c->>'ipa',
    'hangul',c->>'hangul','meaning',c->>'meaning',
    'stress',c->'stress','mnemonic',c->>'mnemonic','scene',c->>'scene',
    'sceneImage',c->>'sceneImage','sceneImageMode',c->>'sceneImageMode',
    'example',c->>'example','translation',c->>'translation','status','published'
  )), '[]'::jsonb)
  FROM public.onecut_state st, lateral jsonb_array_elements(st.cards) c
  WHERE st.id=1 AND c->>'status'='published'
$$;
REVOKE ALL ON FUNCTION public.onecut_public_cards() FROM PUBLIC,anon,authenticated;
GRANT EXECUTE ON FUNCTION public.onecut_public_cards() TO service_role;
