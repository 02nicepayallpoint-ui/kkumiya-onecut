-- 한컷영단어 v3.5.0: 학생별 학습 기록. 기존 콘텐츠 테이블을 변경하지 않습니다.
CREATE TABLE IF NOT EXISTS public.onecut_student_progress (
 user_id uuid PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
 progress jsonb NOT NULL DEFAULT '{}'::jsonb,
 updated_at timestamptz NOT NULL DEFAULT now(),
 CONSTRAINT progress_is_object CHECK (jsonb_typeof(progress)='object'),
 CONSTRAINT progress_size_limit CHECK (pg_column_size(progress)<500000)
);
ALTER TABLE public.onecut_student_progress ENABLE ROW LEVEL SECURITY;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.onecut_student_progress TO authenticated;
DROP POLICY IF EXISTS "Student read own progress" ON public.onecut_student_progress;
DROP POLICY IF EXISTS "Student insert own progress" ON public.onecut_student_progress;
DROP POLICY IF EXISTS "Student update own progress" ON public.onecut_student_progress;
CREATE POLICY "Student read own progress" ON public.onecut_student_progress FOR SELECT TO authenticated
 USING ((select auth.uid()) = user_id);
CREATE POLICY "Student insert own progress" ON public.onecut_student_progress FOR INSERT TO authenticated
 WITH CHECK ((select auth.uid()) = user_id);
CREATE POLICY "Student update own progress" ON public.onecut_student_progress FOR UPDATE TO authenticated
 USING ((select auth.uid()) = user_id) WITH CHECK ((select auth.uid()) = user_id);
-- updated_at 자동 반영
CREATE OR REPLACE FUNCTION public.onecut_touch_progress() RETURNS trigger
LANGUAGE plpgsql SET search_path='' AS $$ BEGIN NEW.updated_at = now(); RETURN NEW; END; $$;
DROP TRIGGER IF EXISTS onecut_student_progress_timestamp ON public.onecut_student_progress;
CREATE TRIGGER onecut_student_progress_timestamp BEFORE UPDATE ON public.onecut_student_progress
 FOR EACH ROW EXECUTE FUNCTION public.onecut_touch_progress();
