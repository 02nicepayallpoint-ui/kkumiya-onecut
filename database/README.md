# Supabase DB 적용 순서

기존 프로젝트는 이미 적용한 SQL을 다시 실행할 필요가 없습니다.

신규 설치 시 숫자 순서대로 적용합니다. v4.10까지 사용 중이었다면 `07_groups_assignments.sql` 다음으로 `08_v5_platform.sql`만 1회 추가 실행합니다.

- 01_cloud_bridge_existing_project.sql
- 02_legacy_import_patch.sql
- 03_allow_unchanged_legacy_cards.sql
- 04_keep_published_on_edit.sql
- 05_admin_final_approval.sql
- 06_student_accounts_progress.sql
- 07_groups_assignments.sql
- 08_v5_platform.sql

`legacy_v2_6_DO_NOT_RUN.sql`은 보관용이며 실행하지 않습니다.
