# orom-ep

오롬(orom.co.kr) 네이버 쇼핑 EP 자동 생성.

- **네이버 DBURL**: https://oromadmin.github.io/orom-ep/naver_ep.tsv
- 상태: https://oromadmin.github.io/orom-ep/status.json · 로그: https://oromadmin.github.io/orom-ep/build.log
- 원천: 구글시트 "오롬 EP 허브" (team@b2b.orom.co.kr) — Apps Script가 Cafe24에서 매시간 동기화
- 실행: `.github/workflows/build-ep.yml` (하루 5회, KST 00:13·09:13·13:13·17:13·21:13 + 수동)
- 검증 실패 시 커밋하지 않음 → 네이버는 직전 정상 파일을 계속 수집

Secrets: `OROM_SHEET_PRODUCTS_CSV`, `OROM_SHEET_OVERRIDES_CSV` (시트 '웹에 게시' CSV 주소), `OROM_SHEET_SNAPSHOT_CSV`(선택)
