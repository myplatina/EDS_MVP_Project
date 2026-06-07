# ED's MVP PWA v0.8.62

## 실행

```cmd
npm install
npm run dev
```

## 배포

```cmd
vercel --prod
```

---

## Changelog

### v0.8.62 (2026-06-07)

**보안 레이어 구축**
- GAS 모든 KIS AppKey/Secret 하드코딩 제거 → `PropertiesService.getScriptProperties()` 참조로 전환
- `api.ts`: 모든 GET/POST 요청에 `x_eds_app_token` 필드 자동 주입 (`VITE_EDS_APP_TOKEN` env)
- `EdsMvpAPI.gs`: `doGet/doPost` 최상단에 `edApi_validateAppToken_` 인터셉터 추가 (ScriptProperties `EDS_APP_TOKEN`)
- 중복 `refreshKrxPrices` case 제거

**홈 화면 UX 개편**
- 홈 화면 최상단에 **Quick Action Bar** 추가 (⚡ 현재가 갱신 / 📊 차트 갱신 + Last Updated 실시간 바인딩)
- 설정 메뉴에 **Advanced Settings 아코디언** 추가 — 버전/상태, 새로고침 로그, PWA 설치 섹션을 기본 접힘으로 이동

**위험 모니터링 엔진 (Risk Monitoring Engine)**
- `utils.ts`에 `computeRiskFlags()` 추가
  - **Drop Warning**: 총수익률 < -20% → 종목명 옆 붉은 ● 표시
  - **Concentration Warning**: 전체 비중 > 15% → 주황 ● 표시
  - **Weight Deviation Warning**: |계좌비중 - 목표비중| / 목표비중 ≥ 30% → 황색 ● 표시
- `utils.ts`에 `computeDeficitCandidates()` 추가
  - 괴리율 -30% 이하 부족 종목 추출
- 홈 화면 포트폴리오 리스트 최상단에 **"Deficit: ₩[Gap] [종목명] 추가 매수 후보"** Action Card 동적 렌더링

**버전 자동화**
- `bump-version.js`: 모든 GAS 파일 + `api.ts` 버전 일괄 치환 (12개 파일 자동 갱신)
- `vite-env.d.ts`: `VITE_EDS_APP_TOKEN`, `VITE_EDS_API_URL` 환경변수 타입 추가

---

### v0.8.62 (2026-06-07)
- 종목 단일가 동기화 속도 개선 (병렬 fetchAll 전환)
- 커밋 자동화 코드 및 버전 로그 생성 매크로 추가

### v0.8.62
- 설정 탭에 `국내 현재가 자동 갱신` 옵션 추가
- 설정 탭에 `원본 시트 가격 반영(KIS)`, `국내 일봉 차트 갱신(KIS)` 버튼 추가
- 리밸런싱 탭에 `리스크 신호` 섹션 추가
- © 삼평동불나방 · myplatina@gmail.com
- 2026-06-07 : [v0.8.62] feat: v0.8.62 홈 UX 개편, 보안 인터셉터 및 위험 모니터링 엔진 전면 이식
- 2026-06-07 : [v0.8.62] chore: 버전 오버랩 싱크 조정 (v0.8.62 최종 정렬)