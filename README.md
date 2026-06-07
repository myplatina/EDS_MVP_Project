# ED's MVP PWA v0.8.53

## v0.8.53 변경점

- 리밸런싱 탭에 `리스크 신호` 섹션 추가
- 손실 확대, 추가매수 주의, 이익 실현 후보, 비중 과집중, 데이터 이상 신호 표시
- 캐시된 국내 일봉 차트가 있는 종목은 고점 대비 낙폭 신호도 표시
- 상단 버전 표기 뒤 제작자/연락처 표시 추가
  - © 삼평동불나방 · myplatina@gmail.com
- 기존 v0.8.53 리밸런싱 매수/매도 제안 유지

## 실행

```cmd
npm install
npm run dev
```

## 배포

```cmd
vercel --prod
```

v0.8.53은 설정 탭에 국내 일봉 차트 갱신(KIS) 메뉴를 복구한 버전이다. Apps Script에는 refreshKrxDailyCharts action 추가가 필요하다.

## v0.8.53 변경 내용

- 설정 탭에 `국내 현재가 자동 갱신` 옵션 추가
  - 끄기 / 1분 / 3분 / 5분 / 10분
  - 국내 장중에만 갱신 옵션 포함
  - 앱이 열려 있고 화면이 활성 상태일 때만 `refreshKrxPrices` 호출
- 설정 탭에 `원본 시트 가격 반영(KIS)` 버튼 추가
  - App_Prices / App_Output 갱신
  - 원본 `2. 종목현황`의 국내 종목 K열 현재가, M열 전일 대비 등락률을 KIS 값으로 기록
  - 기존 GOOGLEFINANCE / NAVER 수식이 값으로 대체될 수 있으므로 실행 전 백업 권장
- Apps Script 백엔드 패치 필요
  - `EdsMvpKisPriceRefresh_v0_8_4.gs`로 교체
  - `EdsMvpApi.gs`에 `refreshKrxPricesToMainSheet` action 추가


## v0.8.53 추가 변경

- 설정 탭에 `국내 일봉 차트 갱신(KIS)` 버튼 추가
- 새로고침 로그에 `종목 차트 일괄 갱신` 표시
- action: `refreshKrxDailyCharts` 호출

# ED's MVP PWA v0.8.53
## v0.8.53 변경점
- 2026-06-07 : 종목 단일가의 동기화 속도 개선 코드 작업
- 2026-06-07 : 커밋 자동화 코드 추가 및 버전 로그 생성 매크로 추가

- 2026-06-07 : fix: 프로젝트 버전 v0.8.53 동기화 및 README 규격 정립

- 2026-06-07 : [v0.8.53] readme 최종 규격 강제 정렬
- 2026-06-07 : [v0.8.53] chore: package.json 버전 규격 표준화 및 파이프라인 정상화 테스트