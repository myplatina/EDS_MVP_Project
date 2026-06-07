# EDS MVP - Google Sheet Ledger Schema (Source of Truth)

## 1. Core Ledgers (원장 시트)

### 🔹 2. 종목현황
* **역할:** 보유 종목의 원화/달러 자산 실적 및 비중 관리 원장
* **스키마 (주요 컬럼):**
  * `증권사` | `계좌` | `국가` | `종목코드 (티커)` | `종목명` | `수량` | `평단가(원화)` | `평단가(달러)` | `현재가(원화)` | `현재가(달러)` | `전일 대비 등락` | `평가액 [원화]` | `투자비중 (전체 자산 내)` | `투자비중 (계좌 내)` | `목표비중 (계좌 내)` | `누적배당금 [원화]` | `누적수익 [원화]` | `총수익률 [평가+배당]`

### 🔹 6. 배당내역
* **역할:** 전체 배당금 수령 히스토리 유지 원장 (삭제 불가 원칙)
* **스키마:**
  * `일자` | `연도` | `월` | `일` | `증권사` | `계좌` | `종목코드` | `종목명` | `원화 배당금` | `외화 배당금` | `원화환산`

### 🔹 5. 입금내역
* **역할:** 자산 계좌별 실제 투자 원금 입출금 기록
* **스키마:**
  * `일자` | `연도` | `월` | `일` | `증권사` | `계좌` | `입금` | `비고`

---

## 2. App Cache & View Layers (App 전용 시트)

### 🔹 App_Output
* **역할:** PWA 메인 홈 및 포트폴리오 탭 렌더링용 핵심 응답 데이터 API 캐시 공간
* **스키마:**
  * `account_id` | `account_name` | `broker` | `account_type` | `asset_id` | `ticker` | `asset_name` | `quantity` | `avg_price` | `price` | `invested_amount` | `valuation_amount` | `profit_amount` | `profit_rate` | `account_weight` | `total_weight` | `target_weight_account` | `target_gap_rate` | `target_gap_amount` | `currency` | `price_source` | `price_fetched_at` | `updated_at`

### 🔹 App_Prices
* **역할:** KIS API로부터 수집된 국내/외 종목 최신 시세 캐시 레이어
* **스키마:**
  * `asset_id` | `ticker` | `price` | `prev_close` | `change_amount` | `change_rate` | `currency` | `source` | `fetched_at` | `updated_at` | `source_price` | `source_currency` | `fx_rate` | `price_market` | `price_exchange`

### 🔹 App_Holdings
* **역할:** 가공용 계좌별 자산 보유 현황 데이터 마스터
* **스키마:**
  * `holding_id` | `account_id` | `asset_id` | `ticker` | `asset_name` | `quantity` | `avg_price` | `target_weight_account` | `memo` | `enabled` | `created_at` | `updated_at`

### 🔹 App_Dividends
* **역할:** PWA 배당 대시보드 출력 데이터
* **스키마:**
  * `dividend_id` | `source_sheet` | `source_row` | `dividend_date` | `year` | `month` | `day` | `broker` | `account_name` | `account_type` | `account_id` | `ticker` | `asset_name` | `market` | `asset_id` | `krw_amount` | `foreign_amount` | `net_amount_krw` | `currency` | `holding_status` | `current_holding_id` | `memo` | `enabled` | `created_at` | `updated_at`

### 🔹 App_PriceRefreshResult
* **역할:** 가격 갱신 자동화 배치 처리 로그 및 모니터링 메트릭
* **스키마:**
  * `asset_id` | `ticker` | `asset_name` | `market` | `exchange` | `status` | `price_krw` | `source_price` | `source_currency` | `fx_rate` | `change_amount_krw` | `change_rate` | `raw_status` | `message`

### 🔹 App_Settings
* **역할:** 전역 앱 환경 설정 값 매핑 정보
* **스키마:**
  * `key` | `value` | `description` | `updated_at`