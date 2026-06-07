/*******************************************************
 * ED's MVP - Chart Data Cache + KIS Integration v0.8.5.2
 *
 * 목적:
 * - 한국투자증권 Open API로 국내 ETF/주식 일/주/月 차트 데이터를 조회
 * - App_ChartPrices 시트에 캐시
 * - PWA가 getChartData / refreshChartData 형태로 사용할 수 있는 데이터 구조 제공
 *
 * 전제:
 * - EdsMvpKisChartTest.gs v0.1에서 testKisDailyChart_458730 등 성공 확인 완료
 * - setupKisCredentials()로 KIS appkey/appsecret 저장 완료
 * - EdsMvpApi.gs에는 getChartData / refreshChartData action 연결 패치 필요
 *******************************************************/

const ED_MVP_CHART = {
  sheets: {
    chartPrices: "App_ChartPrices",
    assets: "App_Assets",
    syncLog: "App_SyncLog",
  },

  source: {
    kis: "kis",
  },

  interval: {
    daily: "D",
    weekly: "W",
    monthly: "M",
  },

  defaultLimit: {
    D: 120,
    W: 104,
    M: 60,
  },
};

/**
 * 최초 1회 실행:
 * App_ChartPrices 시트 생성/정비
 */
function setupChartPricesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ED_MVP_CHART.sheets.chartPrices);

  if (!sheet) {
    sheet = ss.insertSheet(ED_MVP_CHART.sheets.chartPrices);
  }

  const headers = [
    "chart_id",
    "asset_id",
    "ticker",
    "interval",
    "date",
    "open",
    "high",
    "low",
    "close",
    "volume",
    "source",
    "fetched_at",
    "created_at",
    "updated_at",
  ];

  const firstRow = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0]
    : [];

  if (String(firstRow[0] || "") !== "chart_id") {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#137333")
    .setHorizontalAlignment("center");

  SpreadsheetApp.getActiveSpreadsheet().toast("App_ChartPrices 시트 정비 완료", "ED's MVP", 5);
}

/**
 * 테스트용: 458730 일봉을 KIS에서 조회 후 App_ChartPrices에 저장
 */
function refreshChartData_458730_D() {
  const result = refreshKrxChartDataFromKis({
    asset_id: "KRX_458730",
    ticker: "458730",
    interval: "D",
    limit: 120,
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `458730 D chart 저장 완료: ${result.count}건`,
    "ED's MVP",
    8
  );
}

function refreshChartData_458730_W() {
  const result = refreshKrxChartDataFromKis({
    asset_id: "KRX_458730",
    ticker: "458730",
    interval: "W",
    limit: 104,
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `458730 W chart 저장 완료: ${result.count}건`,
    "ED's MVP",
    8
  );
}

function refreshChartData_458730_M() {
  const result = refreshKrxChartDataFromKis({
    asset_id: "KRX_458730",
    ticker: "458730",
    interval: "M",
    limit: 60,
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `458730 M chart 저장 완료: ${result.count}건`,
    "ED's MVP",
    8
  );
}

/**
 * App_Assets에 등록된 KRX 종목 전체 일봉 갱신
 * 호출량을 줄이기 위해 수동 실행 권장
 */
function refreshAllKrxDailyChartsFromKis() {
  setupChartPricesSheet();

  const assets = edChart_readSheetAsObjects_(ED_MVP_CHART.sheets.assets)
    .filter((row) => String(row.enabled || "").toUpperCase() === "TRUE")
    .filter((row) => String(row.market || "") === "KRX")
    .filter((row) => String(row.ticker || "").toUpperCase().match(/^[0-9A-Z]{6}$/));

  const results = [];

  assets.forEach((asset) => {
    try {
      const result = refreshKrxChartDataFromKis({
        asset_id: asset.asset_id,
        ticker: asset.ticker,
        interval: "D",
        limit: 120,
      });
      results.push([asset.asset_id, asset.ticker, "success", result.count, ""]);
      Utilities.sleep(250);
    } catch (e) {
      results.push([asset.asset_id, asset.ticker, "error", 0, e && e.message ? e.message : String(e)]);
    }
  });

  edChart_writeRefreshSummary_("App_ChartRefreshResult", results);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `KRX 일봉 일괄 갱신 완료: ${assets.length}종목`,
    "ED's MVP",
    8
  );
}

/**
 * KIS에서 국내 차트 데이터를 가져와 App_ChartPrices에 캐시
 * payload: { asset_id, ticker, interval, limit }
 */
function refreshKrxChartDataFromKis(payload) {
  setupChartPricesSheet();

  const assetId = String(payload && payload.asset_id || "").trim();
  const ticker = edChart_normalizeTicker_(payload && payload.ticker || edChart_tickerFromAssetId_(assetId));
  const interval = edChart_normalizeInterval_(payload && payload.interval || "D");
  const limit = Number(payload && payload.limit || ED_MVP_CHART.defaultLimit[interval] || 120);

  if (!assetId) throw new Error("asset_id가 필요합니다.");
  if (!ticker) throw new Error("ticker가 필요합니다.");
  if (!/^[0-9A-Z]{6}$/.test(ticker)) {
  throw new Error(`KIS 국내 차트는 6자리 KRX 단축코드만 지원: ${ticker}`);
}

  const today = new Date();
  const end = edKis_formatDate_(today);
  const start = edKis_formatDate_(edKis_addDays_(today, -edKis_periodLookbackDays_(interval, limit)));

  const result = edKis_fetchDomesticChart_(ticker, interval, start, end);
  const items = result.items.slice(-limit);

  edChart_replaceChartRows_(assetId, ticker, interval, items, ED_MVP_CHART.source.kis);

  return {
    asset_id: assetId,
    ticker,
    interval,
    count: items.length,
    source: ED_MVP_CHART.source.kis,
    fetched_at: new Date(),
    rt_cd: result.raw && result.raw.rt_cd,
    msg_cd: result.raw && result.raw.msg_cd,
    msg1: result.raw && result.raw.msg1,
  };
}

/**
 * 캐시된 차트 데이터 조회
 * payload: { asset_id, interval, limit }
 */
function getChartData(payload) {
  const assetId = String(payload && payload.asset_id || "").trim();
  const interval = edChart_normalizeInterval_(payload && payload.interval || "D");
  const limit = Number(payload && payload.limit || ED_MVP_CHART.defaultLimit[interval] || 120);

  if (!assetId) throw new Error("asset_id가 필요합니다.");

  const rows = edChart_readSheetAsObjects_(ED_MVP_CHART.sheets.chartPrices)
    .filter((row) => String(row.asset_id || "") === assetId)
    .filter((row) => String(row.interval || "") === interval)
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")))
    .slice(-limit)
    .map((row) => ({
      date: row.date,
      open: edChart_num_(row.open),
      high: edChart_num_(row.high),
      low: edChart_num_(row.low),
      close: edChart_num_(row.close),
      volume: edChart_num_(row.volume),
      source: row.source,
      fetched_at: row.fetched_at,
    }));

  return {
    asset_id: assetId,
    interval,
    count: rows.length,
    items: rows,
  };
}

/**
 * API용: 캐시가 비어 있으면 KIS 갱신 후 반환
 * payload: { asset_id, ticker, market, interval, limit, force }
 */
function getOrRefreshChartData(payload) {
  const assetId = String(payload && payload.asset_id || "").trim();
  const interval = edChart_normalizeInterval_(payload && payload.interval || "D");
  const force = Boolean(payload && payload.force);

  if (!assetId) throw new Error("asset_id가 필요합니다.");

  let cached = getChartData({
    asset_id: assetId,
    interval,
    limit: payload && payload.limit,
  });

  if (force || cached.count === 0) {
    const market = String(payload && payload.market || edChart_marketFromAssetId_(assetId));

    if (market !== "KRX") {
      return {
        asset_id: assetId,
        interval,
        count: 0,
        items: [],
        message: "KRX가 아닌 종목은 Apps Script 자체 차트 대상이 아닙니다.",
      };
    }

    refreshKrxChartDataFromKis({
      asset_id: assetId,
      ticker: payload && payload.ticker,
      interval,
      limit: payload && payload.limit,
    });

    cached = getChartData({
      asset_id: assetId,
      interval,
      limit: payload && payload.limit,
    });
  }

  return cached;
}

function edChart_replaceChartRows_(assetId, ticker, interval, items, source) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_CHART.sheets.chartPrices);
  if (!sheet) throw new Error("App_ChartPrices 시트를 찾을 수 없습니다.");

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const assetCol = headers.indexOf("asset_id");
  const intervalCol = headers.indexOf("interval");

  if (assetCol < 0 || intervalCol < 0) throw new Error("App_ChartPrices 헤더 오류");

  // 기존 동일 asset_id + interval 행 삭제. 아래에서 위로 삭제.
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][assetCol]) === assetId && String(values[i][intervalCol]) === interval) {
      sheet.deleteRow(i + 1);
    }
  }

  if (!items || items.length === 0) return;

  const now = new Date();
  const rows = items.map((item) => [
    `${assetId}_${interval}_${item.date}`,
    assetId,
    ticker,
    interval,
    item.date,
    item.open,
    item.high,
    item.low,
    item.close,
    item.volume,
    source,
    now,
    now,
    now,
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function edChart_readSheetAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  return values.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = edChart_normalizeValue_(row[index]);
    });
    return obj;
  });
}

function edChart_writeRefreshSummary_(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  const output = [["asset_id", "ticker", "status", "count", "message"]].concat(rows);
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, output[0].length);
  sheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold").setBackground("#137333").setFontColor("#ffffff");
}

function edChart_normalizeInterval_(value) {
  const text = String(value || "D").trim().toUpperCase();
  if (text === "1D" || text === "DAY" || text === "DAILY") return "D";
  if (text === "1W" || text === "WEEK" || text === "WEEKLY") return "W";
  if (text === "1M" || text === "MONTH" || text === "MONTHLY") return "M";
  if (["D", "W", "M"].indexOf(text) >= 0) return text;
  return "D";
}

function edChart_normalizeTicker_(value) {
  const text = String(value || "").trim().toUpperCase();
  if (/^\d+$/.test(text) && text.length < 6) return text.padStart(6, "0");
  return text;
}

function edChart_tickerFromAssetId_(assetId) {
  const parts = String(assetId || "").split("_");
  return parts.length >= 2 ? parts.slice(1).join("_") : "";
}

function edChart_marketFromAssetId_(assetId) {
  return String(assetId || "").split("_")[0] || "";
}

function edChart_num_(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  const n = Number(String(value).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function edChart_normalizeValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  return value;
}

function refreshChartData_0060H0_D() {
  const result = refreshKrxChartDataFromKis({
    asset_id: "KRX_0060H0",
    ticker: "0060H0",
    interval: "D",
    limit: 120,
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `0060H0 D chart 저장 완료: ${result.count}건`,
    "ED's MVP",
    8
  );
}

/*******************************************************
 * ED's MVP - Fast KRX Daily Chart Refresh Patch
 *
 * 목적:
 * - 기존 refreshAllKrxDailyChartsFromKis()의 timeout 방지
 * - deleteRow 반복 제거
 * - KRX 일봉 전체를 한 번에 재작성
 *******************************************************/

function refreshAllKrxDailyChartsFromKisFast() {
  setupChartPricesSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startedAt = new Date();

  const assets = edChart_readSheetAsObjects_(ED_MVP_CHART.sheets.assets)
    .filter((row) => String(row.enabled || "").toUpperCase() === "TRUE")
    .filter((row) => String(row.market || "") === "KRX")
    .filter((row) => String(row.ticker || "").toUpperCase().match(/^[0-9A-Z]{6}$/));

  const interval = "D";
  const limit = ED_MVP_CHART.defaultLimit[interval] || 120;

  const today = new Date();
  const end = edKis_formatDate_(today);
  const start = edKis_formatDate_(
    edKis_addDays_(today, -edKis_periodLookbackDays_(interval, limit))
  );

  const targetAssetIds = new Set();
  const newChartRows = [];
  const results = [];

  assets.forEach((asset, index) => {
    const assetId = String(asset.asset_id || "").trim();
    const ticker = edChart_normalizeTicker_(asset.ticker);

    if (!assetId || !ticker) {
      results.push([assetId, ticker, interval, "skip", 0, "asset_id 또는 ticker 없음"]);
      return;
    }

    targetAssetIds.add(assetId);

    try {
      const result = edChart_fetchKisChartWithRetry_(ticker, interval, start, end);
      const items = result.items.slice(-limit);
      const now = new Date();

      items.forEach((item) => {
        newChartRows.push([
          `${assetId}_${interval}_${item.date}`,
          assetId,
          ticker,
          interval,
          item.date,
          item.open,
          item.high,
          item.low,
          item.close,
          item.volume,
          ED_MVP_CHART.source.kis,
          now,
          now,
          now,
        ]);
      });

      results.push([
        assetId,
        ticker,
        interval,
        "success",
        items.length,
        result.raw && result.raw.msg1 ? result.raw.msg1 : ""
      ]);

      // KIS 호출 제한 완화
      Utilities.sleep(650);
    } catch (e) {
      results.push([
        assetId,
        ticker,
        interval,
        "error",
        0,
        e && e.message ? e.message : String(e)
      ]);

      Utilities.sleep(1000);
    }
  });

  edChart_replaceDailyRowsInBulk_(targetAssetIds, newChartRows);
  edChart_writeRefreshSummaryAllIntervals_("App_ChartRefreshResult", results);

  ss.toast(
    `KRX 일봉 빠른 갱신 완료: target=${assets.length}, rows=${newChartRows.length}`,
    "ED's MVP",
    8
  );

  return {
    started_at: startedAt,
    finished_at: new Date(),
    target_count: assets.length,
    row_count: newChartRows.length,
    success_count: results.filter((r) => r[3] === "success").length,
    error_count: results.filter((r) => r[3] === "error").length,
  };
}

function edChart_replaceDailyRowsInBulk_(targetAssetIds, newRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_CHART.sheets.chartPrices);

  if (!sheet) throw new Error("App_ChartPrices 시트를 찾을 수 없습니다.");

  const headers = [
    "chart_id",
    "asset_id",
    "ticker",
    "interval",
    "date",
    "open",
    "high",
    "low",
    "close",
    "volume",
    "source",
    "fetched_at",
    "created_at",
    "updated_at",
  ];

  const lastRow = sheet.getLastRow();

  let keptRows = [];

  if (lastRow >= 2) {
    const values = sheet.getRange(1, 1, lastRow, headers.length).getValues();
    const existingHeader = values[0];

    const assetCol = existingHeader.indexOf("asset_id");
    const intervalCol = existingHeader.indexOf("interval");

    if (assetCol < 0 || intervalCol < 0) {
      throw new Error("App_ChartPrices 헤더 오류: asset_id 또는 interval 없음");
    }

    keptRows = values.slice(1).filter((row) => {
      const assetId = String(row[assetCol] || "");
      const interval = String(row[intervalCol] || "");

      // 이번 갱신 대상 KRX 종목의 D 일봉만 교체
      if (interval === "D" && targetAssetIds.has(assetId)) {
        return false;
      }

      return true;
    });
  }

  const output = [headers].concat(keptRows).concat(newRows);

  sheet.clear();
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#137333")
    .setHorizontalAlignment("center");
}

function edChart_fetchKisChartWithRetry_(ticker, interval, start, end) {
  const maxAttempts = 3;
  const baseSleepMs = 1500;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return edKis_fetchDomesticChart_(ticker, interval, start, end);
    } catch (e) {
      lastError = e;
      const message = e && e.message ? e.message : String(e);

      // KIS 초당 호출 제한이면 대기 후 재시도
      if (
        message.indexOf("EGW00201") >= 0 ||
        message.indexOf("초당 거래건수") >= 0
      ) {
        Utilities.sleep(baseSleepMs * attempt);
        continue;
      }

      throw e;
    }
  }

  throw lastError;
}

function edChart_writeRefreshSummaryAllIntervals_(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();

  const output = [["asset_id", "ticker", "interval", "status", "count", "message"]].concat(rows);

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, output[0].length);
  sheet
    .getRange(1, 1, 1, output[0].length)
    .setFontWeight("bold")
    .setBackground("#137333")
    .setFontColor("#ffffff");
}

/*******************************************************
 * ED's MVP - KIS API Core Helpers
 *
 * 원래 위치: EdsMvpKisChartTest.gs
 * 이관 이유: EdsMvpChartData.gs 및 EdsMvpChartDailyRefresh.gs에서
 *            실전 차트 조회 시 edKis_ 접두어 함수들을 직접 참조.
 *            테스트 전용 파일 삭제 후에도 참조가 끊기지 않도록
 *            이 파일 하단으로 이관.
 *
 * 포함 함수:
 *   edKis_fetchDomesticChart_     - KIS 국내주식 기간별 시세 조회
 *   edKis_getAccessToken_         - 액세스 토큰 반환 (캐시 우선)
 *   edKis_issueAccessToken_       - KIS OAuth 토큰 신규 발급
 *   edKis_normalizeChartRow_      - KIS 응답 행 정규화
 *   edKis_getBaseUrl_             - ScriptProperty 기반 기본 URL 반환
 *   edKis_clearToken_             - 저장된 토큰 삭제
 *   edKis_normalizeTicker_        - 종목코드 6자리 정규화
 *   edKis_toQueryString_          - 쿼리스트링 직렬화
 *   edKis_formatDate_             - Date → yyyyMMdd 문자열
 *   edKis_addDays_                - 날짜 덧셈
 *   edKis_periodLookbackDays_     - 기간별 조회 일수 계산
 *   edKis_num_                    - 숫자 정규화
 *   edKis_mask_                   - 민감정보 마스킹
 *******************************************************/

// KIS API 엔드포인트 및 속성 키 상수
const ED_KIS_HELPER = {
  baseUrl: "https://openapi.koreainvestment.com:9443",
  chartPath: "/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice",
  tokenPath: "/oauth2/tokenP",
  chartTrId: "FHKST03010100",
  properties: {
    appKey: "KIS_APP_KEY",
    appSecret: "KIS_APP_SECRET",
    accessToken: "KIS_ACCESS_TOKEN",
    accessTokenExpiredAt: "KIS_ACCESS_TOKEN_EXPIRED_AT",
    baseUrl: "KIS_BASE_URL",
  },
};

/**
 * KIS 국내주식 기간별 시세(일/주/월봉) 조회
 * @param {string} ticker - 종목코드 (6자리)
 * @param {string} periodDivCode - 기간구분코드 (D/W/M)
 * @param {string} startDate - 조회 시작일 (yyyyMMdd)
 * @param {string} endDate   - 조회 종료일 (yyyyMMdd)
 * @returns {{ raw: Object, items: Array }}
 */
function edKis_fetchDomesticChart_(ticker, periodDivCode, startDate, endDate) {
  const token = edKis_getAccessToken_();
  const props = PropertiesService.getScriptProperties();
  const appKey = props.getProperty(ED_KIS_HELPER.properties.appKey);
  const appSecret = props.getProperty(ED_KIS_HELPER.properties.appSecret);

  if (!appKey || !appSecret) {
    throw new Error("KIS appkey/appsecret이 없습니다. setupKisCredentials()를 먼저 실행하세요.");
  }

  const query = {
    FID_COND_MRKT_DIV_CODE: "J",
    FID_INPUT_ISCD: edKis_normalizeTicker_(ticker),
    FID_INPUT_DATE_1: startDate,
    FID_INPUT_DATE_2: endDate,
    FID_PERIOD_DIV_CODE: periodDivCode,
    FID_ORG_ADJ_PRC: "0",
  };

  const url = edKis_getBaseUrl_() + ED_KIS_HELPER.chartPath + "?" + edKis_toQueryString_(query);

  const res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      authorization: "Bearer " + token,
      appkey: appKey,
      appsecret: appSecret,
      tr_id: ED_KIS_HELPER.chartTrId,
      custtype: "P",
    },
  });

  const status = res.getResponseCode();
  const text = res.getContentText();
  let json;

  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error(`KIS chart JSON parse 실패. HTTP=${status}, body=${text.slice(0, 500)}`);
  }

  if (status < 200 || status >= 300) {
    throw new Error(`KIS chart HTTP 오류. HTTP=${status}, body=${text.slice(0, 1000)}`);
  }

  if (String(json.rt_cd || "") !== "0") {
    throw new Error(`KIS chart API 오류. rt_cd=${json.rt_cd}, msg_cd=${json.msg_cd}, msg=${json.msg1}, body=${text.slice(0, 500)}`);
  }

  const rawItems = Array.isArray(json.output2) ? json.output2 : [];
  const items = rawItems
    .map((row) => edKis_normalizeChartRow_(row))
    .filter((row) => row.date);

  // 날짜 오름차순 정렬
  items.sort((a, b) => String(a.date).localeCompare(String(b.date)));

  return { raw: json, items };
}

/**
 * 액세스 토큰 반환 (캐시된 토큰이 유효하면 재사용, 만료 10분 전이면 신규 발급)
 */
function edKis_getAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const savedToken = props.getProperty(ED_KIS_HELPER.properties.accessToken);
  const expiredAtText = props.getProperty(ED_KIS_HELPER.properties.accessTokenExpiredAt);

  if (savedToken && expiredAtText) {
    const expiredAt = new Date(expiredAtText);
    if (expiredAt.getTime() - Date.now() > 10 * 60 * 1000) {
      return savedToken;
    }
  }

  return edKis_issueAccessToken_();
}

/**
 * KIS OAuth 토큰 신규 발급 후 ScriptProperties에 저장
 */
function edKis_issueAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const appKey = props.getProperty(ED_KIS_HELPER.properties.appKey);
  const appSecret = props.getProperty(ED_KIS_HELPER.properties.appSecret);

  if (!appKey || !appSecret) {
    throw new Error("KIS appkey/appsecret이 없습니다. setupKisCredentials()를 먼저 실행하세요.");
  }

  const url = edKis_getBaseUrl_() + ED_KIS_HELPER.tokenPath;
  const payload = {
    grant_type: "client_credentials",
    appkey: appKey,
    appsecret: appSecret,
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify(payload),
  });

  const status = res.getResponseCode();
  const text = res.getContentText();
  let json;

  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error(`KIS token JSON parse 실패. HTTP=${status}, body=${text.slice(0, 500)}`);
  }

  if (status < 200 || status >= 300 || !json.access_token) {
    throw new Error(`KIS token 발급 실패. HTTP=${status}, body=${text.slice(0, 1000)}`);
  }

  const expiresIn = Number(json.expires_in || 86400);
  const expiredAt = new Date(Date.now() + expiresIn * 1000);

  props.setProperty(ED_KIS_HELPER.properties.accessToken, json.access_token);
  props.setProperty(ED_KIS_HELPER.properties.accessTokenExpiredAt, expiredAt.toISOString());

  return json.access_token;
}

/**
 * KIS API 응답의 개별 캔들 행을 정규화하여 공통 형식으로 변환
 */
function edKis_normalizeChartRow_(row) {
  const date = row.stck_bsop_date || row.date || "";

  return {
    date,
    open: edKis_num_(row.stck_oprc || row.open),
    high: edKis_num_(row.stck_hgpr || row.high),
    low: edKis_num_(row.stck_lwpr || row.low),
    close: edKis_num_(row.stck_clpr || row.close),
    volume: edKis_num_(row.acml_vol || row.volume),
    rawKeys: Object.keys(row).join(","),
  };
}

/**
 * ScriptProperty에 저장된 KIS 기본 URL 반환 (없으면 기본값 사용)
 */
function edKis_getBaseUrl_() {
  return (
    PropertiesService.getScriptProperties().getProperty(ED_KIS_HELPER.properties.baseUrl) ||
    ED_KIS_HELPER.baseUrl
  );
}

/**
 * 저장된 KIS 액세스 토큰 삭제 (도메인 전환 시 호출)
 */
function edKis_clearToken_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(ED_KIS_HELPER.properties.accessToken);
  props.deleteProperty(ED_KIS_HELPER.properties.accessTokenExpiredAt);
}

/**
 * 종목코드를 6자리로 정규화 (숫자만으로 이루어진 짧은 코드는 앞에 0 패딩)
 */
function edKis_normalizeTicker_(ticker) {
  const t = String(ticker || "").trim().toUpperCase();
  if (/^\d+$/.test(t) && t.length < 6) return t.padStart(6, "0");
  return t;
}

/**
 * 쿼리스트링 직렬화 (URL 인코딩 포함)
 */
function edKis_toQueryString_(obj) {
  return Object.keys(obj)
    .map((key) => encodeURIComponent(key) + "=" + encodeURIComponent(obj[key]))
    .join("&");
}

/**
 * Date 객체를 KIS API 날짜 형식(yyyyMMdd)으로 변환
 */
function edKis_formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd");
}

/**
 * Date에 일수를 더한 새 Date 반환
 */
function edKis_addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

/**
 * 기간구분코드(D/W/M)와 조회 캔들 수를 기반으로 조회 필요 일수 계산
 */
function edKis_periodLookbackDays_(periodDivCode, count) {
  if (periodDivCode === "M") return Math.max(365 * 5, count * 31);
  if (periodDivCode === "W") return Math.max(365 * 2, count * 7);
  return Math.max(365, count * 2);
}

/**
 * 문자열/숫자 값을 안전하게 숫자로 변환 (쉼표, 공백 제거)
 */
function edKis_num_(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  const n = Number(String(value).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

/**
 * 민감정보 마스킹 (앞 4자리 + **** + 뒤 4자리)
 */
function edKis_mask_(value) {
  const text = String(value || "");
  if (!text) return "";
  if (text.length <= 8) return "********";
  return text.slice(0, 4) + "****" + text.slice(-4);
}