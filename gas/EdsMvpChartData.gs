/*******************************************************
 * ED's MVP - Chart Data Cache + KIS Integration v0.8.7
 *
 * 목적:
 * - 한국투자증권 Open API로 국내 ETF/주식 일/주/月 차트 데이터를 조회
 * - App_ChartPrices 시트에 캐시
 * - PWA가 getChartData / refreshChartData 형태로 사용할 수 있는 데이터 구조 제공
 *
 * 전제:
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
 * 특정 asset_id의 일/주/월봉(D/W/M) 3가지 차트 데이터를 한 번에 조회하여 반환
 * payload: { asset_id, limitD, limitW, limitM }
 */
function fetchSingleChartData(payload) {
  const assetId = String(payload && payload.asset_id || "").trim();
  if (!assetId) throw new Error("asset_id가 필요합니다.");

  const limitD = Number(payload && payload.limitD || ED_MVP_CHART.defaultLimit["D"] || 120);
  const limitW = Number(payload && payload.limitW || ED_MVP_CHART.defaultLimit["W"] || 104);
  const limitM = Number(payload && payload.limitM || ED_MVP_CHART.defaultLimit["M"] || 60);
  const force = Boolean(payload && payload.force);

  const allRows = edChart_readSheetAsObjects_(ED_MVP_CHART.sheets.chartPrices)
    .filter((row) => String(row.asset_id || "") === assetId)
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
  const mapToItem = (row) => ({
    date: row.date,
    open: edChart_num_(row.open),
    high: edChart_num_(row.high),
    low: edChart_num_(row.low),
    close: edChart_num_(row.close),
    volume: edChart_num_(row.volume),
    source: row.source,
    fetched_at: row.fetched_at,
  });
  let rowsD = allRows.filter((r) => String(r.interval || "") === "D").slice(-limitD).map(mapToItem);
  let rowsW = allRows.filter((r) => String(r.interval || "") === "W").slice(-limitW).map(mapToItem);
  let rowsM = allRows.filter((r) => String(r.interval || "") === "M").slice(-limitM).map(mapToItem);

  const market = edChart_marketFromAssetId_(assetId);
  const ticker = edChart_normalizeTicker_(edChart_tickerFromAssetId_(assetId));
  const isKrx = (market === "KRX" && /^[0-9A-Z]{6}$/.test(ticker));

  const intervalsToFetch = [];
  if (isKrx && force) {
    intervalsToFetch.push("D", "W", "M");
  }

  if (intervalsToFetch.length > 0) {
    const fetchResults = {};
    const today = new Date();
    const end = edKis_formatDate_(today);
    const token = edKis_getAccessToken_();
    const props = PropertiesService.getScriptProperties();
    const appKey = props.getProperty(ED_KIS_HELPER.properties.appKey);
    const appSecret = props.getProperty(ED_KIS_HELPER.properties.appSecret);
    if (!appKey || !appSecret) {
      throw new Error("KIS appkey/appsecret이 없습니다. setupKisCredentials()를 먼저 실행하세요.");
    }

    try {
      const requests = intervalsToFetch.map((intv) => {
        const limit = intv === "D" ? limitD : (intv === "W" ? limitW : limitM);
        const start = edKis_formatDate_(edKis_addDays_(today, -edKis_periodLookbackDays_(intv, limit)));
        const query = {
          FID_COND_MRKT_DIV_CODE: "J",
          FID_INPUT_ISCD: ticker,
          FID_INPUT_DATE_1: start,
          FID_INPUT_DATE_2: end,
          FID_PERIOD_DIV_CODE: intv,
          FID_ORG_ADJ_PRC: "0",
        };
        const url = edKis_getBaseUrl_() + ED_KIS_HELPER.chartPath + "?" + edKis_toQueryString_(query);
        return {
          url: url,
          method: "get",
          muteHttpExceptions: true,
          headers: {
            authorization: "Bearer " + token,
            appkey: appKey,
            appsecret: appSecret,
            tr_id: ED_KIS_HELPER.chartTrId,
            custtype: "P",
          }
        };
      });

      const responses = UrlFetchApp.fetchAll(requests);
      intervalsToFetch.forEach((intv, index) => {
        const res = responses[index];
        const status = res.getResponseCode();
        const text = res.getContentText();
        let json;
        try {
          json = JSON.parse(text);
        } catch (e) {
          throw new Error(`JSON parsing failed. HTTP=${status}`);
        }
        if (status < 200 || status >= 300 || String(json.rt_cd || "") !== "0") {
          throw new Error(`API error. rt_cd=${json.rt_cd}, msg=${json.msg1}`);
        }
        const rawItems = Array.isArray(json.output2) ? json.output2 : [];
        const items = rawItems
          .map((row) => edKis_normalizeChartRow_(row))
          .filter((row) => row.date);
        items.sort((a, b) => String(a.date).localeCompare(String(b.date)));
        
        const limit = intv === "D" ? limitD : (intv === "W" ? limitW : limitM);
        fetchResults[intv] = items.slice(-limit);
      });
    } catch (parallelError) {
      Logger.log(`[fetchSingleChartData] Parallel fetch failed: ${parallelError.message}. Falling back to sequential...`);
      intervalsToFetch.forEach((intv) => {
        const limit = intv === "D" ? limitD : (intv === "W" ? limitW : limitM);
        const start = edKis_formatDate_(edKis_addDays_(today, -edKis_periodLookbackDays_(intv, limit)));
        const result = edChart_fetchKisChartWithRetry_(ticker, intv, start, end);
        const items = result.items.slice(-limit);
        fetchResults[intv] = items;
        Utilities.sleep(500);
      });
    }

    edChart_replaceChartRowsMultiple_(assetId, ticker, fetchResults, ED_MVP_CHART.source.kis);
    const now = new Date();
    const nowStr = edChart_normalizeValue_(now);
    const mapFetchedToItem = (item) => ({
      date: item.date,
      open: item.open,
      high: item.high,
      low: item.low,
      close: item.close,
      volume: item.volume,
      source: ED_MVP_CHART.source.kis,
      fetched_at: nowStr,
    });
    if (fetchResults["D"]) rowsD = fetchResults["D"].map(mapFetchedToItem);
    if (fetchResults["W"]) rowsW = fetchResults["W"].map(mapFetchedToItem);
    if (fetchResults["M"]) rowsM = fetchResults["M"].map(mapFetchedToItem);
  }

  return {
    asset_id: assetId,
    D: { interval: "D", count: rowsD.length, items: rowsD },
    W: { interval: "W", count: rowsW.length, items: rowsW },
    M: { interval: "M", count: rowsM.length, items: rowsM },
  };
}

/**
 * 복수 인터벌의 차트 데이터를 단일 트랜잭션으로 시트에 벌크 치환하여 캐시 저장
 */
function edChart_replaceChartRowsMultiple_(assetId, ticker, intervalItemsMap, source) {
  setupChartPricesSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_CHART.sheets.chartPrices);
  if (!sheet) throw new Error("App_ChartPrices 시트를 찾을 수 없습니다.");
  const lastRow = sheet.getLastRow();
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
  let keptRows = [];
  const intervalsToReplace = Object.keys(intervalItemsMap);

  if (lastRow >= 2) {
    const values = sheet.getRange(1, 1, lastRow, headers.length).getValues();
    const existingHeader = values[0];
    const assetCol = existingHeader.indexOf("asset_id");
    const intervalCol = existingHeader.indexOf("interval");
    if (assetCol < 0 || intervalCol < 0) {
      throw new Error("App_ChartPrices 헤더 오류: asset_id 또는 interval 열이 없습니다.");
    }

    keptRows = values.slice(1).filter((row) => {
      const rowAssetId = String(row[assetCol] || "");
      const rowInterval = String(row[intervalCol] || "");
      if (rowAssetId === assetId && intervalsToReplace.indexOf(rowInterval) >= 0) {
        return false;
      }
      return true;
    });
  }

  const now = new Date();
  const newRows = [];
  intervalsToReplace.forEach((interval) => {
    const items = intervalItemsMap[interval] || [];
    items.forEach((item) => {
      newRows.push([
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
    });
  });
  const output = [headers].concat(keptRows).concat(newRows);
  sheet.clear();
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);
}

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
  if (force) {
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
      if (message.indexOf("EGW00201") >= 0 || message.indexOf("초당 거래건수") >= 0) {
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
  sheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold").setBackground("#137333").setFontColor("#ffffff");
}

/*******************************************************
 * ED's MVP - KIS API Core Helpers
 *******************************************************/

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
  items.sort((a, b) => String(a.date).localeCompare(String(b.date)));

  return { raw: json, items };
}

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

function edKis_getBaseUrl_() {
  return (
    PropertiesService.getScriptProperties().getProperty(ED_KIS_HELPER.properties.baseUrl) ||
    ED_KIS_HELPER.baseUrl
  );
}

function edKis_clearToken_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(ED_KIS_HELPER.properties.accessToken);
  props.deleteProperty(ED_KIS_HELPER.properties.accessTokenExpiredAt);
}

function edKis_normalizeTicker_(ticker) {
  const t = String(ticker || "").trim().toUpperCase();
  if (/^\d+$/.test(t) && t.length < 6) return t.padStart(6, "0");
  return t;
}

function edKis_toQueryString_(obj) {
  return Object.keys(obj)
    .map((key) => encodeURIComponent(key) + "=" + encodeURIComponent(obj[key]))
    .join("&");
}

function edKis_formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd");
}

function edKis_addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function edKis_periodLookbackDays_(periodDivCode, count) {
  if (periodDivCode === "M") return Math.max(365 * 5, count * 31);
  if (periodDivCode === "W") return Math.max(365 * 2, count * 7);
  return Math.max(365, count * 2);
}

function edKis_num_(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  const n = Number(String(value).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function edKis_mask_(value) {
  const text = String(value || "");
  if (!text) return "";
  if (text.length <= 8) return "********";
  return text.slice(0, 4) + "****" + text.slice(-4);
}