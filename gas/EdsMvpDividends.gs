/*******************************************************
 * ED's MVP - Dividends Import + Dashboard v0.8.61
 *
 * 목적:
 * - 기존 원장 "6. 배당내역"을 기준 데이터로 사용
 * - "6. 배당내역" → App_Dividends 이관
 * - 계좌/종목/보유상태별 배당 대시보드 생성
 * - PWA v0.7.1의 getDividendDashboard / refreshDividends API 지원
 *
 * 적용 순서:
 * 1) Apps Script에 EdsMvpDividends.gs 파일 생성 또는 기존 파일 전체 교체
 * 2) 이 코드 전체 붙여넣기
 * 3) resetAndImportDividendSheetToApp() 실행
 * 4) EdsMvpApi.gs switch(action)에 dividend case 추가
 * 5) Apps Script 웹앱 새 버전 배포
 *******************************************************/

const ED_MVP_DIVIDENDS = {
  sourceSheetNameCandidates: [
    "6. 배당내역",
    "6.배당내역",
    "배당내역",
  ],

  scanHeaderRows: 30,
  minHeaderScore: 5,

  sheets: {
    dividends: "App_Dividends",
    holdings: "App_Holdings",
    accounts: "App_Accounts",
    output: "App_Output",
    syncLog: "App_SyncLog",
  },
};

function setupDividendsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ED_MVP_DIVIDENDS.sheets.dividends);
  if (!sheet) sheet = ss.insertSheet(ED_MVP_DIVIDENDS.sheets.dividends);

  const headers = edDiv_headers_();
  const first = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0]
    : [];

  if (String(first[0] || "") !== "dividend_id") {
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

  SpreadsheetApp.getActiveSpreadsheet().toast("App_Dividends 시트 정비 완료", "ED's MVP", 5);
}

function resetAndImportDividendSheetToApp() {
  setupDividendsSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet = ss.getSheetByName(ED_MVP_DIVIDENDS.sheets.dividends);
  const headers = edDiv_headers_();

  if (appSheet.getLastRow() > 1) {
    appSheet.getRange(2, 1, appSheet.getLastRow() - 1, appSheet.getLastColumn()).clearContent();
  }

  const result = importDividendSheetToApp();

  edDiv_writeLog_(
    "dividend_import",
    "success",
    result.source_sheet,
    ED_MVP_DIVIDENDS.sheets.dividends,
    "",
    `배당내역 이관 완료: imported=${result.imported_count}, skipped=${result.skipped_count}`
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `배당내역 이관 완료: ${result.imported_count}건`,
    "ED's MVP",
    8
  );

  return result;
}

function importDividendSheetToApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = edDiv_findSourceSheet_(ss);
  if (!sourceSheet) throw new Error("기존 원장 '6. 배당내역' 시트를 찾을 수 없습니다.");

  const headerInfo = edDiv_detectHeader_(sourceSheet);
  if (!headerInfo) throw new Error("6. 배당내역 헤더를 찾을 수 없습니다.");

  const appSheet = ss.getSheetByName(ED_MVP_DIVIDENDS.sheets.dividends);
  if (!appSheet) throw new Error("App_Dividends 시트를 찾을 수 없습니다. setupDividendsSheet()를 먼저 실행하세요.");

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  const dataStartRow = headerInfo.headerRow + 1;
  const rowCount = Math.max(lastRow - dataStartRow + 1, 0);

  if (rowCount <= 0) {
    return {
      source_sheet: sourceSheet.getName(),
      imported_count: 0,
      skipped_count: 0,
      message: "데이터 행 없음",
    };
  }

  const values = sourceSheet.getRange(dataStartRow, 1, rowCount, lastCol).getValues();
  const displayValues = sourceSheet.getRange(dataStartRow, 1, rowCount, lastCol).getDisplayValues();

  const currentIndex = edDiv_buildCurrentHoldingIndex_();
  const rows = [];
  let skipped = 0;
  const now = new Date();

  for (let i = 0; i < values.length; i++) {
    const sourceRow = dataStartRow + i;
    const raw = values[i];
    const display = displayValues[i];

    const dividendDate = edDiv_normalizeDate_(edDiv_getCellRaw_(raw, display, headerInfo.colMap.date));
    const broker = edDiv_getCell_(raw, display, headerInfo.colMap.broker);
    const accountName = edDiv_getCell_(raw, display, headerInfo.colMap.accountName);
    const ticker = edDiv_normalizeTicker_(edDiv_getCell_(raw, display, headerInfo.colMap.ticker));
    const assetName = edDiv_getCell_(raw, display, headerInfo.colMap.assetName);
    const krwAmount = edDiv_parseNumber_(edDiv_getCellRaw_(raw, display, headerInfo.colMap.krwAmount));
    const foreignAmount = edDiv_parseNumber_(edDiv_getCellRaw_(raw, display, headerInfo.colMap.foreignAmount));
    const netAmountKrw = edDiv_parseNumber_(edDiv_getCellRaw_(raw, display, headerInfo.colMap.netAmountKrw));
    const memo = edDiv_getCell_(raw, display, headerInfo.colMap.memo);

    if (!dividendDate || !ticker || !assetName || !broker || !accountName || netAmountKrw === 0) {
      skipped++;
      continue;
    }

    const accountType = edDiv_inferAccountType_(accountName);
    const accountId = edDiv_makeAccountId_(broker, accountName, accountType);
    const market = edDiv_inferMarket_(ticker);
    const assetId = edDiv_makeAssetId_(market, ticker);
    const year = Number(dividendDate.slice(0, 4)) || 0;
    const month = Number(dividendDate.slice(5, 7)) || 0;
    const day = Number(dividendDate.slice(8, 10)) || 0;
    const currency = foreignAmount > 0 ? "USD" : "KRW";

    let holdingStatus = "sold";
    let currentHoldingId = "";

    const sameAccountKey = `${accountId}__${assetId}`;
    if (currentIndex.byAccountAsset.has(sameAccountKey)) {
      holdingStatus = "current";
      currentHoldingId = currentIndex.byAccountAsset.get(sameAccountKey).holding_id || "";
    } else if (currentIndex.byAsset.has(assetId)) {
      holdingStatus = "other_account_current";
      currentHoldingId = currentIndex.byAsset.get(assetId).holding_id || "";
    }

    const dividendId = `DIV_${sourceSheet.getSheetId()}_${sourceRow}`;

    rows.push([
      dividendId,
      sourceSheet.getName(),
      sourceRow,
      dividendDate,
      year,
      month,
      day,
      broker,
      accountName,
      accountType,
      accountId,
      ticker,
      assetName,
      market,
      assetId,
      krwAmount,
      foreignAmount,
      netAmountKrw,
      currency,
      holdingStatus,
      currentHoldingId,
      memo,
      "TRUE",
      now,
      now,
    ]);
  }

  if (rows.length > 0) {
    appSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  return {
    source_sheet: sourceSheet.getName(),
    header_row: headerInfo.headerRow,
    data_start_row: dataStartRow,
    imported_count: rows.length,
    skipped_count: skipped,
    fetched_at: new Date(),
  };
}

function getDividendDashboard() {
  setupDividendsSheet();

  const records = edDiv_readSheetAsObjects_(ED_MVP_DIVIDENDS.sheets.dividends)
    .filter((row) => String(row.enabled || "TRUE").toUpperCase() === "TRUE")
    .map((row) => edDiv_normalizeRecord_(row));

  const now = new Date();
  const thisYear = now.getFullYear();
  const lastYear = thisYear - 1;

  const summary = edDiv_makeSummary_(records, thisYear, lastYear);
  const byAsset = edDiv_groupByAsset_(records, thisYear, lastYear);
  const byAccount = edDiv_groupByAccount_(records, thisYear, lastYear);
  const monthly = edDiv_groupByMonth_(records);
  const recent = records
    .slice()
    .sort((a, b) => String(b.dividend_date).localeCompare(String(a.dividend_date)))
    .slice(0, 30);

  return {
    summary,
    by_asset: byAsset,
    by_account: byAccount,
    monthly,
    recent,
    records: records.slice().sort((a, b) => String(b.dividend_date).localeCompare(String(a.dividend_date))),
  };
}

function edDiv_headers_() {
  return [
    "dividend_id",
    "source_sheet",
    "source_row",
    "dividend_date",
    "year",
    "month",
    "day",
    "broker",
    "account_name",
    "account_type",
    "account_id",
    "ticker",
    "asset_name",
    "market",
    "asset_id",
    "krw_amount",
    "foreign_amount",
    "net_amount_krw",
    "currency",
    "holding_status",
    "current_holding_id",
    "memo",
    "enabled",
    "created_at",
    "updated_at",
  ];
}

function edDiv_findSourceSheet_(ss) {
  for (const name of ED_MVP_DIVIDENDS.sourceSheetNameCandidates) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  return null;
}

function edDiv_detectHeader_(sheet) {
  const scanRows = Math.min(sheet.getLastRow(), ED_MVP_DIVIDENDS.scanHeaderRows);
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();

  let best = null;

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const colMap = edDiv_buildColMap_(row);
    const score = Object.keys(colMap).filter((key) => colMap[key] >= 0).length;
    if (!best || score > best.score) {
      best = { headerRow: r + 1, colMap, score };
    }
  }

  if (!best || best.score < ED_MVP_DIVIDENDS.minHeaderScore) return null;
  return best;
}

function edDiv_buildColMap_(row) {
  const headers = row.map((v) => edDiv_normalizeHeader_(v));

  function find(patterns, excludes) {
    excludes = excludes || [];
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (!h) continue;
      if (excludes.some((ex) => h.indexOf(ex) >= 0)) continue;
      if (patterns.some((p) => h === p || h.indexOf(p) >= 0)) return i;
    }
    return -1;
  }

  return {
    date: find(["일자", "배당일", "지급일"]),
    year: find(["연도"]),
    month: find(["월"]),
    day: find(["일"], ["일자"]),
    broker: find(["증권사"]),
    accountName: find(["계좌", "계좌명"], ["계좌형식", "계좌유형"]),
    ticker: find(["종목코드", "티커"]),
    assetName: find(["종목명"]),
    krwAmount: find(["원화배당금", "원화"]),
    foreignAmount: find(["외화배당금", "외화", "달러"]),
    netAmountKrw: find(["원화환산", "최종원화환산값", "환산"]),
    memo: find(["메모", "비고"]),
  };
}

function edDiv_normalizeHeader_(value) {
  return String(value || "")
    .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
    .replace(/\s/g, "")
    .replace(/[()\[\]{}<>]/g, "")
    .trim();
}

function edDiv_buildCurrentHoldingIndex_() {
  const holdings = edDiv_readSheetAsObjects_(ED_MVP_DIVIDENDS.sheets.holdings)
    .filter((row) => String(row.enabled || "").toUpperCase() === "TRUE");

  const byAccountAsset = new Map();
  const byAsset = new Map();

  holdings.forEach((row) => {
    const accountId = String(row.account_id || "");
    const assetId = String(row.asset_id || "");
    if (!accountId || !assetId) return;
    byAccountAsset.set(`${accountId}__${assetId}`, row);
    if (!byAsset.has(assetId)) byAsset.set(assetId, row);
  });

  return { byAccountAsset, byAsset };
}

function edDiv_makeSummary_(records, thisYear, lastYear) {
  const totalNet = edDiv_sum_(records, "net_amount_krw");
  const thisYearNet = edDiv_sum_(records.filter((r) => r.year === thisYear), "net_amount_krw");
  const lastYearNet = edDiv_sum_(records.filter((r) => r.year === lastYear), "net_amount_krw");
  const currentNet = edDiv_sum_(records.filter((r) => r.holding_status === "current"), "net_amount_krw");
  const otherAccountNet = edDiv_sum_(records.filter((r) => r.holding_status === "other_account_current"), "net_amount_krw");
  const soldNet = edDiv_sum_(records.filter((r) => r.holding_status === "sold"), "net_amount_krw");
  const recent12mNet = edDiv_sumRecentMonths_(records, 12);

  return {
    total_gross: totalNet,
    total_tax: 0,
    total_net: totalNet,
    total_net_krw: totalNet,
    this_year_net: thisYearNet,
    last_year_net: lastYearNet,
    recent_12m_net: recent12mNet,
    monthly_average_12m: recent12mNet / 12,
    current_holdings_net: currentNet,
    other_account_holdings_net: otherAccountNet,
    sold_holdings_net: soldNet,
    yoy_growth_rate: lastYearNet > 0 ? (thisYearNet - lastYearNet) / lastYearNet : 0,
    record_count: records.length,
  };
}

function edDiv_groupByAsset_(records, thisYear, lastYear) {
  const map = new Map();
  records.forEach((row) => {
    const key = `${row.account_id}__${row.asset_id}`;
    if (!map.has(key)) {
      map.set(key, {
        account_id: row.account_id,
        account_name: row.account_name,
        broker: row.broker,
        account_type: row.account_type,
        asset_id: row.asset_id,
        ticker: row.ticker,
        asset_name: row.asset_name,
        holding_status: row.holding_status,
        total_gross: 0,
        total_tax: 0,
        total_net: 0,
        total_net_krw: 0,
        this_year_net: 0,
        last_year_net: 0,
        yoy_growth_rate: 0,
        record_count: 0,
        latest_date: "",
      });
    }
    const item = map.get(key);
    const amount = Number(row.net_amount_krw || 0);
    item.total_gross += amount;
    item.total_net += amount;
    item.total_net_krw += amount;
    item.record_count += 1;
    if (!item.latest_date || row.dividend_date > item.latest_date) item.latest_date = row.dividend_date;
    if (row.year === thisYear) item.this_year_net += amount;
    if (row.year === lastYear) item.last_year_net += amount;
  });

  return Array.from(map.values()).map((item) => {
    item.yoy_growth_rate = item.last_year_net > 0 ? (item.this_year_net - item.last_year_net) / item.last_year_net : 0;
    return item;
  }).sort((a, b) => b.total_net - a.total_net);
}

function edDiv_groupByAccount_(records, thisYear, lastYear) {
  const map = new Map();
  records.forEach((row) => {
    const key = row.account_id;
    if (!map.has(key)) {
      map.set(key, {
        account_id: row.account_id,
        account_name: row.account_name,
        broker: row.broker,
        account_type: row.account_type,
        total_net: 0,
        total_net_krw: 0,
        this_year_net: 0,
        last_year_net: 0,
        yoy_growth_rate: 0,
        record_count: 0,
      });
    }
    const item = map.get(key);
    const amount = Number(row.net_amount_krw || 0);
    item.total_net += amount;
    item.total_net_krw += amount;
    item.record_count += 1;
    if (row.year === thisYear) item.this_year_net += amount;
    if (row.year === lastYear) item.last_year_net += amount;
  });

  return Array.from(map.values()).map((item) => {
    item.yoy_growth_rate = item.last_year_net > 0 ? (item.this_year_net - item.last_year_net) / item.last_year_net : 0;
    return item;
  }).sort((a, b) => b.total_net - a.total_net);
}

function edDiv_groupByMonth_(records) {
  const map = new Map();
  records.forEach((row) => {
    const key = `${row.year}-${String(row.month).padStart(2, "0")}`;
    if (!map.has(key)) {
      map.set(key, {
        month_key: key,
        year: row.year,
        month: row.month,
        total_net: 0,
        current_net: 0,
        other_account_net: 0,
        sold_net: 0,
        record_count: 0,
      });
    }
    const item = map.get(key);
    const amount = Number(row.net_amount_krw || 0);
    item.total_net += amount;
    item.record_count += 1;
    if (row.holding_status === "current") item.current_net += amount;
    else if (row.holding_status === "other_account_current") item.other_account_net += amount;
    else item.sold_net += amount;
  });

  return Array.from(map.values()).sort((a, b) => String(a.month_key).localeCompare(String(b.month_key)));
}

function edDiv_sumRecentMonths_(records, months) {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - months + 1, 1);
  return records
    .filter((row) => {
      const d = edDiv_dateObject_(row.dividend_date);
      return d && d >= start;
    })
    .reduce((sum, row) => sum + Number(row.net_amount_krw || 0), 0);
}

function edDiv_normalizeRecord_(row) {
  return {
    dividend_id: row.dividend_id || "",
    source_sheet: row.source_sheet || "",
    source_row: Number(row.source_row || 0),
    dividend_date: edDiv_normalizeDate_(row.dividend_date),
    year: Number(row.year || String(row.dividend_date || "").slice(0, 4) || 0),
    month: Number(row.month || 0),
    day: Number(row.day || 0),
    broker: row.broker || "",
    account_name: row.account_name || "",
    account_type: row.account_type || "",
    account_id: row.account_id || "",
    ticker: row.ticker || "",
    asset_name: row.asset_name || "",
    market: row.market || "",
    asset_id: row.asset_id || "",
    krw_amount: edDiv_parseNumber_(row.krw_amount),
    foreign_amount: edDiv_parseNumber_(row.foreign_amount),
    net_amount: edDiv_parseNumber_(row.net_amount_krw),
    net_amount_krw: edDiv_parseNumber_(row.net_amount_krw),
    gross_amount: edDiv_parseNumber_(row.net_amount_krw),
    tax_amount: 0,
    currency: row.currency || "KRW",
    holding_status: row.holding_status || "sold",
    current_holding_id: row.current_holding_id || "",
    memo: row.memo || "",
    enabled: row.enabled || "TRUE",
    created_at: row.created_at || "",
    updated_at: row.updated_at || "",
  };
}

function edDiv_getCell_(row, displayRow, colIndex) {
  if (colIndex === undefined || colIndex < 0) return "";
  return String(displayRow[colIndex] || row[colIndex] || "").trim();
}

function edDiv_getCellRaw_(row, displayRow, colIndex) {
  if (colIndex === undefined || colIndex < 0) return "";
  if (row[colIndex] !== "" && row[colIndex] !== null && row[colIndex] !== undefined) return row[colIndex];
  return displayRow[colIndex];
}

function edDiv_readSheetAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = edDiv_normalizeValue_(row[index]);
    });
    return obj;
  });
}

function edDiv_normalizeValue_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return value;
}

function edDiv_normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const text = String(value || "").trim();
  if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(text)) {
    const parts = text.split("/");
    return `${parts[0]}-${parts[1].padStart(2, "0")}-${parts[2].padStart(2, "0")}`;
  }
  if (/^\d{8}$/.test(text)) return `${text.slice(0, 4)}-${text.slice(4, 6)}-${text.slice(6, 8)}`;
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
    const parts = text.split("-");
    return `${parts[0]}-${parts[1].padStart(2, "0")}-${parts[2].padStart(2, "0")}`;
  }
  return text;
}

function edDiv_dateObject_(dateValue) {
  const text = edDiv_normalizeDate_(dateValue);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  return new Date(Number(text.slice(0, 4)), Number(text.slice(5, 7)) - 1, Number(text.slice(8, 10)));
}

function edDiv_parseNumber_(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  const text = String(value)
    .replace(/[,$₩￥€£\s]/g, "")
    .replace(/원/g, "")
    .replace(/\+/g, "")
    .replace(/−/g, "-")
    .trim();
  const num = Number(text);
  return isNaN(num) ? 0 : num;
}

function edDiv_normalizeTicker_(value) {
  const t = String(value || "").trim().toUpperCase().replace(/\s/g, "");
  if (/^\d+$/.test(t) && t.length < 6) return t.padStart(6, "0");
  return t;
}

function edDiv_inferAccountType_(accountName) {
  const text = String(accountName || "").trim();
  if (text.indexOf("ISA") >= 0 || text.indexOf("isa") >= 0) return "ISA";
  if (text.indexOf("IRP") >= 0 || text.indexOf("irp") >= 0) return "IRP";
  if (text.indexOf("연금") >= 0) return "개인연금";
  if (text.indexOf("일반") >= 0) return "일반";
  return "기타";
}

function edDiv_inferMarket_(ticker) {
  const t = String(ticker || "").toUpperCase();
  if (/^[0-9A-Z]{6}$/.test(t) && /\d/.test(t)) return "KRX";
  if (["QQQ", "QQQM", "TQQQ", "ONEQ"].indexOf(t) >= 0) return "NASDAQ";
  if (["SPY", "VOO", "VTI", "SCHD", "DIA"].indexOf(t) >= 0) return "NYSE";
  if (["SPYM"].indexOf(t) >= 0) return "AMEX";
  return "기타";
}

function edDiv_makeAccountId_(broker, accountName, accountType) {
  return "ACC_" + edDiv_slug_(`${broker}_${accountName}_${accountType}`);
}

function edDiv_makeAssetId_(market, ticker) {
  return `${market}_${edDiv_slug_(ticker)}`;
}

function edDiv_slug_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^\w가-힣]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();
}

function edDiv_sum_(rows, field) {
  return rows.reduce((sum, row) => sum + Number(row[field] || 0), 0);
}

function edDiv_writeLog_(syncType, status, sourceSheet, targetSheet, targetKey, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_DIVIDENDS.sheets.syncLog);
  if (!sheet) return;
  const now = new Date();
  const logId = `LOG_${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${Math.floor(Math.random() * 10000)}`;
  sheet.appendRow([logId, syncType, status, sourceSheet, targetSheet, targetKey, message, now]);
}

/*******************************************************
 * EdsMvpApi.gs switch(action) 패치
 *
 * 기존 getDividendDashboard/addDividend/updateDividend/disableDividend case가 있으면
 * 아래 형태로 교체 권장.
 *
 * case "getDividendDashboard":
 *   return edApi_json_(edApi_success_(action, getDividendDashboard()));
 *
 * case "refreshDividends":
 *   return edApi_json_(edApi_success_(action, resetAndImportDividendSheetToApp()));
 *******************************************************/
