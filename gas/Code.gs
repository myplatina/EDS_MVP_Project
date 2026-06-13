/**
 * 네이버 금융 모바일 API를 통해 실시간 주가 정보를 가져옵니다.
 * @param {string} code 종목코드 (예: "005930", "329200")
 * @param {string} type [선택] 가져올 데이터 타입 ("price": 현재가(기본값), "rate": 등락률, "name": 종목명)
 * @return 실시간 데이터
 * @customfunction
 */
function GET_NAVER_REALTIME(code, type) {
  if (!code) return "Code Missing";

  // 핵심 수정: type 인자가 없거나 비어있으면 "price"로 강제 설정 (에러 방지)
  if (!type) {
    type = "price";
  }
  
  // 숫자로만 된 코드일 경우 문자로 변환
  code = code.toString(); 
  while (code.length < 6) {
    code = "0" + code;
  }

  try {
    var url = "https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:" + code;
    
    var options = {
      'muteHttpExceptions': true,
      'headers': {
        'Cache-Control': 'no-cache'
      }
    };

    var response = UrlFetchApp.fetch(url, options);
    var jsonText = response.getContentText();
    var data = JSON.parse(jsonText);
    
    if (!data.result || !data.result.areas || !data.result.areas[0].datas) {
      return "Data Error";
    }

    var stockData = data.result.areas[0].datas[0];

    // type이 이제 무조건 문자열이므로 안전하게 변환 가능
    switch (type.toString().toLowerCase()) {
      case "price": // 현재가
        return stockData.nv; 
      case "rate":  // 등락률 (%)
        return stockData.cr; 
      case "name":  // 종목명
        return stockData.nm;
      case "amount": // 등락액
        return stockData.cv;
      default:
        return stockData.nv;
    }

  } catch (e) {
    return "Error: " + e.message;
  }
}


/**
 * 업비트 API를 통해 코인 정보를 가져옵니다.
 * @param {string} ticker 코인 티커 (예: "KRW-BTC", "BTC")
 * @param {string} type [선택] 가져올 데이터 타입 ("price": 현재가(기본값), "rate": 등락률)
 * @return 실시간 데이터
 * @customfunction
 */
function GET_UPBIT_REALTIME(ticker, type) {
  if (!ticker) return "Ticker Missing";

  // type이 비어있으면 기본값 "price"로 설정
  if (!type) {
    type = "price";
  }

  // "BTC"만 입력해도 "KRW-BTC"로 자동 변환해주는 센스
  if (ticker.indexOf("-") === -1) {
    ticker = "KRW-" + ticker;
  }

  try {
    var url = "https://api.upbit.com/v1/ticker?markets=" + ticker;
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    var jsonText = response.getContentText();
    var data = JSON.parse(jsonText);

    if (!data || data.length === 0) {
      return "Ticker Error";
    }

    var coinData = data[0];

    switch (type.toString().toLowerCase()) {
      case "price": // 현재가
        return coinData.trade_price;
      case "rate":  // 전일 대비 등락률 (소수점 형태, 예: 0.015)
        return coinData.signed_change_rate; 
      case "amount": // 전일 대비 등락액 (원)
        return coinData.signed_change_price;
      default:
        return coinData.trade_price;
    }

  } catch (e) {
    return "Error: " + e.message;
  }
}



  /*******************************************************
 * ED's MVP - Google Sheets Initial Setup Script v0.8.7
 *
 * 목적:
 * - 기존 "2.종목현황" 시트는 수정하지 않음
 * - 앱 전용 App_* 시트 생성
 * - 헤더, 기본 서식, 기본 설정값 생성
 *
 * 실행 함수:
 * - setupEdsMvpSheets()
 *******************************************************/

const ED_MVP = {
  appName: "ED's MVP",
  sheets: {
    accounts: "App_Accounts",
    assets: "App_Assets",
    holdings: "App_Holdings",
    prices: "App_Prices",
    accountValues: "App_AccountValues",
    output: "App_Output",
    settings: "App_Settings",
    syncLog: "App_SyncLog",
  },
};

const ED_MVP_SCHEMAS = {
  App_Accounts: [
    "account_id",
    "account_name",
    "broker",
    "account_type",
    "currency_base",
    "display_order",
    "enabled",
    "created_at",
    "updated_at",
  ],

  App_Assets: [
    "asset_id",
    "ticker",
    "asset_name",
    "country",
    "market",
    "asset_class",
    "currency",
    "is_etf",
    "enabled",
    "created_at",
    "updated_at",
  ],

  App_Holdings: [
    "holding_id",
    "account_id",
    "asset_id",
    "ticker",
    "asset_name",
    "quantity",
    "avg_price",
    "target_weight_account",
    "memo",
    "enabled",
    "created_at",
    "updated_at",
  ],

  App_Prices: [
    "asset_id",
    "ticker",
    "price",
    "prev_close",
    "change_amount",
    "change_rate",
    "currency",
    "source",
    "fetched_at",
    "updated_at",
  ],

  App_AccountValues: [
    "value_id",
    "year_month",
    "account_id",
    "account_value",
    "cash_value",
    "invested_value",
    "memo",
    "created_at",
    "updated_at",
  ],

  App_Output: [
    "account_id",
    "account_name",
    "broker",
    "account_type",
    "asset_id",
    "ticker",
    "asset_name",
    "quantity",
    "avg_price",
    "price",
    "invested_amount",
    "valuation_amount",
    "profit_amount",
    "profit_rate",
    "account_weight",
    "total_weight",
    "target_weight_account",
    "target_gap_rate",
    "target_gap_amount",
    "currency",
    "price_source",
    "price_fetched_at",
    "updated_at",
  ],

  App_Settings: [
    "key",
    "value",
    "description",
    "updated_at",
  ],

  App_SyncLog: [
    "log_id",
    "sync_type",
    "status",
    "source_sheet",
    "target_sheet",
    "target_key",
    "message",
    "created_at",
  ],
};

/**
 * 스프레드시트 열릴 때 메뉴 생성
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ED's MVP")
    .addItem("초기 시트 생성/정비", "setupEdsMvpSheets")
    .addToUi();
}

/**
 * 메인 실행 함수
 */
function setupEdsMvpSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(ED_MVP_SCHEMAS).forEach((sheetName) => {
    const headers = ED_MVP_SCHEMAS[sheetName];
    const sheet = ensureSheet_(ss, sheetName, headers);
    applyBaseFormat_(sheet, headers);
    applyValidations_(sheetName, sheet, headers);
    applyNumberFormats_(sheetName, sheet, headers);
  });

  initializeSettings_(ss);
  initializeAccountTemplates_(ss);

  SpreadsheetApp.getUi().alert(
    "ED's MVP 앱용 시트 생성/정비 완료.\n기존 2.종목현황 시트는 수정하지 않았습니다."
  );
}

/**
 * 시트 생성 또는 헤더 정비
 */
function ensureSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * 기본 서식
 */
function applyBaseFormat_(sheet, headers) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#137333")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(1, 32);

  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  const maxRows = Math.max(sheet.getMaxRows(), 1000);
  const bodyRange = sheet.getRange(2, 1, maxRows - 1, headers.length);

  bodyRange
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

/**
 * 데이터 검증
 */
function applyValidations_(sheetName, sheet, headers) {
  if (sheetName === ED_MVP.sheets.accounts) {
    applyListValidation_(sheet, headers, "account_type", [
      "일반",
      "ISA",
      "개인연금",
      "IRP",
      "기타",
    ]);

    applyListValidation_(sheet, headers, "currency_base", [
      "KRW",
      "USD",
    ]);

    applyBooleanValidation_(sheet, headers, "enabled");
  }

  if (sheetName === ED_MVP.sheets.assets) {
    applyListValidation_(sheet, headers, "country", [
      "한국",
      "미국",
      "기타",
    ]);

    applyListValidation_(sheet, headers, "market", [
      "KRX",
      "NYSE",
      "NASDAQ",
      "AMEX",
      "기타",
    ]);

    applyListValidation_(sheet, headers, "asset_class", [
      "주식",
      "주식 ETF",
      "채권",
      "채권 ETF",
      "혼합 ETF",
      "현금",
      "기타",
    ]);

    applyListValidation_(sheet, headers, "currency", [
      "KRW",
      "USD",
    ]);

    applyBooleanValidation_(sheet, headers, "is_etf");
    applyBooleanValidation_(sheet, headers, "enabled");
  }

  if (sheetName === ED_MVP.sheets.holdings) {
    applyBooleanValidation_(sheet, headers, "enabled");
  }

  if (sheetName === ED_MVP.sheets.prices) {
    applyListValidation_(sheet, headers, "currency", [
      "KRW",
      "USD",
    ]);

    applyListValidation_(sheet, headers, "source", [
      "manual",
      "sheet",
      "googlefinance",
      "naver",
      "api",
    ]);
  }

  if (sheetName === ED_MVP.sheets.syncLog) {
    applyListValidation_(sheet, headers, "sync_type", [
      "setup",
      "app_to_main",
      "main_to_app",
      "price_update",
      "manual",
    ]);

    applyListValidation_(sheet, headers, "status", [
      "success",
      "warning",
      "error",
    ]);
  }
}

/**
 * 숫자/날짜 서식
 */
function applyNumberFormats_(sheetName, sheet, headers) {
  const maxRows = Math.max(sheet.getMaxRows(), 1000);

  const numberFields = [
    "quantity",
    "avg_price",
    "price",
    "prev_close",
    "change_amount",
    "account_value",
    "cash_value",
    "invested_value",
    "invested_amount",
    "valuation_amount",
    "profit_amount",
    "target_gap_amount",
    "display_order",
  ];

  const percentFields = [
    "target_weight_account",
    "change_rate",
    "profit_rate",
    "account_weight",
    "total_weight",
    "target_gap_rate",
  ];

  const dateTimeFields = [
    "created_at",
    "updated_at",
    "fetched_at",
    "price_fetched_at",
  ];

  const yearMonthFields = [
    "year_month",
  ];

  numberFields.forEach((field) => {
    const col = headers.indexOf(field) + 1;
    if (col > 0) {
      sheet.getRange(2, col, maxRows - 1, 1).setNumberFormat("#,##0.########");
    }
  });

  percentFields.forEach((field) => {
    const col = headers.indexOf(field) + 1;
    if (col > 0) {
      sheet.getRange(2, col, maxRows - 1, 1).setNumberFormat("0.00%");
    }
  });

  dateTimeFields.forEach((field) => {
    const col = headers.indexOf(field) + 1;
    if (col > 0) {
      sheet.getRange(2, col, maxRows - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }
  });

  yearMonthFields.forEach((field) => {
    const col = headers.indexOf(field) + 1;
    if (col > 0) {
      sheet.getRange(2, col, maxRows - 1, 1).setNumberFormat("@");
    }
  });
}

/**
 * TRUE/FALSE 검증
 */
function applyBooleanValidation_(sheet, headers, fieldName) {
  const col = headers.indexOf(fieldName) + 1;
  if (col <= 0) return;

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(["TRUE", "FALSE"], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
}

/**
 * 리스트 검증
 */
function applyListValidation_(sheet, headers, fieldName, values) {
  const col = headers.indexOf(fieldName) + 1;
  if (col <= 0) return;

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
}

/**
 * 기본 설정값 입력
 * 기존 값이 있으면 덮어쓰지 않음
 */
function initializeSettings_(ss) {
  const sheet = ss.getSheetByName(ED_MVP.sheets.settings);
  const now = new Date();

  const defaults = [
    ["app_name", "ED's MVP", "앱 이름", now],
    ["base_currency", "KRW", "기준 통화", now],
    ["default_account_filter", "ALL", "기본 계좌 필터", now],
    ["price_mode", "sheet", "현재가 처리 방식: manual/sheet/googlefinance/naver/api", now],
    ["dividend_enabled", "FALSE", "1차 MVP 배당 기능 제외", now],
    ["main_sheet_name", "2.종목현황", "기존 기준 원장 시트명", now],
    ["sync_mode", "manual", "기존 원장 동기화 방식: manual/immediate/scheduled", now],
    ["last_sync_at", "", "마지막 동기화 시각", now],
  ];

  upsertRowsByKey_(sheet, "key", defaults);
}

/**
 * 기본 계좌 템플릿 입력
 * 기존 account_id가 있으면 덮어쓰지 않음
 */
function initializeAccountTemplates_(ss) {
  const sheet = ss.getSheetByName(ED_MVP.sheets.accounts);
  const now = new Date();

  const rows = [
    ["ACC_GENERAL_01", "일반계좌1", "", "일반", "KRW", 1, "TRUE", now, now],
    ["ACC_GENERAL_02", "일반계좌2", "", "일반", "KRW", 2, "TRUE", now, now],
    ["ACC_ISA_01", "ISA", "", "ISA", "KRW", 3, "TRUE", now, now],
    ["ACC_PENSION_01", "개인연금1", "", "개인연금", "KRW", 4, "TRUE", now, now],
    ["ACC_PENSION_02", "개인연금2", "", "개인연금", "KRW", 5, "TRUE", now, now],
    ["ACC_IRP_01", "IRP 1", "", "IRP", "KRW", 6, "TRUE", now, now],
  ];

  upsertRowsByKey_(sheet, "account_id", rows);
}

/**
 * 특정 키 기준 upsert
 * 기존 키가 있으면 유지, 없으면 append
 */
function upsertRowsByKey_(sheet, keyHeader, rows) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyCol = headers.indexOf(keyHeader) + 1;

  if (keyCol <= 0) {
    throw new Error(`Key header not found: ${keyHeader}`);
  }

  const lastRow = sheet.getLastRow();
  let existingKeys = new Set();

  if (lastRow >= 2) {
    const keyValues = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
    existingKeys = new Set(
      keyValues
        .flat()
        .filter((value) => value !== "" && value !== null)
        .map((value) => String(value))
    );
  }

  const rowsToAppend = rows.filter((row) => {
    const keyValue = String(row[keyCol - 1]);
    return !existingKeys.has(keyValue);
  });

  if (rowsToAppend.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }
}

/**
 * 동기화 로그 기록용 함수
 * 이후 app_to_main 동기화에서 사용
 */
function writeSyncLog_(syncType, status, sourceSheet, targetSheet, targetKey, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP.sheets.syncLog);

  if (!sheet) return;

  const now = new Date();
  const logId = `LOG_${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${Math.floor(Math.random() * 10000)}`;

  sheet.appendRow([
    logId,
    syncType,
    status,
    sourceSheet,
    targetSheet,
    targetKey,
    message,
    now,
  ]);
}

