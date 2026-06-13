/*******************************************************
 * ED's MVP - Web API Script v0.8.63
 *
 * 목적:
 * - PWA/모바일 웹앱에서 Google Sheets의 App_* 데이터를 읽고 수정하기 위한 API
 * - 기존 2. 종목현황 직접 수정은 syncToMain 액션에서만 수행
 * - 기본은 읽기 중심, 쓰기 액션은 API 토큰 검증 후 수행
 *
 * 주요 함수:
 * - doGet(e)
 * - doPost(e)
 *
 * 전제:
 * - setupEdsMvpSheets() 실행 완료
 * - EdsMvpImport.gs v0.4 존재
 * - EdsMvpSyncPreview.gs v0.2 존재
 * - EdsMvpSyncActual.gs v0.1 존재
 *******************************************************/

const ED_MVP_API = {
  sheets: {
    accounts: "App_Accounts",
    assets: "App_Assets",
    holdings: "App_Holdings",
    prices: "App_Prices",
    output: "App_Output",
    settings: "App_Settings",
    syncPreview: "App_SyncPreview",
    syncResult: "App_SyncResult",
    syncLog: "App_SyncLog",
  },

  settings: {
    apiTokenKey: "api_token",
  },
};

/**
 * GET API
 * 예:
 * /exec?action=ping
 * /exec?action=getDashboard&token=...
 * /exec?action=getPortfolioOutput&token=...
 */
function doGet(e) {
  return edApi_handleRequest_("GET", e);
}

/**
 * POST API
 * body 예:
 * {
 *   "action": "updateHolding",
 *   "token": "...",
 *   "payload": { ... }
 * }
 */
function doPost(e) {
  return edApi_handleRequest_("POST", e);
}

function edApi_handleRequest_(method, e) {
  try {
    const request = edApi_parseRequest_(method, e);
    const action = String(request.action || "").trim();

    if (!action) {
      return edApi_json_({ ok: false, error: "missing_action" });
    }

    // ping은 토큰 없이 허용
    if (action === "ping") {
      return edApi_json_({
        ok: true,
        action,
        app: "ED's MVP",
        timestamp: new Date(),
      });
    }

    const authResult = edApi_validateToken_(request.token);
    if (!authResult.ok) {
      return edApi_json_({
        ok: false,
        action,
        error: authResult.error,
      });
    }

    // --- 앱 토큰 보안 인터셉터 ---
    // ScriptProperties에 EDS_APP_TOKEN이 설정된 경우, 요청의 x_eds_app_token과 반드시 일치해야 함
    // 미설정 시는 개발/테스트 편의를 위해 통과
    const appTokenInterceptResult = edApi_validateAppToken_(request.appToken);
    if (!appTokenInterceptResult.ok) {
      return edApi_json_({
        ok: false,
        action,
        error: appTokenInterceptResult.error,
        message: 'EDS 앱 토큰 검증 실패. VITE_EDS_APP_TOKEN과 ScriptProperties EDS_APP_TOKEN을 확인하세요.',
      });
    }

    switch (action) {
      case "getSettings":
        return edApi_json_(edApi_success_(action, edApi_getSettings_()));

      case "getAccounts":
        return edApi_json_(edApi_success_(action, edApi_getAccounts_()));

      case "getAssets":
        return edApi_json_(edApi_success_(action, edApi_getAssets_()));

      case "getHoldings":
        return edApi_json_(edApi_success_(action, edApi_getHoldings_()));

      case "getPrices":
        return edApi_json_(edApi_success_(action, edApi_getPrices_()));

      case "getPortfolioOutput":
        return edApi_json_(edApi_success_(action, edApi_getPortfolioOutput_()));

      case "refreshKrxPrices":
        return edApi_json_(edApi_success_(action, refreshKrxPricesFromKis(request.payload)));

      case "getAppStatus":
        return edApi_json_(edApi_success_(action, getAppStatus(request.payload)));

      case "refreshKrxPricesToMainSheet":
        return edApi_json_(edApi_success_(action, refreshKrxPricesToMainSheetFromKis(request.payload)));

       case "fetchSingleChartData":
        return edApi_json_(
          edApi_success_(
            action,
            fetchSingleChartData(request.payload)
          )
        );

       case "getChartData":
        return edApi_json_(
          edApi_success_(
            action,
            getOrRefreshChartData(request.payload)
          )
        );

      case "refreshChartData": {
        const payload = Object.assign({}, request.payload || {}, { force: true });
        return edApi_json_(
          edApi_success_(
            action,
            getOrRefreshChartData(payload)
          )
        );
      }

      case "refreshKrxDailyCharts":
        return edApi_json_(edApi_success_(action, refreshKrxDailyChartsFromKis()));

      case "getDividendDashboard":
        return edApi_json_(edApi_success_(action, getDividendDashboard()));

      case "refreshDividends":
        return edApi_json_(edApi_success_(action, resetAndImportDividendSheetToApp()));      

      case "getDashboard":
        return edApi_json_(edApi_success_(action, edApi_getDashboard_()));

      case "refreshOutput":
        return edApi_json_(edApi_success_(action, edApi_refreshOutput_()));

      case "previewSync":
        return edApi_json_(edApi_success_(action, edApi_previewSync_()));

      case "syncToMain":
        return edApi_json_(edApi_success_(action, edApi_syncToMain_(request.payload)));

      case "updateHolding":
        return edApi_json_(edApi_success_(action, edApi_updateHolding_(request.payload)));

      case "upsertHolding":
        return edApi_json_(edApi_success_(action, edApi_upsertHolding_(request.payload)));

      case "disableHolding":
        return edApi_json_(edApi_success_(action, edApi_disableHolding_(request.payload)));

      case "upsertPrice":
        return edApi_json_(edApi_success_(action, edApi_upsertPrice_(request.payload)));

      default:
        return edApi_json_({
          ok: false,
          action,
          error: "unknown_action",
        });
    }
  } catch (err) {
    return edApi_json_({
      ok: false,
      error: "exception",
      message: err && err.message ? err.message : String(err),
      stack: err && err.stack ? err.stack : "",
    });
  }
}

function edApi_parseRequest_(method, e) {
  const params = (e && e.parameter) ? e.parameter : {};

  if (method === "GET") {
    return {
      action: params.action || "",
      token: params.token || "",
      appToken: params.x_eds_app_token || "",
      payload: params.payload ? JSON.parse(params.payload) : {},
    };
  }

  let body = {};

  if (e && e.postData && e.postData.contents) {
    const contentType = String(e.postData.type || "");

    if (contentType.indexOf("application/json") >= 0 || e.postData.contents.trim().charAt(0) === "{") {
      body = JSON.parse(e.postData.contents);
    } else {
      body = params;
    }
  }

  return {
    action: body.action || params.action || "",
    token: body.token || params.token || "",
    appToken: body.x_eds_app_token || params.x_eds_app_token || "",
    payload: body.payload || {},
  };
}

function edApi_success_(action, data) {
  return {
    ok: true,
    action,
    timestamp: new Date(),
    data,
  };
}

function edApi_json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/*******************************************************
 * Auth / Settings
 *******************************************************/

function edApi_validateToken_(token) {
  const settings = edApi_getSettingsMap_();
  const savedToken = String(settings[ED_MVP_API.settings.apiTokenKey] || "").trim();

  // 아직 토큰이 설정되지 않았으면 초기 개발 편의를 위해 허용.
  // PWA 연결 전에는 setDefaultApiToken() 실행 권장.
  if (!savedToken) {
    return { ok: true, warning: "api_token_not_set" };
  }

  if (!token) {
    return { ok: false, error: "missing_token" };
  }

  if (String(token) !== savedToken) {
    return { ok: false, error: "invalid_token" };
  }

  return { ok: true };
}

/**
 * EDS 앱 토큰 인터셉터 검증
 * ScriptProperties 'EDS_APP_TOKEN'이 설정된 경우만 검증을 수행함.
 * 미설정 시에는 열린 통로(전환기 모드) 유지.
 */
function edApi_validateAppToken_(appToken) {
  const props = PropertiesService.getScriptProperties();
  const savedAppToken = String(props.getProperty('EDS_APP_TOKEN') || '').trim();

  // 설정 안 된 경우 전환기 모드로 허용
  if (!savedAppToken) {
    return { ok: true, warning: 'eds_app_token_not_set' };
  }

  if (!appToken) {
    return { ok: false, error: 'missing_app_token' };
  }

  if (String(appToken).trim() !== savedAppToken) {
    return { ok: false, error: 'invalid_app_token' };
  }

  return { ok: true };
}

function setDefaultApiToken() {
  const token = Utilities.getUuid().replace(/-/g, "");
  edApi_upsertSetting_(ED_MVP_API.settings.apiTokenKey, token, "PWA API 호출용 토큰");

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "api_token 생성 완료. App_Settings에서 값을 확인하세요.",
    "ED's MVP",
    8
  );
}

function edApi_getSettings_() {
  const rows = edApi_readSheetAsObjects_(ED_MVP_API.sheets.settings);

  return rows.map((row) => ({
    key: row.key,
    value: row.key === ED_MVP_API.settings.apiTokenKey ? "********" : row.value,
    description: row.description,
    updated_at: row.updated_at,
  }));
}

function edApi_getSettingsMap_() {
  const rows = edApi_readSheetAsObjects_(ED_MVP_API.sheets.settings);
  const map = {};

  rows.forEach((row) => {
    if (row.key !== "") {
      map[String(row.key)] = row.value;
    }
  });

  return map;
}

function edApi_upsertSetting_(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_API.sheets.settings);

  if (!sheet) throw new Error("App_Settings 시트를 찾을 수 없습니다.");

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const keyCol = headers.indexOf("key");
  const valueCol = headers.indexOf("value");
  const descCol = headers.indexOf("description");
  const updatedAtCol = headers.indexOf("updated_at");

  if (keyCol < 0 || valueCol < 0) throw new Error("App_Settings 헤더 오류");

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(key)) {
      sheet.getRange(i + 1, valueCol + 1).setValue(value);
      if (descCol >= 0) sheet.getRange(i + 1, descCol + 1).setValue(description || "");
      if (updatedAtCol >= 0) sheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date());
      return;
    }
  }

  const row = new Array(headers.length).fill("");
  row[keyCol] = key;
  row[valueCol] = value;
  if (descCol >= 0) row[descCol] = description || "";
  if (updatedAtCol >= 0) row[updatedAtCol] = new Date();

  sheet.appendRow(row);
}

/*******************************************************
 * Read APIs
 *******************************************************/

function edApi_getAccounts_() {
  return edApi_readSheetAsObjects_(ED_MVP_API.sheets.accounts)
    .filter((row) => String(row.enabled || "").toUpperCase() === "TRUE")
    .sort((a, b) => edApi_num_(a.display_order) - edApi_num_(b.display_order));
}

function edApi_getAssets_() {
  return edApi_readSheetAsObjects_(ED_MVP_API.sheets.assets)
    .filter((row) => String(row.enabled || "").toUpperCase() === "TRUE");
}

function edApi_getHoldings_() {
  return edApi_readSheetAsObjects_(ED_MVP_API.sheets.holdings)
    .filter((row) => String(row.enabled || "").toUpperCase() === "TRUE");
}

function edApi_getPrices_() {
  return edApi_readSheetAsObjects_(ED_MVP_API.sheets.prices);
}

function edApi_getPortfolioOutput_() {
  return edApi_readSheetAsObjects_(ED_MVP_API.sheets.output);
}

function edApi_getDashboard_() {
  const accounts = edApi_getAccounts_();
  const output = edApi_getPortfolioOutput_();

  const totalValuation = output.reduce((sum, row) => sum + edApi_num_(row.valuation_amount), 0);
  const totalInvested = output.reduce((sum, row) => sum + edApi_num_(row.invested_amount), 0);
  const totalProfit = totalValuation - totalInvested;
  const totalProfitRate = totalInvested > 0 ? totalProfit / totalInvested : 0;

  const accountMap = new Map();

  output.forEach((row) => {
    const accountId = String(row.account_id || "");

    if (!accountMap.has(accountId)) {
      accountMap.set(accountId, {
        account_id: accountId,
        account_name: row.account_name || "",
        broker: row.broker || "",
        account_type: row.account_type || "",
        valuation_amount: 0,
        invested_amount: 0,
        profit_amount: 0,
        profit_rate: 0,
        weight: 0,
        holding_count: 0,
      });
    }

    const item = accountMap.get(accountId);
    item.valuation_amount += edApi_num_(row.valuation_amount);
    item.invested_amount += edApi_num_(row.invested_amount);
    item.holding_count += 1;
  });

  const accountSummaries = Array.from(accountMap.values()).map((item) => {
    item.profit_amount = item.valuation_amount - item.invested_amount;
    item.profit_rate = item.invested_amount > 0 ? item.profit_amount / item.invested_amount : 0;
    item.weight = totalValuation > 0 ? item.valuation_amount / totalValuation : 0;
    return item;
  });

  const topHoldings = output
    .slice()
    .sort((a, b) => edApi_num_(b.valuation_amount) - edApi_num_(a.valuation_amount))
    .slice(0, 10);

  return {
    summary: {
      total_valuation: totalValuation,
      total_invested: totalInvested,
      total_profit: totalProfit,
      total_profit_rate: totalProfitRate,
      account_count: accounts.length,
      holding_count: output.length,
    },
    accounts: accountSummaries,
    top_holdings: topHoldings,
    output,
  };
}

/*******************************************************
 * Write / Action APIs
 *******************************************************/

function edApi_refreshOutput_() {
  buildAppOutputFromAppSheets();
  return {
    message: "App_Output refreshed",
    output_count: edApi_readSheetAsObjects_(ED_MVP_API.sheets.output).length,
  };
}

function edApi_previewSync_() {
  previewSyncAppHoldingsToMainSheet();
  return {
    message: "Sync preview created",
    preview_summary: edApi_readSummarySheet_(ED_MVP_API.sheets.syncPreview),
  };
}

function edApi_syncToMain_(payload) {
  // payload.skipValidation === true 일 때 고속 숏컷 모드로 분기
  // 그 외(undefined, false, null 등) 모든 경우는 안전 모드로 동작
  const skipValidation = !!(payload && payload.skipValidation === true);

  try {
    syncAppHoldingsToMainSheet({ skipValidation: skipValidation });
  } catch (err) {
    // 고속 모드 실패 시 안전 모드로 자동 폴백
    if (skipValidation) {
      Logger.log("[edApi_syncToMain_] 고속 모드 실패, 안전 모드로 폴백: " + err.message);
      syncAppHoldingsToMainSheet({ skipValidation: false });
    } else {
      throw err;
    }
  }

  return {
    message: skipValidation ? "Fast sync to main completed (skip validation)" : "Sync to main completed",
    mode: skipValidation ? "fast" : "safe",
    sync_result_summary: edApi_readSummarySheet_(ED_MVP_API.sheets.syncResult),
  };
}

function edApi_updateHolding_(payload) {
  if (!payload || !payload.holding_id) throw new Error("holding_id가 필요합니다.");

  const allowedFields = [
    "quantity",
    "avg_price",
    "target_weight_account",
    "memo",
    "enabled",
  ];

  const updated = edApi_updateRowByKey_(
    ED_MVP_API.sheets.holdings,
    "holding_id",
    payload.holding_id,
    payload,
    allowedFields
  );

  buildAppOutputFromAppSheets();

  return {
    updated,
    holding_id: payload.holding_id,
  };
}

function edApi_upsertHolding_(payload) {
  if (!payload) throw new Error("payload가 필요합니다.");

  const required = ["holding_id", "account_id", "asset_id", "ticker", "asset_name"];
  required.forEach((field) => {
    if (!payload[field]) throw new Error(`${field}가 필요합니다.`);
  });

  const sheetName = ED_MVP_API.sheets.holdings;
  const result = edApi_upsertRowByKey_(sheetName, "holding_id", payload.holding_id, payload);

  buildAppOutputFromAppSheets();

  return result;
}

function edApi_disableHolding_(payload) {
  if (!payload || !payload.holding_id) throw new Error("holding_id가 필요합니다.");

  const updated = edApi_updateRowByKey_(
    ED_MVP_API.sheets.holdings,
    "holding_id",
    payload.holding_id,
    {
      enabled: "FALSE",
      updated_at: new Date(),
    },
    ["enabled", "updated_at"]
  );

  buildAppOutputFromAppSheets();

  return {
    updated,
    holding_id: payload.holding_id,
  };
}

function edApi_upsertPrice_(payload) {
  if (!payload || !payload.asset_id) throw new Error("asset_id가 필요합니다.");

  const result = edApi_upsertRowByKey_(ED_MVP_API.sheets.prices, "asset_id", payload.asset_id, payload);
  buildAppOutputFromAppSheets();

  return result;
}

/*******************************************************
 * Sheet Helpers
 *******************************************************/

function edApi_readSheetAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  return values.slice(1).map((row) => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = edApi_normalizeValue_(row[index]);
    });

    return obj;
  });
}

function edApi_readSummarySheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) return {};

  const maxRows = Math.min(sheet.getLastRow(), 30);
  const values = sheet.getRange(1, 1, maxRows, 2).getValues();
  const result = {};

  values.forEach((row) => {
    const key = String(row[0] || "");
    if (!key || key === "metric") return;
    result[key] = edApi_normalizeValue_(row[1]);
  });

  return result;
}

function edApi_updateRowByKey_(sheetName, keyHeader, keyValue, payload, allowedFields) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) throw new Error(`${sheetName} 시트를 찾을 수 없습니다.`);

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const keyCol = headers.indexOf(keyHeader);

  if (keyCol < 0) throw new Error(`${keyHeader} 헤더를 찾을 수 없습니다.`);

  const allowedSet = new Set(allowedFields || []);

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(keyValue)) {
      const rowIndex = i + 1;
      let changedCount = 0;

      headers.forEach((header, colIndex) => {
        if (!allowedSet.has(header)) return;
        if (!(header in payload)) return;

        sheet.getRange(rowIndex, colIndex + 1).setValue(payload[header]);
        changedCount++;
      });

      const updatedAtCol = headers.indexOf("updated_at");
      if (updatedAtCol >= 0) {
        sheet.getRange(rowIndex, updatedAtCol + 1).setValue(new Date());
      }

      return {
        row: rowIndex,
        changed_count: changedCount,
      };
    }
  }

  throw new Error(`${sheetName}에서 ${keyHeader}=${keyValue} 행을 찾을 수 없습니다.`);
}

function edApi_upsertRowByKey_(sheetName, keyHeader, keyValue, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) throw new Error(`${sheetName} 시트를 찾을 수 없습니다.`);

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const keyCol = headers.indexOf(keyHeader);

  if (keyCol < 0) throw new Error(`${keyHeader} 헤더를 찾을 수 없습니다.`);

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(keyValue)) {
      const rowIndex = i + 1;

      headers.forEach((header, colIndex) => {
        if (header in payload) {
          sheet.getRange(rowIndex, colIndex + 1).setValue(payload[header]);
        }
      });

      const updatedAtCol = headers.indexOf("updated_at");
      if (updatedAtCol >= 0) {
        sheet.getRange(rowIndex, updatedAtCol + 1).setValue(new Date());
      }

      return {
        mode: "update",
        row: rowIndex,
        key: keyValue,
      };
    }
  }

  const newRow = new Array(headers.length).fill("");

  headers.forEach((header, index) => {
    if (header in payload) {
      newRow[index] = payload[header];
    }
  });

  const createdAtCol = headers.indexOf("created_at");
  const updatedAtCol = headers.indexOf("updated_at");
  const enabledCol = headers.indexOf("enabled");

  if (createdAtCol >= 0 && !newRow[createdAtCol]) newRow[createdAtCol] = new Date();
  if (updatedAtCol >= 0 && !newRow[updatedAtCol]) newRow[updatedAtCol] = new Date();
  if (enabledCol >= 0 && !newRow[enabledCol]) newRow[enabledCol] = "TRUE";

  sheet.appendRow(newRow);

  return {
    mode: "insert",
    row: sheet.getLastRow(),
    key: keyValue,
  };
}

function edApi_normalizeValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
  }

  return value;
}

function edApi_num_(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;

  const text = String(value)
    .replace(/,/g, "")
    .replace(/원/g, "")
    .replace(/주/g, "")
    .replace(/%/g, "")
    .replace(/\+/g, "")
    .replace(/−/g, "-")
    .trim();

  const num = Number(text);
  return isNaN(num) ? 0 : num;
}
