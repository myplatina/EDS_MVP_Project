/*******************************************************
 * ED's MVP - App Data Validation Script v0.8.63
 *
 * 목적:
 * - 기존 원장에서 App_* 시트로 이관된 데이터 검산
 * - 기존 2. 종목현황 시트는 수정하지 않음
 *
 * 실행 함수:
 * - validateEdsMvpAppData()
 *******************************************************/

const ED_MVP_VALIDATE = {
  sheets: {
    accounts: "App_Accounts",
    assets: "App_Assets",
    holdings: "App_Holdings",
    prices: "App_Prices",
    output: "App_Output",
    report: "App_ValidationReport",
  },

  allowed: {
    country: ["한국", "미국", "기타"],
    market: ["KRX", "NYSE", "NASDAQ", "AMEX", "기타"],
    assetClass: ["주식", "주식 ETF", "채권", "채권 ETF", "혼합 ETF", "현금", "기타"],
    currency: ["KRW", "USD"],
    accountType: ["일반", "ISA", "개인연금", "IRP", "기타"],
    priceSource: ["manual", "sheet", "googlefinance", "naver", "api"],
  },
};

function validateEdsMvpAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issues = [];

  const accountsSheet = ss.getSheetByName(ED_MVP_VALIDATE.sheets.accounts);
  const assetsSheet = ss.getSheetByName(ED_MVP_VALIDATE.sheets.assets);
  const holdingsSheet = ss.getSheetByName(ED_MVP_VALIDATE.sheets.holdings);
  const pricesSheet = ss.getSheetByName(ED_MVP_VALIDATE.sheets.prices);
  const outputSheet = ss.getSheetByName(ED_MVP_VALIDATE.sheets.output);

  validateSheetExists_(issues, accountsSheet, "App_Accounts");
  validateSheetExists_(issues, assetsSheet, "App_Assets");
  validateSheetExists_(issues, holdingsSheet, "App_Holdings");
  validateSheetExists_(issues, pricesSheet, "App_Prices");
  validateSheetExists_(issues, outputSheet, "App_Output");

  if (issues.some((issue) => issue.severity === "ERROR")) {
    writeValidationReport_(ss, issues, []);
    ss.toast("검산 중단: 필수 시트 누락", "ED's MVP", 8);
    return;
  }

  const accounts = readSheetAsObjectsWithRow_(accountsSheet);
  const assets = readSheetAsObjectsWithRow_(assetsSheet);
  const holdings = readSheetAsObjectsWithRow_(holdingsSheet);
  const prices = readSheetAsObjectsWithRow_(pricesSheet);
  const outputs = readSheetAsObjectsWithRow_(outputSheet);

  const accountIds = new Set(accounts.map((r) => String(r.account_id || "")));
  const assetIds = new Set(assets.map((r) => String(r.asset_id || "")));
  const priceAssetIds = new Set(prices.map((r) => String(r.asset_id || "")));

  validateDuplicateKey_(issues, accounts, "App_Accounts", "account_id");
  validateDuplicateKey_(issues, assets, "App_Assets", "asset_id");
  validateDuplicateKey_(issues, holdings, "App_Holdings", "holding_id");
  validateDuplicateKey_(issues, prices, "App_Prices", "asset_id");

  validateAccounts_(issues, accounts);
  validateAssets_(issues, assets);
  validatePrices_(issues, prices);
  validateHoldings_(issues, holdings, accountIds, assetIds, priceAssetIds);
  validateOutput_(issues, holdings, outputs);
  validateAccountTargetWeights_(issues, holdings, accounts);

  const summary = buildValidationSummary_(issues, accounts, assets, holdings, prices, outputs);
  writeValidationReport_(ss, issues, summary);

  const errorCount = issues.filter((i) => i.severity === "ERROR").length;
  const warnCount = issues.filter((i) => i.severity === "WARNING").length;

  ss.toast(
    `검산 완료: ERROR ${errorCount}건 / WARNING ${warnCount}건`,
    "ED's MVP",
    8
  );
}

function validateSheetExists_(issues, sheet, sheetName) {
  if (!sheet) {
    issues.push({
      severity: "ERROR",
      issue_type: "missing_sheet",
      sheet: sheetName,
      row: "",
      key: "",
      message: `${sheetName} 시트를 찾을 수 없음`,
      value: "",
    });
  }
}

function validateDuplicateKey_(issues, rows, sheetName, keyField) {
  const seen = new Map();

  rows.forEach((row) => {
    const key = String(row[keyField] || "");

    if (!key) {
      issues.push({
        severity: "ERROR",
        issue_type: "missing_key",
        sheet: sheetName,
        row: row.__row,
        key: keyField,
        message: `${keyField} 값 없음`,
        value: "",
      });
      return;
    }

    if (seen.has(key)) {
      issues.push({
        severity: "ERROR",
        issue_type: "duplicate_key",
        sheet: sheetName,
        row: row.__row,
        key,
        message: `${keyField} 중복. 최초 행=${seen.get(key)}`,
        value: key,
      });
    } else {
      seen.set(key, row.__row);
    }
  });
}

function validateAccounts_(issues, accounts) {
  accounts.forEach((row) => {
    const type = String(row.account_type || "");
    const currency = String(row.currency_base || "");
    const enabled = String(row.enabled || "").toUpperCase();

    if (!ED_MVP_VALIDATE.allowed.accountType.includes(type)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_account_type",
        sheet: "App_Accounts",
        row: row.__row,
        key: row.account_id,
        message: "허용되지 않은 account_type",
        value: type,
      });
    }

    if (!ED_MVP_VALIDATE.allowed.currency.includes(currency)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_currency_base",
        sheet: "App_Accounts",
        row: row.__row,
        key: row.account_id,
        message: "허용되지 않은 currency_base",
        value: currency,
      });
    }

    if (enabled !== "TRUE" && enabled !== "FALSE") {
      issues.push({
        severity: "WARNING",
        issue_type: "invalid_enabled",
        sheet: "App_Accounts",
        row: row.__row,
        key: row.account_id,
        message: "enabled 값은 TRUE/FALSE 권장",
        value: row.enabled,
      });
    }
  });
}

function validateAssets_(issues, assets) {
  assets.forEach((row) => {
    const country = String(row.country || "");
    const market = String(row.market || "");
    const assetClass = String(row.asset_class || "");
    const currency = String(row.currency || "");

    if (!ED_MVP_VALIDATE.allowed.country.includes(country)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_country",
        sheet: "App_Assets",
        row: row.__row,
        key: row.asset_id,
        message: "허용되지 않은 country",
        value: country,
      });
    }

    if (!ED_MVP_VALIDATE.allowed.market.includes(market)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_market",
        sheet: "App_Assets",
        row: row.__row,
        key: row.asset_id,
        message: "허용되지 않은 market",
        value: market,
      });
    }

    if (!ED_MVP_VALIDATE.allowed.assetClass.includes(assetClass)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_asset_class",
        sheet: "App_Assets",
        row: row.__row,
        key: row.asset_id,
        message: "허용되지 않은 asset_class",
        value: assetClass,
      });
    }

    if (!ED_MVP_VALIDATE.allowed.currency.includes(currency)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_asset_currency",
        sheet: "App_Assets",
        row: row.__row,
        key: row.asset_id,
        message: "허용되지 않은 currency",
        value: currency,
      });
    }
  });
}

function validatePrices_(issues, prices) {
  prices.forEach((row) => {
    const price = parseNumber_(row.price);
    const currency = String(row.currency || "");
    const source = String(row.source || "");

    if (price <= 0) {
      issues.push({
        severity: "WARNING",
        issue_type: "missing_or_zero_price",
        sheet: "App_Prices",
        row: row.__row,
        key: row.asset_id,
        message: "현재가가 0 이하. 평가액이 0으로 계산될 수 있음",
        value: row.price,
      });
    }

    if (!ED_MVP_VALIDATE.allowed.currency.includes(currency)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_price_currency",
        sheet: "App_Prices",
        row: row.__row,
        key: row.asset_id,
        message: "허용되지 않은 currency",
        value: currency,
      });
    }

    if (!ED_MVP_VALIDATE.allowed.priceSource.includes(source)) {
      issues.push({
        severity: "ERROR",
        issue_type: "invalid_price_source",
        sheet: "App_Prices",
        row: row.__row,
        key: row.asset_id,
        message: "허용되지 않은 source",
        value: source,
      });
    }
  });
}

function validateHoldings_(issues, holdings, accountIds, assetIds, priceAssetIds) {
  holdings.forEach((row) => {
    const enabled = String(row.enabled || "").toUpperCase();
    const quantity = parseNumber_(row.quantity);
    const avgPrice = parseNumber_(row.avg_price);
    const targetWeight = parsePercent_(row.target_weight_account);

    if (enabled !== "TRUE") return;

    if (!accountIds.has(String(row.account_id || ""))) {
      issues.push({
        severity: "ERROR",
        issue_type: "missing_account_reference",
        sheet: "App_Holdings",
        row: row.__row,
        key: row.holding_id,
        message: "App_Accounts에 없는 account_id 참조",
        value: row.account_id,
      });
    }

    if (!assetIds.has(String(row.asset_id || ""))) {
      issues.push({
        severity: "ERROR",
        issue_type: "missing_asset_reference",
        sheet: "App_Holdings",
        row: row.__row,
        key: row.holding_id,
        message: "App_Assets에 없는 asset_id 참조",
        value: row.asset_id,
      });
    }

    if (!priceAssetIds.has(String(row.asset_id || ""))) {
      issues.push({
        severity: "WARNING",
        issue_type: "missing_price_reference",
        sheet: "App_Holdings",
        row: row.__row,
        key: row.holding_id,
        message: "App_Prices에 가격 데이터 없음",
        value: row.asset_id,
      });
    }

    if (quantity <= 0) {
      issues.push({
        severity: "WARNING",
        issue_type: "zero_quantity",
        sheet: "App_Holdings",
        row: row.__row,
        key: row.holding_id,
        message: "수량이 0 이하",
        value: row.quantity,
      });
    }

    if (avgPrice < 0) {
      issues.push({
        severity: "ERROR",
        issue_type: "negative_avg_price",
        sheet: "App_Holdings",
        row: row.__row,
        key: row.holding_id,
        message: "평단가가 음수",
        value: row.avg_price,
      });
    }

    if (targetWeight < 0 || targetWeight > 1) {
      issues.push({
        severity: "WARNING",
        issue_type: "target_weight_out_of_range",
        sheet: "App_Holdings",
        row: row.__row,
        key: row.holding_id,
        message: "목표비중이 0~100% 범위를 벗어남",
        value: row.target_weight_account,
      });
    }
  });
}

function validateOutput_(issues, holdings, outputs) {
  const activeHoldings = holdings.filter((h) => {
    return String(h.enabled || "").toUpperCase() === "TRUE";
  });

  if (outputs.length !== activeHoldings.length) {
    issues.push({
      severity: "WARNING",
      issue_type: "output_row_count_mismatch",
      sheet: "App_Output",
      row: "",
      key: "",
      message: `활성 보유종목 수(${activeHoldings.length})와 App_Output 행 수(${outputs.length}) 불일치`,
      value: `${activeHoldings.length} vs ${outputs.length}`,
    });
  }

  outputs.forEach((row) => {
    const quantity = parseNumber_(row.quantity);
    const avgPrice = parseNumber_(row.avg_price);
    const price = parseNumber_(row.price);
    const investedAmount = parseNumber_(row.invested_amount);
    const valuationAmount = parseNumber_(row.valuation_amount);

    const expectedInvested = quantity * avgPrice;
    const expectedValuation = quantity * price;

    if (Math.abs(investedAmount - expectedInvested) > 1) {
      issues.push({
        severity: "WARNING",
        issue_type: "invested_amount_mismatch",
        sheet: "App_Output",
        row: row.__row,
        key: row.asset_id,
        message: "매입금액 계산값 불일치",
        value: `${investedAmount} vs ${expectedInvested}`,
      });
    }

    if (Math.abs(valuationAmount - expectedValuation) > 1) {
      issues.push({
        severity: "WARNING",
        issue_type: "valuation_amount_mismatch",
        sheet: "App_Output",
        row: row.__row,
        key: row.asset_id,
        message: "평가액 계산값 불일치",
        value: `${valuationAmount} vs ${expectedValuation}`,
      });
    }
  });
}

function validateAccountTargetWeights_(issues, holdings, accounts) {
  const accountNameById = {};
  accounts.forEach((a) => {
    accountNameById[a.account_id] = a.account_name;
  });

  const sumByAccount = {};

  holdings.forEach((h) => {
    if (String(h.enabled || "").toUpperCase() !== "TRUE") return;

    const accountId = String(h.account_id || "");
    const targetWeight = parsePercent_(h.target_weight_account);

    if (!sumByAccount[accountId]) sumByAccount[accountId] = 0;
    sumByAccount[accountId] += targetWeight;
  });

  Object.keys(sumByAccount).forEach((accountId) => {
    const sum = sumByAccount[accountId];

    if (sum > 1.05 || sum < 0.95) {
      issues.push({
        severity: "WARNING",
        issue_type: "account_target_weight_sum_not_100",
        sheet: "App_Holdings",
        row: "",
        key: accountId,
        message: `계좌 목표비중 합계가 100%에서 벗어남: ${accountNameById[accountId] || accountId}`,
        value: sum,
      });
    }
  });
}

function buildValidationSummary_(issues, accounts, assets, holdings, prices, outputs) {
  const errorCount = issues.filter((i) => i.severity === "ERROR").length;
  const warningCount = issues.filter((i) => i.severity === "WARNING").length;

  const activeHoldings = holdings.filter((h) => {
    return String(h.enabled || "").toUpperCase() === "TRUE";
  });

  const totalValuation = outputs.reduce((sum, row) => {
    return sum + parseNumber_(row.valuation_amount);
  }, 0);

  return [
    ["metric", "value"],
    ["검산시각", new Date()],
    ["ERROR", errorCount],
    ["WARNING", warningCount],
    ["계좌 수", accounts.length],
    ["종목 마스터 수", assets.length],
    ["보유종목 수", holdings.length],
    ["활성 보유종목 수", activeHoldings.length],
    ["가격 데이터 수", prices.length],
    ["App_Output 행 수", outputs.length],
    ["App_Output 총 평가액", totalValuation],
  ];
}

function writeValidationReport_(ss, issues, summary) {
  let sheet = ss.getSheetByName(ED_MVP_VALIDATE.sheets.report);

  if (!sheet) {
    sheet = ss.insertSheet(ED_MVP_VALIDATE.sheets.report);
  }

  sheet.clear();

  if (summary && summary.length > 0) {
    sheet.getRange(1, 1, summary.length, summary[0].length).setValues(summary);
  }

  const startRow = summary && summary.length > 0 ? summary.length + 3 : 1;

  const headers = [
    "severity",
    "issue_type",
    "sheet",
    "row",
    "key",
    "message",
    "value",
  ];

  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);

  const rows = issues.map((issue) => [
    issue.severity,
    issue.issue_type,
    issue.sheet,
    issue.row,
    issue.key,
    issue.message,
    issue.value,
  ]);

  if (rows.length > 0) {
    sheet.getRange(startRow + 1, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(startRow);
  sheet.autoResizeColumns(1, headers.length);

  sheet
    .getRange(startRow, 1, 1, headers.length)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#137333")
    .setHorizontalAlignment("center");
}

function readSheetAsObjectsWithRow_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  return values.slice(1).map((row, index) => {
    const obj = {};

    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    obj.__row = index + 2;

    return obj;
  });
}

function parseNumber_(value) {
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

function parsePercent_(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return value > 1 ? value / 100 : value;
  }

  const text = String(value).trim();

  if (text.indexOf("%") >= 0) {
    return parseNumber_(text) / 100;
  }

  const num = parseNumber_(text);

  return num > 1 ? num / 100 : num;
}