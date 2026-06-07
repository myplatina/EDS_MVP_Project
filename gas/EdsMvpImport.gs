/*******************************************************
 * ED's MVP - Main Sheet to App Sheets Import Script v0.8.61
 *
 * 목적:
 * - 기존 "2. 종목현황" 시트 데이터를 App_* 시트로 초기 이관
 * - 기존 원장 시트는 수정하지 않음
 * - 원화/달러 2중 가격 컬럼 구조 대응
 * - 미국 ETF의 달러 현재가 + 원화 평가액 구조를 앱용 원화 평가단가로 변환
 * - 하단 보조/요약 영역을 보유종목으로 잘못 이관하지 않도록 필터 유지
 *
 * 주요 실행 함수:
 * - resetAndRunInitialImportAndBuildOutput()
 * - runInitialImportAndBuildOutput()
 * - importMainSheetToApp()
 * - buildAppOutputFromAppSheets()
 *******************************************************/

const ED_MVP_IMPORT = {
  mainSheetNameCandidates: [
    "2. 종목현황",
    "2.종목현황",
    "종목현황",
  ],

  scanRowsForHeader: 120,
  minHeaderScore: 5,

  /**
   * 현재 원본 2. 종목현황에서 실제 보유종목은 9~41행 영역.
   * 하단 100행 이후에는 포트폴리오 차트/보조 데이터가 있어 오인식 방지 필요.
   * 향후 원장 구조가 바뀌면 이 값을 조정.
   */
  maxMainDataRow: 90,

  /**
   * 실제 보유종목을 한 번이라도 찾은 뒤,
   * 유효하지 않은 행이 이 횟수 이상 연속되면 보유종목 테이블 종료로 간주.
   */
  stopAfterInvalidStreakAfterFirstHolding: 25,

  sheets: {
    accounts: "App_Accounts",
    assets: "App_Assets",
    holdings: "App_Holdings",
    prices: "App_Prices",
    output: "App_Output",
    syncLog: "App_SyncLog",
  },
};

/**
 * 추천 실행 함수:
 * 기존 이관 데이터 초기화 → 기존 원장 이관 → App_Output 생성
 */
function resetAndRunInitialImportAndBuildOutput() {
  edImport_clearImportedAppData_();
  importMainSheetToApp();
  buildAppOutputFromAppSheets();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "초기화 후 이관 및 App_Output 생성 완료",
    "ED's MVP",
    5
  );
}

/**
 * 기존 데이터는 유지하면서 이관/출력만 재실행
 */
function runInitialImportAndBuildOutput() {
  importMainSheetToApp();
  buildAppOutputFromAppSheets();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "초기 이관 및 App_Output 생성 완료",
    "ED's MVP",
    5
  );
}

/**
 * App_* 이관 대상 시트의 2행 이하 데이터 삭제
 * 헤더는 유지
 */
function edImport_clearImportedAppData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheets = [
    ED_MVP_IMPORT.sheets.accounts,
    ED_MVP_IMPORT.sheets.assets,
    ED_MVP_IMPORT.sheets.holdings,
    ED_MVP_IMPORT.sheets.prices,
    ED_MVP_IMPORT.sheets.output,
  ];

  targetSheets.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow > 1 && lastCol > 0) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
  });

  edImport_writeLog_(
    "manual",
    "success",
    "App_*",
    "App_*",
    "",
    "이관 대상 App_* 시트 2행 이하 데이터 초기화 완료"
  );
}

/**
 * 기존 2. 종목현황 데이터를 App_* 시트로 이관
 */
function importMainSheetToApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = edImport_findMainSheet_(ss);

  if (!mainSheet) {
    ss.toast("기존 원장 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("main_to_app", "error", "", "", "", "기존 원장 시트 없음");
    return;
  }

  const headerInfo = edImport_detectHeader_(mainSheet);

  if (!headerInfo) {
    ss.toast("기존 원장 헤더를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("main_to_app", "error", mainSheet.getName(), "", "", "헤더 탐지 실패");
    return;
  }

  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();

  if (lastRow < headerInfo.dataStartRow) {
    ss.toast("기존 원장에 데이터 행이 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("main_to_app", "warning", mainSheet.getName(), "App_*", "", "데이터 행 없음");
    return;
  }

  const scanEndRow = Math.min(lastRow, ED_MVP_IMPORT.maxMainDataRow);
  const rowCount = scanEndRow - headerInfo.dataStartRow + 1;

  if (rowCount <= 0) {
    ss.toast("스캔 가능한 원장 데이터 행이 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("main_to_app", "warning", mainSheet.getName(), "App_*", "", "스캔 행 없음");
    return;
  }

  const values = mainSheet
    .getRange(headerInfo.dataStartRow, 1, rowCount, lastCol)
    .getValues();

  const displayValues = mainSheet
    .getRange(headerInfo.dataStartRow, 1, rowCount, lastCol)
    .getDisplayValues();

  const accountsMap = new Map();
  const assetsMap = new Map();
  const holdingsMap = new Map();
  const pricesMap = new Map();

  let lastBroker = "";
  let lastAccountName = "";
  let lastAccountType = "";
  let lastCountry = "";

  let foundHoldingOnce = false;
  let invalidStreakAfterFirstHolding = 0;
  let skippedInvalidTicker = 0;
  let skippedInvalidAssetName = 0;
  let skippedNonHoldingRow = 0;
  let skippedOutOfRange = 0;
  let krwPriceCount = 0;
  let fxDerivedPriceCount = 0;
  let zeroPriceCount = 0;

  for (let i = 0; i < values.length; i++) {
    const absoluteRow = headerInfo.dataStartRow + i;

    if (absoluteRow > ED_MVP_IMPORT.maxMainDataRow) {
      skippedOutOfRange++;
      break;
    }

    const row = values[i];
    const displayRow = displayValues[i];

    const broker = edImport_getCell_(row, displayRow, headerInfo.colMap.broker);
    const accountName = edImport_getCell_(row, displayRow, headerInfo.colMap.account);
    const accountTypeRaw = edImport_getCell_(row, displayRow, headerInfo.colMap.accountType);
    const countryRaw = edImport_getCell_(row, displayRow, headerInfo.colMap.country);

    const ticker = edImport_normalizeTicker_(
      edImport_getCell_(row, displayRow, headerInfo.colMap.ticker)
    );

    const assetName = edImport_getCell_(row, displayRow, headerInfo.colMap.assetName);

    const quantity = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.quantity)
    );

    const avgPriceKrw = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.avgPriceKrw)
    );

    const avgPriceUsd = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.avgPriceUsd)
    );

    const priceKrw = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.priceKrw)
    );

    const priceUsd = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.priceUsd)
    );

    const valuationAmountKrw = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.valuationAmount)
    );

    const changeAmount = edImport_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.changeAmount)
    );

    const changeRate = edImport_parsePercent_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.changeRate)
    );

    const targetWeight = edImport_parsePercent_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.targetWeight)
    );

    if (broker) lastBroker = broker;
    if (accountName) lastAccountName = accountName;
    if (accountTypeRaw) lastAccountType = accountTypeRaw;
    if (countryRaw) lastCountry = countryRaw;

    const effectiveBroker = broker || lastBroker;
    const effectiveAccountName = accountName || lastAccountName;
    const effectiveAccountTypeRaw =
      accountTypeRaw ||
      lastAccountType ||
      edImport_inferAccountType_(effectiveAccountName);

    const effectiveCountryRaw = countryRaw || lastCountry;

    const normalizedAccountType = edImport_normalizeAccountType_(effectiveAccountTypeRaw);
    const normalizedCountry = edImport_normalizeCountry_(effectiveCountryRaw, ticker, assetName);

    const effectivePriceInfo = edImport_getEffectiveKrwPriceInfo_({
      quantity,
      priceKrw,
      priceUsd,
      valuationAmountKrw,
    });

    const priceForApp = effectivePriceInfo.priceKrw;

    const avgPriceForApp = edImport_getEffectiveKrwAvgPrice_({
      avgPriceKrw,
      avgPriceUsd,
      priceUsd,
      effectiveKrwPrice: priceForApp,
    });

    const validation = edImport_validateHoldingCandidate_({
      absoluteRow,
      ticker,
      assetName,
      quantity,
      avgPrice: avgPriceForApp,
      price: priceForApp,
      effectiveAccountName,
      normalizedAccountType,
    });

    if (!validation.valid) {
      if (validation.reason === "invalid_ticker") skippedInvalidTicker++;
      if (validation.reason === "invalid_asset_name") skippedInvalidAssetName++;
      if (validation.reason === "non_holding_row") skippedNonHoldingRow++;

      if (foundHoldingOnce) {
        invalidStreakAfterFirstHolding++;
        if (invalidStreakAfterFirstHolding >= ED_MVP_IMPORT.stopAfterInvalidStreakAfterFirstHolding) {
          break;
        }
      }

      continue;
    }

    foundHoldingOnce = true;
    invalidStreakAfterFirstHolding = 0;

    if (effectivePriceInfo.source === "krw_price") krwPriceCount++;
    if (effectivePriceInfo.source === "valuation_div_quantity") fxDerivedPriceCount++;
    if (priceForApp <= 0) zeroPriceCount++;

    const market = edImport_inferMarket_(normalizedCountry, ticker);
    const assetCurrency = edImport_inferCurrency_(normalizedCountry, market);
    const assetClass = edImport_normalizeAssetClass_(edImport_inferAssetClass_(assetName));
    const isEtf = edImport_inferIsEtf_(assetName);

    const accountId = edImport_makeAccountId_(
      effectiveBroker,
      effectiveAccountName,
      normalizedAccountType
    );

    const assetId = edImport_makeAssetId_(market, ticker);
    const holdingId = `HLD_${accountId}_${assetId}`;
    const now = new Date();

    accountsMap.set(accountId, [
      accountId,
      effectiveAccountName || "미분류 계좌",
      effectiveBroker || "",
      normalizedAccountType,
      "KRW",
      accountsMap.size + 1,
      "TRUE",
      now,
      now,
    ]);

    assetsMap.set(assetId, [
      assetId,
      ticker,
      assetName,
      normalizedCountry,
      market,
      assetClass,
      assetCurrency,
      isEtf ? "TRUE" : "FALSE",
      "TRUE",
      now,
      now,
    ]);

    holdingsMap.set(holdingId, [
      holdingId,
      accountId,
      assetId,
      ticker,
      assetName,
      quantity || 0,
      avgPriceForApp || 0,
      targetWeight || 0,
      effectivePriceInfo.source === "valuation_div_quantity"
        ? "원본 달러 현재가 + 원화 평가액 기준 원화 평가단가 역산"
        : "",
      "TRUE",
      now,
      now,
    ]);

    pricesMap.set(assetId, [
      assetId,
      ticker,
      priceForApp || 0,
      "",
      changeAmount || 0,
      changeRate || 0,
      "KRW",
      "sheet",
      now,
      now,
    ]);
  }

  edImport_upsertRowsByKey_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.accounts),
    "account_id",
    Array.from(accountsMap.values()),
    true
  );

  edImport_upsertRowsByKey_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.assets),
    "asset_id",
    Array.from(assetsMap.values()),
    true
  );

  edImport_upsertRowsByKey_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.holdings),
    "holding_id",
    Array.from(holdingsMap.values()),
    true
  );

  edImport_upsertRowsByKey_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.prices),
    "asset_id",
    Array.from(pricesMap.values()),
    true
  );

  const message =
    `이관 완료: accounts=${accountsMap.size}, ` +
    `assets=${assetsMap.size}, holdings=${holdingsMap.size}, prices=${pricesMap.size}, ` +
    `krw_price=${krwPriceCount}, fx_derived_price=${fxDerivedPriceCount}, zero_price=${zeroPriceCount}, ` +
    `skipped_invalid_ticker=${skippedInvalidTicker}, ` +
    `skipped_invalid_asset_name=${skippedInvalidAssetName}, ` +
    `skipped_non_holding=${skippedNonHoldingRow}`;

  edImport_writeLog_("main_to_app", "success", mainSheet.getName(), "App_*", "", message);

  ss.toast(message, "ED's MVP", 8);
}

/**
 * App_Holdings / App_Accounts / App_Prices 기준으로 App_Output 생성
 */
function buildAppOutputFromAppSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const accounts = edImport_readSheetAsObjects_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.accounts)
  );

  const holdings = edImport_readSheetAsObjects_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.holdings)
  );

  const prices = edImport_readSheetAsObjects_(
    ss.getSheetByName(ED_MVP_IMPORT.sheets.prices)
  );

  const activeHoldings = holdings.filter((h) => {
    return String(h.enabled).toUpperCase() === "TRUE";
  });

  const rows = [];
  const accountTotals = {};
  let totalValuation = 0;

  activeHoldings.forEach((h) => {
    const priceRow = prices.find((p) => String(p.asset_id) === String(h.asset_id));
    const price = edImport_parseNumber_(priceRow ? priceRow.price : 0);
    const quantity = edImport_parseNumber_(h.quantity);
    const valuationAmount = quantity * price;

    if (!accountTotals[h.account_id]) accountTotals[h.account_id] = 0;

    accountTotals[h.account_id] += valuationAmount;
    totalValuation += valuationAmount;
  });

  activeHoldings.forEach((h) => {
    const account =
      accounts.find((a) => String(a.account_id) === String(h.account_id)) || {};

    const priceRow =
      prices.find((p) => String(p.asset_id) === String(h.asset_id)) || {};

    const quantity = edImport_parseNumber_(h.quantity);
    const avgPrice = edImport_parseNumber_(h.avg_price);
    const price = edImport_parseNumber_(priceRow.price);

    const investedAmount = quantity * avgPrice;
    const valuationAmount = quantity * price;
    const profitAmount = valuationAmount - investedAmount;
    const profitRate = investedAmount > 0 ? profitAmount / investedAmount : 0;

    const accountTotal = accountTotals[h.account_id] || 0;
    const accountWeight = accountTotal > 0 ? valuationAmount / accountTotal : 0;
    const totalWeight = totalValuation > 0 ? valuationAmount / totalValuation : 0;

    const targetWeight = edImport_parsePercent_(h.target_weight_account);
    const targetGapRate = accountWeight - targetWeight;
    const targetAmount = accountTotal * targetWeight;
    const targetGapAmount = targetAmount - valuationAmount;

    rows.push([
      h.account_id,
      account.account_name || "",
      account.broker || "",
      account.account_type || "",
      h.asset_id,
      h.ticker,
      h.asset_name,
      quantity,
      avgPrice,
      price,
      investedAmount,
      valuationAmount,
      profitAmount,
      profitRate,
      accountWeight,
      totalWeight,
      targetWeight,
      targetGapRate,
      targetGapAmount,
      priceRow.currency || "KRW",
      priceRow.source || "",
      priceRow.fetched_at || "",
      new Date(),
    ]);
  });

  const outputSheet = ss.getSheetByName(ED_MVP_IMPORT.sheets.output);

  if (!outputSheet) {
    throw new Error("App_Output 시트를 찾을 수 없습니다.");
  }

  const headers = outputSheet
    .getRange(1, 1, 1, outputSheet.getLastColumn())
    .getValues()[0];

  if (outputSheet.getLastRow() > 1) {
    outputSheet
      .getRange(2, 1, outputSheet.getLastRow() - 1, outputSheet.getLastColumn())
      .clearContent();
  }

  if (rows.length > 0) {
    outputSheet
      .getRange(2, 1, rows.length, headers.length)
      .setValues(rows);
  }

  edImport_writeLog_(
    "manual",
    "success",
    "App_Holdings/App_Prices",
    "App_Output",
    "",
    `App_Output 생성 완료: ${rows.length}건`
  );

  ss.toast(`App_Output 생성 완료: ${rows.length}건`, "ED's MVP", 5);
}

/**
 * 기존 원장 시트 찾기
 */
function edImport_findMainSheet_(ss) {
  for (const candidate of ED_MVP_IMPORT.mainSheetNameCandidates) {
    const sheet = ss.getSheetByName(candidate);
    if (sheet) return sheet;
  }

  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    const normalized = sheet.getName().replace(/\s/g, "");
    if (normalized === "2.종목현황" || normalized.includes("종목현황")) {
      return sheet;
    }
  }

  return null;
}

/**
 * 헤더 행 자동 탐지
 * - 2중 헤더 구조 대응
 * - headerRow: 상위 헤더 행
 * - subHeaderRow: 원화/달러 보조 헤더 행 가능
 * - dataStartRow: 실제 데이터 시작 행
 */
function edImport_detectHeader_(sheet) {
  const scanRows = Math.min(sheet.getLastRow(), ED_MVP_IMPORT.scanRowsForHeader);
  const lastCol = sheet.getLastColumn();

  const displayValues = sheet
    .getRange(1, 1, scanRows, lastCol)
    .getDisplayValues();

  let best = null;

  for (let r = 0; r < displayValues.length; r++) {
    const headerRow = displayValues[r];
    const subHeaderRow = r + 1 < displayValues.length ? displayValues[r + 1] : [];
    const colMap = edImport_buildColMap_(headerRow, subHeaderRow);
    const score = edImport_scoreColMap_(colMap);

    if (!best || score > best.score) {
      const hasSubHeader = edImport_rowLooksLikeSubHeader_(subHeaderRow);
      best = {
        headerRow: r + 1,
        subHeaderRow: hasSubHeader ? r + 2 : null,
        dataStartRow: hasSubHeader ? r + 3 : r + 2,
        colMap,
        score,
      };
    }
  }

  if (!best || best.score < ED_MVP_IMPORT.minHeaderScore) {
    return null;
  }

  return best;
}

function edImport_scoreColMap_(colMap) {
  const required = [
    "broker",
    "account",
    "country",
    "ticker",
    "assetName",
    "quantity",
  ];

  const useful = [
    "avgPriceKrw",
    "avgPriceUsd",
    "priceKrw",
    "priceUsd",
    "valuationAmount",
    "targetWeight",
  ];

  let score = 0;

  required.forEach((key) => {
    if (colMap[key] >= 0) score += 2;
  });

  useful.forEach((key) => {
    if (colMap[key] >= 0) score += 1;
  });

  return score;
}

function edImport_rowLooksLikeSubHeader_(row) {
  if (!row || row.length === 0) return false;

  const normalized = row.map((v) => String(v || "").replace(/\s/g, ""));
  const joined = normalized.join("|");

  return (
    joined.indexOf("원화") >= 0 ||
    joined.indexOf("달러") >= 0 ||
    joined.indexOf("USD") >= 0 ||
    joined.indexOf("KRW") >= 0
  );
}

/**
 * 헤더명 기준 컬럼 매핑
 * 반환값은 0-based index
 */
function edImport_buildColMap_(headerRow, subHeaderRow) {
  const normalizedHeaders = headerRow.map((h) =>
    String(h || "").replace(/\s/g, "")
  );

  const normalizedSubHeaders = (subHeaderRow || []).map((h) =>
    String(h || "").replace(/\s/g, "")
  );

  function findCol(patterns, excludePatterns) {
    excludePatterns = excludePatterns || [];

    for (let i = 0; i < normalizedHeaders.length; i++) {
      const h = normalizedHeaders[i];

      for (const p of patterns) {
        if (h === p) return i;
      }
    }

    for (let i = 0; i < normalizedHeaders.length; i++) {
      const h = normalizedHeaders[i];

      let excluded = false;
      for (const ex of excludePatterns) {
        if (h.indexOf(ex) >= 0) {
          excluded = true;
          break;
        }
      }

      if (excluded) continue;

      for (const p of patterns) {
        if (h.indexOf(p) >= 0) return i;
      }
    }

    return -1;
  }

  function findGroupStart(patterns) {
    return findCol(patterns, []);
  }

  function findSubColumnNear(groupStart, subPatterns, maxWidth, fallbackOffset) {
    if (groupStart < 0) return -1;

    const width = maxWidth || 3;
    const fallback = groupStart + (fallbackOffset || 0);

    for (let i = groupStart; i < Math.min(groupStart + width, normalizedSubHeaders.length); i++) {
      const h = normalizedSubHeaders[i];
      for (const p of subPatterns) {
        if (h === p || h.indexOf(p) >= 0) return i;
      }
    }

    return fallback < normalizedHeaders.length ? fallback : -1;
  }

  const avgGroupStart = findGroupStart(["평단가", "평균단가"]);
  const priceGroupStart = findGroupStart(["현재가"]);

  const avgPriceKrw = findSubColumnNear(avgGroupStart, ["원화", "KRW"], 3, 0);
  const avgPriceUsd = findSubColumnNear(avgGroupStart, ["달러", "USD"], 3, 1);

  const priceKrw = findSubColumnNear(priceGroupStart, ["원화", "KRW"], 3, 0);
  const priceUsd = findSubColumnNear(priceGroupStart, ["달러", "USD"], 3, 1);

  const map = {
    broker: findCol(["증권사"]),
    account: findCol(["계좌"], ["계좌형식", "계좌유형", "투자비중", "목표비중"]),
    accountType: findCol(["계좌형식", "계좌유형"]),
    country: findCol(["국가"]),
    ticker: findCol(["종목코드", "티커"]),
    assetName: findCol(["종목명"]),
    quantity: findCol(["수량"]),
    avgPriceKrw,
    avgPriceUsd,
    priceKrw,
    priceUsd,
    valuationAmount: findCol(["평가액"]),
    changeAmount: findCol(["전일대비등락", "등락액", "전일대비"]),
    changeRate: findCol(["등락률", "전일대비율"]),
    accountWeight: findCol(["투자비중(계좌내)", "계좌내"]),
    totalWeight: findCol(["투자비중(전체자산내)", "전체자산내"]),
    targetWeight: findCol(["목표비중"]),
  };

  // 기존 Preview/Sync 스크립트 호환용 alias.
  map.avgPrice = map.avgPriceKrw;
  map.price = map.priceKrw;

  return map;
}

function edImport_getEffectiveKrwPriceInfo_(args) {
  const quantity = edImport_parseNumber_(args.quantity);
  const priceKrw = edImport_parseNumber_(args.priceKrw);
  const priceUsd = edImport_parseNumber_(args.priceUsd);
  const valuationAmountKrw = edImport_parseNumber_(args.valuationAmountKrw);

  if (priceKrw > 0) {
    return {
      priceKrw,
      source: "krw_price",
    };
  }

  if (quantity > 0 && valuationAmountKrw > 0) {
    return {
      priceKrw: valuationAmountKrw / quantity,
      source: priceUsd > 0 ? "valuation_div_quantity" : "valuation_div_quantity_no_native_price",
    };
  }

  return {
    priceKrw: 0,
    source: "missing_price",
  };
}

function edImport_getEffectiveKrwAvgPrice_(args) {
  const avgPriceKrw = edImport_parseNumber_(args.avgPriceKrw);
  const avgPriceUsd = edImport_parseNumber_(args.avgPriceUsd);
  const priceUsd = edImport_parseNumber_(args.priceUsd);
  const effectiveKrwPrice = edImport_parseNumber_(args.effectiveKrwPrice);

  if (avgPriceKrw > 0) return avgPriceKrw;

  // 달러 평단만 있고, 달러 현재가와 원화 환산 현재가가 있으면 동일 환율로 원화 평단 추정.
  // 1차 MVP용 추정값이며, 2차에서는 native_price/native_currency/fx_rate 분리 권장.
  if (avgPriceUsd > 0 && priceUsd > 0 && effectiveKrwPrice > 0) {
    const impliedFxRate = effectiveKrwPrice / priceUsd;
    return avgPriceUsd * impliedFxRate;
  }

  return 0;
}

/**
 * 보유종목 후보 행 검증
 */
function edImport_validateHoldingCandidate_(candidate) {
  const ticker = String(candidate.ticker || "").trim();
  const assetName = String(candidate.assetName || "").trim();
  const accountName = String(candidate.effectiveAccountName || "").trim();
  const accountType = String(candidate.normalizedAccountType || "").trim();

  if (edImport_isNonHoldingRow_(ticker, assetName)) {
    return { valid: false, reason: "non_holding_row" };
  }

  if (!edImport_isValidTicker_(ticker)) {
    return { valid: false, reason: "invalid_ticker" };
  }

  if (!edImport_isValidAssetName_(assetName)) {
    return { valid: false, reason: "invalid_asset_name" };
  }

  if (!accountName) {
    return { valid: false, reason: "non_holding_row" };
  }

  if (!edImport_isValidAccountTypeForHolding_(accountType)) {
    return { valid: false, reason: "non_holding_row" };
  }

  if (!candidate.quantity && !candidate.avgPrice && !candidate.price) {
    return { valid: false, reason: "non_holding_row" };
  }

  return { valid: true, reason: "valid" };
}

/**
 * 허용 종목코드:
 * - 미국 티커: 알파벳 1~5자
 * - 한국 정규 종목코드: 숫자 6자리
 * - 한국 ETF 특수코드: 숫자+영문 혼합 6자리
 */
function edImport_isValidTicker_(ticker) {
  const t = String(ticker || "").trim().toUpperCase();

  if (!t) return false;

  if (/^[A-Z]{1,5}$/.test(t)) return true;
  if (/^\d{6}$/.test(t)) return true;
  if (/^[0-9A-Z]{6}$/.test(t) && /\d/.test(t) && /[A-Z]/.test(t)) return true;

  return false;
}

/**
 * 종목명 검증:
 * - 빈 값 제외
 * - 숫자만 있는 값 제외
 * - 금액/비율처럼 보이는 값 제외
 */
function edImport_isValidAssetName_(assetName) {
  const n = String(assetName || "").trim();

  if (!n) return false;

  const numericLike = n
    .replace(/,/g, "")
    .replace(/원/g, "")
    .replace(/%/g, "")
    .replace(/\s/g, "");

  if (/^\d+(\.\d+)?$/.test(numericLike)) return false;

  return true;
}

function edImport_isValidAccountTypeForHolding_(accountType) {
  const allowed = ["일반", "ISA", "개인연금", "IRP"];
  return allowed.indexOf(String(accountType || "")) >= 0;
}

function edImport_getCell_(row, displayRow, colIndex) {
  if (colIndex === undefined || colIndex < 0) return "";

  const displayValue = displayRow[colIndex];
  const rawValue = row[colIndex];

  return String(displayValue || rawValue || "").trim();
}

function edImport_getCellRaw_(row, displayRow, colIndex) {
  if (colIndex === undefined || colIndex < 0) return "";

  const rawValue = row[colIndex];

  if (rawValue !== "" && rawValue !== null && rawValue !== undefined) {
    return rawValue;
  }

  return displayRow[colIndex];
}

function edImport_normalizeTicker_(value) {
  let ticker = String(value || "").trim();

  if (!ticker) return "";

  ticker = ticker.replace(/\s/g, "");

  if (/^\d+$/.test(ticker) && ticker.length < 6) {
    ticker = ticker.padStart(6, "0");
  }

  return ticker.toUpperCase();
}

function edImport_parseNumber_(value) {
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

function edImport_parsePercent_(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return value > 1 ? value / 100 : value;
  }

  const text = String(value).trim();

  if (text.indexOf("%") >= 0) {
    return edImport_parseNumber_(text) / 100;
  }

  const num = edImport_parseNumber_(text);

  return num > 1 ? num / 100 : num;
}

function edImport_isNonHoldingRow_(ticker, assetName) {
  const t = String(ticker || "").trim();
  const n = String(assetName || "").trim();

  if (!t && !n) return true;
  if (t === "현금" || n === "현금") return true;
  if (n.indexOf("합계") >= 0 || n.indexOf("총계") >= 0) return true;
  if (t.indexOf("합계") >= 0 || t.indexOf("총계") >= 0) return true;

  return false;
}

function edImport_inferAccountType_(accountName) {
  return edImport_normalizeAccountType_(accountName);
}

/**
 * App_Accounts.account_type 검증 규칙 대응:
 * 허용값 = 일반, ISA, 개인연금, IRP, 기타
 */
function edImport_normalizeAccountType_(value) {
  const text = String(value || "").trim();

  if (!text) return "기타";

  if (text.indexOf("ISA") >= 0 || text.indexOf("isa") >= 0) return "ISA";
  if (text.indexOf("IRP") >= 0 || text.indexOf("irp") >= 0) return "IRP";
  if (text.indexOf("연금") >= 0) return "개인연금";
  if (text.indexOf("일반") >= 0) return "일반";

  return "기타";
}

/**
 * App_Assets.country 검증 규칙 대응:
 * 허용값 = 한국, 미국, 기타
 */
function edImport_normalizeCountry_(value, ticker, assetName) {
  const raw = String(value || "").trim().toLowerCase();
  const t = String(ticker || "").trim().toUpperCase();
  const n = String(assetName || "").trim().toLowerCase();

  const combined = `${raw} ${t} ${n}`;

  if (
    combined.indexOf("미국") >= 0 ||
    combined.indexOf("usa") >= 0 ||
    combined.indexOf("unitedstates") >= 0 ||
    combined.indexOf("united states") >= 0 ||
    combined.indexOf("s&p") >= 0 ||
    combined.indexOf("sp500") >= 0 ||
    combined.indexOf("nasdaq") >= 0 ||
    combined.indexOf("나스닥") >= 0 ||
    combined.indexOf("다우") >= 0 ||
    combined.indexOf("schd") >= 0 ||
    combined.indexOf("qqq") >= 0 ||
    combined.indexOf("spy") >= 0
  ) {
    return "미국";
  }

  if (
    combined.indexOf("한국") >= 0 ||
    combined.indexOf("국내") >= 0 ||
    combined.indexOf("korea") >= 0 ||
    combined.indexOf("krx") >= 0 ||
    /^\d{6}$/.test(t)
  ) {
    return "한국";
  }

  return "기타";
}

/**
 * App_Assets.market 검증 규칙 대응:
 * 허용값 = KRX, NYSE, NASDAQ, AMEX, 기타
 */
function edImport_inferMarket_(country, ticker) {
  const c = String(country || "");
  const t = String(ticker || "").toUpperCase();

  if (/^\d{6}$/.test(t)) return "KRX";
  if (/^[0-9A-Z]{6}$/.test(t) && /\d/.test(t) && /[A-Z]/.test(t)) return "KRX";
  if (c.indexOf("한국") >= 0) return "KRX";

  if (c.indexOf("미국") >= 0) {
    if (["QQQ", "QQQM", "TQQQ", "ONEQ"].indexOf(t) >= 0) return "NASDAQ";
    if (["SPY", "VOO", "VTI", "SCHD", "DIA"].indexOf(t) >= 0) return "NYSE";
    if (["SPYM"].indexOf(t) >= 0) return "AMEX";
    return "기타";
  }

  return "기타";
}

/**
 * App_Assets.currency 기준:
 * - App_Assets.currency는 실제 거래 통화
 * - App_Prices.currency는 앱 평가용 단가 통화로 KRW 고정
 */
function edImport_inferCurrency_(country, market) {
  if (market === "KRX") return "KRW";

  if (String(country || "").indexOf("미국") >= 0) return "USD";

  return "KRW";
}

/**
 * App_Assets.asset_class 검증 규칙 대응:
 * 허용값 = 주식, 주식 ETF, 채권, 채권 ETF, 혼합 ETF, 현금, 기타
 */
function edImport_inferAssetClass_(assetName) {
  const name = String(assetName || "");

  if (name.indexOf("현금") >= 0) return "현금";
  if (name.indexOf("채권혼합") >= 0) return "혼합 ETF";
  if (name.indexOf("혼합") >= 0) return "혼합 ETF";
  if (name.indexOf("채권") >= 0) return "채권 ETF";

  if (
    name.indexOf("ETF") >= 0 ||
    name.indexOf("TIGER") >= 0 ||
    name.indexOf("KODEX") >= 0 ||
    name.indexOf("RISE") >= 0 ||
    name.indexOf("ACE") >= 0 ||
    name.indexOf("SOL") >= 0 ||
    name.indexOf("HANARO") >= 0 ||
    name.indexOf("KOSEF") >= 0 ||
    name.indexOf("TIMEFOLIO") >= 0 ||
    name.indexOf("SCHD") >= 0 ||
    name.indexOf("QQQ") >= 0 ||
    name.indexOf("SPY") >= 0
  ) {
    return "주식 ETF";
  }

  return "주식";
}

function edImport_normalizeAssetClass_(value) {
  const text = String(value || "").trim();

  const allowed = [
    "주식",
    "주식 ETF",
    "채권",
    "채권 ETF",
    "혼합 ETF",
    "현금",
    "기타",
  ];

  if (allowed.indexOf(text) >= 0) return text;

  return "기타";
}

function edImport_inferIsEtf_(assetName) {
  const name = String(assetName || "");

  return (
    name.indexOf("ETF") >= 0 ||
    name.indexOf("TIGER") >= 0 ||
    name.indexOf("KODEX") >= 0 ||
    name.indexOf("RISE") >= 0 ||
    name.indexOf("ACE") >= 0 ||
    name.indexOf("SOL") >= 0 ||
    name.indexOf("HANARO") >= 0 ||
    name.indexOf("KOSEF") >= 0 ||
    name.indexOf("TIMEFOLIO") >= 0 ||
    name.indexOf("SCHD") >= 0 ||
    name.indexOf("QQQ") >= 0 ||
    name.indexOf("SPY") >= 0
  );
}

function edImport_makeAccountId_(broker, accountName, accountType) {
  const raw = `${broker}_${accountName}_${accountType}`;
  return "ACC_" + edImport_slug_(raw);
}

function edImport_makeAssetId_(market, ticker) {
  return `${market}_${edImport_slug_(ticker)}`;
}

function edImport_slug_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^\w가-힣]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();
}

/**
 * key 기준 upsert
 * overwrite=true면 기존 행도 업데이트
 */
function edImport_upsertRowsByKey_(sheet, keyHeader, rows, overwrite) {
  if (!sheet) throw new Error(`Sheet not found for upsert: ${keyHeader}`);
  if (!rows || rows.length === 0) return;

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const keyCol = headers.indexOf(keyHeader);

  if (keyCol < 0) {
    throw new Error(`Key header not found: ${keyHeader}`);
  }

  const lastRow = sheet.getLastRow();
  const existingMap = new Map();

  if (lastRow >= 2) {
    const existingValues = sheet
      .getRange(2, 1, lastRow - 1, headers.length)
      .getValues();

    existingValues.forEach((row, idx) => {
      const key = String(row[keyCol] || "");
      if (key) existingMap.set(key, idx + 2);
    });
  }

  const appendRows = [];

  rows.forEach((row) => {
    const key = String(row[keyCol] || "");

    if (existingMap.has(key)) {
      if (overwrite) {
        const rowIndex = existingMap.get(key);
        sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      }
    } else {
      appendRows.push(row);
    }
  });

  if (appendRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, appendRows.length, appendRows[0].length)
      .setValues(appendRows);
  }
}

function edImport_readSheetAsObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  return values.slice(1).map((row) => {
    const obj = {};

    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    return obj;
  });
}

function edImport_writeLog_(syncType, status, sourceSheet, targetSheet, targetKey, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_IMPORT.sheets.syncLog);

  if (!sheet) return;

  const now = new Date();

  const logId =
    `LOG_${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_` +
    `${Math.floor(Math.random() * 10000)}`;

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
