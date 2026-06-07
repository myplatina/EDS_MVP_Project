/*******************************************************
 * ED's MVP - App to Main Sheet Sync Preview v0.8.61
 *
 * 목적:
 * - App_Holdings 데이터를 기존 "2. 종목현황"에 반영하기 전 미리보기 생성
 * - 기존 2. 종목현황 시트는 수정하지 않음
 * - EdsMvpImport.gs v0.4의 원화/달러 2중 가격 컬럼 구조와 동일한 방식으로 비교
 *
 * 주요 변경점 v0.2:
 * - 기존 원장의 평단가 원화/달러, 현재가 원화/달러, 평가액[원화] 구조 반영
 * - 미국 ETF의 앱용 가격은 평가액[원화] ÷ 수량 기준으로 비교
 * - App_Prices.price와 기존 원장의 앱용 원화 평가단가 비교 추가
 * - App_Holdings.avg_price와 기존 원장의 앱용 원화 평단가 비교
 *
 * 실행 함수:
 * - previewSyncAppHoldingsToMainSheet()
 *
 * 전제:
 * - EdsMvpImport.gs v0.4가 같은 Apps Script 프로젝트에 존재해야 함
 *******************************************************/

const ED_MVP_SYNC_PREVIEW = {
  sheets: {
    holdings: "App_Holdings",
    prices: "App_Prices",
    preview: "App_SyncPreview",
    syncLog: "App_SyncLog",
  },

  tolerance: {
    quantity: 0.000001,
    price: 1,
    percent: 0.000001,
  },
};

/**
 * App_Holdings → 2. 종목현황 동기화 미리보기
 * 실제 원장은 수정하지 않음
 */
function previewSyncAppHoldingsToMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mainSheet = edImport_findMainSheet_(ss);
  if (!mainSheet) {
    ss.toast("기존 원장 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("manual", "error", "App_Holdings", "2. 종목현황", "", "기존 원장 시트 없음");
    return;
  }

  const headerInfo = edImport_detectHeader_(mainSheet);
  if (!headerInfo) {
    ss.toast("기존 원장 헤더를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("manual", "error", "App_Holdings", mainSheet.getName(), "", "기존 원장 헤더 탐지 실패");
    return;
  }

  const holdingsSheet = ss.getSheetByName(ED_MVP_SYNC_PREVIEW.sheets.holdings);
  const pricesSheet = ss.getSheetByName(ED_MVP_SYNC_PREVIEW.sheets.prices);

  if (!holdingsSheet) {
    ss.toast("App_Holdings 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("manual", "error", "App_Holdings", mainSheet.getName(), "", "App_Holdings 없음");
    return;
  }

  if (!pricesSheet) {
    ss.toast("App_Prices 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("manual", "error", "App_Prices", mainSheet.getName(), "", "App_Prices 없음");
    return;
  }

  const appHoldings = edSyncPreview_readSheetAsObjects_(holdingsSheet);
  const appPrices = edSyncPreview_readSheetAsObjects_(pricesSheet);
  const appPriceByAssetId = edSyncPreview_makeMapByKey_(appPrices, "asset_id");

  const mainIndex = edSyncPreview_buildMainSheetIndex_(mainSheet, headerInfo);

  const previewRows = [];
  let matchedCount = 0;
  let unmatchedCount = 0;
  let duplicateCount = 0;
  let updateCount = 0;
  let noChangeCount = 0;
  let priceMismatchCount = 0;

  appHoldings.forEach((holding) => {
    if (String(holding.enabled || "").toUpperCase() !== "TRUE") return;

    const holdingId = String(holding.holding_id || "");
    const matches = mainIndex.byHoldingId.get(holdingId) || [];

    const appPriceRow = appPriceByAssetId.get(String(holding.asset_id || "")) || {};

    const appQuantity = edSyncPreview_parseNumber_(holding.quantity);
    const appAvgPriceKrw = edSyncPreview_parseNumber_(holding.avg_price);
    const appTargetWeight = edSyncPreview_parsePercent_(holding.target_weight_account);
    const appPriceKrw = edSyncPreview_parseNumber_(appPriceRow.price);

    if (matches.length === 0) {
      unmatchedCount++;

      previewRows.push([
        "UNMATCHED_APP_HOLDING",
        "NOT_FOUND",
        "",
        holdingId,
        holding.account_id || "",
        holding.asset_id || "",
        holding.ticker || "",
        holding.asset_name || "",
        "",
        appQuantity,
        "",
        appAvgPriceKrw,
        "",
        appPriceKrw,
        "",
        appTargetWeight,
        "",
        "",
        "기존 2. 종목현황에서 대응 행을 찾지 못함",
      ]);

      return;
    }

    if (matches.length > 1) {
      duplicateCount++;
    }

    const match = matches[0];

    const quantityChanged = Math.abs(match.quantity - appQuantity) > ED_MVP_SYNC_PREVIEW.tolerance.quantity;
    const avgPriceChanged = Math.abs(match.avgPriceKrwForApp - appAvgPriceKrw) > ED_MVP_SYNC_PREVIEW.tolerance.price;
    const appPriceChanged = Math.abs(match.priceKrwForApp - appPriceKrw) > ED_MVP_SYNC_PREVIEW.tolerance.price;
    const targetWeightChanged = Math.abs(match.targetWeight - appTargetWeight) > ED_MVP_SYNC_PREVIEW.tolerance.percent;

    const changed = quantityChanged || avgPriceChanged || targetWeightChanged;

    if (appPriceChanged) {
      priceMismatchCount++;
    }

    if (changed) {
      updateCount++;
    } else {
      noChangeCount++;
    }

    matchedCount++;

    const diffSummary = edSyncPreview_buildDiffSummary_({
      quantityChanged,
      avgPriceChanged,
      appPriceChanged,
      targetWeightChanged,
      duplicate: matches.length > 1,
      mainPriceSource: match.priceSource,
    });

    previewRows.push([
      changed ? "UPDATE_CANDIDATE" : "NO_CHANGE",
      matches.length > 1 ? "DUPLICATE_MATCH" : "MATCHED",
      match.rowIndex,
      holdingId,
      holding.account_id || "",
      holding.asset_id || "",
      holding.ticker || "",
      holding.asset_name || "",
      match.quantity,
      appQuantity,
      match.avgPriceKrwForApp,
      appAvgPriceKrw,
      match.priceKrwForApp,
      appPriceKrw,
      match.targetWeight,
      appTargetWeight,
      match.accountName,
      match.priceSource,
      diffSummary,
    ]);
  });

  edSyncPreview_writePreviewSheet_(ss, previewRows, {
    mainSheetName: mainSheet.getName(),
    matchedCount,
    unmatchedCount,
    duplicateCount,
    updateCount,
    noChangeCount,
    priceMismatchCount,
  });

  const message =
    `동기화 미리보기 완료: matched=${matchedCount}, ` +
    `unmatched=${unmatchedCount}, duplicate=${duplicateCount}, ` +
    `update=${updateCount}, noChange=${noChangeCount}, ` +
    `price_mismatch=${priceMismatchCount}`;

  edImport_writeLog_("manual", "success", "App_Holdings/App_Prices", mainSheet.getName(), "", message);

  ss.toast(message, "ED's MVP", 8);
}

/**
 * 기존 2. 종목현황의 보유종목 행을 holding_id 기준으로 인덱싱
 * EdsMvpImport v0.4와 동일한 원화 환산 가격 로직 사용
 */
function edSyncPreview_buildMainSheetIndex_(mainSheet, headerInfo) {
  const lastRow = Math.min(mainSheet.getLastRow(), ED_MVP_IMPORT.maxMainDataRow);
  const lastCol = mainSheet.getLastColumn();

  const rowCount = lastRow - headerInfo.dataStartRow + 1;
  const byHoldingId = new Map();

  if (rowCount <= 0) {
    return { byHoldingId };
  }

  const values = mainSheet
    .getRange(headerInfo.dataStartRow, 1, rowCount, lastCol)
    .getValues();

  const displayValues = mainSheet
    .getRange(headerInfo.dataStartRow, 1, rowCount, lastCol)
    .getDisplayValues();

  let lastBroker = "";
  let lastAccountName = "";
  let lastAccountType = "";
  let lastCountry = "";
  let foundHoldingOnce = false;
  let invalidStreakAfterFirstHolding = 0;

  for (let i = 0; i < values.length; i++) {
    const absoluteRow = headerInfo.dataStartRow + i;
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

    const quantity = edSyncPreview_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.quantity)
    );

    const avgPriceKrw = edSyncPreview_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.avgPriceKrw)
    );

    const avgPriceUsd = edSyncPreview_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.avgPriceUsd)
    );

    const priceKrw = edSyncPreview_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.priceKrw)
    );

    const priceUsd = edSyncPreview_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.priceUsd)
    );

    const valuationAmountKrw = edSyncPreview_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.valuationAmount)
    );

    const targetWeight = edSyncPreview_parsePercent_(
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

    const priceKrwForApp = effectivePriceInfo.priceKrw;

    const avgPriceKrwForApp = edImport_getEffectiveKrwAvgPrice_({
      avgPriceKrw,
      avgPriceUsd,
      priceUsd,
      effectiveKrwPrice: priceKrwForApp,
    });

    const validation = edImport_validateHoldingCandidate_({
      absoluteRow,
      ticker,
      assetName,
      quantity,
      avgPrice: avgPriceKrwForApp,
      price: priceKrwForApp,
      effectiveAccountName,
      normalizedAccountType,
    });

    if (!validation.valid) {
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

    const market = edImport_inferMarket_(normalizedCountry, ticker);

    const accountId = edImport_makeAccountId_(
      effectiveBroker,
      effectiveAccountName,
      normalizedAccountType
    );

    const assetId = edImport_makeAssetId_(market, ticker);
    const holdingId = `HLD_${accountId}_${assetId}`;

    const item = {
      holdingId,
      rowIndex: absoluteRow,
      accountId,
      assetId,
      broker: effectiveBroker,
      accountName: effectiveAccountName,
      accountType: normalizedAccountType,
      country: normalizedCountry,
      ticker,
      assetName,
      quantity,
      avgPriceKrw,
      avgPriceUsd,
      avgPriceKrwForApp,
      priceKrw,
      priceUsd,
      valuationAmountKrw,
      priceKrwForApp,
      priceSource: effectivePriceInfo.source,
      targetWeight,
    };

    if (!byHoldingId.has(holdingId)) {
      byHoldingId.set(holdingId, []);
    }

    byHoldingId.get(holdingId).push(item);
  }

  return { byHoldingId };
}

/**
 * 미리보기 시트 작성
 */
function edSyncPreview_writePreviewSheet_(ss, rows, summary) {
  let sheet = ss.getSheetByName(ED_MVP_SYNC_PREVIEW.sheets.preview);

  if (!sheet) {
    sheet = ss.insertSheet(ED_MVP_SYNC_PREVIEW.sheets.preview);
  }

  sheet.clear();

  const summaryRows = [
    ["metric", "value"],
    ["생성시각", new Date()],
    ["대상 원장 시트", summary.mainSheetName],
    ["matched", summary.matchedCount],
    ["unmatched", summary.unmatchedCount],
    ["duplicate", summary.duplicateCount],
    ["update_candidate", summary.updateCount],
    ["no_change", summary.noChangeCount],
    ["price_mismatch", summary.priceMismatchCount],
  ];

  sheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows);

  const startRow = summaryRows.length + 3;

  const headers = [
    "action",
    "match_status",
    "main_row",
    "holding_id",
    "account_id",
    "asset_id",
    "ticker",
    "asset_name",
    "main_quantity",
    "app_quantity",
    "main_avg_price_krw_for_app",
    "app_avg_price_krw",
    "main_price_krw_for_app",
    "app_price_krw",
    "main_target_weight",
    "app_target_weight",
    "main_account_name",
    "main_price_source",
    "diff_summary",
  ];

  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);

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

  if (rows.length > 0) {
    sheet.getRange(startRow + 1, 9, rows.length, 6).setNumberFormat("#,##0.########");
    sheet.getRange(startRow + 1, 15, rows.length, 2).setNumberFormat("0.00%");
  }
}

/**
 * 시트 → 객체 배열
 */
function edSyncPreview_readSheetAsObjects_(sheet) {
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

function edSyncPreview_makeMapByKey_(rows, keyField) {
  const map = new Map();

  rows.forEach((row) => {
    const key = String(row[keyField] || "");
    if (key) map.set(key, row);
  });

  return map;
}

function edSyncPreview_buildDiffSummary_(flags) {
  const diffs = [];

  if (flags.quantityChanged) diffs.push("수량 변경");
  if (flags.avgPriceChanged) diffs.push("평단가 변경");
  if (flags.appPriceChanged) diffs.push("앱 평가단가 불일치");
  if (flags.targetWeightChanged) diffs.push("목표비중 변경");
  if (flags.duplicate) diffs.push("기존 원장 중복 매칭");

  if (flags.mainPriceSource === "valuation_div_quantity") {
    diffs.push("원화 평가액÷수량 기준 가격");
  }

  if (diffs.length === 0) return "변경 없음";

  return diffs.join(", ");
}

function edSyncPreview_parseNumber_(value) {
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

function edSyncPreview_parsePercent_(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return value > 1 ? value / 100 : value;
  }

  const text = String(value).trim();

  if (text.indexOf("%") >= 0) {
    return edSyncPreview_parseNumber_(text) / 100;
  }

  const num = edSyncPreview_parseNumber_(text);

  return num > 1 ? num / 100 : num;
}
