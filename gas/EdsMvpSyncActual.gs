/*******************************************************
 * ED's MVP - App to Main Sheet Actual Sync v0.8.5.2
 *
 * 목적:
 * - App_Holdings의 입력값을 기존 "2. 종목현황"에 실제 반영
 * - 기존 원장 수정 전 App_SyncBackup에 변경 전 값을 백업
 * - App_SyncResult에 실제 반영 결과 기록
 *
 * v0.2 주요 변경 (속도 최적화):
 * - skipValidation 옵션 도입
 *   - false(기본): 기존 전체 인덱싱 + 검증 + 안전장치 유지 (안전 모드)
 *   - true         : 인덱싱·검증 완전 생략, 2차원 배열 Bulk Write 고속 모드
 * - 기존 setValue() 반복 루프 → 2차원 배열 조작 후 1회 setValues() 일괄 쓰기
 *
 * 중요 정책 v0.2:
 * - 실제 반영 대상: 수량, 목표비중
 * - 평단가 반영 대상: 원화 평단가 구조의 행만 반영
 * - USD 종목의 달러 평단가는 자동 반영하지 않음
 *   이유: App_Holdings.avg_price는 앱 평가용 KRW 환산 평단이므로,
 *         원장 달러 평단 컬럼에 그대로 쓰면 안 됨
 *
 * 실행 함수:
 * - syncAppHoldingsToMainSheet()                 → 안전 모드 (기본)
 * - syncAppHoldingsToMainSheet({skipValidation: true}) → 고속 숏컷 모드
 *
 * 전제:
 * - EdsMvpImport.gs v0.4 존재
 * - skipValidation=false 시: EdsMvpSyncPreview.gs로 preview 확인 완료 권장
 * - skipValidation=true 시: 1인 단독 사용 환경, 데이터 신뢰 전제
 *******************************************************/

const ED_MVP_SYNC_ACTUAL = {
  sheets: {
    holdings: "App_Holdings",
    backup: "App_SyncBackup",
    result: "App_SyncResult",
    syncLog: "App_SyncLog",
  },

  tolerance: {
    quantity: 0.000001,
    price: 1,
    percent: 0.000001,
  },

  // 원장 고속 반영 대상 컬럼 인덱스 (0-based, 기본값)
  // 실제 원장 구조가 다르면 edSyncActual_detectWriteColumns_ 함수가 재탐지
  mainSheet: {
    dataStartRow: 9,   // 원장 데이터 시작 행 (headerInfo가 있으면 덮어씀)
    tickerColZero: 5,  // F열 (0-based)
  },
};

/*******************************************************
 * Public Entry Point
 *******************************************************/

/**
 * App_Holdings → 기존 2. 종목현황 실제 반영
 *
 * @param {Object} [options]
 * @param {boolean} [options.skipValidation=false]
 *   true:  고속 숏컷 모드 (인덱싱·검증 생략, Bulk Write)
 *   false: 안전 모드 (기존 전체 인덱싱 + 검증 + Bulk Write)
 */
function syncAppHoldingsToMainSheet(options) {
  const skipValidation = Boolean(options && options.skipValidation);

  if (skipValidation) {
    return edSyncActual_fastSync_();
  }

  return edSyncActual_safeSync_();
}

/*******************************************************
 * 고속 숏컷 모드 (skipValidation = true)
 *
 * 전략:
 * 1. App_Holdings 시트를 읽어 holding_id → 변경값 Map 구성
 * 2. 원장 시트의 데이터 범위를 2차원 배열로 1회 getValues()
 * 3. 메모리상에서 필요한 셀만 수정
 * 4. 단 1회 setValues()로 일괄 반영
 * 5. 최소한의 안전장치: holding_id 매칭 실패 시 해당 행 Skip (원장 불변)
 *******************************************************/

function edSyncActual_fastSync_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const syncId = edSyncActual_makeSyncId_();

  // --- 1. 원장 시트 탐지 ---
  const mainSheet = edImport_findMainSheet_(ss);
  if (!mainSheet) {
    ss.toast("기존 원장 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    return { status: "ERROR", message: "기존 원장 시트 없음" };
  }

  const headerInfo = edImport_detectHeader_(mainSheet);
  if (!headerInfo) {
    ss.toast("기존 원장 헤더를 찾을 수 없습니다.", "ED's MVP", 8);
    return { status: "ERROR", message: "기존 원장 헤더 탐지 실패" };
  }

  // --- 2. App_Holdings 로드 ---
  const holdingsSheet = ss.getSheetByName(ED_MVP_SYNC_ACTUAL.sheets.holdings);
  if (!holdingsSheet) {
    ss.toast("App_Holdings 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    return { status: "ERROR", message: "App_Holdings 없음" };
  }

  const appHoldings = edSyncActual_readSheetAsObjects_(holdingsSheet)
    .filter((h) => String(h.enabled || "").toUpperCase() === "TRUE");

  if (appHoldings.length === 0) {
    ss.toast("App_Holdings에 활성 보유종목이 없습니다.", "ED's MVP", 5);
    return { status: "NO_DATA", message: "활성 holding 없음" };
  }

  // --- 3. 원장 전체를 메모리에 적재 (단 1회 getValues, getDisplayValues 호출 없음) ---
  const lastRow = Math.min(mainSheet.getLastRow(), ED_MVP_IMPORT.maxMainDataRow);
  const lastCol = mainSheet.getLastColumn();
  const dataStartRow = headerInfo.dataStartRow;
  const rowCount = lastRow - dataStartRow + 1;

  if (rowCount <= 0) {
    ss.toast("원장에 데이터 행이 없습니다.", "ED's MVP", 5);
    return { status: "NO_DATA", message: "원장 데이터 없음" };
  }

  // values: 0-indexed 2차원 배열, values[i] = i번째 데이터행
  const range = mainSheet.getRange(dataStartRow, 1, rowCount, lastCol);
  const values = range.getValues();

  // --- 4. 원장에서 holding_id 역인덱스 빠르게 구성 ---
  // holdingId → 행 배열 인덱스 (여러 개 매칭 가능)
  const mainIndexByHoldingId = edSyncActual_buildFastIndex_(
    values,
    headerInfo,
    dataStartRow
  );

  // --- 5. App_Holdings 데이터를 기반으로 메모리 배열 수정 ---
  const colMap = headerInfo.colMap;
  const backupRows = [];
  const resultRows = [];

  let matchedCount = 0;
  let unmatchedCount = 0;
  let changedHoldingCount = 0;
  let noChangeCount = 0;
  let skippedAvgPriceCount = 0;
  let writeCellCount = 0;

  appHoldings.forEach((holding) => {
    const holdingId = String(holding.holding_id || "");
    const matches = mainIndexByHoldingId.get(holdingId) || [];

    if (matches.length === 0) {
      unmatchedCount++;
      resultRows.push(edSyncActual_makeResultRow_(
        syncId, "SKIP", "UNMATCHED", "", holding, "원장에서 대응 행을 찾지 못함 (고속 모드)"
      ));
      return;
    }

    // 고속 모드에서는 duplicate 시 첫 번째 매칭만 처리 (경고는 결과에 기록)
    const match = matches[0];
    const isDuplicate = matches.length > 1;
    matchedCount++;

    const appQuantity = edSyncActual_parseNumber_(holding.quantity);
    const appAvgPriceKrw = edSyncActual_parseNumber_(holding.avg_price);
    const appTargetWeight = edSyncActual_parsePercent_(holding.target_weight_account);

    // values 배열 인덱스 (0-based)
    const rowIdx = match.arrayIndex;
    let changed = false;

    // (A) 수량 반영
    if (colMap.quantity !== undefined && colMap.quantity >= 0) {
      if (Math.abs(match.quantity - appQuantity) > ED_MVP_SYNC_ACTUAL.tolerance.quantity) {
        backupRows.push(edSyncActual_makeBackupRowDirect_(
          syncId, mainSheet.getName(), dataStartRow + rowIdx, colMap.quantity + 1,
          "quantity", holding, match.quantity, appQuantity
        ));
        values[rowIdx][colMap.quantity] = appQuantity;
        writeCellCount++;
        changed = true;
      }
    }

    // (B) 평단가 반영 (원화 평단 구조인 경우만)
    if (colMap.avgPriceKrw !== undefined && colMap.avgPriceKrw >= 0) {
      const avgPriceDiff =
        Math.abs(match.avgPriceKrwForApp - appAvgPriceKrw) > ED_MVP_SYNC_ACTUAL.tolerance.price;
      if (avgPriceDiff) {
        const canWrite =
          !(match.avgPriceUsd > 0 && match.avgPriceKrw <= 0); // USD 원장 구조 제외
        if (canWrite) {
          backupRows.push(edSyncActual_makeBackupRowDirect_(
            syncId, mainSheet.getName(), dataStartRow + rowIdx, colMap.avgPriceKrw + 1,
            "avg_price_krw", holding, match.avgPriceKrw, appAvgPriceKrw
          ));
          values[rowIdx][colMap.avgPriceKrw] = appAvgPriceKrw;
          writeCellCount++;
          changed = true;
        } else {
          skippedAvgPriceCount++;
        }
      }
    }

    // (C) 목표비중 반영
    if (colMap.targetWeight !== undefined && colMap.targetWeight >= 0) {
      if (Math.abs(match.targetWeight - appTargetWeight) > ED_MVP_SYNC_ACTUAL.tolerance.percent) {
        backupRows.push(edSyncActual_makeBackupRowDirect_(
          syncId, mainSheet.getName(), dataStartRow + rowIdx, colMap.targetWeight + 1,
          "target_weight_account", holding, match.targetWeight, appTargetWeight
        ));
        values[rowIdx][colMap.targetWeight] = appTargetWeight;
        writeCellCount++;
        changed = true;
      }
    }

    if (changed) {
      changedHoldingCount++;
      const warn = isDuplicate ? " [DUPLICATE_MATCH: 첫 번째 행만 반영]" : "";
      resultRows.push(edSyncActual_makeResultRow_(
        syncId, "UPDATE", isDuplicate ? "DUPLICATE_FIRST" : "MATCHED",
        dataStartRow + rowIdx, holding, `고속 Bulk Write 완료${warn}`
      ));
    } else {
      noChangeCount++;
      resultRows.push(edSyncActual_makeResultRow_(
        syncId, "NO_CHANGE", "MATCHED", dataStartRow + rowIdx, holding, "변경 없음"
      ));
    }
  });

  // --- 6. 백업 기록 (변경사항이 있을 때만) ---
  if (backupRows.length > 0) {
    edSyncActual_appendBackupRows_(ss, backupRows);
  }

  // --- 7. 핵심: 단 1회 setValues()로 원장 전체 범위를 일괄 반영 ---
  if (writeCellCount > 0) {
    range.setValues(values);
  }

  // --- 8. 결과 시트 기록 ---
  const summary = {
    syncId,
    mainSheetName: mainSheet.getName(),
    status: "SUCCESS_FAST",
    matchedCount,
    unmatchedCount,
    duplicateCount: 0,
    changedHoldingCount,
    noChangeCount,
    writeCellCount,
    skippedAvgPriceCount,
    message: `고속 Bulk Write 완료 (skipValidation=true, setValues 1회 호출)`,
  };

  edSyncActual_writeResultSheet_(ss, resultRows, summary);

  const message =
    `[고속] 동기화 완료: changed=${changedHoldingCount}, cells=${writeCellCount}, ` +
    `no_change=${noChangeCount}, unmatched=${unmatchedCount}, skipped_avg=${skippedAvgPriceCount}`;

  edImport_writeLog_("manual", "success", "App_Holdings", mainSheet.getName(), syncId, message);
  ss.toast(message, "ED's MVP", 10);

  return summary;
}

/*******************************************************
 * 안전 모드 (skipValidation = false, 기존 로직 유지 + Bulk Write 최적화)
 *******************************************************/

function edSyncActual_safeSync_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const syncId = edSyncActual_makeSyncId_();

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

  const requiredCols = [
    ["quantity", headerInfo.colMap.quantity],
    ["targetWeight", headerInfo.colMap.targetWeight],
  ];

  const missingCols = requiredCols.filter(([, col]) => col === undefined || col < 0);
  if (missingCols.length > 0) {
    const msg = "필수 원장 컬럼 누락: " + missingCols.map(([name]) => name).join(", ");
    ss.toast(msg, "ED's MVP", 8);
    edImport_writeLog_("manual", "error", "App_Holdings", mainSheet.getName(), "", msg);
    return;
  }

  const holdingsSheet = ss.getSheetByName(ED_MVP_SYNC_ACTUAL.sheets.holdings);
  if (!holdingsSheet) {
    ss.toast("App_Holdings 시트를 찾을 수 없습니다.", "ED's MVP", 8);
    edImport_writeLog_("manual", "error", "App_Holdings", mainSheet.getName(), "", "App_Holdings 없음");
    return;
  }

  const appHoldings = edSyncActual_readSheetAsObjects_(holdingsSheet);
  const activeHoldings = appHoldings.filter((h) => String(h.enabled || "").toUpperCase() === "TRUE");

  const mainIndex = edSyncActual_buildMainSheetIndex_(mainSheet, headerInfo);

  // --- 원장 전체를 메모리에 1회 로드 ---
  const dataStartRow = headerInfo.dataStartRow;
  const lastRow = Math.min(mainSheet.getLastRow(), ED_MVP_IMPORT.maxMainDataRow);
  const lastCol = mainSheet.getLastColumn();
  const rowCount = lastRow - dataStartRow + 1;
  const fullRange = rowCount > 0
    ? mainSheet.getRange(dataStartRow, 1, rowCount, lastCol)
    : null;
  const fullValues = fullRange ? fullRange.getValues() : [];

  const backupRows = [];
  const resultRows = [];
  // 실제 쓰기가 필요한 (arrayIndex, colZero, newValue) 목록
  const writeQueue = [];

  let matchedCount = 0;
  let unmatchedCount = 0;
  let duplicateCount = 0;
  let changedHoldingCount = 0;
  let noChangeCount = 0;
  let skippedAvgPriceCount = 0;

  activeHoldings.forEach((holding) => {
    const holdingId = String(holding.holding_id || "");
    const matches = mainIndex.byHoldingId.get(holdingId) || [];

    const appQuantity = edSyncActual_parseNumber_(holding.quantity);
    const appAvgPriceKrw = edSyncActual_parseNumber_(holding.avg_price);
    const appTargetWeight = edSyncActual_parsePercent_(holding.target_weight_account);

    if (matches.length === 0) {
      unmatchedCount++;
      resultRows.push(edSyncActual_makeResultRow_(syncId, "SKIP", "UNMATCHED", "", holding, "기존 원장에서 대응 행을 찾지 못함"));
      return;
    }

    if (matches.length > 1) {
      duplicateCount++;
      resultRows.push(edSyncActual_makeResultRow_(syncId, "SKIP", "DUPLICATE", "", holding, `기존 원장 중복 매칭 ${matches.length}건`));
      return;
    }

    matchedCount++;
    const match = matches[0];
    const arrayIdx = match.rowIndex - dataStartRow; // fullValues 인덱스
    const changesForHolding = [];

    // 1) 수량 반영
    if (Math.abs(match.quantity - appQuantity) > ED_MVP_SYNC_ACTUAL.tolerance.quantity) {
      changesForHolding.push({ field: "quantity", col: headerInfo.colMap.quantity, old: match.quantity, new_: appQuantity });
    }

    // 2) 평단가 반영
    const avgPriceDiff =
      Math.abs(match.avgPriceKrwForApp - appAvgPriceKrw) > ED_MVP_SYNC_ACTUAL.tolerance.price;
    if (avgPriceDiff) {
      if (edSyncActual_canWriteAvgPriceKrw_(match, headerInfo)) {
        changesForHolding.push({ field: "avg_price_krw", col: headerInfo.colMap.avgPriceKrw, old: match.avgPriceKrw, new_: appAvgPriceKrw });
      } else {
        skippedAvgPriceCount++;
        resultRows.push(edSyncActual_makeResultRow_(syncId, "SKIP_FIELD", "AVG_PRICE_USD_NATIVE", match.rowIndex, holding, "USD 원장 구조: App의 KRW 환산 평단을 달러 평단 컬럼에 반영하지 않음"));
      }
    }

    // 3) 목표비중 반영
    if (Math.abs(match.targetWeight - appTargetWeight) > ED_MVP_SYNC_ACTUAL.tolerance.percent) {
      changesForHolding.push({ field: "target_weight_account", col: headerInfo.colMap.targetWeight, old: match.targetWeight, new_: appTargetWeight });
    }

    if (changesForHolding.length === 0) {
      noChangeCount++;
      resultRows.push(edSyncActual_makeResultRow_(syncId, "NO_CHANGE", "MATCHED", match.rowIndex, holding, "변경 없음"));
      return;
    }

    changedHoldingCount++;

    changesForHolding.forEach((change) => {
      backupRows.push(edSyncActual_makeBackupRowDirect_(
        syncId, mainSheet.getName(), match.rowIndex, change.col + 1,
        change.field, holding, change.old, change.new_
      ));
      // 메모리 배열 수정
      if (arrayIdx >= 0 && arrayIdx < fullValues.length) {
        fullValues[arrayIdx][change.col] = change.new_;
      }
      writeQueue.push({ arrayIdx, col: change.col, value: change.new_ });
    });

    resultRows.push(edSyncActual_makeResultRow_(
      syncId, "UPDATE", "MATCHED", match.rowIndex, holding,
      `변경 필드: ${changesForHolding.map((c) => c.field).join(", ")}`
    ));
  });

  // 안전장치: unmatched 또는 duplicate가 있으면 실제 쓰기 중단
  if (unmatchedCount > 0 || duplicateCount > 0) {
    edSyncActual_writeResultSheet_(ss, resultRows, {
      syncId,
      mainSheetName: mainSheet.getName(),
      status: "ABORTED",
      matchedCount,
      unmatchedCount,
      duplicateCount,
      changedHoldingCount,
      noChangeCount,
      writeCellCount: 0,
      skippedAvgPriceCount,
      message: "unmatched 또는 duplicate 존재로 실제 반영 중단",
    });

    edImport_writeLog_("manual", "error", "App_Holdings", mainSheet.getName(), "", `동기화 중단: unmatched=${unmatchedCount}, duplicate=${duplicateCount}`);
    ss.toast(`동기화 중단: unmatched=${unmatchedCount}, duplicate=${duplicateCount}`, "ED's MVP", 8);
    return;
  }

  // 백업 먼저 기록
  if (backupRows.length > 0) {
    edSyncActual_appendBackupRows_(ss, backupRows);
  }

  // --- 핵심: 1회 setValues()로 일괄 반영 ---
  if (writeQueue.length > 0 && fullRange) {
    fullRange.setValues(fullValues);
  }

  edSyncActual_writeResultSheet_(ss, resultRows, {
    syncId,
    mainSheetName: mainSheet.getName(),
    status: "SUCCESS",
    matchedCount,
    unmatchedCount,
    duplicateCount,
    changedHoldingCount,
    noChangeCount,
    writeCellCount: writeQueue.length,
    skippedAvgPriceCount,
    message: "동기화 완료 (Bulk Write)",
  });

  const message =
    `동기화 완료: holdings_changed=${changedHoldingCount}, ` +
    `cells_written=${writeQueue.length}, no_change=${noChangeCount}, ` +
    `avg_price_skipped=${skippedAvgPriceCount}`;

  edImport_writeLog_("manual", "success", "App_Holdings", mainSheet.getName(), syncId, message);
  ss.toast(message, "ED's MVP", 8);
}

/*******************************************************
 * 고속 인덱스 구성 (holding_id → { arrayIndex, 필드값들 })
 * getDisplayValues 의존성 없음. values(raw) 만으로 인덱스 구성.
 * 숫자/사후 포맷팅은 V8 엔진에서 직접 캐스팅 처리.
 *******************************************************/

function edSyncActual_buildFastIndex_(values, headerInfo, dataStartRow) {
  const byHoldingId = new Map();
  const colMap = headerInfo.colMap;

  // raw value를 문자열로 변환하는 인라인 헬퍼
  function rawStr(row, col) {
    if (col === undefined || col < 0) return '';
    const v = row[col];
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  let lastBroker = '';
  let lastAccountName = '';
  let lastAccountType = '';
  let lastCountry = '';
  let foundHoldingOnce = false;
  let invalidStreakAfterFirstHolding = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const absoluteRow = dataStartRow + i;

    const broker = rawStr(row, colMap.broker);
    const accountName = rawStr(row, colMap.account);
    const accountTypeRaw = rawStr(row, colMap.accountType);
    const countryRaw = rawStr(row, colMap.country);
    const ticker = edImport_normalizeTicker_(rawStr(row, colMap.ticker));
    const assetName = rawStr(row, colMap.assetName);

    const quantity = edSyncActual_parseNumber_(row[colMap.quantity]);
    const avgPriceKrw = edSyncActual_parseNumber_(row[colMap.avgPriceKrw]);
    const avgPriceUsd = edSyncActual_parseNumber_(row[colMap.avgPriceUsd]);
    const priceKrw = edSyncActual_parseNumber_(row[colMap.priceKrw]);
    const priceUsd = edSyncActual_parseNumber_(row[colMap.priceUsd]);
    const valuationAmountKrw = edSyncActual_parseNumber_(row[colMap.valuationAmount]);
    const targetWeight = edSyncActual_parsePercent_(row[colMap.targetWeight]);

    if (broker) lastBroker = broker;
    if (accountName) lastAccountName = accountName;
    if (accountTypeRaw) lastAccountType = accountTypeRaw;
    if (countryRaw) lastCountry = countryRaw;

    const effectiveBroker = broker || lastBroker;
    const effectiveAccountName = accountName || lastAccountName;
    const effectiveAccountTypeRaw = accountTypeRaw || lastAccountType || edImport_inferAccountType_(effectiveAccountName);
    const effectiveCountryRaw = countryRaw || lastCountry;

    const normalizedAccountType = edImport_normalizeAccountType_(effectiveAccountTypeRaw);
    const normalizedCountry = edImport_normalizeCountry_(effectiveCountryRaw, ticker, assetName);

    const effectivePriceInfo = edImport_getEffectiveKrwPriceInfo_({ quantity, priceKrw, priceUsd, valuationAmountKrw });
    const priceKrwForApp = effectivePriceInfo.priceKrw;
    const avgPriceKrwForApp = edImport_getEffectiveKrwAvgPrice_({ avgPriceKrw, avgPriceUsd, priceUsd, effectiveKrwPrice: priceKrwForApp });

    const validation = edImport_validateHoldingCandidate_({
      absoluteRow, ticker, assetName, quantity,
      avgPrice: avgPriceKrwForApp, price: priceKrwForApp,
      effectiveAccountName, normalizedAccountType,
    });

    if (!validation.valid) {
      if (foundHoldingOnce) {
        invalidStreakAfterFirstHolding++;
        if (invalidStreakAfterFirstHolding >= ED_MVP_IMPORT.stopAfterInvalidStreakAfterFirstHolding) break;
      }
      continue;
    }

    foundHoldingOnce = true;
    invalidStreakAfterFirstHolding = 0;

    const market = edImport_inferMarket_(normalizedCountry, ticker);
    const accountId = edImport_makeAccountId_(effectiveBroker, effectiveAccountName, normalizedAccountType);
    const assetId = edImport_makeAssetId_(market, ticker);
    const holdingId = `HLD_${accountId}_${assetId}`;

    const item = {
      holdingId,
      arrayIndex: i,          // values 배열 인덱스 (0-based)
      rowIndex: absoluteRow,   // 시트 실제 행 번호 (1-based)
      accountId,
      assetId,
      ticker,
      assetName,
      quantity,
      avgPriceKrw,
      avgPriceUsd,
      avgPriceKrwForApp,
      priceKrwForApp,
      targetWeight,
    };

    if (!byHoldingId.has(holdingId)) byHoldingId.set(holdingId, []);
    byHoldingId.get(holdingId).push(item);
  }

  return byHoldingId;
}

/*******************************************************
 * 기존 2. 종목현황의 보유종목 행을 holding_id 기준으로 인덱싱
 * (안전 모드에서 사용 - EdsMvpImport v0.4와 동일한 원화 환산 가격 로직)
 *******************************************************/

function edSyncActual_buildMainSheetIndex_(mainSheet, headerInfo) {
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

    const quantity = edSyncActual_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.quantity)
    );

    const avgPriceKrw = edSyncActual_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.avgPriceKrw)
    );

    const avgPriceUsd = edSyncActual_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.avgPriceUsd)
    );

    const priceKrw = edSyncActual_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.priceKrw)
    );

    const priceUsd = edSyncActual_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.priceUsd)
    );

    const valuationAmountKrw = edSyncActual_parseNumber_(
      edImport_getCellRaw_(row, displayRow, headerInfo.colMap.valuationAmount)
    );

    const targetWeight = edSyncActual_parsePercent_(
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

function edSyncActual_canWriteAvgPriceKrw_(match, headerInfo) {
  if (headerInfo.colMap.avgPriceKrw === undefined || headerInfo.colMap.avgPriceKrw < 0) return false;

  // 달러 평단 구조: USD 평단이 있고 KRW 평단이 비어있던 행은 자동 반영 금지.
  if (match.avgPriceUsd > 0 && match.avgPriceKrw <= 0) return false;

  return true;
}

/*******************************************************
 * 백업/결과 행 빌더
 *******************************************************/

/**
 * 고속/안전 모드 공용 백업 행 생성 (직접 파라미터 전달 방식)
 */
function edSyncActual_makeBackupRowDirect_(
  syncId, sheetName, rowIndex, colIndex1based, fieldName, holding, oldValue, newValue
) {
  return [
    syncId,
    new Date(),
    sheetName,
    rowIndex,
    colIndex1based,
    edSyncActual_columnToLetter_(colIndex1based),
    fieldName,
    holding.holding_id || "",
    holding.account_id || "",
    holding.asset_id || "",
    holding.ticker || "",
    holding.asset_name || "",
    oldValue,
    newValue,
  ];
}

function edSyncActual_makeResultRow_(syncId, action, status, mainRow, holding, message) {
  return [
    syncId,
    new Date(),
    action,
    status,
    mainRow || "",
    holding.holding_id || "",
    holding.account_id || "",
    holding.asset_id || "",
    holding.ticker || "",
    holding.asset_name || "",
    message,
  ];
}

function edSyncActual_appendBackupRows_(ss, rows) {
  let sheet = ss.getSheetByName(ED_MVP_SYNC_ACTUAL.sheets.backup);

  if (!sheet) {
    sheet = ss.insertSheet(ED_MVP_SYNC_ACTUAL.sheets.backup);
  }

  const headers = [
    "sync_id",
    "timestamp",
    "sheet_name",
    "row",
    "col",
    "col_letter",
    "field_name",
    "holding_id",
    "account_id",
    "asset_id",
    "ticker",
    "asset_name",
    "old_value",
    "new_value",
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const existingHeader = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    if (String(existingHeader[0] || "") !== "sync_id") {
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#137333")
    .setHorizontalAlignment("center");
}

function edSyncActual_writeResultSheet_(ss, rows, summary) {
  let sheet = ss.getSheetByName(ED_MVP_SYNC_ACTUAL.sheets.result);

  if (!sheet) {
    sheet = ss.insertSheet(ED_MVP_SYNC_ACTUAL.sheets.result);
  }

  sheet.clear();

  const summaryRows = [
    ["metric", "value"],
    ["sync_id", summary.syncId],
    ["생성시각", new Date()],
    ["상태", summary.status],
    ["대상 원장 시트", summary.mainSheetName],
    ["matched", summary.matchedCount],
    ["unmatched", summary.unmatchedCount],
    ["duplicate", summary.duplicateCount],
    ["changed_holdings", summary.changedHoldingCount],
    ["no_change", summary.noChangeCount],
    ["write_cells", summary.writeCellCount],
    ["skipped_avg_price", summary.skippedAvgPriceCount],
    ["message", summary.message],
  ];

  sheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows);

  const startRow = summaryRows.length + 3;

  const headers = [
    "sync_id",
    "timestamp",
    "action",
    "status",
    "main_row",
    "holding_id",
    "account_id",
    "asset_id",
    "ticker",
    "asset_name",
    "message",
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
}

/*******************************************************
 * 유틸리티
 *******************************************************/

function edSyncActual_readSheetAsObjects_(sheet) {
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

function edSyncActual_parseNumber_(value) {
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

function edSyncActual_parsePercent_(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return value > 1 ? value / 100 : value;
  }

  const text = String(value).trim();

  if (text.indexOf("%") >= 0) {
    return edSyncActual_parseNumber_(text) / 100;
  }

  const num = edSyncActual_parseNumber_(text);

  return num > 1 ? num / 100 : num;
}

function edSyncActual_makeSyncId_() {
  const now = new Date();
  return `SYNC_${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${Math.floor(Math.random() * 10000)}`;
}

function edSyncActual_columnToLetter_(column) {
  let temp = "";
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}
