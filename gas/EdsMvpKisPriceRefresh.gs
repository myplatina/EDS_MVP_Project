/*******************************************************
 * ED's MVP - KIS Price Refresh v0.8.61
 *
 * 목적:
 * 1) KIS Open API로 보유종목 전체 현재가 갱신
 *    - 국내 KRX: 국내주식 현재가 API
 *    - 미국 NASDAQ/NYSE/AMEX: 해외주식 현재체결가 API
 * 2) App_Prices 갱신
 *    - price: App_Output 계산용 원화 환산 단가
 *    - source_price: 원천 통화 가격(KRW 또는 USD)
 *    - source_currency: 원천 통화(KRW/USD)
 *    - fx_rate: USD/KRW 적용 환율. 국내는 1
 * 3) App_Output 재계산
 * 4) 선택 시 원본 '2. 종목현황' 반영
 *    - 국내: K열 현재가, M열 전일 대비 등락률
 *    - 미국: L열 달러 현재가, M열 전일 대비 등락률
 *
 * 정리 원칙:
 * - test* 함수 없음
 * - 실제 실행 함수 2개만 유지
 *   a. refreshKrxPricesFromKis()              : App_* 데이터만 갱신. 이름은 API 호환을 위해 유지
 *   b. refreshKrxPricesToMainSheetFromKis()   : App_* + 원본 2.종목현황 반영. 이름은 API 호환을 위해 유지
 * - 함수명은 기존 PWA/API 호환 때문에 KRX를 유지하지만, 실제 대상은 국내+미국 전체
 *******************************************************/

const ED_MVP_PRICE_REFRESH = {
  version: '0.8.5.2',
  sheets: {
    appPrices: 'App_Prices',
    appOutput: 'App_Output',
    priceRefreshResult: 'App_PriceRefreshResult',
    mainSheet: '2. 종목현황',
    settings: 'App_Settings',
  },
  kis: {
    baseUrlProperty: 'KIS_BASE_URL',
    appKeyProperty: 'KIS_APP_KEY',
    appSecretProperty: 'KIS_APP_SECRET',
    accessTokenProperty: 'KIS_ACCESS_TOKEN',
    accessTokenExpiredAtProperty: 'KIS_ACCESS_TOKEN_EXPIRED_AT',
    defaultBaseUrl: 'https://openapi.koreainvestment.com:9443',
    tokenPath: '/oauth2/tokenP',
    domesticPricePath: '/uapi/domestic-stock/v1/quotations/inquire-price',
    domesticPriceTrId: 'FHKST01010100',
    overseasPricePath: '/uapi/overseas-price/v1/quotations/price',
    overseasPriceTrId: 'HHDFS00000300',
  },
  request: {
    delayMs: 100,          // KIS 실전계좌 초당 20건 제한 기준, 100ms = 안전 마진 포함 실질 초당 ~10건
    retryMaxAttempts: 4,
    retryBaseSleepMs: 800,  // 레이트리밋 감지 시 800ms, 1600ms, ... 지수 백오프
  },
  mainSheet: {
    dataStartRow: 9,
    tickerCol: 6, // F
    krwPriceCol: 11, // K
    usdPriceCol: 12, // L
    changeRateCol: 13, // M
  },
  allowedPriceSource: 'api',
  optionalPriceHeaders: [
    'source_price',
    'source_currency',
    'fx_rate',
    'price_market',
    'price_exchange',
  ],
};

/*******************************************************
 * Public API-compatible functions
 *******************************************************/

/**
 * App_Prices / App_Output만 갱신.
 * 원본 '2. 종목현황'은 수정하지 않음.
 *
 * NOTE: 기존 PWA action 호환을 위해 함수명은 refreshKrxPricesFromKis 유지.
 * 실제 처리 대상은 국내 KRX + 미국 NASDAQ/NYSE/AMEX 전체.
 */
function refreshKrxPricesFromKis() {
  return edPrice_refreshAllMarketPricesCore_({ updateMainSheet: false });
}

/**
 * App_Prices / App_Output 갱신 후 원본 '2. 종목현황' 가격열까지 반영.
 * 국내 KRX: K/M열, 미국: L/M열.
 *
 * NOTE: 기존 PWA action 호환을 위해 함수명은 refreshKrxPricesToMainSheetFromKis 유지.
 */
function refreshKrxPricesToMainSheetFromKis() {
  return edPrice_refreshAllMarketPricesCore_({ updateMainSheet: true });
}

/*******************************************************
 * Google Sheet menu helpers
 *******************************************************/

function edMvpMenuRefreshKrxPricesAppOnly() {
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert(
    '현재가 갱신 - 앱 데이터만',
    'KIS API로 보유종목 전체 현재가를 갱신합니다.\n\n대상:\n- 국내 KRX 종목\n- 미국 NASDAQ/NYSE/AMEX 종목\n\n반영 대상:\n- App_Prices\n- App_Output\n\n원본 2. 종목현황은 수정하지 않습니다.',
    ui.ButtonSet.OK_CANCEL
  );
  if (choice !== ui.Button.OK) return;

  const result = refreshKrxPricesFromKis();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `현재가 갱신 완료: success=${result.success_count}, error=${result.error_count}, output=${result.output_count}`,
    "ED's MVP",
    8
  );
}

function edMvpMenuRefreshKrxPricesToMainSheet() {
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert(
    '원장 가격 반영(KIS)',
    "KIS API 현재가를 원본 '2. 종목현황'에 반영합니다.\n\n반영 위치:\n- 국내: K열 현재가, M열 전일 대비 등락\n- 미국: L열 달러 현재가, M열 전일 대비 등락\n\n기존 GOOGLEFINANCE / NAVER 수식은 값으로 대체될 수 있습니다.\n진행할까요?",
    ui.ButtonSet.OK_CANCEL
  );
  if (choice !== ui.Button.OK) return;

  const result = refreshKrxPricesToMainSheetFromKis();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `원장 가격 반영 완료: success=${result.success_count}, error=${result.error_count}, main=${result.main_sheet_updated_count}`,
    "ED's MVP",
    8
  );
}

function edMvpAddSheetMenu() {
  SpreadsheetApp.getUi()
    .createMenu("ED's MVP")
    .addItem('현재가 갱신 - 앱 데이터만', 'edMvpMenuRefreshKrxPricesAppOnly')
    .addItem('원장 가격 반영(KIS)', 'edMvpMenuRefreshKrxPricesToMainSheet')
    .addToUi();
}

/*******************************************************
 * Core refresh
 *******************************************************/

function edPrice_refreshAllMarketPricesCore_(options) {
  const startedAt = new Date();
  const updateMainSheet = Boolean(options && options.updateMainSheet);
  const usdKrwRate = edPrice_getUsdKrwRate_();

  const appPriceContext = edPrice_readAppPricesContext_();
  const targets = appPriceContext.rows
    .filter((row) => row.rowIndex > 1)
    .filter((row) => edPrice_isEnabledPriceRow_(row))
    .map((row) => edPrice_classifyTarget_(row))
    .filter((row) => row.refreshable);

  const results = [];
  const priceResultsByTicker = new Map();
  let successCount = 0;
  let errorCount = 0;
  let updatedPriceCount = 0;
  let domesticCount = 0;
  let overseasCount = 0;

  // --- fetchAll 병렬 배치 처리 ---
  // 1. 토큰/자격증명을 루프 전 1회만 취득 (직렬 루프에서 매번 취득하던 방식 제거)
  const token = edPrice_getKisAccessToken_();
  const creds = edPrice_getKisCredentials_();
  const baseUrl = edPrice_getKisBaseUrl_();

  // 2. 각 종목의 request 객체 배열 조립
  //    미국 종목은 exchange fallback 후보를 첫 번째 후보로 사용 (fetchAll 1차 시도)
  const requestObjects = targets.map((target) => {
    const headers = {
      Authorization: 'Bearer ' + token,
      appkey: creds.appKey,
      appsecret: creds.appSecret,
      custtype: 'P',
    };

    if (target.price_market === 'KRX') {
      const query = {
        FID_COND_MRKT_DIV_CODE: 'J',
        FID_INPUT_ISCD: target.ticker,
      };
      return {
        url: baseUrl + ED_MVP_PRICE_REFRESH.kis.domesticPricePath + '?' + edPrice_toQueryString_(query),
        method: 'get',
        headers: Object.assign({}, headers, { tr_id: ED_MVP_PRICE_REFRESH.kis.domesticPriceTrId }),
        muteHttpExceptions: true,
      };
    } else {
      // 미국: exchange 첫 번째 후보로 1차 병렬 요청
      const exchangeCandidates = edPrice_getUsExchangeCandidates_(target.price_exchange, target.ticker);
      const exchange = exchangeCandidates[0] || 'NAS';
      const query = { AUTH: '', EXCD: exchange, SYMB: target.ticker };
      return {
        url: baseUrl + ED_MVP_PRICE_REFRESH.kis.overseasPricePath + '?' + edPrice_toQueryString_(query),
        method: 'get',
        headers: Object.assign({}, headers, { tr_id: ED_MVP_PRICE_REFRESH.kis.overseasPriceTrId }),
        muteHttpExceptions: true,
        _exchange: exchange,         // 응답 처리 시 참조용 (fetchAll options에 포함되지 않음)
        _exchangeCandidates: exchangeCandidates,
      };
    }
  });

  // fetchAll용 options 배열 (GAS UrlFetchApp.fetchAll은 순수 request spec만 받아야 함)
  const fetchOptions = requestObjects.map((req) => ({
    url: req.url,
    method: req.method,
    headers: req.headers,
    muteHttpExceptions: req.muteHttpExceptions,
  }));

  // 3. 병렬 발사 — 20종목 이하: 1회 전송. 초과: 15개씩 chunk 나눠 chunk 사이 150ms 대기
  const CHUNK_SIZE = 15;
  const allResponses = [];
  for (let chunkStart = 0; chunkStart < fetchOptions.length; chunkStart += CHUNK_SIZE) {
    if (chunkStart > 0) Utilities.sleep(150);
    const chunkResponses = UrlFetchApp.fetchAll(fetchOptions.slice(chunkStart, chunkStart + CHUNK_SIZE));
    chunkResponses.forEach((r) => allResponses.push(r));
  }

  // 4. 응답 일괄 처리
  allResponses.forEach((res, i) => {
    const target = targets[i];
    const reqMeta = requestObjects[i];
    try {
      const status = res.getResponseCode();
      const text = res.getContentText();
      let json;
      try { json = JSON.parse(text); } catch (e) {
        throw new Error(`KIS price JSON parse 실패. HTTP=${status}, body=${text.slice(0, 400)}`);
      }

      // 미국 종목: rt_cd=0 + 가격 0이면 exchange fallback 단건 재시도
      if (target.price_market === 'US') {
        const outputRaw = json.output || {};
        const output = Array.isArray(outputRaw) ? (outputRaw[0] || {}) : outputRaw;
        const firstPrice = edPrice_firstNumber_(output.last, output.ovrs_prpr, output.price, output.stck_prpr, output.clpr);

        if (!firstPrice || firstPrice <= 0 || String(json.rt_cd || '') !== '0') {
          // fallback: 나머지 exchange 후보 단건 순차 시도
          const candidates = (reqMeta._exchangeCandidates || []).slice(1);
          let fetched = null;
          for (const exchange of candidates) {
            try {
              fetched = edPrice_fetchKisOverseasPrice_(Object.assign({}, target, { price_exchange: exchange }));
              if (fetched) break;
            } catch (fe) { /* continue */ }
          }
          if (!fetched) throw new Error(`KIS overseas price 모든 거래소 후보 실패. ticker=${target.ticker}`);
          const normalized = edPrice_normalizeFetchedPriceForApp_(target, fetched, usdKrwRate);
          edPrice_applyPriceToAppPriceRow_(appPriceContext, target, normalized);
          priceResultsByTicker.set(target.ticker, normalized);
          results.push({ asset_id: target.asset_id, ticker: target.ticker, asset_name: target.asset_name, market: target.price_market, exchange: normalized.price_exchange || target.price_exchange, status: 'success', price: normalized.price, source_price: normalized.source_price, source_currency: normalized.source_currency, fx_rate: normalized.fx_rate, change_amount: normalized.change_amount, change_rate: normalized.change_rate, raw_status: normalized.raw_status, message: normalized.message });
          successCount += 1; updatedPriceCount += 1; overseasCount += 1;
          return;
        }
      }

      if (String(json.rt_cd || '') !== '0') {
        throw new Error(`KIS price API 오류. rt_cd=${json.rt_cd}, msg_cd=${json.msg_cd}, msg=${json.msg1}`);
      }

      // 정상 응답 처리
      let fetched;
      if (target.price_market === 'KRX') {
        const output = json.output || {};
        const sourcePrice = edPrice_num_(output.stck_prpr);
        const prevClose = edPrice_num_(output.stck_sdpr || output.prdy_clpr || 0);
        const sign = output.prdy_vrss_sign || '';
        const calc = edPrice_calcChangeFromPricesOrApiPercent_(sourcePrice, prevClose, output.prdy_vrss, output.prdy_ctrt, sign);
        fetched = { ticker: target.ticker, price_market: 'KRX', price_exchange: 'KRX', source_price: sourcePrice, source_currency: 'KRW', prev_close: prevClose, change_amount: calc.change_amount, change_rate: calc.change_rate, raw_status: json.msg_cd || '', message: json.msg1 || '', raw: json };
      } else {
        const outputRaw = json.output || {};
        const output = Array.isArray(outputRaw) ? (outputRaw[0] || {}) : outputRaw;
        const exchange = reqMeta._exchange || target.price_exchange;
        const sourcePrice = edPrice_firstNumber_(output.last, output.ovrs_prpr, output.price, output.stck_prpr, output.clpr);
        const prevClose = edPrice_firstNumber_(output.base, output.prev, output.prdy_clpr, output.pclose, output.basp);
        const apiChangeAmount = edPrice_firstNumberAllowBlank_(output.diff, output.prdy_vrss, output.change, output.vrss);
        const apiRatePercent = edPrice_firstNumberAllowBlank_(output.rate, output.prdy_ctrt, output.change_rate, output.ctrt);
        const sign = output.sign || output.prdy_vrss_sign || '';
        const calc = edPrice_calcChangeFromPricesOrApiPercent_(sourcePrice, prevClose, apiChangeAmount, apiRatePercent, sign);
        fetched = { ticker: target.ticker, price_market: 'US', price_exchange: exchange, source_price: sourcePrice, source_currency: edPrice_str_(output.curr || output.currency || 'USD').toUpperCase() || 'USD', prev_close: prevClose, change_amount_source: calc.change_amount, change_amount: calc.change_amount, change_rate: calc.change_rate, raw_status: json.msg_cd || '', message: json.msg1 || '', raw: json };
      }

      const normalized = edPrice_normalizeFetchedPriceForApp_(target, fetched, usdKrwRate);
      edPrice_applyPriceToAppPriceRow_(appPriceContext, target, normalized);
      priceResultsByTicker.set(target.ticker, normalized);
      results.push({ asset_id: target.asset_id, ticker: target.ticker, asset_name: target.asset_name, market: target.price_market, exchange: normalized.price_exchange || target.price_exchange, status: 'success', price: normalized.price, source_price: normalized.source_price, source_currency: normalized.source_currency, fx_rate: normalized.fx_rate, change_amount: normalized.change_amount, change_rate: normalized.change_rate, raw_status: normalized.raw_status, message: normalized.message });
      successCount += 1; updatedPriceCount += 1;
      if (target.price_market === 'KRX') domesticCount += 1; else overseasCount += 1;

    } catch (e) {
      const message = e && e.message ? e.message : String(e);
      results.push({ asset_id: target.asset_id, ticker: target.ticker, asset_name: target.asset_name, market: target.price_market, exchange: target.price_exchange, status: 'error', price: '', source_price: '', source_currency: '', fx_rate: target.price_market === 'US' ? usdKrwRate : 1, change_amount: '', change_rate: '', raw_status: '', message });
      errorCount += 1;
    }
  }); // fetchAll 응답 처리 종료

  edPrice_writeAppPricesContext_(appPriceContext);

  let outputCount = 0;
  if (typeof buildAppOutputFromAppSheets === 'function') {
    const output = buildAppOutputFromAppSheets();
    outputCount = edPrice_extractOutputCount_(output);
  }

  let mainSheetUpdatedCount = 0;
  if (updateMainSheet) {
    mainSheetUpdatedCount = edPrice_updateMainSheetPrices_(priceResultsByTicker);
  }

  const summary = {
    version: ED_MVP_PRICE_REFRESH.version,
    fetched_at: new Date(),
    started_at: startedAt,
    finished_at: new Date(),
    target_count: targets.length,
    domestic_count: domesticCount,
    overseas_count: overseasCount,
    success_count: successCount,
    error_count: errorCount,
    skipped_count: 0,
    updated_price_count: updatedPriceCount,
    output_count: outputCount,
    main_sheet_updated_count: mainSheetUpdatedCount,
    source: ED_MVP_PRICE_REFRESH.allowedPriceSource,
    update_main_sheet: updateMainSheet,
    usd_krw_rate: usdKrwRate,
  };

  edPrice_writeRefreshResult_(summary, results);
  return summary;
}

/*******************************************************
 * Context / sheet IO
 *******************************************************/

function edPrice_readAppPricesContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_PRICE_REFRESH.sheets.appPrices);
  if (!sheet) throw new Error('App_Prices 시트를 찾을 수 없습니다.');

  edPrice_ensureAppPricesOptionalHeaders_(sheet);

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) throw new Error('App_Prices 데이터가 비어 있습니다.');

  const headers = values[0].map((h) => String(h || '').trim());
  const col = edPrice_makeHeaderIndex_(headers);

  const required = ['asset_id', 'ticker', 'price', 'change_amount', 'change_rate', 'currency', 'source'];
  required.forEach((key) => {
    if (col[key] === undefined) throw new Error(`App_Prices 필수 헤더 누락: ${key}`);
  });

  const rows = values.slice(1).map((row, i) => ({
    rowIndex: i + 2,
    row,
    asset_id: edPrice_str_(row[col.asset_id]),
    ticker: edPrice_normalizeTicker_(row[col.ticker]),
    asset_name: col.asset_name !== undefined ? edPrice_str_(row[col.asset_name]) : '',
    market: col.market !== undefined ? edPrice_str_(row[col.market]) : edPrice_marketFromAssetId_(row[col.asset_id]),
    enabled: col.enabled !== undefined ? edPrice_str_(row[col.enabled]) : 'TRUE',
  }));

  return { sheet, values, headers, col, rows };
}

function edPrice_ensureAppPricesOptionalHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
  let nextCol = headers.length + 1;
  let changed = false;

  ED_MVP_PRICE_REFRESH.optionalPriceHeaders.forEach((header) => {
    if (headers.indexOf(header) < 0) {
      sheet.getRange(1, nextCol).setValue(header);
      headers.push(header);
      nextCol += 1;
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#137333').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function edPrice_applyPriceToAppPriceRow_(context, target, fetched) {
  const rowArray = context.values[target.rowIndex - 1];
  const col = context.col;
  const now = new Date();

  rowArray[col.price] = fetched.price;
  rowArray[col.change_amount] = fetched.change_amount;
  rowArray[col.change_rate] = fetched.change_rate;
  rowArray[col.currency] = 'KRW';
  rowArray[col.source] = ED_MVP_PRICE_REFRESH.allowedPriceSource;

  if (col.source_price !== undefined) rowArray[col.source_price] = fetched.source_price;
  if (col.source_currency !== undefined) rowArray[col.source_currency] = fetched.source_currency;
  if (col.fx_rate !== undefined) rowArray[col.fx_rate] = fetched.fx_rate;
  if (col.price_market !== undefined) rowArray[col.price_market] = target.price_market;
  if (col.price_exchange !== undefined) rowArray[col.price_exchange] = target.price_exchange;
  if (col.prev_close !== undefined) rowArray[col.prev_close] = fetched.prev_close || '';
  if (col.fetched_at !== undefined) rowArray[col.fetched_at] = now;
  if (col.updated_at !== undefined) rowArray[col.updated_at] = now;
}

function edPrice_writeAppPricesContext_(context) {
  context.sheet.getRange(1, 1, context.values.length, context.values[0].length).setValues(context.values);
}

function edPrice_updateMainSheetPrices_(priceResultsByTicker) {
  if (!priceResultsByTicker || priceResultsByTicker.size === 0) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_PRICE_REFRESH.sheets.mainSheet);
  if (!sheet) throw new Error("원본 '2. 종목현황' 시트를 찾을 수 없습니다.");

  const lastRow = sheet.getLastRow();
  const startRow = ED_MVP_PRICE_REFRESH.mainSheet.dataStartRow;
  if (lastRow < startRow) return 0;

  const rowCount = lastRow - startRow + 1;
  const tickerValues = sheet.getRange(startRow, ED_MVP_PRICE_REFRESH.mainSheet.tickerCol, rowCount, 1).getValues();
  const krwPriceRange = sheet.getRange(startRow, ED_MVP_PRICE_REFRESH.mainSheet.krwPriceCol, rowCount, 1);
  const usdPriceRange = sheet.getRange(startRow, ED_MVP_PRICE_REFRESH.mainSheet.usdPriceCol, rowCount, 1);
  const changeRange = sheet.getRange(startRow, ED_MVP_PRICE_REFRESH.mainSheet.changeRateCol, rowCount, 1);
  const krwPriceValues = krwPriceRange.getValues();
  const usdPriceValues = usdPriceRange.getValues();
  const changeValues = changeRange.getValues();

  let updated = 0;

  tickerValues.forEach((row, index) => {
    const ticker = edPrice_normalizeTicker_(row[0]);
    if (!priceResultsByTicker.has(ticker)) return;

    const item = priceResultsByTicker.get(ticker);
    if (item.price_market === 'KRX') {
      krwPriceValues[index][0] = item.source_price;
    } else if (item.price_market === 'US') {
      usdPriceValues[index][0] = item.source_price;
    }
    changeValues[index][0] = item.change_rate;
    updated += 1;
  });

  if (updated > 0) {
    krwPriceRange.setValues(krwPriceValues);
    usdPriceRange.setValues(usdPriceValues);
    changeRange.setValues(changeValues);
  }

  return updated;
}

/*******************************************************
 * Target classification
 *******************************************************/

function edPrice_classifyTarget_(row) {
  const marketRaw = edPrice_str_(row.market || edPrice_marketFromAssetId_(row.asset_id)).toUpperCase();
  const ticker = edPrice_normalizeTicker_(row.ticker);

  if (marketRaw === 'KRX') {
    return Object.assign({}, row, {
      ticker,
      refreshable: /^[0-9A-Z]{6}$/.test(ticker),
      price_market: 'KRX',
      price_exchange: 'KRX',
    });
  }

  const usExchange = edPrice_toKisUsExchangeCode_(marketRaw);
  if (usExchange) {
    return Object.assign({}, row, {
      ticker,
      refreshable: ticker.length > 0,
      price_market: 'US',
      price_exchange: usExchange,
    });
  }

  // 일부 구버전 App_Prices에 market 컬럼이 없으면 asset_id prefix로 재판정.
  const prefix = edPrice_marketFromAssetId_(row.asset_id).toUpperCase();
  const prefixExchange = edPrice_toKisUsExchangeCode_(prefix);
  if (prefix === 'KRX') {
    return Object.assign({}, row, {
      ticker,
      refreshable: /^[0-9A-Z]{6}$/.test(ticker),
      price_market: 'KRX',
      price_exchange: 'KRX',
    });
  }
  if (prefixExchange) {
    return Object.assign({}, row, {
      ticker,
      refreshable: ticker.length > 0,
      price_market: 'US',
      price_exchange: prefixExchange,
    });
  }

  return Object.assign({}, row, {
    ticker,
    refreshable: false,
    price_market: 'UNSUPPORTED',
    price_exchange: '',
  });
}

function edPrice_toKisUsExchangeCode_(market) {
  const m = edPrice_str_(market).toUpperCase();
  if (['NASDAQ', 'NASD', 'NAS'].indexOf(m) >= 0) return 'NAS';
  if (['NYSE', 'NYS'].indexOf(m) >= 0) return 'NYS';
  if (['AMEX', 'AMS'].indexOf(m) >= 0) return 'AMS';
  return '';
}

/*******************************************************
 * KIS fetch - domestic / overseas
 *******************************************************/

function edPrice_fetchKisDomesticPriceWithRetry_(ticker) {
  return edPrice_fetchWithRetry_(() => edPrice_fetchKisDomesticPrice_(ticker));
}

function edPrice_fetchKisOverseasPriceWithRetry_(target) {
  return edPrice_fetchWithRetry_(() => edPrice_fetchKisOverseasPrice_(target));
}

function edPrice_fetchWithRetry_(fetcher) {
  let lastError = null;
  for (let attempt = 1; attempt <= ED_MVP_PRICE_REFRESH.request.retryMaxAttempts; attempt++) {
    try {
      const result = fetcher();
      if (attempt > 1) result.message = `${result.message || ''} / retry=${attempt - 1}`.trim();
      return result;
    } catch (e) {
      lastError = e;
      const message = e && e.message ? e.message : String(e);
      if (edPrice_isRateLimitError_(message)) {
        Utilities.sleep(ED_MVP_PRICE_REFRESH.request.retryBaseSleepMs * attempt);
        continue;
      }
      throw e;
    }
  }
  throw lastError;
}

function edPrice_fetchKisDomesticPrice_(ticker) {
  const token = edPrice_getKisAccessToken_();
  const creds = edPrice_getKisCredentials_();
  const query = {
    FID_COND_MRKT_DIV_CODE: 'J',
    FID_INPUT_ISCD: edPrice_normalizeTicker_(ticker),
  };

  const url = edPrice_getKisBaseUrl_()
    + ED_MVP_PRICE_REFRESH.kis.domesticPricePath
    + '?'
    + edPrice_toQueryString_(query);

  const json = edPrice_fetchKisJson_(url, token, creds, ED_MVP_PRICE_REFRESH.kis.domesticPriceTrId, 'KIS domestic price');
  const output = json.output || {};

  const sourcePrice = edPrice_num_(output.stck_prpr);
  const prevClose = edPrice_num_(output.stck_sdpr || output.prdy_clpr || 0);
  const sign = output.prdy_vrss_sign || '';

  // 등락/등락률은 API의 rate 필드보다 "현재가 - 전일종가" 기준을 우선한다.
  // 이렇게 해야 국내/해외 모두 M열이 "전일 종가 대비 현재가"로 동일해진다.
  const calc = edPrice_calcChangeFromPricesOrApiPercent_(
    sourcePrice,
    prevClose,
    output.prdy_vrss,
    output.prdy_ctrt,
    sign
  );

  return {
    ticker: edPrice_normalizeTicker_(ticker),
    price_market: 'KRX',
    price_exchange: 'KRX',
    source_price: sourcePrice,
    source_currency: 'KRW',
    prev_close: prevClose,
    change_amount: calc.change_amount,
    change_rate: calc.change_rate,
    raw_status: json.msg_cd || '',
    message: json.msg1 || '정상처리 되었습니다.',
    raw: json,
  };
}

function edPrice_fetchKisOverseasPrice_(target) {
  const token = edPrice_getKisAccessToken_();
  const creds = edPrice_getKisCredentials_();
  const exchangeCandidates = edPrice_getUsExchangeCandidates_(target.price_exchange, target.ticker);
  let lastError = null;

  for (let i = 0; i < exchangeCandidates.length; i++) {
    const exchange = exchangeCandidates[i];
    const query = {
      AUTH: '',
      EXCD: exchange,
      SYMB: target.ticker,
    };

    const url = edPrice_getKisBaseUrl_()
      + ED_MVP_PRICE_REFRESH.kis.overseasPricePath
      + '?'
      + edPrice_toQueryString_(query);

    try {
      const json = edPrice_fetchKisJson_(url, token, creds, ED_MVP_PRICE_REFRESH.kis.overseasPriceTrId, 'KIS overseas price');
      const outputRaw = json.output || {};
      const output = Array.isArray(outputRaw) ? (outputRaw[0] || {}) : outputRaw;

      const sourcePrice = edPrice_firstNumber_(output.last, output.ovrs_prpr, output.price, output.stck_prpr, output.clpr);

      // KIS 해외 API는 잘못된 거래소 코드에서도 rt_cd=0으로 내려오면서
      // 현재가 필드가 0/공백인 케이스가 있다. SCHD 같은 NYSE Arca 계열 ETF가 대표적.
      // 이 경우 success로 처리하면 App_Prices가 0으로 오염되므로 다음 거래소 후보로 fallback한다.
      if (!sourcePrice || sourcePrice <= 0) {
        lastError = new Error(
          `KIS overseas price returned zero. ticker=${target.ticker}, exchange=${exchange}, msg_cd=${json.msg_cd || ''}, msg=${json.msg1 || ''}`
        );
        if (i < exchangeCandidates.length - 1) continue;
        throw lastError;
      }

      const prevClose = edPrice_firstNumber_(output.base, output.prev, output.prdy_clpr, output.pclose, output.basp);
      const apiChangeAmount = edPrice_firstNumberAllowBlank_(output.diff, output.prdy_vrss, output.change, output.vrss);
      const apiRatePercent = edPrice_firstNumberAllowBlank_(output.rate, output.prdy_ctrt, output.change_rate, output.ctrt);
      const sign = output.sign || output.prdy_vrss_sign || output.diff_sign || '';

      // 해외 API의 rate 필드는 퍼센트 단위인데, -0.07 같은 작은 값도 "-0.07%"이다.
      // 이전 버전의 휴리스틱은 -0.07을 -7%로 해석할 수 있었다.
      // 따라서 전일종가가 있으면 무조건 (현재가 - 전일종가) / 전일종가로 재계산한다.
      const calc = edPrice_calcChangeFromPricesOrApiPercent_(
        sourcePrice,
        prevClose,
        apiChangeAmount,
        apiRatePercent,
        sign
      );

      return {
        ticker: target.ticker,
        price_market: 'US',
        price_exchange: exchange,
        source_price: sourcePrice,
        source_currency: edPrice_str_(output.curr || output.currency || 'USD').toUpperCase() || 'USD',
        prev_close: prevClose,
        change_amount_source: calc.change_amount,
        change_amount: calc.change_amount,
        change_rate: calc.change_rate,
        raw_status: json.msg_cd || '',
        message: json.msg1 || '정상처리 되었습니다.',
        raw: json,
      };
    } catch (e) {
      lastError = e;
      // 거래소 코드 fallback만 시도. 마지막 후보까지 실패하면 throw.
      if (i < exchangeCandidates.length - 1) continue;
    }
  }

  throw lastError;
}

function edPrice_fetchKisJson_(url, token, creds, trId, label) {
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      authorization: 'Bearer ' + token,
      appkey: creds.appKey,
      appsecret: creds.appSecret,
      tr_id: trId,
      custtype: 'P',
    },
  });

  const status = res.getResponseCode();
  const text = res.getContentText();
  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error(`${label} JSON parse 실패. HTTP=${status}, body=${text.slice(0, 500)}`);
  }

  if (status < 200 || status >= 300 || String(json.rt_cd || '') !== '0') {
    throw new Error(`${label} HTTP 오류. HTTP=${status}, body=${text.slice(0, 1000)}`);
  }

  return json;
}

function edPrice_getUsExchangeCandidates_(exchange, ticker) {
  const e = edPrice_str_(exchange).toUpperCase();
  const t = edPrice_str_(ticker).toUpperCase();
  const candidates = [];

  function add(value) {
    const v = edPrice_str_(value).toUpperCase();
    if (v && candidates.indexOf(v) < 0) candidates.push(v);
  }

  // ETF 중 일부는 원장/TradingView 표기는 NYSE처럼 보이지만 KIS 가격 API에서는
  // NYSE Arca/AMEX 계열 코드(AMS)로 조회해야 정상 가격이 나온다.
  // SCHD는 NYSE Arca 상장 ETF라 AMS 우선 후보로 둔다.
  if (t === 'SCHD') add('AMS');

  if (e === 'NAS' || e === 'NASDAQ' || e === 'NASD') add('NAS');
  if (e === 'NYS' || e === 'NYSE') add('NYS');
  if (e === 'AMS' || e === 'AMEX') add('AMS');

  // 잘못된 거래소 코드에서 rt_cd=0 + 가격 0이 나올 수 있으므로
  // 주요 미국 거래소 후보를 모두 fallback한다.
  add('AMS');
  add('NAS');
  add('NYS');

  return candidates;
}

function edPrice_normalizeFetchedPriceForApp_(target, fetched, usdKrwRate) {
  if (target.price_market === 'KRX') {
    return {
      ticker: target.ticker,
      price_market: 'KRX',
      price_exchange: 'KRX',
      price: fetched.source_price,
      source_price: fetched.source_price,
      source_currency: 'KRW',
      fx_rate: 1,
      prev_close: fetched.prev_close,
      change_amount: fetched.change_amount,
      change_rate: fetched.change_rate,
      raw_status: fetched.raw_status,
      message: fetched.message,
      raw: fetched.raw,
    };
  }

  const fx = Number(usdKrwRate || 0);
  if (!fx || fx <= 0) throw new Error('USD/KRW 환율을 확인할 수 없어 미국 종목 원화 평가액을 계산할 수 없습니다.');

  return {
    ticker: target.ticker,
    price_market: 'US',
    price_exchange: fetched.price_exchange || target.price_exchange,
    price: fetched.source_price * fx,
    source_price: fetched.source_price,
    source_currency: fetched.source_currency || 'USD',
    fx_rate: fx,
    prev_close: fetched.prev_close ? fetched.prev_close * fx : '',
    change_amount: (fetched.change_amount_source !== undefined ? fetched.change_amount_source : fetched.change_amount) * fx,
    change_rate: fetched.change_rate,
    raw_status: fetched.raw_status,
    message: fetched.message,
    raw: fetched.raw,
  };
}

/*******************************************************
 * KIS auth
 *******************************************************/

function edPrice_getKisCredentials_() {
  const props = PropertiesService.getScriptProperties();
  const appKey = props.getProperty(ED_MVP_PRICE_REFRESH.kis.appKeyProperty);
  const appSecret = props.getProperty(ED_MVP_PRICE_REFRESH.kis.appSecretProperty);
  if (!appKey || !appSecret) {
    throw new Error('KIS appkey/appsecret이 없습니다. setupKisCredentials()를 먼저 실행하세요.');
  }
  return { appKey, appSecret };
}

function edPrice_getKisAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const savedToken = props.getProperty(ED_MVP_PRICE_REFRESH.kis.accessTokenProperty);
  const expiredAtText = props.getProperty(ED_MVP_PRICE_REFRESH.kis.accessTokenExpiredAtProperty);

  if (savedToken && expiredAtText) {
    const expiredAt = new Date(expiredAtText);
    if (expiredAt.getTime() - Date.now() > 10 * 60 * 1000) return savedToken;
  }
  return edPrice_issueKisAccessToken_();
}

function edPrice_issueKisAccessToken_() {
  const creds = edPrice_getKisCredentials_();
  const props = PropertiesService.getScriptProperties();
  const url = edPrice_getKisBaseUrl_() + ED_MVP_PRICE_REFRESH.kis.tokenPath;

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      grant_type: 'client_credentials',
      appkey: creds.appKey,
      appsecret: creds.appSecret,
    }),
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
  props.setProperty(ED_MVP_PRICE_REFRESH.kis.accessTokenProperty, json.access_token);
  props.setProperty(ED_MVP_PRICE_REFRESH.kis.accessTokenExpiredAtProperty, expiredAt.toISOString());
  return json.access_token;
}

/*******************************************************
 * Refresh result log
 *******************************************************/

function edPrice_writeRefreshResult_(summary, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ED_MVP_PRICE_REFRESH.sheets.priceRefreshResult);
  if (!sheet) sheet = ss.insertSheet(ED_MVP_PRICE_REFRESH.sheets.priceRefreshResult);

  sheet.clear();

  const summaryRows = [
    ['metric', 'value'],
    ['version', summary.version],
    ['fetched_at', summary.fetched_at],
    ['target_count', summary.target_count],
    ['domestic_count', summary.domestic_count],
    ['overseas_count', summary.overseas_count],
    ['success_count', summary.success_count],
    ['error_count', summary.error_count],
    ['skipped_count', summary.skipped_count],
    ['updated_price_count', summary.updated_price_count],
    ['output_count', summary.output_count],
    ['main_sheet_updated_count', summary.main_sheet_updated_count],
    ['source', summary.source],
    ['update_main_sheet', summary.update_main_sheet],
    ['usd_krw_rate', summary.usd_krw_rate],
  ];

  sheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#137333').setFontColor('#ffffff');

  const detailHeader = [
    'asset_id',
    'ticker',
    'asset_name',
    'market',
    'exchange',
    'status',
    'price_krw',
    'source_price',
    'source_currency',
    'fx_rate',
    'change_amount_krw',
    'change_rate',
    'raw_status',
    'message',
  ];

  const detailRows = rows.map((row) => [
    row.asset_id,
    row.ticker,
    row.asset_name,
    row.market,
    row.exchange,
    row.status,
    row.price,
    row.source_price,
    row.source_currency,
    row.fx_rate,
    row.change_amount,
    row.change_rate,
    row.raw_status,
    row.message,
  ]);

  const startRow = 18;
  sheet.getRange(startRow, 1, 1, detailHeader.length).setValues([detailHeader]);
  sheet.getRange(startRow, 1, 1, detailHeader.length).setFontWeight('bold').setBackground('#137333').setFontColor('#ffffff');

  if (detailRows.length > 0) {
    sheet.getRange(startRow + 1, 1, detailRows.length, detailHeader.length).setValues(detailRows);
  }

  sheet.autoResizeColumns(1, detailHeader.length);
}

/*******************************************************
 * FX rate
 *******************************************************/

function edPrice_getUsdKrwRate_() {
  const fromSettings = edPrice_getUsdKrwRateFromSettings_();
  if (fromSettings > 0) return fromSettings;

  const fromMain = edPrice_getUsdKrwRateFromMainSheet_();
  if (fromMain > 0) return fromMain;

  throw new Error("USD/KRW 환율을 찾을 수 없습니다. App_Settings에 usd_krw_rate를 넣거나 '2. 종목현황' L5 환율 표시를 확인하세요.");
}

function edPrice_getUsdKrwRateFromSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_PRICE_REFRESH.sheets.settings);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const key = edPrice_str_(values[i][0]).toLowerCase();
    if (['usd_krw_rate', 'usdkrw_rate', 'fx_usdkrw', 'usdkrw'].indexOf(key) >= 0) {
      const n = edPrice_num_(values[i][1]);
      if (n > 0) return n;
    }
  }
  return 0;
}

function edPrice_getUsdKrwRateFromMainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ED_MVP_PRICE_REFRESH.sheets.mainSheet);
  if (!sheet) return 0;

  const candidates = [
    sheet.getRange(5, 12).getDisplayValue(), // L5: "$1 = ￦..." 표시 셀
    sheet.getRange(5, 12).getValue(),
  ];

  for (let i = 0; i < candidates.length; i++) {
    const n = edPrice_parseLargestNumber_(candidates[i]);
    if (n > 0) return n;
  }
  return 0;
}

/*******************************************************
 * Utilities
 *******************************************************/

function edPrice_makeHeaderIndex_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[String(header || '').trim()] = index;
  });
  return map;
}

function edPrice_isEnabledPriceRow_(row) {
  const enabled = String(row.enabled || 'TRUE').toUpperCase();
  return enabled !== 'FALSE' && enabled !== '0' && enabled !== 'N';
}

function edPrice_normalizeTicker_(value) {
  const text = String(value || '').trim().toUpperCase();
  if (/^\d+$/.test(text) && text.length < 6) return text.padStart(6, '0');
  return text;
}

function edPrice_marketFromAssetId_(assetId) {
  return String(assetId || '').split('_')[0] || '';
}

function edPrice_str_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function edPrice_num_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const n = Number(String(value).replace(/,/g, '').replace(/%/g, '').replace(/₩/g, '').replace(/\$/g, '').trim());
  return isNaN(n) ? 0 : n;
}

function edPrice_firstNumber_() {
  for (let i = 0; i < arguments.length; i++) {
    const value = arguments[i];
    if (value === null || value === undefined || value === '') continue;
    const n = edPrice_num_(value);
    if (!isNaN(n) && n !== 0) return n;
  }
  return 0;
}

function edPrice_firstNumberAllowBlank_() {
  for (let i = 0; i < arguments.length; i++) {
    const value = arguments[i];
    if (value === null || value === undefined || value === '') continue;
    const n = edPrice_num_(value);
    if (!isNaN(n)) return n;
  }
  return null;
}

function edPrice_parseLargestNumber_(value) {
  const text = String(value || '');
  const matches = text.match(/[0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]+(?:\.[0-9]+)?/g);
  if (!matches || matches.length === 0) return 0;
  const nums = matches.map((m) => edPrice_num_(m)).filter((n) => n > 0);
  if (nums.length === 0) return 0;
  return Math.max.apply(null, nums);
}


function edPrice_calcChangeFromPricesOrApiPercent_(currentPrice, prevClose, apiChangeAmount, apiRatePercent, sign) {
  const current = edPrice_num_(currentPrice);
  const prev = edPrice_num_(prevClose);

  if (current > 0 && prev > 0) {
    const changeAmount = current - prev;
    return {
      change_amount: changeAmount,
      change_rate: changeAmount / prev,
    };
  }

  let amount = edPrice_firstNumberAllowBlank_(apiChangeAmount);
  if (amount === null) amount = 0;
  amount = edPrice_signedNumBySignOrValue_(amount, sign);

  let rate = edPrice_firstNumberAllowBlank_(apiRatePercent);
  if (rate === null) {
    rate = 0;
  } else {
    // KIS의 등락률 필드는 국내/해외 모두 퍼센트 단위로 취급한다.
    // 예: 0.23 => 0.23% => 0.0023, -0.07 => -0.07% => -0.0007
    rate = edPrice_signedPercentToDecimal_(rate, sign);
  }

  return {
    change_amount: amount,
    change_rate: rate,
  };
}

function edPrice_signedPercentToDecimal_(percentValue, sign) {
  const raw = edPrice_num_(percentValue);
  if (raw === 0) return 0;
  const decimal = Math.abs(raw) / 100;
  const s = String(sign || '');
  const negativeByValue = String(percentValue || '').trim().startsWith('-') || raw < 0;
  if (negativeByValue || s === '4' || s === '5' || s === '-' || s.toUpperCase() === 'D') return -decimal;
  if (s === '3' || s === '0') return 0;
  return decimal;
}

function edPrice_signedNumBySign_(value, sign) {
  const raw = Math.abs(edPrice_num_(value));
  const s = String(sign || '');
  if (s === '4' || s === '5' || s === '-' || s.toUpperCase() === 'D') return -raw;
  if (s === '3' || s === '0') return 0;
  return raw;
}

function edPrice_signedNumBySignOrValue_(value, sign) {
  const text = String(value || '');
  if (text.trim().startsWith('-')) return edPrice_num_(value);
  return edPrice_signedNumBySign_(value, sign);
}

function edPrice_signedRateBySign_(percentValue, sign) {
  const rawRate = Math.abs(edPrice_num_(percentValue)) / 100;
  const s = String(sign || '');
  if (s === '4' || s === '5' || s === '-' || s.toUpperCase() === 'D') return -rawRate;
  if (s === '3' || s === '0') return 0;
  return rawRate;
}

function edPrice_normalizeRate_(value, sign) {
  const raw = edPrice_num_(value);
  if (raw === 0) return 0;
  const decimal = Math.abs(raw) > 0.2 ? Math.abs(raw) / 100 : Math.abs(raw);
  const s = String(sign || '');
  const negativeByValue = String(value || '').trim().startsWith('-') || raw < 0;
  if (negativeByValue || s === '4' || s === '5' || s === '-' || s.toUpperCase() === 'D') return -decimal;
  if (s === '3' || s === '0') return 0;
  return decimal;
}

function edPrice_isRateLimitError_(message) {
  const text = String(message || '');
  return text.indexOf('EGW00201') >= 0 || text.indexOf('초당 거래건수') >= 0;
}

function edPrice_getKisBaseUrl_() {
  return PropertiesService.getScriptProperties().getProperty(ED_MVP_PRICE_REFRESH.kis.baseUrlProperty)
    || ED_MVP_PRICE_REFRESH.kis.defaultBaseUrl;
}

function edPrice_toQueryString_(obj) {
  return Object.keys(obj)
    .map((key) => encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]))
    .join('&');
}

function edPrice_extractOutputCount_(output) {
  if (!output) return 0;
  if (typeof output === 'number') return output;
  if (output.output_count !== undefined) return Number(output.output_count) || 0;
  if (output.row_count !== undefined) return Number(output.row_count) || 0;
  if (output.count !== undefined) return Number(output.count) || 0;
  return 0;
}
