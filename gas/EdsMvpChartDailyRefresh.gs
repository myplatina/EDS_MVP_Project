/*******************************************************
 * ED's MVP - KRX Daily Chart Bulk Refresh v0.8.61
 *
 * 목적:
 * - PWA 설정 탭의 `국내 일봉 차트 갱신(KIS)` 버튼용 백엔드
 * - KRX 보유종목의 D 일봉 차트를 KIS에서 조회
 * - App_ChartPrices의 해당 KRX D 일봉만 일괄 교체
 * - 기존 refreshAllKrxDailyChartsFromKis()의 deleteRow 반복/timeout 문제 회피
 *
 * 전제:
 * - setupKisCredentials() 완료
 * - EdsMvpKisChartTest.gs 또는 동등 KIS helper 존재
 *   edKis_fetchDomesticChart_ / edKis_formatDate_ / edKis_addDays_ / edKis_periodLookbackDays_
 * - App_Assets, App_ChartPrices 존재
 *******************************************************/

function refreshKrxDailyChartsFromKis() {
  return refreshAllKrxDailyChartsFromKisFast();
}

function refreshAllKrxDailyChartsFromKisFast() {
  if (typeof setupChartPricesSheet === 'function') {
    setupChartPricesSheet();
  } else {
    edDaily_setupChartPricesSheet_();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startedAt = new Date();

  const assets = edDaily_readSheetAsObjects_('App_Assets')
    .filter((row) => String(row.enabled || '').toUpperCase() === 'TRUE')
    .filter((row) => String(row.market || '') === 'KRX')
    .filter((row) => String(row.ticker || '').toUpperCase().match(/^[0-9A-Z]{6}$/));

  const interval = 'D';
  const limit = 120;
  const today = new Date();
  const end = edKis_formatDate_(today);
  const start = edKis_formatDate_(edKis_addDays_(today, -edKis_periodLookbackDays_(interval, limit)));

  const targetAssetIds = new Set();
  const newChartRows = [];
  const results = [];

  // --- fetchAll 병렬 배치 처리 ---
  // 1. KIS 토큰/자격증명 1회 취득
  const token = edKis_getAccessToken_();
  const props = PropertiesService.getScriptProperties();
  const appKey = props.getProperty(ED_KIS_HELPER.properties.appKey);
  const appSecret = props.getProperty(ED_KIS_HELPER.properties.appSecret);

  if (!appKey || !appSecret) {
    throw new Error('KIS appkey/appsecret이 없습니다. setupKisCredentials()를 먼저 실행하세요.');
  }

  // 2. 마유효 종목만 대상 필터링
  const validAssets = assets.filter((asset) => {
    const assetId = String(asset.asset_id || '').trim();
    const ticker = edDaily_normalizeTicker_(asset.ticker);
    if (!assetId || !ticker) {
      results.push([assetId, ticker, interval, 'skip', 0, 'asset_id 또는 ticker 없음']);
      return false;
    }
    targetAssetIds.add(assetId);
    return true;
  });

  // 3. 요청 객체 배열 조립
  const kisBaseUrl = edKis_getBaseUrl_();
  const requestSpecs = validAssets.map((asset) => {
    const ticker = edDaily_normalizeTicker_(asset.ticker);
    const query = {
      FID_COND_MRKT_DIV_CODE: 'J',
      FID_INPUT_ISCD: ticker,
      FID_INPUT_DATE_1: start,
      FID_INPUT_DATE_2: end,
      FID_PERIOD_DIV_CODE: interval,
      FID_ORG_ADJ_PRC: '0',
    };
    return {
      url: kisBaseUrl + ED_KIS_HELPER.chartPath + '?' + edKis_toQueryString_(query),
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token,
        appkey: appKey,
        appsecret: appSecret,
        tr_id: ED_KIS_HELPER.chartTrId,
        custtype: 'P',
      },
      muteHttpExceptions: true,
    };
  });

  // 4. 병렬 발사 (15개씩 chunk, chunk 사이 150ms 대기)
  const CHUNK_SIZE = 15;
  const allResponses = [];
  for (let chunkStart = 0; chunkStart < requestSpecs.length; chunkStart += CHUNK_SIZE) {
    if (chunkStart > 0) Utilities.sleep(150);
    const chunkRes = UrlFetchApp.fetchAll(requestSpecs.slice(chunkStart, chunkStart + CHUNK_SIZE));
    chunkRes.forEach((r) => allResponses.push(r));
  }

  // 5. 응답 일괄 처리
  const now = new Date();
  allResponses.forEach((res, i) => {
    const asset = validAssets[i];
    const assetId = String(asset.asset_id || '').trim();
    const ticker = edDaily_normalizeTicker_(asset.ticker);
    try {
      const status = res.getResponseCode();
      const text = res.getContentText();
      let json;
      try { json = JSON.parse(text); } catch (e) {
        throw new Error(`KIS chart JSON parse 실패. HTTP=${status}, body=${text.slice(0, 400)}`);
      }
      if (status < 200 || status >= 300) {
        throw new Error(`KIS chart HTTP 오류. HTTP=${status}, body=${text.slice(0, 400)}`);
      }
      if (String(json.rt_cd || '') !== '0') {
        throw new Error(`KIS chart API 오류. rt_cd=${json.rt_cd}, msg_cd=${json.msg_cd}, msg=${json.msg1}`);
      }

      const rawItems = Array.isArray(json.output2) ? json.output2 : [];
      const parsedItems = rawItems
        .map((row) => edKis_normalizeChartRow_(row))
        .filter((row) => row.date)
        .sort((a, b) => String(a.date).localeCompare(String(b.date)))
        .slice(-limit);

      parsedItems.forEach((item) => {
        newChartRows.push([
          `${assetId}_${interval}_${item.date}`,
          assetId, ticker, interval,
          item.date, item.open, item.high, item.low, item.close, item.volume,
          'kis', now, now, now,
        ]);
      });

      results.push([assetId, ticker, interval, 'success', parsedItems.length,
        json.msg1 || '']);
    } catch (e) {
      results.push([assetId, ticker, interval, 'error', 0,
        e && e.message ? e.message : String(e)]);
    }
  });

  edDaily_replaceDailyRowsInBulk_(targetAssetIds, newChartRows);
  edDaily_writeRefreshSummary_('App_ChartRefreshResult', results);

  const finishedAt = new Date();
  const summary = {
    started_at: startedAt,
    finished_at: finishedAt,
    target_count: assets.length,
    row_count: newChartRows.length,
    success_count: results.filter((r) => r[3] === 'success').length,
    error_count: results.filter((r) => r[3] === 'error').length,
    skipped_count: results.filter((r) => r[3] === 'skip').length,
  };

  ss.toast(
    `KRX 일봉 차트 갱신 완료: success=${summary.success_count}, error=${summary.error_count}, rows=${summary.row_count}`,
    "ED's MVP",
    8
  );

  return summary;
}

function edDaily_replaceDailyRowsInBulk_(targetAssetIds, newRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('App_ChartPrices');
  if (!sheet) {
    edDaily_setupChartPricesSheet_();
    sheet = ss.getSheetByName('App_ChartPrices');
  }
  if (!sheet) throw new Error('App_ChartPrices 시트를 찾을 수 없습니다.');

  const headers = edDaily_chartHeaders_();
  const lastRow = sheet.getLastRow();
  let keptRows = [];

  if (lastRow >= 2) {
    const width = Math.max(sheet.getLastColumn(), headers.length);
    const values = sheet.getRange(1, 1, lastRow, width).getValues();
    const existingHeader = values[0];
    const assetCol = existingHeader.indexOf('asset_id');
    const intervalCol = existingHeader.indexOf('interval');

    if (assetCol < 0 || intervalCol < 0) {
      throw new Error('App_ChartPrices 헤더 오류: asset_id 또는 interval 없음');
    }

    keptRows = values.slice(1).filter((row) => {
      const assetId = String(row[assetCol] || '');
      const interval = String(row[intervalCol] || '');
      if (interval === 'D' && targetAssetIds.has(assetId)) return false;
      return true;
    }).map((row) => {
      const normalized = row.slice(0, headers.length);
      while (normalized.length < headers.length) normalized.push('');
      return normalized;
    });
  }

  const output = [headers].concat(keptRows).concat(newRows);
  sheet.clear();
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#137333')
    .setHorizontalAlignment('center');
}

function edDaily_fetchKisChartWithRetry_(ticker, interval, start, end) {
  const maxAttempts = 3;
  const baseSleepMs = 1500;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return edKis_fetchDomesticChart_(ticker, interval, start, end);
    } catch (e) {
      lastError = e;
      const message = e && e.message ? e.message : String(e);
      if (message.indexOf('EGW00201') >= 0 || message.indexOf('초당 거래건수') >= 0) {
        Utilities.sleep(baseSleepMs * attempt);
        continue;
      }
      throw e;
    }
  }

  throw lastError;
}

function edDaily_setupChartPricesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('App_ChartPrices');
  if (!sheet) sheet = ss.insertSheet('App_ChartPrices');

  const headers = edDaily_chartHeaders_();
  const firstRow = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0]
    : [];

  if (String(firstRow[0] || '') !== 'chart_id') {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function edDaily_chartHeaders_() {
  return [
    'chart_id',
    'asset_id',
    'ticker',
    'interval',
    'date',
    'open',
    'high',
    'low',
    'close',
    'volume',
    'source',
    'fetched_at',
    'created_at',
    'updated_at',
  ];
}

function edDaily_readSheetAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  return values.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function edDaily_writeRefreshSummary_(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  const output = [['asset_id', 'ticker', 'interval', 'status', 'count', 'message']].concat(rows);
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, output[0].length);
  sheet.getRange(1, 1, 1, output[0].length).setFontWeight('bold').setBackground('#137333').setFontColor('#ffffff');
}

function edDaily_normalizeTicker_(value) {
  const text = String(value || '').trim().toUpperCase();
  if (/^\d+$/.test(text) && text.length < 6) return text.padStart(6, '0');
  return text;
}
