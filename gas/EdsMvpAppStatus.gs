/*******************************************************
 * ED's MVP - App Status API Helper v0.8.8
 *
 * 목적:
 * - PWA 설정 탭에서 Apps Script 배포/스프레드시트 상태를 확인
 * - PWA v0.7.4의 getAppStatus action에서 호출
 *
 * 적용:
 * 1) Apps Script에 EdsMvpAppStatus.gs 새 파일 생성
 * 2) 이 코드 전체 붙여넣기
 * 3) EdsMvpApi.gs switch(action)에 getAppStatus case 추가
 * 4) 웹앱 새 버전 배포
 *******************************************************/

const ED_MVP_VERSION_TAG = "ED's MVP - App Status API Helper v0.8.8";

const ED_MVP_APP_STATUS = {
  appsScriptVersion: ED_MVP_VERSION_TAG.split(' v')[1],
  appName: "ED's MVP",
  watchedSheets: [
    '2. 종목현황',
    '6. 배당내역',
    'App_Accounts',
    'App_Assets',
    'App_Holdings',
    'App_Prices',
    'App_Output',
    'App_Dividends',
    'App_ChartPrices',
    'App_PriceRefreshResult',
    'App_SyncLog',
  ],
};

function getAppStatus(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {};

  ED_MVP_APP_STATUS.watchedSheets.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheets[name] = { exists: false };
      return;
    }
    sheets[name] = {
      exists: true,
      last_row: sheet.getLastRow(),
      last_column: sheet.getLastColumn(),
    };
  });

  return {
    app: ED_MVP_APP_STATUS.appName,
    pwa_version: payload && payload.pwa_version ? payload.pwa_version : '',
    apps_script_version: ED_MVP_APP_STATUS.appsScriptVersion,
    spreadsheet_name: ss.getName(),
    spreadsheet_id: ss.getId(),
    spreadsheet_url: ss.getUrl(),
    api_time: new Date(),
    timezone: Session.getScriptTimeZone(),
    sheets,
  };
}
