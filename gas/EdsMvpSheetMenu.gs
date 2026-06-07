/*******************************************************
 * ED's MVP - Sheet Menu / Button Actions
 * 목적:
 * - 구글 시트 상단 메뉴에서 주요 작업 실행
 * - 시트 내 버튼에 연결할 안전 wrapper 제공
 *******************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ED's MVP")
    .addItem("국내 현재가 갱신 - 앱 데이터만", "menuRefreshKrxPricesAppOnly")
    .addItem("국내 현재가 갱신 - 원장 K/M열 반영", "menuRefreshKrxPricesToMainSheet")
    .addSeparator()
    .addItem("원장 → 앱 데이터 재이관", "menuResetAndImportMainSheet")
    .addItem("배당내역 다시 가져오기", "menuResetAndImportDividends")
    .addSeparator()
    .addItem("검산 실행", "menuValidateEdsMvp")
    .addToUi();
}

/**
 * App_Prices / App_Output만 갱신
 * 원본 2. 종목현황은 수정하지 않음
 */
function menuRefreshKrxPricesAppOnly() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "국내 현재가 갱신",
    "KIS API로 국내 현재가를 갱신합니다.\n\n대상: App_Prices / App_Output\n원본 2. 종목현황은 수정하지 않습니다.",
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  const output = refreshKrxPricesFromKis();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `앱 데이터 갱신 완료: success=${output.success_count}, error=${output.error_count}`,
    "ED's MVP",
    8
  );
}

/**
 * 원본 2. 종목현황 K/M열까지 갱신
 * 수식이 값으로 대체될 수 있으므로 확인창 필수
 */
function menuRefreshKrxPricesToMainSheet() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "원장 가격 반영",
    "KIS API 현재가를 원본 2. 종목현황에 반영합니다.\n\n반영 위치:\n- K열 현재가\n- M열 전일 대비 등락\n\n기존 GOOGLEFINANCE / NAVER 수식이 값으로 대체될 수 있습니다.\n진행할까요?",
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  const output = refreshKrxPricesToMainSheetFromKis();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `원장 가격 반영 완료: success=${output.success_count}, error=${output.error_count}, main=${output.main_sheet_updated_count}`,
    "ED's MVP",
    8
  );
}

/**
 * 2. 종목현황 → App_* 재이관
 */
function menuResetAndImportMainSheet() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "원장 재이관",
    "2. 종목현황 기준으로 App_* 시트를 다시 생성합니다.\n\n앱에서만 수정하고 아직 원장에 반영하지 않은 값은 사라질 수 있습니다.\n진행할까요?",
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  const output = resetAndRunInitialImportAndBuildOutput();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "원장 → 앱 데이터 재이관 완료",
    "ED's MVP",
    8
  );

  return output;
}

/**
 * 6. 배당내역 → App_Dividends 재이관
 */
function menuResetAndImportDividends() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "배당내역 다시 가져오기",
    "6. 배당내역 기준으로 App_Dividends를 다시 생성합니다.\n진행할까요?",
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  const output = resetAndImportDividendSheetToApp();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `배당내역 가져오기 완료: imported=${output.imported_count || 0}, skipped=${output.skipped_count || 0}, error=${output.error_count || 0}`,
    "ED's MVP",
    8
  );

  return output;
}

/**
 * 검산 실행
 */
function menuValidateEdsMvp() {
  const output = validateEdsMvpAppData();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "검산 완료. App_ValidationReport를 확인하세요.",
    "ED's MVP",
    8
  );

  return output;
}
