/*******************************************************
 * ED's MVP - Main Sheet Mapping Inspector v0.8.8
 *
 * 목적:
 * - 기존 "2. 종목현황" 시트 구조 분석
 * - 시트명 공백 차이 자동 보정
 * - alert() 대신 toast() 사용
 *
 * 실행 함수:
 * - listAllSheetNamesForDebug()
 * - inspectMainPortfolioSheet()
 *******************************************************/

const ED_MVP_MAPPING = {
  mainSheetName: "2. 종목현황",
  mainSheetNameCandidates: [
    "2. 종목현황",
    "2.종목현황",
    "종목현황",
  ],
  mapSheetName: "App_MainSheetMap",
  scanRows: 120,
  scanCols: 80,
};

function listAllSheetNamesForDebug() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ss.getSheets().map((sheet) => [sheet.getName()]);

  let debugSheet = ss.getSheetByName("App_Debug_SheetNames");
  if (!debugSheet) {
    debugSheet = ss.insertSheet("App_Debug_SheetNames");
  }

  debugSheet.clear();
  debugSheet.getRange(1, 1, 1, 1).setValues([["sheet_name"]]);

  if (names.length > 0) {
    debugSheet.getRange(2, 1, names.length, 1).setValues(names);
  }

  debugSheet.setFrozenRows(1);
  debugSheet.autoResizeColumns(1, 1);

  ss.toast("시트명 목록 생성 완료: App_Debug_SheetNames", "ED's MVP", 5);
}

function inspectMainPortfolioSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = findMainSheet_(ss);

  if (!mainSheet) {
    writeMappingErrorSheet_(ss, "기존 원장 시트를 찾을 수 없습니다.");
    ss.toast("기존 원장 시트를 찾을 수 없습니다. App_Debug_SheetNames를 확인하세요.", "ED's MVP", 8);
    listAllSheetNamesForDebug();
    return;
  }

  let mapSheet = ss.getSheetByName(ED_MVP_MAPPING.mapSheetName);
  if (!mapSheet) {
    mapSheet = ss.insertSheet(ED_MVP_MAPPING.mapSheetName);
  }

  mapSheet.clear();

  const lastRow = Math.min(mainSheet.getLastRow(), ED_MVP_MAPPING.scanRows);
  const lastCol = Math.min(mainSheet.getLastColumn(), ED_MVP_MAPPING.scanCols);

  const values = mainSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  const output = [];
  output.push([
    "source_sheet",
    "row",
    "col",
    "col_letter",
    "cell_value",
    "header_candidate_score",
    "matched_keyword",
  ]);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cellValue = String(values[r][c] || "").trim();
      if (!cellValue) continue;

      const match = scoreHeaderCandidate_(cellValue);

      output.push([
        mainSheet.getName(),
        r + 1,
        c + 1,
        columnToLetter_(c + 1),
        cellValue,
        match.score,
        match.keyword,
      ]);
    }
  }

  mapSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  mapSheet.setFrozenRows(1);
  mapSheet.autoResizeColumns(1, output[0].length);

  const headerRange = mapSheet.getRange(1, 1, 1, output[0].length);
  headerRange
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#137333")
    .setHorizontalAlignment("center");

  ss.toast(
    "2. 종목현황 구조 분석 완료: App_MainSheetMap 확인",
    "ED's MVP",
    5
  );
}

function findMainSheet_(ss) {
  for (const candidate of ED_MVP_MAPPING.mainSheetNameCandidates) {
    const sheet = ss.getSheetByName(candidate);
    if (sheet) return sheet;
  }

  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    const name = sheet.getName().replace(/\s/g, "");
    if (name === "2.종목현황" || name.includes("종목현황")) {
      return sheet;
    }
  }

  return null;
}

function writeMappingErrorSheet_(ss, message) {
  let mapSheet = ss.getSheetByName(ED_MVP_MAPPING.mapSheetName);
  if (!mapSheet) {
    mapSheet = ss.insertSheet(ED_MVP_MAPPING.mapSheetName);
  }

  mapSheet.clear();
  mapSheet.getRange(1, 1, 1, 2).setValues([["error", "message"]]);
  mapSheet.getRange(2, 1, 1, 2).setValues([["main_sheet_not_found", message]]);
  mapSheet.autoResizeColumns(1, 2);
}

function scoreHeaderCandidate_(value) {
  const normalized = value.replace(/\s/g, "");

  const keywords = [
    "증권사",
    "계좌",
    "계좌형식",
    "국가",
    "종목코드",
    "티커",
    "종목명",
    "수량",
    "평단가",
    "현재가",
    "전일대비",
    "등락",
    "평가액",
    "투자비중",
    "목표비중",
    "수익",
    "수익률",
    "누적배당",
    "총수익",
  ];

  let bestKeyword = "";
  let score = 0;

  keywords.forEach((keyword) => {
    const normalizedKeyword = keyword.replace(/\s/g, "");

    if (normalized === normalizedKeyword) {
      if (score < 100) {
        score = 100;
        bestKeyword = keyword;
      }
    } else if (normalized.indexOf(normalizedKeyword) >= 0) {
      if (score < 70) {
        score = 70;
        bestKeyword = keyword;
      }
    }
  });

  return {
    score,
    keyword: bestKeyword,
  };
}

function columnToLetter_(column) {
  let temp = "";
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}