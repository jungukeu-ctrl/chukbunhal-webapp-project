/** * [고도화 통합 버전] 
 * 1. 성능 최적화: getAllData() 추가 (전체 데이터 1회 로드)
 * 2. 기존 로직 유지: importLatestCSV() 동기화 기능 보존
 * 3. 스마트 매칭을 위한 준비 완료
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSV 관리')
    .addItem('최신 운동기록 불러오기', 'importLatestCSV')
    .addToUi();
}

/**
 * [신규] 앱 시작 시 필요한 모든 데이터를 한 번에 가져옵니다. (성능 최적화 핵심)
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 운동 기록 (시트1) 파싱
  const historySheet = ss.getSheetByName('시트1');
  const historyData = historySheet ? historySheet.getDataRange().getValues() : [];
  const historyHeaders = historyData.length > 0 ? historyData.shift() : [];

  // 2. 루틴 프로그램 파싱
  const programSheet = ss.getSheetByName('축분할정규화프로그램');
  const programData = programSheet ? programSheet.getDataRange().getValues() : [];

  // 3. 대체 운동 목록 파싱
  const exerciseSheet = ss.getSheetByName('부위별운동');
  const exerciseData = exerciseSheet ? exerciseSheet.getDataRange().getValues() : [];

  return {
    history: {
      headers: historyHeaders,
      rows: historyData
    },
    programs: programData,
    exercises: exerciseData,
    serverTime: new Date().toISOString()
  };
}

function doGet() {
  var output = HtmlService.createTemplateFromFile('Index').evaluate();
  output.setTitle('축분할 코칭 보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

/* ==============================================
   PART 2. CSV 동기화 (기존 로직 유지)
   ============================================== */
function importLatestCSV() {
  try {
    const searchPattern = "RepCount Excel Export(CSV)";
    const files = DriveApp.searchFiles("title contains '" + searchPattern + "' and trashed = false");
    let latestFile = null;
    let latestTime = 0;
    while (files.hasNext()) {
      const file = files.next();
      const lastUpdated = file.getLastUpdated().getTime();
      if (lastUpdated > latestTime) { latestFile = file; latestTime = lastUpdated; }
    }
    if (!latestFile) return "❌ 오류: 파일 없음";
    const blob = latestFile.getBlob();
    const csvData = Utilities.parseCsv(blob.getDataAsString('UTF-8'));

    let startRow = 0;
    for (let i = 0; i < Math.min(csvData.length, 20); i++) {
      const rowStr = csvData[i].join(',').toLowerCase();
      if (rowStr.includes('exercise') || rowStr.includes('운동')) { startRow = i; break; }
    }
    const cleanData = csvData.slice(startRow);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('시트1');
    if (!sheet) sheet = ss.insertSheet('시트1');
    sheet.clear();
    if (cleanData.length > 0) sheet.getRange(1, 1, cleanData.length, cleanData[0].length).setValues(cleanData);

    const resultMsg = "✅ 동기화 완료! (" + cleanData.length + "행)";
    try { SpreadsheetApp.getUi().alert(resultMsg); } catch (e) { }
    return resultMsg;
  } catch (e) { return "❌ 에러: " + e.toString(); }
}