/** * [최종 통합본] 고성능 데이터 로드 + 기존 CSV 동기화 유지
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSV 관리')
    .addItem('최신 운동기록 불러오기', 'importLatestCSV')
    .addToUi();
}

/**
 * 앱 시작 시 모든 데이터를 한 번에 가져오는 함수
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const historySheet = ss.getSheetByName('시트1');
  const historyData = historySheet ? historySheet.getDataRange().getValues() : [];
  const historyHeaders = historyData.length > 0 ? historyData.shift() : [];

  const programSheet = ss.getSheetByName('축분할정규화프로그램');
  const programData = programSheet ? programSheet.getDataRange().getValues() : [];

  const exerciseSheet = ss.getSheetByName('부위별운동');
  const exerciseData = exerciseSheet ? exerciseSheet.getDataRange().getValues() : [];

  return {
    history: {
      headers: historyHeaders,
      rows: historyData
    },
    programs: programData,
    exercises: exerciseData
  };
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('축분할 코칭 보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* 기존 importLatestCSV 로직 유지 */
function importLatestCSV() {
  // 사용자님이 제공해주신 PART 2 코드를 여기에 그대로 유지하세요.
}