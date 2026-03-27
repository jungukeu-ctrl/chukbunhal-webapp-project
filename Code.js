/** * [최종 통합본] 성능 최적화 + 지능형 RPE 분석
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('CSV 관리')
    .addItem('최신 운동기록 불러오기', 'importLatestCSV').addToUi();
}

function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('시트1');
  const historyData = historySheet ? historySheet.getDataRange().getValues() : [];
  const historyHeaders = historyData.length > 0 ? historyData.shift() : [];

  return {
    history: { headers: historyHeaders, rows: historyData },
    programs: ss.getSheetByName('축분할정규화프로그램').getDataRange().getValues(),
    exercises: ss.getSheetByName('부위별운동') ? ss.getSheetByName('부위별운동').getDataRange().getValues() : [],
    serverTime: new Date().toISOString()
  };
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('축분할 코칭 보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function importLatestCSV() {
  // 기존 동기화 로직 유지 (사용자님의 기존 코드를 여기에 넣어주세요)
}