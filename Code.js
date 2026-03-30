/** * [최종 완결판] 
 * 1. 하이픈(-) 및 물결표(~) 반복수 범위 모두 지원
 * 2. 범위 밖 기록도 보여주는 Fallback 로직 탑재 (0kg 오류 및 기록 누락 방지)
 * 3. 모바일 스크롤 멈춤 해결 및 한글 헤더 완벽 지원
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSV 관리')
    .addItem('최신 운동기록 불러오기', 'importLatestCSV')
    .addToUi();
}

function doGet() {
  var output = HtmlService.createHtmlOutputFromFile('Index');
  output.setTitle('축분할 코칭 보드')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0') 
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

/* ==============================================
   PART 2. CSV 동기화 (시트1 저장)
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
    try { SpreadsheetApp.getUi().alert(resultMsg); } catch(e) {}
    return resultMsg;
  } catch (e) { return "❌ 에러: " + e.toString(); }
}

/* ==============================================
   PART 3. 데이터 분석 (지능형 분석 엔진)
   ============================================== */
function getRoutineData(program, week, day) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progSheet = ss.getSheetByName('축분할정규화프로그램');
  const historySheet = ss.getSheetByName('시트1');
  const typeSheet = ss.getSheetByName('부위별운동');

  if (!historySheet) return { error: true, message: "시트1 없음. 동기화 필요." };
  
  const progData = progSheet.getDataRange().getValues();
  const historyData = historySheet.getDataRange().getValues();
  const typeData = typeSheet ? typeSheet.getDataRange().getValues() : [];
  
  let replacementMap = {};
  if (typeData.length > 1) {
    for (let i = 1; i < typeData.length; i++) {
      const key = String(typeData[i][0]) + "_" + String(typeData[i][1]);
      if (!replacementMap[key]) replacementMap[key] = [];
      replacementMap[key].push(String(typeData[i][2]));
    }
  }

  const headers = historyData.length > 0 ? historyData[0].map(h => String(h).toLowerCase().trim().replace(/\s+/g, '')) : [];
  const excludeList = ['body', 'target', 'goal', 'plan', 'unit', '체중', '목표'];
  const colMap = {
    weight: findColIndex(headers, ['웨이트', 'weight', '무게', 'kg', 'lbs'], excludeList),
    name: findColIndex(headers, ['운동', 'exercise', 'name'], []),
    reps: findColIndex(headers, ['반복수', 'reps'], []),
    date: findColIndex(headers, ['헬스시작', '헬스 시작', '날짜', 'date'], []),
    note: findColIndex(headers, ['메모', 'notes'], [])
  };

  const targetWeek = String(week).replace(/[^0-9]/g, "");
  const targetDay = String(day).replace(/[^0-9]/g, "");
  let routine = [];
  
  for (let i = 1; i < progData.length; i++) {
    const row = progData[i];
    if (String(row[0]) === program && String(row[1]) === targetWeek && String(row[2]) === targetDay) {
      const name = row[5]; const reps = String(row[7]); const rpe = row[8];
      let sets = parseInt(row[6]); if (isNaN(sets)) sets = 1; 

      let isMerged = false;
      if (routine.length > 0) {
        let lastItem = routine[routine.length - 1];
        if (lastItem.name === name && lastItem.reps === reps && lastItem.rpe === rpe) {
          lastItem.sets += sets; isMerged = true;
        }
      }

      if (!isMerged) {
        const suggestion = findBestRecord(historyData, colMap, name, reps);
        routine.push({
          part: row[3], group: row[4], name: name, sets: sets, reps: reps, rpe: rpe,
          suggestion: suggestion, replacementKey: row[3] + "_" + row[4]
        });
      }
    }
  }
  return { routine: routine, replacements: replacementMap };
}

// [핵심] 기록 찾기 로직: 범위가 달라도 기록이 있으면 무조건 표시
function findBestRecord(historyData, idx, exName, targetRepsStr) {
  if (historyData.length < 2 || idx.name === -1) return "기록 없음";

  let minR = 0, maxR = 999;
  targetRepsStr = String(targetRepsStr);
  if (targetRepsStr.includes('~')) {
    const p = targetRepsStr.split('~'); minR = parseInt(p[0]); maxR = parseInt(p[1]);
  } else if (targetRepsStr.includes('-')) {
    const p = targetRepsStr.split('-'); minR = parseInt(p[0]); maxR = parseInt(p[1]);
  } else {
    minR = maxR = parseInt(targetRepsStr);
  }

  const targetNameNormalized = String(exName).trim().toLowerCase().replace(/\s+/g, '');
  let validCandidates = []; // 범위 내
  let allCandidates = [];   // 전체

  for (let k = 1; k < historyData.length; k++) {
    const histNameNorm = String(historyData[k][idx.name]).trim().toLowerCase().replace(/\s+/g, '');
    if (histNameNorm !== targetNameNormalized) continue;

    let w = 0; if (idx.weight > -1) w = parseFloat(String(historyData[k][idx.weight]).replace(/[^0-9.]/g, ""));
    let r = 0; if (idx.reps > -1) r = parseInt(historyData[k][idx.reps]);
    const note = (idx.note > -1 && historyData[k][idx.note]) ? ` [${historyData[k][idx.note]}]` : "";
    let date = (idx.date > -1 && historyData[k][idx.date]) ? new Date(historyData[k][idx.date]) : new Date(0);

    if (!isNaN(r) && !isNaN(w) && w > 0) {
      const rec = { weight: w, reps: r, memo: note, date: date, dateString: date.toDateString() };
      allCandidates.push(rec);
      if (r >= minR && r <= maxR) validCandidates.push(rec);
    }
  }

  // 1. 범위 내 최신 기록
  if (validCandidates.length > 0) {
    validCandidates.sort((a, b) => b.date - a.date);
    const top = validCandidates.filter(c => c.dateString === validCandidates[0].dateString).sort((a, b) => b.weight - a.weight)[0];
    return `${top.weight}kg x ${top.reps}회${top.memo}`;
  }
  // 2. (Fallback) 범위 밖이라도 가장 최근 기록 표시
  if (allCandidates.length > 0) {
    allCandidates.sort((a, b) => b.date - a.date);
    const top = allCandidates.filter(c => c.dateString === allCandidates[0].dateString).sort((a, b) => b.weight - a.weight)[0];
    return `${top.weight}kg x ${top.reps}회${top.memo} (범위밖)`;
  }
  return "기록 없음";
}

function findColIndex(headers, keywords, excludeKeywords) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    let isEx = false;
    for (let j = 0; j < excludeKeywords.length; j++) { if (h.includes(excludeKeywords[j])) { isEx = true; break; } }
    if (isEx) continue;
    for (let k = 0; k < keywords.length; k++) { if (h.includes(keywords[k])) return i; }
  }
  return -1;
}

function getSingleExSuggestion(exName, targetRepsStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('시트1');
  const historyData = historySheet.getDataRange().getValues();
  const headers = historyData.length > 0 ? historyData[0].map(h => String(h).toLowerCase().trim().replace(/\s+/g, '')) : [];
  const colMap = {
    weight: findColIndex(headers, ['웨이트', 'weight', '무게', 'kg'], ['body']),
    name: findColIndex(headers, ['운동', 'exercise', 'name'], []),
    reps: findColIndex(headers, ['반복수', 'reps'], []),
    date: findColIndex(headers, ['헬스시작', '헬스 시작', '날짜', 'date'], []),
    note: findColIndex(headers, ['메모', 'notes'], [])
  };
  return findBestRecord(historyData, colMap, exName, targetRepsStr);
}

/* ==============================================
   PART 4. 전체 데이터 1회 로드 (고도화용)
   ============================================== */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progSheet = ss.getSheetByName('축분할정규화프로그램');
  const historySheet = ss.getSheetByName('시트1');
  const typeSheet = ss.getSheetByName('부위별운동');

  if (!historySheet) return { error: true, message: "시트1 없음. 동기화 필요." };
  if (!progSheet)    return { error: true, message: "축분할정규화프로그램 시트 없음." };

  const progData    = progSheet.getDataRange().getValues();
  const historyData = historySheet.getDataRange().getValues();
  const typeData    = typeSheet ? typeSheet.getDataRange().getValues() : [];

  // 대체운동 맵
  const replacementMap = {};
  for (let i = 1; i < typeData.length; i++) {
    const key = String(typeData[i][0]) + "_" + String(typeData[i][1]);
    if (!replacementMap[key]) replacementMap[key] = [];
    replacementMap[key].push(String(typeData[i][2]));
  }

  // 기록 시트 컬럼 인덱스
  const headers = historyData.length > 0
    ? historyData[0].map(h => String(h).toLowerCase().trim().replace(/\s+/g, ''))
    : [];
  const excludeList = ['body', 'target', 'goal', 'plan', 'unit', '체중', '목표'];
  const colMap = {
    weight: findColIndex(headers, ['웨이트', 'weight', '무게', 'kg', 'lbs'], excludeList),
    name:   findColIndex(headers, ['운동', 'exercise', 'name'], []),
    reps:   findColIndex(headers, ['반복수', 'reps'], []),
    date:   findColIndex(headers, ['헬스시작', '헬스 시작', '날짜', 'date'], []),
    note:   findColIndex(headers, ['메모', 'notes'], [])
  };

  // 기록 배열 (유효한 행만)
  const historyRecords = [];
  for (let k = 1; k < historyData.length; k++) {
    const row = historyData[k];
    if (colMap.name === -1) continue;
    const name   = String(row[colMap.name]).trim();
    if (!name) continue;
    const weight = colMap.weight > -1 ? parseFloat(String(row[colMap.weight]).replace(/[^0-9.]/g, "")) : NaN;
    const reps   = colMap.reps   > -1 ? parseInt(row[colMap.reps]) : NaN;
    const note   = (colMap.note  > -1 && row[colMap.note]) ? String(row[colMap.note]) : "";
    const dateVal = colMap.date  > -1 ? row[colMap.date] : null;
    const date   = dateVal ? new Date(dateVal) : new Date(0);
    if (isNaN(weight) || weight <= 0 || isNaN(reps) || reps <= 0) continue;
    historyRecords.push({
      name:       name,
      weight:     weight,
      reps:       reps,
      note:       note,
      dateMs:     date.getTime(),
      dateString: date.toDateString()
    });
  }

  // 루틴 배열
  const routineRows = [];
  for (let i = 1; i < progData.length; i++) {
    const row = progData[i];
    routineRows.push({
      program: String(row[0]),
      week:    String(row[1]).replace(/[^0-9]/g, ""),
      day:     String(row[2]).replace(/[^0-9]/g, ""),
      part:    String(row[3]),
      group:   String(row[4]),
      name:    String(row[5]),
      sets:    parseInt(row[6]) || 1,
      reps:    String(row[7]),
      rpe:     String(row[8])
    });
  }

  return {
    routineRows:    routineRows,
    historyRecords: historyRecords,
    replacements:   replacementMap
  };
}