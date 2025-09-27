/** Code.gs (날짜별 시트 저장 버전) **/
const DEFAULT_HEADERS = ['이름','결과 타이틀','결과 내용','점수','타임스탬프'];
const HEADER_KEYS = {
  name:    ['이름','name'],
  title:   ['결과 타이틀','결과타이틀','title','resulttitle'],
  content: ['결과 내용','결과내용','content','resultcontent'],
  score:   ['점수','score'],
  ts:      ['타임스탬프','타임스태프','timestamp','시간','작성시각']
};

function norm_(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,''); }

// 날짜 기반 시트명 생성 함수
function getDateSheetName_(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// 날짜별 시트 가져오기 또는 생성
function getOrCreateDateSheet_(ss, date) {
  const sheetName = getDateSheetName_(date);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // 새 시트 생성
    sheet = ss.insertSheet(sheetName);
    
    // 시트 탭 색상 설정 (선택사항)
    const colors = ['#ff9999', '#99ccff', '#99ff99', '#ffcc99', '#cc99ff', '#ffff99'];
    const colorIndex = Math.floor(Math.random() * colors.length);
    sheet.setTabColor(colors[colorIndex]);
  }
  
  return sheet;
}

function ensureHeaders_(sh){
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,DEFAULT_HEADERS.length).setValues([DEFAULT_HEADERS]);
    sh.getRange(1,1,1,DEFAULT_HEADERS.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    return;
  }
  const lastCol = sh.getLastColumn();
  const row1 = sh.getRange(1,1,1,lastCol).getValues()[0];
  const normRow = row1.map(norm_);

  // 각 키의 대표 헤더가 없으면 추가, 비표준 변형이면 표준 명칭으로 교체
  const want = {
    name: '이름', title: '결과 타이틀', content: '결과 내용', score: '점수', ts: '타임스탬프'
  };
  Object.entries(HEADER_KEYS).forEach(([k, alts])=>{
    const idx = normRow.findIndex(h => alts.includes(h));
    if (idx === -1) {
      // 맨 뒤에 새로 추가
      sh.getRange(1, sh.getLastColumn()+1).setValue(want[k]);
      sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
    } else {
      // 표준 명칭으로 교체(보기 깔끔, 차트 스크립트 호환)
      sh.getRange(1, idx+1).setValue(want[k]);
    }
  });
}

function findCol_(sh, key){ // key in HEADER_KEYS
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const normHeaders = headers.map(norm_);
  const alts = HEADER_KEYS[key];
  const idx = normHeaders.findIndex(h => alts.includes(h));
  return idx === -1 ? null : idx+1; // 1-based
}

function looksScore_(v){ return /\d+\s*(점|\/|\d)/.test(String(v||'')); }

function doPost(e){
  try{
    const SPREADSHEET_ID = '1dDDeUp4rGpn9WMzV0rmD6Obq4KoIZq_M9O1-qkqu_Ts';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 현재 날짜로 시트 결정
    const currentDate = new Date();
    const sh = getOrCreateDateSheet_(ss, currentDate);

    ensureHeaders_(sh);

    // 페이로드 파싱
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try{ data = JSON.parse(e.postData.contents); }
      catch{ data = e.parameter || {}; }
    } else {
      data = e ? (e.parameter || {}) : {};
    }

    // 값 준비
    const rawScore = data.score || data.resultScore || data['점수'] ||
                     (looksScore_(data.timestamp) ? data.timestamp : '');
    const rawTs = (!looksScore_(data.timestamp) && data.timestamp) ? data.timestamp
                 : new Date().toLocaleString('ko-KR');

    // 열 찾기
    const nameCol = findCol_(sh,'name');
    const titleCol = findCol_(sh,'title');
    const contentCol = findCol_(sh,'content');
    const scoreCol = findCol_(sh,'score');
    const tsCol = findCol_(sh,'ts');

    const rowLen = sh.getLastColumn();
    const row = new Array(rowLen).fill('');

    if (nameCol)   row[nameCol-1]   = data.name || '';
    if (titleCol)  row[titleCol-1]  = data.resultTitle || '';
    if (contentCol)row[contentCol-1]= data.resultContent || '';
    if (scoreCol)  row[scoreCol-1]  = rawScore || '';
    if (tsCol)     row[tsCol-1]     = rawTs;

    sh.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status:'success', 
      sheetName: sh.getName(),
      message: `${sh.getName()} 시트에 저장되었습니다.`
    })).setMimeType(ContentService.MimeType.JSON);
    
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({status:'error', message:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// (선택) CORS 프리플라이트 허용
function doOptions(e){
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

// 디버그용
function doGet(){
  return ContentService.createTextOutput('Apps Script is working!')
    .setMimeType(ContentService.MimeType.TEXT);
}

// 모든 날짜 시트의 데이터를 통합하여 요약 시트 생성
function createSummarySheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheetName = '전체_요약';
    
    // 기존 요약 시트 삭제 후 새로 생성
    let summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) {
      ss.deleteSheet(summarySheet);
    }
    summarySheet = ss.insertSheet(summarySheetName);
    
    // 요약 시트 헤더 설정
    const summaryHeaders = ['날짜', '이름', '결과 타이틀', '결과 내용', '점수', '타임스탬프'];
    summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
    summarySheet.getRange(1, 1, 1, summaryHeaders.length)
      .setFontWeight('bold')
      .setBackground('#34495e')
      .setFontColor('white');
    
    const allSheets = ss.getSheets();
    const dateSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return /^\d{4}-\d{2}-\d{2}$/.test(name); // YYYY-MM-DD 형식만
    });
    
    let summaryRow = 2;
    
    // 각 날짜 시트에서 데이터 수집
    dateSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) { // 헤더 외에 데이터가 있는 경우
        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        
        data.forEach(row => {
          if (row.some(cell => cell !== '')) { // 빈 행이 아닌 경우
            const summaryRowData = [sheetName, ...row];
            summarySheet.getRange(summaryRow, 1, 1, summaryRowData.length).setValues([summaryRowData]);
            summaryRow++;
          }
        });
      }
    });
    
    // 요약 시트를 첫 번째 위치로 이동
    ss.moveSheet(summarySheet, 1);
    
    SpreadsheetApp.getUi().alert(`요약 시트가 생성되었습니다!\n총 ${dateSheets.length}개의 날짜 시트에서 데이터를 통합했습니다.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('요약 시트 생성 중 오류 발생: ' + error.toString());
  }
}

// 특정 날짜의 데이터만 삭제하는 함수
function deleteDataByDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('날짜별 데이터 삭제', 'YYYY-MM-DD 형식으로 삭제할 날짜를 입력하세요:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const dateInput = response.getResponseText().trim();
    
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateInput)) {
      ui.alert('올바른 날짜 형식(YYYY-MM-DD)을 입력해주세요.');
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dateInput);
    
    if (sheet) {
      const confirmResponse = ui.alert(`${dateInput} 시트를 삭제하시겠습니까?`, ui.ButtonSet.YES_NO);
      if (confirmResponse == ui.Button.YES) {
        ss.deleteSheet(sheet);
        ui.alert(`${dateInput} 시트가 삭제되었습니다.`);
      }
    } else {
      ui.alert(`${dateInput} 시트를 찾을 수 없습니다.`);
    }
  }
}

/** 한 번 실행하면, 점수가 타임스탬프 칸에 들어간 기록을 교정 */
function fixMisplacedScoresOnce(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const dateSheets = allSheets.filter(sheet => {
    const name = sheet.getName();
    return /^\d{4}-\d{2}-\d{2}$/.test(name) || name === 'Responses';
  });
  
  let totalChanged = 0;
  
  dateSheets.forEach(sh => {
    ensureHeaders_(sh);
    const scoreCol = findCol_(sh,'score');
    const tsCol = findCol_(sh,'ts');
    if (!scoreCol || !tsCol) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const scores = sh.getRange(2, scoreCol, lastRow-1, 1).getValues();
    const tsVals = sh.getRange(2, tsCol,   lastRow-1, 1).getValues();

    let changed = 0;
    for (let i=0;i<scores.length;i++){
      if (!scores[i][0] && looksScore_(tsVals[i][0])) {
        scores[i][0] = tsVals[i][0];
        tsVals[i][0] = '';
        changed++;
      }
    }
    if (changed){
      sh.getRange(2, scoreCol, lastRow-1, 1).setValues(scores);
      sh.getRange(2, tsCol,   lastRow-1, 1).setValues(tsVals);
      totalChanged += changed;
    }
  });
  
  SpreadsheetApp.getUi().alert(`정리 완료: 총 ${totalChanged}건 수정`);
}
