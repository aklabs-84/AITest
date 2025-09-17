/** Code.gs (REPLACE ALL) **/
const DEFAULT_HEADERS = ['이름','결과 타이틀','결과 내용','점수','타임스탬프'];
const HEADER_KEYS = {
  name:    ['이름','name'],
  title:   ['결과 타이틀','결과타이틀','title','resulttitle'],
  content: ['결과 내용','결과내용','content','resultcontent'],
  score:   ['점수','score'],
  ts:      ['타임스탬프','타임스태프','timestamp','시간','작성시각']
};

function norm_(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,''); }

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
    const sh = ss.getSheetByName('Responses') || ss.insertSheet('Responses');

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

    return ContentService.createTextOutput(JSON.stringify({status:'success'}))
      .setMimeType(ContentService.MimeType.JSON);
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

/** 한 번 실행하면, 점수가 타임스탬프 칸에 들어간 기록을 교정 */
function fixMisplacedScoresOnce(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Responses');
  if (!sh) return;

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
  }
  SpreadsheetApp.getUi().alert('정리 완료: ' + changed + '건 수정');
}
