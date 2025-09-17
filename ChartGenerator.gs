/** ChartGenerator.gs - 전체 교체 **/
/* 헤더 찾기(오타/영문 허용) */
function getColByHeader_(sheet, headerName) {
  const HEADER_ALTS = {
    '이름': ['이름','name'],
    '점수': ['점수','score'],
    '타임스탬프': ['타임스탬프','타임스태프','timestamp','시간','작성시각']
  };
  const alts = HEADER_ALTS[headerName] || [headerName];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];

  const norm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,'');
  const normHeaders = headers.map(norm);
  const normAlts = alts.map(norm);

  const idx = normHeaders.findIndex(h => normAlts.includes(h));
  return idx === -1 ? null : idx + 1; // 1-based
}

/* "60점 / 80점" → 75 같은 퍼센트 정수로 변환 */
function parseScoreToPercent_(cell) {
  if (!cell) return NaN;
  const str = String(cell).trim();
  const nums = str.match(/\d+/g);
  if (!nums) return NaN;

  if (nums.length >= 2) {
    const score = parseInt(nums[0], 10);
    const total = parseInt(nums[1], 10);
    if (total > 0) return Math.round((score / total) * 100);
    return NaN;
  }
  return parseInt(nums[0], 10);
}

/* 막대 차트 생성 */
function createScoreChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Responses');
  if (!sheet) { SpreadsheetApp.getUi().alert('Responses 시트를 찾을 수 없습니다.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('데이터가 충분하지 않습니다.'); return; }

  // 기존 차트 제거
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  // 열 찾기
  const nameCol  = getColByHeader_(sheet, '이름');
  const scoreCol = getColByHeader_(sheet, '점수');
  if (!nameCol || !scoreCol) { SpreadsheetApp.getUi().alert('헤더(이름/점수)를 찾을 수 없습니다.'); return; }

  // 데이터 읽기
  const names  = sheet.getRange(2, nameCol,  lastRow - 1, 1).getValues().flat();
  const scores = sheet.getRange(2, scoreCol, lastRow - 1, 1).getValues().flat();

  // 유효 데이터(퍼센트) 구성
  const validData = [];
  for (let i = 0; i < names.length; i++) {
    const p = parseScoreToPercent_(scores[i]);
    if (names[i] && !isNaN(p)) validData.push([names[i], p]);
  }
  if (validData.length === 0) { SpreadsheetApp.getUi().alert('유효한 점수 데이터가 없습니다.'); return; }

  // 임시 시트
  const tempName = '차트_데이터_임시';
  let temp = ss.getSheetByName(tempName);
  if (temp) ss.deleteSheet(temp);
  temp = ss.insertSheet(tempName);

  temp.getRange(1,1).setValue('이름');
  temp.getRange(1,2).setValue('점수(%)');
  validData.forEach((row, i) => temp.getRange(i + 2, 1, 1, 2).setValues([row]));

  // 차트
  const chart = temp.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(temp.getRange(1, 1, validData.length + 1, 2))
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'AI 활용 점수 현황')
    .setOption('hAxis', { title: '이름', titleTextStyle: { fontSize: 12, bold: true } })
    .setOption('vAxis', { title: '점수(%)', minValue: 0, maxValue: 100, titleTextStyle: { fontSize: 12, bold: true } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 600)
    .setOption('height', 400)
    .setOption('colors', ['#4285F4'])
    .build();

  sheet.insertChart(chart);
  temp.hideSheet();

  SpreadsheetApp.getUi().alert('차트가 성공적으로 생성되었습니다!');
}

/* 원형 차트 생성 */
function createPieChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Responses');
  if (!sheet) { SpreadsheetApp.getUi().alert('Responses 시트를 찾을 수 없습니다.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('데이터가 충분하지 않습니다.'); return; }

  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const nameCol  = getColByHeader_(sheet, '이름');
  const scoreCol = getColByHeader_(sheet, '점수');
  if (!nameCol || !scoreCol) { SpreadsheetApp.getUi().alert('헤더(이름/점수)를 찾을 수 없습니다.'); return; }

  const names  = sheet.getRange(2, nameCol,  lastRow - 1, 1).getValues().flat();
  const scores = sheet.getRange(2, scoreCol, lastRow - 1, 1).getValues().flat();

  const validData = [];
  for (let i = 0; i < names.length; i++) {
    const p = parseScoreToPercent_(scores[i]);
    if (names[i] && !isNaN(p)) validData.push([names[i], p]);
  }
  if (validData.length === 0) { SpreadsheetApp.getUi().alert('유효한 점수 데이터가 없습니다.'); return; }

  const tempName = '차트_데이터_임시';
  let temp = ss.getSheetByName(tempName);
  if (temp) ss.deleteSheet(temp);
  temp = ss.insertSheet(tempName);

  temp.getRange(1,1).setValue('이름');
  temp.getRange(1,2).setValue('점수(%)');
  validData.forEach((row, i) => temp.getRange(i + 2, 1, 1, 2).setValues([row]));

  const chart = temp.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(temp.getRange(1, 1, validData.length + 1, 2))
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'AI 활용 점수 분포')
    .setOption('width', 500)
    .setOption('height', 400)
    .setOption('pieSliceText', 'label')
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .build();

  sheet.insertChart(chart);
  temp.hideSheet();

  SpreadsheetApp.getUi().alert('원형 차트가 생성되었습니다!');
}

/* 메뉴 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 차트 생성')
    .addItem('막대 차트 만들기', 'createScoreChart')
    .addItem('원형 차트 만들기', 'createPieChart')
    .addToUi();
}
