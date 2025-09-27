/** ChartGenerator.gs - 날짜별 시트 대응 버전 **/

/* 헤더 찾기(오타/영문 허용) */
function getColByHeader_(sheet, headerName) {
  const HEADER_ALTS = {
    '이름': ['이름','name'],
    '점수': ['점수','score'],
    '타임스탬프': ['타임스탬프','타임스태프','timestamp','시간','작성시각'],
    '날짜': ['날짜','date']
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

/* 시트 선택 다이얼로그 */
function selectSheetForChart_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 사용 가능한 시트 목록 생성
  const allSheets = ss.getSheets();
  const dateSheets = allSheets.filter(sheet => {
    const name = sheet.getName();
    return /^\d{4}-\d{2}-\d{2}$/.test(name); // 날짜 형식 시트
  });
  
  // 요약 시트 확인
  const summarySheet = ss.getSheetByName('전체_요약');
  const responseSheet = ss.getSheetByName('Responses');
  
  let options = [];
  if (summarySheet) options.push('전체_요약 (모든 날짜 통합)');
  if (responseSheet) options.push('Responses (기존 시트)');
  
  // 최근 10개 날짜 시트만 표시
  const recentDateSheets = dateSheets
    .sort((a, b) => b.getName().localeCompare(a.getName()))
    .slice(0, 10);
  
  recentDateSheets.forEach(sheet => {
    options.push(sheet.getName());
  });
  
  if (options.length === 0) {
    ui.alert('차트를 생성할 수 있는 데이터 시트가 없습니다.');
    return null;
  }
  
  // 시트 선택 다이얼로그
  const response = ui.prompt(
    '차트 생성할 시트 선택', 
    `다음 중 선택하세요:\n${options.map((opt, i) => `${i+1}. ${opt}`).join('\n')}\n\n번호를 입력하세요:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const choice = parseInt(response.getResponseText().trim()) - 1;
    if (choice >= 0 && choice < options.length) {
      const selectedName = options[choice].replace(' (모든 날짜 통합)', '').replace(' (기존 시트)', '');
      return ss.getSheetByName(selectedName);
    }
  }
  
  return null;
}

/* 데이터 수집 함수 */
function collectChartData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const nameCol = getColByHeader_(sheet, '이름');
  const scoreCol = getColByHeader_(sheet, '점수');
  
  if (!nameCol || !scoreCol) return [];

  const names = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues().flat();
  const scores = sheet.getRange(2, scoreCol, lastRow - 1, 1).getValues().flat();

  const validData = [];
  for (let i = 0; i < names.length; i++) {
    const p = parseScoreToPercent_(scores[i]);
    if (names[i] && !isNaN(p)) {
      validData.push([names[i], p]);
    }
  }
  
  return validData;
}

/* 막대 차트 생성 */
function createScoreChart() {
  const sheet = selectSheetForChart_();
  if (!sheet) return;

  const validData = collectChartData_(sheet);
  if (validData.length === 0) {
    SpreadsheetApp.getUi().alert('유효한 점수 데이터가 없습니다.');
    return;
  }

  // 기존 차트 제거
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = '차트_데이터_임시';
  let temp = ss.getSheetByName(tempName);
  if (temp) ss.deleteSheet(temp);
  temp = ss.insertSheet(tempName);

  // 데이터만 입력 (헤더 없이)
  validData.forEach((row, i) => {
    temp.getRange(i + 1, 1).setValue(row[0]); // 이름
    temp.getRange(i + 1, 2).setValue(row[1]); // 점수
  });

  const chartTitle = `AI 활용 점수 현황 - ${sheet.getName()}`;
  
  const chart = temp.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(temp.getRange(1, 1, validData.length, 2)) // 헤더 제외하고 데이터만
    .setPosition(5, 5, 0, 0)
    .setOption('title', chartTitle)
    .setOption('hAxis', { title: '이름', titleTextStyle: { fontSize: 12, bold: true } })
    .setOption('vAxis', { title: '점수(%)', minValue: 0, maxValue: 100, titleTextStyle: { fontSize: 12, bold: true } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 700)
    .setOption('height', 450)
    .setOption('colors', ['#4285F4'])
    .build();

  sheet.insertChart(chart);
  temp.hideSheet();

  SpreadsheetApp.getUi().alert(`${sheet.getName()} 시트에 막대 차트가 생성되었습니다!`);
}

/* 원형 차트 생성 */
function createPieChart() {
  const sheet = selectSheetForChart_();
  if (!sheet) return;

  const validData = collectChartData_(sheet);
  if (validData.length === 0) {
    SpreadsheetApp.getUi().alert('유효한 점수 데이터가 없습니다.');
    return;
  }

  // 기존 차트 제거
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = '차트_데이터_임시';
  let temp = ss.getSheetByName(tempName);
  if (temp) ss.deleteSheet(temp);
  temp = ss.insertSheet(tempName);

  // 헤더 설정
  temp.getRange(1,1).setValue('이름');
  temp.getRange(1,2).setValue('점수(%)');
  
  // 데이터 입력 (2행부터)
  validData.forEach((row, i) => {
    temp.getRange(i + 2, 1).setValue(row[0]); // 이름
    temp.getRange(i + 2, 2).setValue(row[1]); // 점수
  });

  const chartTitle = `AI 활용 점수 분포 - ${sheet.getName()}`;

  const chart = temp.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(temp.getRange(1, 1, validData.length + 1, 2)) // 헤더 포함하여 범위 설정
    .setPosition(5, 5, 0, 0)
    .setOption('title', chartTitle)
    .setOption('width', 600)
    .setOption('height', 450)
    .setOption('pieSliceText', 'label')
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .build();

  sheet.insertChart(chart);
  temp.hideSheet();

  SpreadsheetApp.getUi().alert(`${sheet.getName()} 시트에 원형 차트가 생성되었습니다!`);
}

/* 일별 통계 차트 생성 (요약 시트용) */
function createDailyStatsChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('전체_요약');
  
  if (!summarySheet) {
    SpreadsheetApp.getUi().alert('전체_요약 시트가 없습니다. 먼저 요약 시트를 생성해주세요.');
    return;
  }

  const lastRow = summarySheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('요약 시트에 데이터가 없습니다.');
    return;
  }

  // 날짜별 참여자 수 집계
  const dateCol = getColByHeader_(summarySheet, '날짜');
  if (!dateCol) {
    SpreadsheetApp.getUi().alert('날짜 열을 찾을 수 없습니다.');
    return;
  }

  const dates = summarySheet.getRange(2, dateCol, lastRow - 1, 1).getValues().flat();
  const dateCount = {};
  
  dates.forEach(date => {
    if (date) {
      const dateStr = String(date);
      dateCount[dateStr] = (dateCount[dateStr] || 0) + 1;
    }
  });

  const chartData = Object.entries(dateCount).sort();
  
  if (chartData.length === 0) {
    SpreadsheetApp.getUi().alert('차트를 생성할 데이터가 없습니다.');
    return;
  }

  // 기존 차트 제거
  summarySheet.getCharts().forEach(c => summarySheet.removeChart(c));

  const tempName = '일별통계_임시';
  let temp = ss.getSheetByName(tempName);
  if (temp) ss.deleteSheet(temp);
  temp = ss.insertSheet(tempName);

  temp.getRange(1,1).setValue('날짜');
  temp.getRange(1,2).setValue('참여자 수');
  chartData.forEach((row, i) => temp.getRange(i + 2, 1, 1, 2).setValues([row]));

  const chart = temp.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(temp.getRange(1, 1, chartData.length + 1, 2))
    .setPosition(5, 5, 0, 0)
    .setOption('title', '일별 테스트 참여자 수')
    .setOption('hAxis', { title: '날짜', titleTextStyle: { fontSize: 12, bold: true } })
    .setOption('vAxis', { title: '참여자 수', titleTextStyle: { fontSize: 12, bold: true } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 800)
    .setOption('height', 400)
    .setOption('colors', ['#34495e'])
    .build();

  summarySheet.insertChart(chart);
  temp.hideSheet();

  SpreadsheetApp.getUi().alert('일별 통계 차트가 생성되었습니다!');
}

/* 메뉴 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 차트 & 데이터 관리')
    .addItem('📊 막대 차트 만들기', 'createScoreChart')
    .addItem('🥧 원형 차트 만들기', 'createPieChart')
    .addItem('📈 일별 통계 차트', 'createDailyStatsChart')
    .addSeparator()
    .addItem('📋 전체 요약 시트 생성', 'createSummarySheet')
    .addItem('🗑️ 날짜별 데이터 삭제', 'deleteDataByDate')
    .addToUi();
}
