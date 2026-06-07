/**
 * 조르의 과학 실험실 — Google Apps Script 수집 코드
 *
 * 사용법:
 * 1. Google Sheets에서 확장 프로그램 > Apps Script 열기
 * 2. 이 코드 전체를 붙여넣고 저장
 * 3. 배포 > 새 배포 > 웹 앱 선택
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 후 생성된 URL을 index.html의 APPS_SCRIPT_URL에 붙여넣기
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheet = ss.getSheetByName('학생응답');
    if (!sheet) {
      sheet = ss.insertSheet('학생응답');
      sheet.appendRow([
        '제출시각', '학년', '반', '이름', '최종신뢰도',
        '미션1(용질종류)', '미션2(온도오류)',
        '미션3-Q1', '미션3-Q2', '미션3-Q3',
        '미션4-Q1', '미션4-Q2', '미션4-Q3',
        '설계-물양(mL)', '설계-비커수', '설계-설탕차이',
        '설계Q8(예측서술)', '설계Q8점수'
      ]);
      const header = sheet.getRange(1, 1, 1, 18);
      header.setBackground('#1a237e');
      header.setFontColor('#ffffff');
      header.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.timestamp,
      data.grade,
      data.cls,
      data.name,
      data.trust,
      data.m1,
      data.m2,
      data.m3q1,
      data.m3q2,
      data.m3q3,
      data.m4q1,
      data.m4q2,
      data.m4q3,
      data.dWater,
      data.dBeakers,
      data.dSugar,
      data.dq8,
      data.dq8s
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 테스트용: Apps Script 편집기에서 직접 실행해서 시트 구조 확인
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('스프레드시트 이름: ' + ss.getName());
  Logger.log('시트 목록: ' + ss.getSheets().map(s => s.getName()).join(', '));
}
