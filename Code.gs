/**
 * 따끈따끈 베이커리 — Google Apps Script 서버 코드 (Code.gs)
 *
 * ────────────────────────────────────────────────────────────
 *  사용법
 * ────────────────────────────────────────────────────────────
 *  1) 학생 명단 시트 준비
 *     · 시트 이름: "명단"  (또는 첫 번째 시트가 자동 사용됨)
 *     · A열  : 번호       (1, 2, 3 …)
 *     · B열  : 이름       (학생 이름 한 줄에 한 명)
 *     · 1행  : 헤더(예: "번호", "이름") — 비워둬도 됩니다
 *     · 2행부터 실제 데이터
 *
 *     예)
 *        |  A   |   B    |
 *        | 번호 |  이름   |   ← 1행 (헤더)
 *        |  1  | 김민준  |   ← 2행 부터 학생
 *        |  2  | 이서연  |
 *        |  3  | 박지호  |
 *
 *  2) 스프레드시트에서  확장 프로그램 → Apps Script  열기
 *
 *  3) 기본으로 들어 있는 Code.gs 내용을 비우고 이 파일 전체를 붙여넣기
 *
 *  4) 새 파일 만들기 ▸ HTML  →  파일 이름을 정확히  index  로 지정
 *     →  같은 폴더의 index.html 전체 내용을 붙여넣기
 *
 *  5) 우측 상단  배포  →  새 배포  →  유형: 웹 앱
 *        · 다음으로 실행 :  나(본인 계정)
 *        · 액세스 권한    :  학생이 들어올 수 있는 범위 선택
 *                          (학교 도메인 / 링크가 있는 모든 사용자 등)
 *     →  배포 URL을 학생들에게 공유
 *
 *  6) 학생이 URL을 열면 명단 드롭다운이 자동으로 나타나고,
 *     "1. 김민준" 형식으로 자기 번호·이름을 골라 학습을 시작할 수 있습니다.
 * ────────────────────────────────────────────────────────────
 */


/**
 * 웹 앱 진입점.
 * 학생이 배포 URL에 접속하면 index.html 을 화면에 띄웁니다.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('따끈따끈 베이커리')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


/**
 * 학생 명단 가져오기.
 * "명단" 시트(없으면 첫 번째 시트)의 2행부터 A열(번호) · B열(이름)을 함께 읽어
 * { number, name } 객체 배열로 반환합니다.
 *
 * 클라이언트(index.html)에서  google.script.run.getStudentList()  로 호출.
 */
function getStudentList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('명단') || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // A2:B(lastRow) — A열(번호), B열(이름)
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    var students = [];
    for (var i = 0; i < values.length; i++) {
      var num = values[i][0];
      var name = values[i][1];
      if (name == null) continue;
      var nameStr = String(name).trim();
      if (!nameStr) continue;

      // 번호는 1, 2, 3 … 정수형이면 소수점 제거
      var numStr = '';
      if (num != null && num !== '') {
        numStr = (typeof num === 'number') ? String(parseInt(num, 10)) : String(num).trim();
      }

      students.push({ number: numStr, name: nameStr });
    }
    return students;
  } catch (e) {
    Logger.log('getStudentList error: ' + e);
    return [];
  }
}


/**
 * 비 줄이기 게임 점수 누적 저장
 *  - 클라이언트에서  google.script.run.saveScore({ name, mode, score, time })  으로 호출.
 *  - "기록" 시트에 한 행이 추가됩니다.
 *  - 헤더: 날짜 | 이름 | 도전 모드 | 맞춤 | 소요 시간(초)
 */
function saveScore(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('기록') || ss.insertSheet('기록');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['날짜', '이름', '도전 모드', '맞춤', '소요 시간(초)']);
    }
    sheet.appendRow([
      new Date(),
      (payload && payload.name)  || '',
      (payload && payload.mode)  || '',
      (payload && payload.score) || 0,
      (payload && payload.time)  || 0
    ]);
    return { ok: true };
  } catch (e) {
    Logger.log('saveScore error: ' + e);
    return { ok: false, error: String(e) };
  }
}
