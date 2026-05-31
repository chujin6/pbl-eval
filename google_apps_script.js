// ═══════════════════════════════════════════════════════════════
//  12주차 PBL 상호평가 — Google Apps Script (서버 코드)
//  이 파일 전체를 복사해서 Apps Script 편집기에 붙여넣기 하세요
// ═══════════════════════════════════════════════════════════════

// ── ① 시트 이름 설정 (변경 불필요) ─────────────────────────────
const SHEET_NAME = "PBL평가결과";

// ── ② 웹 앱 진입점: GET 요청 (대시보드용 데이터 조회) ──────────
function doGet(e) {
  const action = e.parameter.action || "";

  if (action === "getData") {
    return sendJSON(getAllData());
  }

  // 기본: 상태 확인용
  return sendJSON({ status: "ok", message: "PBL 평가 서버 정상 동작 중" });
}

// ── ③ 웹 앱 진입점: POST 요청 (평가 데이터 저장) ───────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || "";

    if (action === "submit") {
      saveEval(payload.data);
      return sendJSON({ status: "ok" });
    }

    return sendJSON({ status: "error", message: "알 수 없는 액션" });
  } catch (err) {
    return sendJSON({ status: "error", message: err.toString() });
  }
}

// ── ④ 평가 데이터 저장 ──────────────────────────────────────────
function saveEval(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // 시트가 없으면 자동 생성 + 헤더 추가
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "제출시간",
      "평가자 조",
      "평가 대상 조",
      "점수1_문제정의",
      "점수2_실현가능성",
      "점수3_전달력협업",
      "평균점수",
      "의견1_문제정의",
      "의견2_실현가능성",
      "의견3_전달력협업"
    ]);
    // 헤더 스타일
    const header = sheet.getRange(1, 1, 1, 10);
    header.setBackground("#0f2347");
    header.setFontColor("#ffffff");
    header.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const avg = ((data.scores[0] + data.scores[1] + data.scores[2]) / 3).toFixed(2);

  sheet.appendRow([
    new Date(),
    data.myGroup,
    data.targetGroup,
    data.scores[0],
    data.scores[1],
    data.scores[2],
    parseFloat(avg),
    data.comments[0] || "",
    data.comments[1] || "",
    data.comments[2] || ""
  ]);
}

// ── ⑤ 전체 데이터 조회 ─────────────────────────────────────────
function getAllData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { rows: [] };

  const raw  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const rows = raw.map(r => ({
    timestamp:   Utilities.formatDate(new Date(r[0]), "Asia/Seoul", "yyyy-MM-dd HH:mm"),
    myGroup:     r[1],
    targetGroup: r[2],
    scores:      [r[3], r[4], r[5]],
    avg:         r[6],
    comments:    [r[7], r[8], r[9]]
  }));

  return { rows };
}

// ── ⑥ JSON 응답 헬퍼 ────────────────────────────────────────────
function sendJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
