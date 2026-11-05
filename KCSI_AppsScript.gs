/**
 * K-CSI 커뮤니티 지속가능성 통합 진단 — 응답 축적 + 이메일 알림 (담당자 전용, GitHub에 올리지 마세요)
 * 진단 응답과 오픈네트워크 신청이 들어올 때마다
 *   1) 이 스프레드시트의 시트 탭('진단 응답', '오픈네트워크 신청')에 한 줄씩 쌓고
 *   2) 아래 TO 주소로 접수 내용 + 누적 엑셀(.xlsx) 첨부 메일을 보냅니다.
 */
const TO = 'bkim1229@ksa.or.kr';
const TOKEN = 'ksa-kcsi-d143f577a5';          // index.html CONFIG.sheetToken 과 같아야 합니다
const ATTACH_XLSX = true;          // 매 메일에 누적 엑셀 첨부 (false 로 바꾸면 본문만)
const ALLOWED = ['진단 응답', '오픈네트워크 신청'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data._token !== TOKEN) return ContentService.createTextOutput('denied');
    const name = ALLOWED.includes(data._sheet) ? data._sheet : '진단 응답';
    delete data._token; delete data._sheet; delete data._sent;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(Object.keys(data)); sh.setFrozenRows(1); }
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const extra = Object.keys(data).filter(k => !header.includes(k));
    if (extra.length) { sh.getRange(1, header.length + 1, 1, extra.length).setValues([extra]); header.push(...extra); }
    sh.appendRow(header.map(k => data[k] !== undefined ? String(data[k]) : ''));
    SpreadsheetApp.flush();
    notify(ss, name, data);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function notify(ss, name, data) {
  const who = data['성명'] || data['이름'] || '';
  const org = data['소속'] || data['직장'] || data['조직명'] || '';
  const subject = '[K-CSI 통합진단] ' + name + ' 접수 · ' + who + (org ? ' (' + org + ')' : '');
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const rows = Object.keys(data).filter(k => k !== '결과코드').map(k =>
    '<tr><th align="left" style="padding:5px 10px;background:#eef3f0;white-space:nowrap;font-weight:600">' + esc(k) +
    '</th><td style="padding:5px 10px;border-bottom:1px solid #eee">' + esc(data[k]) + '</td></tr>').join('');
  const count = ss.getSheetByName(name).getLastRow() - 1;
  const html = '<p><b>' + esc(name) + '</b> 1건이 접수되었습니다. (누적 ' + count + '건)</p>' +
    '<table style="border-collapse:collapse;font-size:13px;font-family:sans-serif">' + rows + '</table>' +
    '<p style="margin-top:14px"><a href="' + ss.getUrl() + '">누적 스프레드시트 열기</a></p>';
  const opts = { htmlBody: html, name: 'K-CSI 통합진단' };
  if (ATTACH_XLSX) {
    DriveApp.getFileById(ss.getId()); // 내보내기 권한 확보용
    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
      .getBlob().setName('KCSI_통합진단_누적_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm') + '.xlsx');
    opts.attachments = [blob];
  }
  MailApp.sendEmail(TO, subject, '접수 내용은 HTML 메일로 확인해 주세요.', opts);
}

/** 배포 전에 한 번 실행해 권한을 승인하고, 테스트 메일이 오는지 확인하세요. */
function testNotify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('테스트') || ss.insertSheet('테스트');
  if (sh.getLastRow() === 0) sh.appendRow(['성명', '소속', '관심키워드']);
  sh.appendRow(['테스트', '한국표준협회', '연결 확인']);
  notify(ss, '테스트', { '성명': '테스트', '소속': '한국표준협회', '관심키워드': '연결 확인' });
}
