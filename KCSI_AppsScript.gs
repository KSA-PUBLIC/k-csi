/**
 * 커뮤니티 지속가능성 통합 진단 — 응답 축적 + 이메일 발송 (담당자 전용, GitHub에 올리지 마세요)
 *  · 진단 응답 / 오픈네트워크 신청 / 심화 진단 상담 → 시트 탭에 한 줄씩 저장 + 담당자(TO)에게 알림 메일(누적 엑셀 첨부)
 *  · 진단 응답 → 응답자가 입력한 이메일로 결과지 메일 자동 발송(요약 + 결과지 링크 + PDF)
 *  · 결과 메일 → 결과지 화면의 ‘결과지 메일 다시 받기’ 요청 시 응답자에게 재발송
 */
const TO = 'bkim1229@ksa.or.kr';
const TOKEN = 'ksa-kcsi-d143f577a5';           // index.html CONFIG.sheetToken 과 같아야 합니다
const PAGE_URL = 'https://ksa-public.github.io/k-csi/';
const ATTACH_XLSX = true;           // 담당자 메일에 누적 엑셀 첨부
const SEND_TO_RESPONDENT = true;    // 응답자에게 결과지 메일 발송
const RESULT_PDF = true;            // 응답자 메일에 결과 요약 PDF 첨부
const SHEETS = ['진단 응답', '오픈네트워크 신청', '심화 진단 상담'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data._token !== TOKEN) return ContentService.createTextOutput('denied');
    const kind = data._sheet; delete data._token; delete data._sheet; delete data._sent;
    if (kind === '결과 메일') { sendResult(data); return ContentService.createTextOutput('ok'); }
    const name = SHEETS.includes(kind) ? kind : '진단 응답';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(Object.keys(data)); sh.setFrozenRows(1); }
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const extra = Object.keys(data).filter(k => !header.includes(k));
    if (extra.length) { sh.getRange(1, header.length + 1, 1, extra.length).setValues([extra]); header.push(...extra); }
    sh.appendRow(header.map(k => data[k] !== undefined ? String(data[k]) : ''));
    SpreadsheetApp.flush();
    notify(ss, name, data);
    if (name === '진단 응답' && SEND_TO_RESPONDENT) sendResult(data);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

const esc = v => String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;');

function notify(ss, name, data) {
  const who = data['성명'] || data['이름'] || '';
  const org = data['소속'] || data['직장'] || data['조직명'] || '';
  const subject = '[K-CSI 통합진단] ' + name + ' 접수 · ' + who + (org ? ' (' + org + ')' : '');
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

function sendResult(d) {
  const email = String(d['이메일'] || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
  const link = d['결과코드'] ? PAGE_URL + '?r=' + d['결과코드'] + '#result' : PAGE_URL;
  const row = (k, v) => v === '' || v == null ? '' : '<tr><th align="left" style="padding:8px 12px;background:#eef3f0;white-space:nowrap;font-weight:600;color:#10202B">' + esc(k) + '</th><td style="padding:8px 12px;border-bottom:1px solid #e6ebe7">' + esc(v) + '</td></tr>';
  const imp = ['성과', '수혜자', '규모', '기여', '리스크'].map(k => k + ' ' + (d['IMP ' + k] === '' ? '—' : d['IMP ' + k])).join(' · ');
  const html =
    '<div style="font-family:sans-serif;max-width:640px;color:#10202B">' +
    '<div style="background:#0B4F45;color:#fff;padding:22px 24px;border-radius:12px 12px 0 0">' +
    '<div style="font-size:12px;letter-spacing:.08em;opacity:.85">한국표준협회 공공혁신센터 · 커뮤니티 지속가능성 통합 진단</div>' +
    '<div style="font-size:22px;font-weight:700;margin-top:6px">' + esc(d['조직명'] || '') + ' 진단 결과지</div>' +
    '<div style="font-size:13px;opacity:.85;margin-top:4px">' + esc(d['응답일시'] || '') + ' · 주요 대상지 ' + esc(d['주요대상지'] || '') + '</div></div>' +
    '<div style="border:1px solid #e2e6e0;border-top:0;border-radius:0 0 12px 12px;padding:20px 24px">' +
    '<p style="margin:0 0 14px">' + esc(d['이름'] || '') + '님, 진단에 참여해 주셔서 감사합니다. 결과 요약을 보내 드립니다.</p>' +
    '<table style="border-collapse:collapse;font-size:14px;width:100%">' +
      row('핵심 셀', d['핵심 셀']) + row('변화·가치 점수', d['변화·가치 점수'] === '' ? '' : d['변화·가치 점수'] + ' / 100') +
      row('유형', d['조직 유형 판정']) + row('임팩트 분류', d['임팩트 분류']) + row('지역 적합도', d['지역 적합도']) +
      row('주요 대상지 K-CSI', d['대상지 K-CSI Score'] === '' ? '' : d['대상지 K-CSI Score'] + ' (' + d['대상지 등급'] + '등급, 전국 ' + d['대상지 순위'] + '위)') +
      row('주요 타겟층', d['주요 타겟층']) + row('IMP 5대 요소', imp) + row('관련 SDGs', d['관련 SDGs']) +
      row('연간 순편익(보수 추정)', d['연간 순편익(원)'] === '' ? '' : Number(d['연간 순편익(원)']).toLocaleString('ko-KR') + '원') +
    '</table>' +
    '<p style="margin:20px 0 8px"><a href="' + link + '" style="display:inline-block;background:#0B4F45;color:#fff;text-decoration:none;padding:12px 20px;border-radius:8px;font-weight:700">전체 결과지 보기</a></p>' +
    '<p style="font-size:12px;color:#6B7780;margin:12px 0 0">버튼을 누르면 72셀 핵심 셀, 육각형 모델, 유사 사례, 변화이론, SDGs, 정성 평가가 담긴 전체 결과지가 열립니다. 심화 진단 문의: ' + TO + '</p>' +
    '</div></div>';
  const opts = { htmlBody: html, name: '한국표준협회 공공혁신센터', replyTo: TO };
  if (RESULT_PDF) {
    try { opts.attachments = [Utilities.newBlob(html, MimeType.HTML, 'result.html').getAs(MimeType.PDF).setName('통합진단_결과지_' + (d['조직명'] || '') + '.pdf')]; } catch (err) {}
  }
  MailApp.sendEmail(email, '[한국표준협회] ' + (d['조직명'] || '') + ' 커뮤니티 지속가능성 통합 진단 결과지', '결과지는 HTML 메일로 확인해 주세요: ' + link, opts);
}

/** 배포 전에 실행해 권한을 승인하고 메일 2통(담당자 알림, 응답자 결과지 견본)이 TO로 오는지 확인하세요. */
function testNotify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('테스트') || ss.insertSheet('테스트');
  if (sh.getLastRow() === 0) sh.appendRow(['성명', '소속', '관심키워드']);
  sh.appendRow(['테스트', '한국표준협회', '연결 확인']);
  notify(ss, '테스트', { '성명': '테스트', '소속': '한국표준협회', '관심키워드': '연결 확인' });
  sendResult({ '이메일': TO, '이름': '테스트', '조직명': '견본 조직', '응답일시': Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd HH:mm'), '주요대상지': '충청남도 아산시',
    '핵심 셀': '문화·정체성 × 매력', '변화·가치 점수': 65, '조직 유형 판정': '도심활력형', '임팩트 분류': 'B · 이해관계자 편익', '지역 적합도': 26.8,
    '대상지 K-CSI Score': 25.2, '대상지 등급': 'C', '대상지 순위': 100, '주요 타겟층': '청년; 지역주민; 소상공인',
    'IMP 성과': 68.8, 'IMP 수혜자': 62.5, 'IMP 규모': 66.7, 'IMP 기여': 62.5, 'IMP 리스크': 66.7, '관련 SDGs': 'SDG11; SDG8; SDG4', '연간 순편익(원)': 197600000, '결과코드': '' });
}
