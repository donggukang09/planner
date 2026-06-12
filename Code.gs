/************************************************************
 *  업무 플래너 - Google Apps Script 백엔드
 *  - 시트 읽기/쓰기 API (PWA와 연동)
 *  - 1분마다 알람 시각 확인 후 이메일 발송
 *  DEVELOPED BY DONGGU KANG
 ************************************************************/

const SHEET_ID   = '1XBxfL94QkNBg1W9H14Vejo5D0fgIljH5kQHy2YeDlc8';
const SHEET_NAME = '플래너';                 // 데이터가 저장될 탭 이름
const ALARM_TO   = 'donggu.kang@hiqasc.kr';  // 알람 받을 이메일
const TZ         = 'Asia/Seoul';
const HEADERS = ['id','date','time','text','prio','cat','done','alarm','notified','updated'];

/* ---------- 시트 준비 ---------- */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2,1,last-1,HEADERS.length).getValues();
  return rows.filter(r => r[0] !== '').map(r => ({
    id:String(r[0]), date:String(r[1]), time:String(r[2]), text:String(r[3]),
    prio:String(r[4]||'mid'), cat:String(r[5]||'업무'),
    done: r[6]===true || String(r[6]).toUpperCase()==='TRUE',
    alarm: r[7]===true || String(r[7]).toUpperCase()==='TRUE',
    notified: r[8]===true || String(r[8]).toUpperCase()==='TRUE'
  }));
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2,1,last-1,1).getValues();
  for (let i=0;i<ids.length;i++){ if (String(ids[i][0])===String(id)) return i+2; }
  return -1;
}

function upsert_(t) {
  const sh = getSheet_();
  const row = [t.id, t.date, t.time||'', t.text, t.prio||'mid', t.cat||'업무',
               !!t.done, !!t.alarm, !!t.notified, new Date()];
  const r = findRow_(sh, t.id);
  if (r === -1) sh.appendRow(row);
  else sh.getRange(r,1,1,HEADERS.length).setValues([row]);
}

function remove_(id) {
  const sh = getSheet_();
  const r = findRow_(sh, id);
  if (r !== -1) sh.deleteRow(r);
}

/* ---------- 웹앱 엔드포인트 ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'list') return json_({ ok:true, items: readAll_() });
  return json_({ ok:false, error:'unknown action' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'upsert') { upsert_(body.item); return json_({ ok:true }); }
    if (action === 'delete') { remove_(body.id);     return json_({ ok:true }); }
    if (action === 'bulk') {  // 전체 동기화(여러 건)
      (body.items||[]).forEach(upsert_);
      return json_({ ok:true, count:(body.items||[]).length });
    }
    return json_({ ok:false, error:'unknown action' });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 알람 이메일 (1분 트리거) ---------- */
function checkAlarms() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const now = new Date();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const nowMin = parseInt(Utilities.formatDate(now, TZ, 'H'),10)*60
               + parseInt(Utilities.formatDate(now, TZ, 'm'),10);

  const data = sh.getRange(2,1,last-1,HEADERS.length).getValues();
  let changed = false;
  for (let i=0;i<data.length;i++){
    const r = data[i];
    if (!r[0]) continue;
    const date=String(r[1]), time=String(r[2]);
    const done = r[6]===true||String(r[6]).toUpperCase()==='TRUE';
    const alarm= r[7]===true||String(r[7]).toUpperCase()==='TRUE';
    const notified = r[8]===true||String(r[8]).toUpperCase()==='TRUE';
    if (!alarm || done || notified || !time) continue;
    if (date !== today) continue;
    const parts = time.split(':');
    const tMin = parseInt(parts[0],10)*60 + parseInt(parts[1],10);
    // 알람 시각이 되었고(지났고) 10분 이내면 발송
    if (nowMin >= tMin && nowMin - tMin <= 10) {
      sendAlarmMail_(String(r[3]), time, String(r[5]||''), String(r[4]||'mid'));
      sh.getRange(i+2, 9).setValue(true); // notified = TRUE
      changed = true;
    }
  }
  if (changed) SpreadsheetApp.flush();
}

function sendAlarmMail_(text, time, cat, prio) {
  const pLabel = prio==='high'?'높음':prio==='low'?'낮음':'중간';
  const subject = '⏰ [업무 플래너] ' + time + ' · ' + text;
  const html =
    '<div style="font-family:sans-serif;max-width:480px;margin:0 auto;background:#221d16;color:#ece4d4;border-radius:14px;padding:28px 26px;border:1px solid #e0b252">'
    + '<div style="font-size:11px;letter-spacing:.2em;color:#caa052;text-transform:uppercase">Work Planner Alarm</div>'
    + '<div style="font-size:26px;color:#e0b252;margin:10px 0 4px">⏰ ' + time + '</div>'
    + '<div style="font-size:19px;margin:6px 0 16px">' + escapeHtml_(text) + '</div>'
    + '<div style="font-size:12px;color:#a89b82">분류: ' + escapeHtml_(cat||'-') + ' &nbsp;·&nbsp; 우선순위: ' + pLabel + '</div>'
    + '<div style="margin-top:18px;font-size:10px;color:#6f6553;letter-spacing:.15em">DEVELOPED BY DONGGU KANG</div>'
    + '</div>';
  MailApp.sendEmail({ to: ALARM_TO, subject: subject, htmlBody: html });
}

function escapeHtml_(s){return String(s).replace(/[&<>"']/g,function(m){return({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[m];});}

/* ---------- 최초 1회 실행: 1분 트리거 설치 ---------- */
function setupTrigger() {
  // 기존 트리거 정리
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction()==='checkAlarms') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkAlarms').timeBased().everyMinutes(1).create();
  getSheet_(); // 시트/헤더 초기화
  Logger.log('트리거 설치 완료 — 1분마다 알람 확인');
}
