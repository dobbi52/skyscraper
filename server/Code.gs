/* ============================================================
   스카이스크래퍼 — 기록 저장 서버 (Google Apps Script)
   · 로그인 확인은 생각 오락실 통합 계정 서버에 물어봐요.
   · 기록은 이 스프레드시트의 '기록' 시트에 닉네임마다 한 줄씩 저장해요.
   설치: 새 구글 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여 넣고 저장
         배포 → 새 배포 → 유형: 웹 앱 (실행: 나 / 액세스 권한: 모든 사용자) → 배포
         나온 웹 앱 주소(…/exec)를 게임 index.html 의 SAVE_API 에 넣어요.
   ============================================================ */
var HUB_API = 'https://script.google.com/macros/s/AKfycbzZlmUD9M4rfo4Y4eEYgNjZH0N8SppUA1Otuo9yMEOqcLfklw8XbS5ZmXj1D3KZY4EG/exec';
var SHEET = '기록';
var HEAD = ['닉네임', '저장 시각', '누적 별', '건물', '벽돌', '연습 판', '어려움 연습', '최근 7일 중 푼 날', '데이터'];
var BUILDING_AT = [2, 5, 9, 14, 20, 27, 35, 44, 54, 65];   // 게임의 CONFIG.BUILDING_AT 과 같게
var MAX_DATA = 45000;
var MIN_VER = 1;   // 이보다 옛날 게임 화면에서 온 요청은 받지 않아요(새로고침 안내)

function doGet() { return out({ ok: true, msg: '스카이스크래퍼 서버가 켜져 있어요.' }); }

function doPost(e) {
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if ((+b.ver || 0) < MIN_VER) return out({ ok: false, code: 'OLD', error: '게임이 바뀌었어요. 새로고침해 주세요.' });
    var who = verify(b.nick, b.pin);
    if (!who.ok) return out(who);
    if (b.action === 'load') return out(load(who.nick));
    if (b.action === 'save') return out(store(who.nick, b.data));
    return out({ ok: false, error: '알 수 없는 요청이에요.' });
  } catch (err) {
    return out({ ok: false, error: '서버 오류: ' + err, retry: true });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 통합 계정 확인 (6시간 동안 기억) ---------- */
function verify(nick, pin) {
  nick = String(nick || '').trim(); pin = String(pin || '').trim();
  if (!nick || !/^\d{4}$/.test(pin)) return { ok: false, code: 'PIN', error: '로그인 정보가 없어요.' };
  var cache = CacheService.getScriptCache();
  var key = 'v:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, nick + '|' + pin));
  var hit = cache.get(key);
  if (hit) return { ok: true, nick: hit };
  var r = UrlFetchApp.fetch(HUB_API, {
    method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true, followRedirects: true,
    payload: JSON.stringify({ action: 'me', nick: nick, pin: pin })
  });
  var j = {};
  try { j = JSON.parse(r.getContentText()); } catch (e) { return { ok: false, error: '생각 오락실 서버에 연결하지 못했어요.', retry: true }; }
  if (!j.ok) return { ok: false, code: j.code || 'PIN', error: j.error || '생각 오락실에서 다시 로그인해 주세요.' };
  cache.put(key, j.nick || nick, 21600);
  return { ok: true, nick: j.nick || nick };
}

/* ---------- 시트 ---------- */
function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function findRow(sh, nick) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var names = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (var i = 0; i < names.length; i++) if (names[i][0] === nick) return i + 2;
  return 0;
}

function load(nick) {
  var sh = sheet(), row = findRow(sh, nick);
  if (!row) return { ok: true, nick: nick, data: null };
  var raw = sh.getRange(row, HEAD.length).getValue();
  var data = null;
  try { data = JSON.parse(raw); } catch (e) {}
  return { ok: true, nick: nick, data: data };
}

function store(nick, data) {
  if (!data || typeof data !== 'object') return { ok: false, error: '저장할 기록이 없어요.' };
  var raw = JSON.stringify(data);
  if (raw.length > MAX_DATA) return { ok: false, error: '기록이 너무 커요.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet(), row = findRow(sh, nick);
    if (row) {   // 선생님이 고친 기록(rev가 더 큼)보다 옛날 기록이면 받지 않고 서버 기록을 돌려줘요
      var old = null;
      try { old = JSON.parse(sh.getRange(row, HEAD.length).getValue()); } catch (e) {}
      if (old && (old.rev || 0) > (data.rev || 0)) return { ok: false, code: 'REV', data: old };
    }
    var stars = +data.stars || 0;
    var built = BUILDING_AT.filter(function (t) { return stars >= t; }).length;
    var prac = data.prac || {};
    var days = 0, daily = data.daily || {};
    for (var d in daily) {
      var lv = daily[d] || {};
      for (var k in lv) if (lv[k] && lv[k].status === 'win') { days++; break; }
    }
    var vals = [[nick, new Date(), stars, built, +data.bricks || 0, +prac.total || 0, +prac.hard || 0, days, raw]];
    if (!row) row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, HEAD.length).setValues(vals);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
