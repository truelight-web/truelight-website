/**
 * 참빛 좋은 날 — 참석 회신 저장소 (구글 시트 + Apps Script 웹 앱)
 *
 * 설정 (한 번만):
 *  1. 구글 드라이브에서 새 스프레드시트를 만든다.
 *  2. 메뉴 [확장 프로그램] → [Apps Script] 를 열고, 이 파일 내용을 통째로 붙여 넣고 저장한다.
 *  3. 왼쪽 톱니(프로젝트 설정) → 맨 아래 [스크립트 속성 추가]
 *     속성: ADMIN_PW   값: 관리자 비밀번호 (사이트 관리자 모드에서 쓸 비밀번호)
 *  4. 오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]
 *     실행 계정: 나 / 액세스 권한: 모든 사용자 → [배포] → 권한 승인
 *  5. 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 index.html 의 APPS_SCRIPT_URL 자리에 넣는다.
 *
 * 비밀번호는 스크립트 속성에만 있으므로 사이트 소스에는 드러나지 않는다.
 */

// 1부·2부로 나뉘면서 칸이 바뀌어 새 탭에 쌓는다 (예전 '참석 회신' 탭은 지워도 됨)
const SHEET_NAME = '참석 회신 (1부·2부)';
const HEADER = ['ID', '성함', '연락처', '1부', '1부 인원', '2부', '2부 인원', '개인정보 동의', '회신 일시'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('@'); // 연락처 앞의 0이 사라지지 않게
  }
  return sh;
}

// 축하의 말 — 사이트에 공개된다. 지우고 싶은 글은 시트에서 그 줄을 삭제하면 사이트에서도 사라진다
const CHEER_SHEET = '축하의 말';
const CHEER_HEADER = ['기기 ID', '축하의 말', '남긴 일시'];

function cheerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CHEER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CHEER_SHEET);
    sh.appendRow(CHEER_HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// =, +, -, @ 로 시작하는 값이 수식으로 실행되지 않게 막는다
function safe_(v) {
  const s = String(v == null ? '' : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }

  if (b.action === 'rsvp') {
    const id = String(b.id || '');
    const name = String(b.name || '').trim().slice(0, 30);
    const phone = String(b.phone || '').trim().slice(0, 20);
    // 인원 0 = 불참
    const part = function (v) { return Math.max(0, Math.min(20, Math.floor(Number(v)) || 0)); };
    const p1 = part(b.p1), p2 = part(b.p2);
    if (!/^[A-Za-z0-9-]{8,40}$/.test(id) || !name || phone.replace(/\D/g, '').length < 9 || b.agree !== true) {
      return out_({ ok: false, error: 'invalid' });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = sheet_();
      const row = [id, safe_(name), safe_(phone), p1 > 0 ? '참석' : '불참', p1, p2 > 0 ? '참석' : '불참', p2, '동의', new Date()];
      const last = sh.getLastRow();
      const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
      const at = ids.indexOf(id);
      // 같은 기기에서 다시 회신하면 새 줄을 만들지 않고 고친다
      sh.getRange(at >= 0 ? at + 2 : last + 1, 1, 1, row.length).setValues([row]);
      // 축하의 말도 함께 왔으면 '축하의 말' 탭에 넣는다 (같은 기기면 고친다)
      const msg = String(b.msg || '').trim().slice(0, 300);
      if (msg) {
        const ch = cheerSheet_();
        const cl = ch.getLastRow();
        const cids = cl > 1 ? ch.getRange(2, 1, cl - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
        const ci = cids.indexOf(id);
        const crow = [id, safe_(msg), new Date()];
        ch.getRange(ci >= 0 ? ci + 2 : cl + 1, 1, 1, crow.length).setValues([crow]);
      }
    } finally {
      lock.releaseLock();
    }
    return out_({ ok: true });
  }

  if (b.action === 'list') {
    const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
    if (!pw || String(b.pw || '') !== pw) return out_({ ok: false, error: 'wrong_pw' });
    const values = sheet_().getDataRange().getValues().slice(1);
    const rows = values.filter(function (r) { return r[0]; }).map(function (r) {
      return {
        name: String(r[1]),
        phone: String(r[2]),
        p1: Number(r[4]) || 0,
        p2: Number(r[6]) || 0,
        agree: r[7] === '동의',
        updatedAt: r[8] instanceof Date ? r[8].toISOString() : String(r[8])
      };
    });
    return out_({ ok: true, rows: rows });
  }

  if (b.action === 'cheers') {
    // 최근 글부터 100개까지 (기기 ID는 내보내지 않는다)
    const list = cheerSheet_().getDataRange().getValues().slice(1)
      .filter(function (r) { return r[1]; })
      .map(function (r) {
        return {
          msg: String(r[1]).replace(/^'/, ''),
          at: r[2] instanceof Date ? r[2].toISOString() : String(r[2])
        };
      })
      .reverse().slice(0, 100);
    return out_({ ok: true, list: list });
  }

  return out_({ ok: false, error: 'unknown_action' });
}
