/**
 * 일일 안전보건 현장관리 (S.O.S) — 구글시트 저장용 스크립트
 * 김해신문센트럴아이파크
 *
 * 붙여 넣는 곳: 작업 현황 구글시트 → 확장 프로그램 → Apps Script
 * 처음 한 번: 아래 ADMIN_PASSWORD 를 원하는 관리 비밀번호로 바꾼 뒤 배포하세요.
 */

const ADMIN_PASSWORD = '1234';      // ← 관리 탭 비밀번호 (처음 한 번만 사용, 이후 앱에서 변경 가능)
const DATA_SHEET = '앱원본';         // 앱이 작업을 쌓아 두는 탭 (자동 생성)
const MAP_FOLDER = 'SOS 배치도';     // 배치도 이미지를 보관할 드라이브 폴더 (자동 생성)

const COLS = ['id','date','company','title','trade','status','dongs','floors','loc','workers','start','end',
  'point','equip','highRisk','ptw','ptwTypes','riskUp','hazmat','watch','early','opc','edu','stands','cutters',
  'accident','leader','supCo','supHdc','safeHdc','note','zones','createdAt','updatedAt'];
const JSON_COLS = ['dongs','floors','ptwTypes','zones'];
const BOOL_COLS = ['highRisk','riskUp','opc'];
const NUM_COLS = ['workers','createdAt','updatedAt'];

const DEFAULT_SETTINGS = {
  dongs: ['101동','102동','103동','104동','105동','106동','107동','108동','109동','110동','111동','112동','113동',
          '114동','115동','116동','117동','118동','119동'],
  bmin: 2, fmax: 25,
  trades: ['토사반출','양수작업','토류판','파일하차','항타작업','JSP'],
  sup: [], safe: [], mapId: '', sheetUrl: '',
  lat: 35.1885, lon: 128.8125
};

/* ---------------- 진입점 ---------------- */
function doGet() {
  return out_({ ok: true, app: 'SOS', time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return out_({ ok: false, error: '요청 형식이 잘못되었습니다' }); }
  try {
    const fn = ACTIONS[req.action];
    if (!fn) return out_({ ok: false, error: '알 수 없는 요청: ' + req.action });
    return out_(Object.assign({ ok: true }, fn(req) || {}));
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err), code: err && err.code || '' });
  }
}

const ACTIONS = {
  list: r => ({ works: listWorks_(r.from || '0000', r.to || '9999') }),
  add: r => withLock_(() => ({ id: addWork_(r.work || {}) })),
  update: r => withLock_(() => { updateWork_(r.id, r.patch || {}); return {}; }),
  remove: r => withLock_(() => { removeWork_(r.id); return {}; }),
  getSettings: () => ({ settings: getSettings_() }),
  setSettings: r => { requireAdmin_(r.pw); return withLock_(() => ({ settings: setSettings_(r.settings || {}) })); },
  checkAdmin: r => { requireAdmin_(r.pw); return {}; },
  changePw: r => { requireAdmin_(r.pw); changePw_(r.newPw); return {}; },
  uploadMap: r => { requireAdmin_(r.pw); return withLock_(() => uploadMap_(r)); },
  fillSheet: r => withLock_(() => fillSheet_(r.date))
};

/* ---------------- 공통 ---------------- */
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function err_(msg, code) { const e = new Error(msg); e.code = code || ''; throw e; }

function dataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, sh.getMaxRows(), COLS.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#DCE8FC');
    sh.setFrozenRows(1);
  }
  return sh;
}
function enc_(k, v) {
  if (v === undefined || v === null) return '';
  if (JSON_COLS.indexOf(k) >= 0) return JSON.stringify(v || []);
  if (BOOL_COLS.indexOf(k) >= 0) return v ? 'TRUE' : 'FALSE';
  return String(v);
}
function dec_(k, v) {
  v = v === null || v === undefined ? '' : String(v);
  if (JSON_COLS.indexOf(k) >= 0) { try { return v ? JSON.parse(v) : []; } catch (e) { return []; } }
  if (BOOL_COLS.indexOf(k) >= 0) return v === 'TRUE' || v === 'true';
  if (NUM_COLS.indexOf(k) >= 0) return v === '' ? 0 : Number(v);
  if ((k === 'stands' || k === 'cutters') && v !== '') return Number(v);
  return v;
}
function readAll_() {
  const sh = dataSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, COLS.length).getDisplayValues()
    .map((row, i) => { const o = { _row: i + 2 }; COLS.forEach((k, j) => o[k] = dec_(k, row[j])); return o; })
    .filter(o => o.id);
}
function strip_(w) { const o = Object.assign({}, w); delete o._row; return o; }

/* ---------------- 작업 ---------------- */
function listWorks_(from, to) {
  return readAll_().filter(w => w.date >= from && w.date <= to).map(strip_);
}
function addWork_(w) {
  if (!w.date || !w.company || !w.title) err_('작업 일자·업체·작업명이 필요합니다');
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  w.id = id; w.createdAt = w.createdAt || Date.now(); w.updatedAt = w.updatedAt || '';
  const sh = dataSheet_();
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, COLS.length).setNumberFormat('@').setValues([COLS.map(k => enc_(k, w[k]))]);
  return id;
}
function findRow_(id) {
  const w = readAll_().filter(x => x.id === id)[0];
  if (!w) err_('작업을 찾을 수 없습니다. 이미 삭제되었을 수 있습니다', 'not_found');
  return w;
}
function updateWork_(id, patch) {
  const w = findRow_(id);
  const next = Object.assign({}, strip_(w), patch, { id: id, updatedAt: Date.now() });
  dataSheet_().getRange(w._row, 1, 1, COLS.length).setValues([COLS.map(k => enc_(k, next[k]))]);
}
function removeWork_(id) {
  const w = findRow_(id);
  dataSheet_().deleteRow(w._row);
}

/* ---------------- 설정 · 관리자 ---------------- */
function props_() { return PropertiesService.getScriptProperties(); }
function getSettings_() {
  let s = {};
  try { s = JSON.parse(props_().getProperty('SETTINGS') || '{}'); } catch (e) {}
  return Object.assign({}, DEFAULT_SETTINGS, s);
}
function setSettings_(s) {
  const allowed = ['dongs','bmin','fmax','trades','sup','safe','mapId','sheetUrl','lat','lon'];
  const cur = getSettings_();
  allowed.forEach(k => { if (s[k] !== undefined) cur[k] = s[k]; });
  props_().setProperty('SETTINGS', JSON.stringify(cur));
  return cur;
}
function hash_(pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'sos-kimhae:' + String(pw || ''), Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function requireAdmin_(pw) {
  const saved = props_().getProperty('ADMIN_HASH') || hash_(ADMIN_PASSWORD);
  if (!pw || hash_(pw) !== saved) err_('관리 비밀번호가 맞지 않습니다', 'auth');
}
function changePw_(pw) {
  if (!pw || String(pw).length < 4) err_('비밀번호는 4자 이상이어야 합니다');
  props_().setProperty('ADMIN_HASH', hash_(pw));
}

/* ---------------- 배치도 이미지 ---------------- */
function uploadMap_(r) {
  if (!r.b64) err_('이미지가 없습니다');
  const folders = DriveApp.getFoldersByName(MAP_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(MAP_FOLDER);
  const blob = Utilities.newBlob(Utilities.base64Decode(r.b64), r.mime || 'image/jpeg', r.name || 'site-map.jpg');
  const file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const cur = getSettings_();
  if (cur.mapId && cur.mapId !== file.getId()) { try { DriveApp.getFileById(cur.mapId).setTrashed(true); } catch (e) {} }
  const s = setSettings_({ mapId: file.getId() });
  return { id: file.getId(), settings: s };
}

/* ---------------- 날짜 탭 양식 채우기 (서식 유지) ----------------
 * 날짜 탭 이름: 26.11.05 형식. A열에 업체명이 적힌 줄에 B~M 값만 써 넣습니다.
 * 같은 날 다시 누르면 지난번에 채운 칸을 비우고 새로 채웁니다(사람이 직접 적은 줄은 건드리지 않음).
 */
function fillSheet_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) err_('날짜가 올바르지 않습니다');
  const tab = date.slice(2).replace(/-/g, '.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(tab);
  if (!sh) err_('시트에 "' + tab + '" 탭이 없습니다. 날짜 탭을 먼저 만들어 주세요', 'no_tab');

  const key = 'FILL_' + sh.getSheetId();
  let filled = [];
  try { filled = JSON.parse(props_().getProperty(key) || '[]'); } catch (e) {}
  filled.forEach(r => { try { sh.getRange(r, 2, 1, 12).clearContent(); } catch (e) {} });
  filled = [];

  const works = listWorks_(date, date).filter(w => w.status !== '작업취소');
  const order = [];
  works.slice().sort((a, b) => a.createdAt - b.createdAt).forEach(w => { if (order.indexOf(w.company) < 0) order.push(w.company); });
  works.sort((a, b) => order.indexOf(a.company) - order.indexOf(b.company) || a.createdAt - b.createdAt);

  let end = dataEnd_(sh);
  let added = 0;
  works.forEach(w => {
    const vals = sh.getRange(3, 1, Math.max(1, end - 2), 13).getDisplayValues();
    let target = -1, lastSame = -1, lastAny = 2;
    for (let i = 0; i < vals.length; i++) {
      const r = i + 3, a = String(vals[i][0]).trim();
      if (a) lastAny = r;
      if (a === w.company) {
        lastSame = r;
        if (target < 0 && filled.indexOf(r) < 0 && vals[i].slice(1).every(v => String(v).trim() === '')) target = r;
      }
    }
    if (target < 0) {
      const after = lastSame > 0 ? lastSame : lastAny;
      sh.insertRowAfter(after);
      target = after + 1;
      sh.getRange(after, 1, 1, 13).copyTo(sh.getRange(target, 1, 1, 13), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sh.getRange(after, 1, 1, 13).copyTo(sh.getRange(target, 1, 1, 13), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      sh.setRowHeight(target, sh.getRowHeight(after));
      sh.getRange(target, 1).setValue(w.company);
      filled = filled.map(r => r >= target ? r + 1 : r);
      end++; added++;
    }
    const row = sh.getRange(target, 2, 1, 12);
    row.setValues([sheetValues_(w, sh, target)]);
    filled.push(target);
  });
  props_().setProperty(key, JSON.stringify(filled));
  return { tab: tab, count: works.length, added: added };
}
function dataEnd_(sh) {
  const last = Math.max(3, sh.getLastRow());
  const a = sh.getRange(3, 1, last - 2, 1).getDisplayValues();
  for (let i = 0; i < a.length; i++) {
    if (String(a[i][0]).indexOf('->') >= 0 || String(a[i][0]).indexOf('협력회사') >= 0) return i + 2;
  }
  return last;
}
function pick_(cell, positive, fallback) {
  try {
    const rule = cell.getDataValidation();
    if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const list = rule.getCriteriaValues()[0] || [];
      if (!positive) { const n = list.filter(v => String(v).indexOf('해당없음') >= 0)[0]; return n || fallback; }
      const p = list.filter(v => String(v).indexOf('해당없음') < 0 && String(v).trim() !== '')[0];
      return p || fallback;
    }
  } catch (e) {}
  return fallback;
}
function sheetValues_(w, sh, r) {
  const equip = String(w.equip || '').split(/[,\n]/).map(s => s.trim().replace(/(\d+)\s*대$/, '$1')).filter(String).join('\n');
  const mgr = w.supHdc ? (/M$/.test(w.supHdc) ? w.supHdc : w.supHdc + 'M') : '';
  const ptwOn = String(w.ptw || '').indexOf('해당 (O)') === 0 || String(w.ptw || '') === '해당';
  const loc = w.loc || (w.dongs || []).join('\n');
  let note = w.note || '';
  if (ptwOn && (w.ptwTypes || []).length) note = (note ? note + '\n' : '') + 'PTW: ' + w.ptwTypes.join(', ');
  return [
    loc,
    w.title,
    Number(w.workers) || 0,
    equip,
    pick_(sh.getRange(r, 6), !!w.riskUp, w.riskUp ? '상등급' : '해당없음'),
    pick_(sh.getRange(r, 7), ptwOn, ptwOn ? 'PTW 대상' : '해당없음'),
    w.early || '',
    mgr,
    w.watch === '배치' ? '상주' : '',
    w.edu || pick_(sh.getRange(r, 11), false, '해당없음'),
    pick_(sh.getRange(r, 12), !!w.opc, w.opc ? '해당' : '해당없음'),
    note
  ];
}
