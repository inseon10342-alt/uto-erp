/**
 * 통합 업무관리 · 서버 (Google Apps Script)  — 1차
 *
 * - 원본 데이터: Google 스프레드시트 「업무관리 DB」 (setup 이 만듦)
 * - 전환 전(설정 원본=엑셀): Drive 의 통합관리 엑셀(.xlsm)을 「가져오기」로 DB 에 옮김 (미리보기 → 관리자 승인)
 *                            이 기간에는 엑셀에서 온 표(현장·계약·점검·청구·연락·거래처)는 웹에서 고칠 수 없음
 * - 전환 후(원본=웹): 웹에서 등록·수정, 엑셀은 백업·내보내기
 * - 공유용 모바일 뷰어(현장 조회 프로젝트)와는 완전히 별개 — 그 프로젝트·파일은 건드리지 않음
 * - 백업 파일 이름에는 「통합관리」를 넣지 않음 (공유용 뷰어가 그 이름의 엑셀을 읽기 때문)
 */

// ===================== 설정 =====================
const SERVER_VERSION = '2026-11-16b'; // @uto-erp-server · GitHub server/VERSION 과 같게 맞춤
const CFG = {
  DB_NAME: '업무관리 DB',
  SOURCE_NAME: '통합관리',          // 가져올 엑셀: 이 글자가 들어간 가장 최근 .xlsm (스크립트 속성 XLSM_FILE_ID 로 고정 가능)
  BACKUP_FOLDER: '업무관리 백업',
  TMP_FOLDER: '업무관리 임시',
  DEVICE_DAYS: 180,
  CODE_MIN: 10,
  BACKUP_KEEP: 60,                    // 자동 백업 보관 개수
  APP_TITLE: '업무관리',
};
const P = PropertiesService.getScriptProperties();
const META = ['버전', '수정일', '수정자', '삭제'];
const SCHEMA = {
  현장: ['id', '현장명', '주소', '관할구', '규모', '비고'],
  거래처: ['id', '거래처명', '사업자번호', '대표자', '주소', '업태', '종목', '구분', '전화', '이메일', '비고'],
  계약: ['id', '현장id', '계약명', '연도', 'No', '계약일', '공급가', '부가세', '계약금', 'VAT기준', '종합보고서구분', '종합보고서금액', '계약외', '할증', '이월수금', '미수금', '상태',
    '발주처id', '실제입금처id', '시공사id', '용역코드', '준공예정일', '완료일', '비고', '입금비고', '기존수금', '기존미입금', '상태지정'], // 상태지정: 웹 전용 — 완료 처리 / 진행으로 복원 (비어 있으면 엑셀 수식처럼 수금 기준 자동)
  점검: ['id', '계약id', '항목', '단계', '예정일', '점검일', '기한', '발송일', '시기', '위치', '내부검토', '비고', '작성완료일'], // 작성완료일: 웹 전용 (엑셀에 칸 없음)
  청구: ['id', '계약id', '발행일', '품명', '청구액', '공급가', '부가세', '입금일', '수금액', '입금자', '상태', '연결점검', '계약외', '할증', '비고', '분할입금'],
  연락: ['id', '현장id', '연락일', '방법', '담당자', '공정', '메모', '예상점검', '예상시기', '다음연락', '비고', '작성자'],
  문서: ['id', '종류', '번호', '계약id', '현장id', '거래처id', '일자', '상태', '제목', '내용', '공급가', '세액', '합계', '파일', '작성자'],
  자료요청: ['id', 'NO', '현장id', '요청일', '방법', '담당자', '항목', '점검일', '구분', '원요청', '자료', '기한', '수신일', '상태', '검토', '비고'],
  메일: ['id', 'NO', '현장id', '구분', '항목', '받는사람', '이메일', '보낸날', '제목', '내용', '첨부', '회신필요', '회신상태', '회신일', '회신내용', '회신첨부', '비고'],
  검토: ['id', 'NO', '계약id', '항목', '회차', '요청일', '검토자', '제목', '상태', '회신일', '의견', '조치', '재발송일', '비고'],
  일지: ['id', '날짜', '현장', '구분', '내용', '상태', '후속', '비고', '작성자'],
  택배: ['id', '현장id', '발송일', '항목', '받는사람', '주소', '비고'],
  // 종합보고서: 용역이 끝난 계약에 나중에 들어오는 종합보고서 요청 (웹 전용). 한 줄 = 요청 1회(차수), 재작성 요청은 다음 차수로 새 줄
  종합: ['id', '계약id', '현장id', '차수', '요청일', '요청자', '상태', '작성일', '제출일', '제출처', '파일', '비고'],
  // 외주관리 (엑셀 「외주관리」 A~K 입력 칸). 업체명·현장명은 엑셀에 적힌 글자 그대로 보관 (거래처·계약을 못 찾아도 값이 사라지지 않게)
  외주: ['id', '계산서일', '업체id', '업체명', '계약id', '현장명', '점검일', '내용', '청구금액', '지급금액', '지급일', '비고', '연결점검'],
  // 회차별 단가 (엑셀 「회차별 단가」: 1줄 = 용역 1개의 점검항목(공정) 1개). VAT 별도 계약은 공급가로 저장 (엑셀과 같음)
  단가: ['id', '계약id', '계약명', '항목', '단가', '횟수', '회차', '메모'],
};
const NUMCOLS = ['연도', 'No', '공급가', '부가세', '계약금', '종합보고서금액', '계약외', '할증', '이월수금', '미수금', '청구액', '수금액', '세액', '합계', '버전', '기존수금', '기존미입금', '차수', '청구금액', '지급금액', '횟수']; // 단가(회차별 단가)는 빈 칸과 0원을 구분해야 해서 글자로 보관
const FROM_EXCEL = ['현장', '거래처', '계약', '점검', '청구', '연락', '자료요청', '메일', '검토', '일지', '택배', '외주', '단가'];   // 가져오기가 채우는 표
const ROLE_RANK = { 조회: 0, 일반: 1, 관리자: 2 };
const WRITE_ROLE = { 현장: '일반', 거래처: '일반', 계약: '관리자', 점검: '일반', 청구: '관리자', 연락: '일반', 문서: '일반', 자료요청: '일반', 메일: '일반', 검토: '일반', 일지: '일반', 택배: '일반', 종합: '일반', 외주: '관리자', 단가: '관리자' };
const LABEL = { 현장: '현장', 거래처: '거래처', 계약: '계약', 점검: '점검', 청구: '청구(계산서)', 연락: '연락기록', 문서: '문서', 자료요청: '자료요청', 메일: '메일', 검토: '보고서 검토', 일지: '업무일지', 택배: '택배', 종합: '종합보고서', 외주: '외주', 단가: '회차별 단가' };

// ===================== 진입 =====================
function doGet(e) {
  if (e && e.parameter && e.parameter.cb) {
    const cb = String(e.parameter.cb).replace(/[^A-Za-z0-9_]/g, '').slice(0, 40) || 'cb';
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(dispatch_(e.parameter.q)) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput('업무관리 서버입니다. 화면은 GitHub Pages 주소로 여세요.');
}
function doPost(e) {
  return ContentService.createTextOutput(JSON.stringify(dispatch_(e && e.postData && e.postData.contents))).setMimeType(ContentService.MimeType.JSON);
}
function dispatch_(raw) {
  try {
    const q = JSON.parse(raw || '{}');
    const fns = { requestCode: requestCode, verifyCode: verifyCode, api: api };
    if (!Object.prototype.hasOwnProperty.call(fns, q.fn)) throw new Error('허용되지 않은 요청');
    return fns[q.fn].apply(null, Array.isArray(q.args) ? q.args.slice(0, 3) : []);
  } catch (err) {
    return { ok: false, code: 'error', msg: '요청을 처리하지 못했습니다: ' + err.message };
  }
}

// ===================== 처음 한 번: setup =====================
function setup() {
  if (!P.getProperty('SECRET')) P.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  let id = P.getProperty('DB_ID');
  if (!id || !tryOpen_(id)) {
    const ss = SpreadsheetApp.create(CFG.DB_NAME);
    id = ss.getId();
    P.setProperty('DB_ID', id);
    DB_ = ss;
  }
  const ss = db_();
  Object.keys(SCHEMA).forEach((t) => ensureSheet_(ss, t, SCHEMA[t].concat(META)));
  ensureSheet_(ss, '변경이력', ['시각', '사용자', '표', 'id', '동작', '전', '후']);
  const u = ensureSheet_(ss, '사용자', ['이름', '이메일', '권한(관리자/일반/조회)', '상태(승인/차단)', '메모']);
  if (u.getLastRow() < 2) u.appendRow(['관리자', Session.getEffectiveUser().getEmail(), '관리자', '승인', '설치한 사람']);
  ensureSheet_(ss, '기기', ['기기ID', '이메일', '기기 이름', '토큰 지문(해시)', '등록일', '마지막 접속', '만료일', '상태(사용/해제/차단)', '메모']);
  const s = ensureSheet_(ss, '설정', ['키', '값', '설명']);
  Object.keys(SCHEMA).forEach((t) => { if (db_().getSheetByName(t)) ensureCols_(t); });
  if (!setting_('원본')) setSetting_('원본', '엑셀', '엑셀 = 통합관리 엑셀이 원본(웹은 조회·가져오기) / 웹 = 웹이 원본');
  const have = s.getDataRange().getValues().map((r) => String(r[0]));
  [['회사_상호', '㈜유토엔지니어링건축사사무소'], ['회사_대표', ''], ['회사_사업자번호', ''], ['회사_주소', ''], ['회사_업태', ''], ['회사_종목', ''], ['회사_전화', ''], ['회사_팩스', ''], ['회사_거래명세서_기타사항', '']]
    .forEach(([k, v]) => { if (have.indexOf(k) < 0) s.appendRow([k, v, '거래명세서에 들어가는 회사 정보 (값 칸에 입력)']); });
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'dailyBackup').forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(3).everyDays(1).inTimezone('Asia/Seoul').create();
  folder_(CFG.BACKUP_FOLDER); folder_(CFG.TMP_FOLDER);
  Logger.log('업무관리 DB: https://docs.google.com/spreadsheets/d/' + id);
  return id;
}
function tryOpen_(id) { try { return SpreadsheetApp.openById(id); } catch (e) { return null; } }
function ensureSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), head.length).setNumberFormat('@'); // 날짜를 글자 그대로 보관
  }
  return sh;
}
function folder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// ===================== 로그인 (이메일 코드 + 기기 등록) =====================
function requestCode(email) {
  email = String(email || '').trim().toLowerCase();
  const u = findUser_(email);
  Utilities.sleep(400);
  const msg = '등록된 이메일이면 인증 코드를 보냈습니다. 메일이 오지 않으면 관리자에게 승인을 요청하세요.';
  if (!u || u.status !== '승인') return { ok: true, msg };
  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('code:' + email, JSON.stringify({ h: hash_(code + email), n: 0 }), CFG.CODE_MIN * 60);
  MailApp.sendEmail({ to: email, subject: '[' + CFG.APP_TITLE + '] 로그인 인증 코드', body: `인증 코드: ${code}\n\n${CFG.CODE_MIN}분 안에 화면에 입력하세요. 요청하지 않았다면 이 메일을 무시하세요.` });
  return { ok: true, msg };
}
function verifyCode(email, code, deviceName) {
  email = String(email || '').trim().toLowerCase();
  const c = CacheService.getScriptCache(), key = 'code:' + email, raw = c.get(key);
  if (!raw) return { ok: false, msg: '인증 코드가 만료되었습니다. 다시 요청하세요.' };
  const st = JSON.parse(raw);
  if (st.n >= 5) { c.remove(key); return { ok: false, msg: '여러 번 틀려 코드가 무효가 되었습니다. 다시 요청하세요.' }; }
  if (st.h !== hash_(String(code).trim() + email)) { st.n++; c.put(key, JSON.stringify(st), CFG.CODE_MIN * 60); return { ok: false, msg: '인증 코드가 맞지 않습니다.' }; }
  c.remove(key);
  const u = findUser_(email);
  if (!u || u.status !== '승인') return { ok: false, msg: '승인된 사용자가 아닙니다.' };
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const devId = 'D' + Date.now().toString(36).toUpperCase();
  const now = new Date(), exp = new Date(now.getTime() + CFG.DEVICE_DAYS * 864e5);
  db_().getSheetByName('기기').appendRow([devId, email, String(deviceName || '기기').slice(0, 60), hash_(token), fmtT_(now), fmtT_(now), fmtT_(exp), '사용', '']);
  return { ok: true, token, deviceId: devId, user: { name: u.name, email, role: u.role } };
}
function hash_(s) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s + '|' + P.getProperty('SECRET'), Utilities.Charset.UTF_8);
  return b.map((x) => ((x + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function findUser_(email) {
  const v = db_().getSheetByName('사용자').getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][1]).trim().toLowerCase() === email) {
      const role = String(v[i][2]).trim();
      return { name: String(v[i][0]), email, role: ROLE_RANK[role] == null ? '조회' : role, status: String(v[i][3]).trim(), row: i + 1 };
    }
  }
  return null;
}
function auth_(token) {
  if (!token || String(token).length < 40) return { ok: false, code: 'auth', msg: '로그인이 필요합니다.' };
  const h = hash_(String(token));
  const C = CacheService.getScriptCache(), hit = C.get('auth:' + h);
  if (hit) return JSON.parse(hit);
  const sh = db_().getSheetByName('기기'), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][3] !== h) continue;
    const row = i + 1, status = String(v[i][7]), exp = toDate_(v[i][6]);
    if (status !== '사용') return { ok: false, code: 'auth', msg: status === '차단' ? '이 기기는 관리자가 차단했습니다.' : '이 기기의 로그인이 해제되었습니다. 다시 인증하세요.' };
    if (!exp || exp < new Date()) { sh.getRange(row, 8).setValue('만료'); return { ok: false, code: 'auth', msg: '로그인 기간이 끝났습니다. 다시 인증하세요.' }; }
    const u = findUser_(String(v[i][1]).toLowerCase());
    if (!u || u.status !== '승인') return { ok: false, code: 'auth', msg: '사용 권한이 없습니다. 관리자에게 문의하세요.' };
    const now = new Date();
    if ((exp - now) / 864e5 < CFG.DEVICE_DAYS - 30) sh.getRange(row, 6, 1, 2).setValues([[fmtT_(now), fmtT_(new Date(now.getTime() + CFG.DEVICE_DAYS * 864e5))]]);
    const res = { ok: true, user: { name: u.name, email: u.email, role: u.role }, device: { id: v[i][0], name: v[i][2], row, h } };
    C.put('auth:' + h, JSON.stringify(res), 600); // 10분. 해제·차단·권한 변경 때는 즉시 지움
    return res;
  }
  return { ok: false, code: 'auth', msg: '등록되지 않은 기기입니다. 로그인하세요.' };
}
function can_(who, need) { return ROLE_RANK[who.user.role] >= ROLE_RANK[need]; }

// ===================== 읽기 결과 캐시 (속도) =====================
// 표를 고치면(log_) 데이터 버전이 바뀌어 예전 캐시는 자동으로 안 쓰게 됩니다.
function dataVer_() { return P.getProperty('DATA_VER') || '0'; }
function bumpVer_() { P.setProperty('DATA_VER', Date.now().toString(36)); }
function cached_(name, fn) {
  const C = CacheService.getScriptCache();
  const raw = 'r|' + SERVER_VERSION + '|' + dataVer_() + '|' + today_() + '|' + name; // 서버를 새로 올리면 예전 계산 결과를 쓰지 않음
  const key = 'r:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8));
  try {
    const head = C.get(key);
    if (head) {
      const n = Number(head);
      const keys = []; for (let i = 0; i < n; i++) keys.push(key + ':' + i);
      const got = C.getAll(keys);
      if (keys.every((k) => got[k] != null)) return JSON.parse(keys.map((k) => got[k]).join(''));
    }
  } catch (e) {}
  const val = fn();
  try {
    const str = JSON.stringify(val), size = 25000, parts = {};
    let n = 0; for (let i = 0; i < str.length; i += size) parts[key + ':' + (n++)] = str.slice(i, i + size);
    if (n <= 40) { C.putAll(parts, 21600); C.put(key, String(n), 21600); }
  } catch (e) {}
  return val;
}

// ===================== API 입구 =====================
function api(token, action, arg) {
  arg = arg || {};
  const who = auth_(token);
  if (!who.ok) return who;
  try {
    ensureAll_();
    const f = ACTIONS[action];
    if (!f) return { ok: false, msg: '알 수 없는 요청: ' + action };
    const out = f(who, arg);
    return Object.assign({ ok: true }, out || {});
  } catch (e) {
    return { ok: false, code: 'error', msg: e.message };
  }
}
// 저장 직후 화면이 다시 부를 조회를 같은 요청 안에서 미리 계산해 돌려줌 (왕복 한 번 절약). 조회 전용만 허용
const AFTER_OK = ['finals', 'dash', 'lookups', 'sites', 'site', 'contracts', 'contract', 'schedule', 'contacts', 'requests', 'money', 'reviews', 'journal', 'parcels', 'report', 'partners', 'partner', 'prices', 'outsource'];
function withAfter_(who, a, out) {
  if (!out || out.ok === false || !Array.isArray(a.after)) return out;
  const pre = [];
  a.after.slice(0, 6).forEach((c) => {
    if (!Array.isArray(c) || AFTER_OK.indexOf(c[0]) < 0) return;
    try { pre.push([c[0], c[1] || {}, ACTIONS[c[0]](who, c[1] || {})]); } catch (e) {}
  });
  return Object.assign({}, out, { pre });
}
const WEB_ONLY = { 점검: ['작성완료일'], 계약: ['계약일', '상태지정'] }; // 계약일: 완료된 용역은 엑셀에 칸이 없어 웹에서 넣은 값을 가져오기 때 지우지 않음
// ---- 데이터 점검: 표 사이 연결이 끊겼거나 빠진 값을 찾아 보여줌 (읽기만 하고 아무것도 고치지 않음) ----
function audit_() {
  const S = byId_('현장'), C = byId_('계약'), cs = live_('계약'), insp = live_('점검'), inv = live_('청구'), rv = live_('검토');
  const out = [], add = (구분, 심각, 설명, list) => out.push({ 구분, 심각, 설명, 건수: list.length, 예: list.slice(0, 30) });
  const nm = (c) => (c.계약명 || c.id) + (c.연도 ? ' (' + c.연도 + ')' : '');
  const byC = {}; insp.forEach((x) => { (byC[x.계약id] = byC[x.계약id] || []).push(x); });
  const done = cs.filter((c) => c.상태 && c.상태 !== '진행중'), act = cs.filter((c) => c.상태 === '진행중');
  add('연결', '오류', '계약에 연결된 현장이 없음', cs.filter((c) => !S[c.현장id]).map(nm));
  add('연결', '오류', '점검에 연결된 계약이 없음', insp.filter((x) => !C[x.계약id]).map((x) => x.id));
  add('연결', '오류', '세금계산서에 연결된 계약이 없음', inv.filter((b) => b.계약id && !C[b.계약id]).map((b) => (b.발행일 || '') + ' ' + (b.품명 || b.id)));
  add('연결', '오류', '보고서 검토에 연결된 계약이 없음', rv.filter((v) => !C[v.계약id]).map((v) => v.id));
  add('연결', '확인', '보고서 검토의 점검항목이 그 계약의 점검에 없음 (종합보고서 제외)', rv.filter((v) => C[v.계약id] && key_(v.항목) !== key_('종합보고서') && !(byC[v.계약id] || []).some((x) => key_(x.항목) === key_(v.항목))).map((v) => nm(C[v.계약id]) + ' · ' + v.항목));
  add('연결', '확인', '계약이 하나도 없는 현장', live_('현장').filter((s) => !cs.some((c) => c.현장id === s.id)).map((s) => s.현장명));
  add('상태', '오류', '완료된 용역인데 완료가 아닌 점검이 있음', done.filter((c) => (byC[c.id] || []).some((x) => stageOf_(x) !== '완료')).map((c) => nm(c) + ' · ' + (byC[c.id] || []).filter((x) => stageOf_(x) !== '완료').map((x) => x.항목 + '=' + stageOf_(x)).join(', ')));
  add('상태', '확인', '진행 중 용역인데 점검이 모두 완료됨 (완료 처리 대상일 수 있음)', act.filter((c) => (byC[c.id] || []).length && (byC[c.id] || []).every((x) => stageOf_(x) === '완료')).map(nm));
  add('상태', '확인', '점검이 한 건도 없는 계약', cs.filter((c) => !(byC[c.id] || []).length).map(nm));
  add('누락', '참고', '계약일이 비어 있는 완료 용역 (엑셀 용역대장·완료보관에 계약일 칸이 없음 → 웹에서 입력)', done.filter((c) => !c.계약일).map(nm));
  add('누락', '확인', '계약일이 비어 있는 진행 중 용역', act.filter((c) => !c.계약일).map(nm));
  add('누락', '참고', '완료된 점검인데 시공사 발송일 기록이 없음 (이전 관리파일에서 점검일만 넘어온 건)', insp.filter((x) => stageOf_(x) === '완료' && !x.발송일).map((x) => nm(C[x.계약id] || {}) + ' · ' + x.항목));
  add('누락', '확인', '완료된 점검인데 점검일이 없음', insp.filter((x) => stageOf_(x) === '완료' && !x.점검일).map((x) => nm(C[x.계약id] || {}) + ' · ' + x.항목));
  return { at: fmtT_(new Date()), 수: { 현장: live_('현장').length, 계약: cs.length, 진행: act.length, 완료: done.length, 점검: insp.length, 세금계산서: inv.length, 검토: rv.length }, checks: out };
}
// ---- 종합보고서 ----
function finals_() {
  const C = byId_('계약'), S = byId_('현장'), Pn = byId_('거래처'), insp = live_('점검');
  const pn = (id) => (Pn[id] || {}).거래처명 || '';
  const byC = {}; insp.forEach((x) => { (byC[x.계약id] = byC[x.계약id] || []).push(x); });
  const info = (c) => { const its = byC[c.id] || [], ds = its.map((x) => x.점검일).filter(Boolean).sort();
    return { 계약명: c.계약명 || '', 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || '', 현장id: c.현장id || '', 연도: c.연도, 계약상태: c.상태 || '', 계약일: c.계약일 || '', 계약금: c.계약금 || 0,
      시공사: pn(c.시공사id) || pn(c.발주처id), 발주처: pn(c.발주처id), 종합보고서구분: c.종합보고서구분 || '', 종합보고서금액: c.종합보고서금액 || 0, 점검수: its.length, 점검완료: its.filter((x) => stageOf_(x) === '완료').length, 첫점검: ds[0] || '', 마지막점검: ds[ds.length - 1] || '' }; };
  // 내부 검토: 보고서 검토 관리(검토 표)에 점검항목 「종합보고서」로 기록 — 정기점검 보고서 검토와 같은 방식, 단계 계산에는 영향 없음
  const RVF = {}; live_('검토').filter((v) => key_(v.항목) === key_('종합보고서')).forEach((v) => { (RVF[v.계약id] = RVF[v.계약id] || []).push(clean_(v)); });
  const rvSt = (list) => { if (!list || !list.length) return ''; if (list.some((v) => v.상태 === '검토완료')) return '검토완료'; const l = list.slice().sort((a, b) => (+a.회차 || 0) - (+b.회차 || 0) || String(a.요청일).localeCompare(String(b.요청일))).pop(); return l.상태 === '수정요청' ? (l.재발송일 ? '검토중' : '수정요청') : '검토중'; };
  const rows = live_('종합').map((r) => { const c = C[r.계약id] || {}; let files = []; try { files = JSON.parse(r.파일 || '[]'); } catch (e) {} const rv = (RVF[r.계약id] || []).sort((a, b) => (+a.회차 || 0) - (+b.회차 || 0)); return Object.assign(clean_(r), info(c), { 파일: files, 검토: rv, 검토상태: rvSt(rv) }); })
    .sort((a, b) => String(b.요청일).localeCompare(String(a.요청일)) || (b.차수 - a.차수));
  // 계약별 최신 차수만 「현재 상태」로 셈
  const last = {}; rows.forEach((r) => { if (!last[r.계약id] || r.차수 > last[r.계약id].차수) last[r.계약id] = r; });
  const cur = Object.values(last);
  const done = live_('계약').filter((c) => c.상태 && c.상태 !== '진행중').map((c) => Object.assign({ id: c.id }, info(c), { 요청: last[c.id] ? last[c.id].상태 : '' }))
    .sort((a, b) => (b.연도 || 0) - (a.연도 || 0) || a.현장명.localeCompare(b.현장명));
  return { list: rows, done, kpi: { 완료용역: done.length, 요청: cur.length, 미작성: cur.filter((r) => r.상태 === '미작성').length, 작성중: cur.filter((r) => r.상태 === '작성 중').length, 검토중: cur.filter((r) => r.상태 !== '제출 완료' && (r.검토상태 === '검토중' || r.검토상태 === '수정요청')).length, 제출완료: cur.filter((r) => r.상태 === '제출 완료').length } };
}
function finalDir_() {
  const id = P.getProperty('FINAL_DIR');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed || !f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder('업무관리 첨부 - 종합보고서');
  P.setProperty('FINAL_DIR', f.getId()); return f;
}
function finalFile_(who, a) {
  need_(who, '일반');
  const r = live_('종합').find((x) => x.id === a.id); if (!r) throw new Error('종합보고서 요청을 찾지 못했습니다. 새로고침해 보세요.');
  let files = []; try { files = JSON.parse(r.파일 || '[]'); } catch (e) {}
  if (a.op === 'add') {
    const name = String(a.name || '').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
    if (!/\.(hwp|hwpx|pdf)$/i.test(name)) throw new Error('HWP, HWPX, PDF 파일만 올릴 수 있습니다.');
    const data = String(a.data || ''); if (!data) throw new Error('파일 내용이 비어 있습니다.');
    if (data.length > 40 * 1024 * 1024) throw new Error('30MB 이하 파일만 올릴 수 있습니다.');
    const mime = /\.pdf$/i.test(name) ? 'application/pdf' : 'application/x-hwp';
    const bytes = Utilities.base64Decode(data), f = finalDir_().createFile(Utilities.newBlob(bytes, mime, name));
    files.push({ id: f.getId(), n: name, s: bytes.length, d: today_(), 차수: r.차수, 구분: a.kind || '' , by: who.user.email });
  } else if (a.op === 'del') {
    const k = files.findIndex((x) => x.id === a.fid); if (k < 0) throw new Error('파일을 찾지 못했습니다.');
    try { DriveApp.getFileById(a.fid).setTrashed(true); } catch (e) {} // 드라이브 휴지통으로 (30일 안에 복구 가능)
    files.splice(k, 1);
  } else throw new Error('알 수 없는 파일 요청');
  const out = save_(who, '종합', { id: r.id, 파일: JSON.stringify(files) }, r.버전);
  if (out && out.ok === false) return out;
  return { files };
}
function finalGet_(who, a) {
  need_(who, '조회');
  const r = live_('종합').find((x) => x.id === a.id); if (!r) throw new Error('없음');
  let files = []; try { files = JSON.parse(r.파일 || '[]'); } catch (e) {}
  const f = files.find((x) => x.id === a.fid); if (!f) throw new Error('파일을 찾지 못했습니다.');
  const b = DriveApp.getFileById(f.id).getBlob();
  return { name: f.n, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}
const ACTIONS = {
  boot: (who) => (migrateG_(who.user.email), { user: who.user, source: setting_('원본') || '엑셀', counts: cached_('counts', counts_), dbUrl: can_(who, '관리자') ? 'https://docs.google.com/spreadsheets/d/' + P.getProperty('DB_ID') : '' }),
  refresh: () => { bumpVer_(); return {}; }, // 새로고침 버튼: DB 를 직접 고쳤을 때도 바로 반영
  logout: (who) => { setDevice_(who.device.row, '해제', who.device.h); return {}; },
  dash: () => cached_('dash', dash_),
  lookups: () => cached_('lookups', lookups_),
  sites: () => cached_('sites', () => ({ sites: sitesList_() })),
  site: (who, a) => cached_('site:' + a.id, () => siteDetail_(a.id)),
  contracts: () => cached_('contracts', () => ({ contracts: contractList_() })),
  contract: (who, a) => cached_('contract:' + a.id, () => contractDetail_(a.id)),
  schedule: (who, a) => cached_('schedule:' + a.from + ':' + a.to, () => ({ events: schedule_(a.from, a.to) })),
  contacts: () => cached_('contacts', () => ({ contacts: contactsList_() })),
  requests: () => cached_('requests', () => ({ requests: requestsList_(), mails: mailsList_() })),
  money: () => cached_('money', money_),
  finals: () => cached_('finals', finals_),
  finalFile: (who, a) => finalFile_(who, a),
  finalGet: (who, a) => finalGet_(who, a),
  reviews: () => cached_('reviews', () => {
    const list = reviewsList_(), has = {}; list.forEach((r) => { has[r.계약id + '|' + key_(r.항목)] = 1; });
    const C = byId_('계약'), S = byId_('현장'), Pn = byId_('거래처'), T = today_();
    const writing = live_('점검').filter((x) => isWriting_(stageOf_(x)) && !has[x.계약id + '|' + key_(x.항목)]).map((x) => { const c = C[x.계약id] || {}; return { 작성완료일: x.작성완료일 || '', 점검id: x.id, 계약id: x.계약id, 현장id: c.현장id, 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || '', 시공사: (Pn[c.시공사id] || Pn[c.발주처id] || {}).거래처명 || '', 항목: x.항목, 점검일: x.점검일, 기한: x.기한, 남은날: x.기한 ? dd_(x.기한, T) : null }; })
      .sort((a, b) => (a.남은날 == null ? 9999 : a.남은날) - (b.남은날 == null ? 9999 : b.남은날));
    return { reviews: list, writing };
  }),
  journal: () => cached_('journal', () => ({ jobs: live_('일지').map(clean_).sort((a, b) => String(a.날짜).localeCompare(String(b.날짜))), year: +(setting_('일지연도') || new Date().getFullYear()) })),
  parcels: () => cached_('parcels', parcelsList_),
  prices: () => cached_('prices', pricesList_),
  outsource: () => cached_('outsource', outsourceList_),
  report: () => cached_('report', report_),
  partners: () => cached_('partners', () => ({ partners: partnersList_() })),
  partner: (who, a) => cached_('partner:' + a.id, () => partnerDetail_(a.id)),
  docs: (who, a) => ({ docs: live_('문서').filter((d) => !a.종류 || d.종류 === a.종류).sort((x, y) => String(y.일자).localeCompare(String(x.일자))).map(clean_) }),
  doc: (who, a) => { const d = live_('문서').find((x) => x.id === a.id); return { doc: d ? clean_(d) : null }; },
  history: (who, a) => ({ history: history_(a.t, a.id) }),
  save: (who, a) => withAfter_(who, a, save_(who, a.t, a.row || {}, a.ver)),
  del: (who, a) => withAfter_(who, a, save_(who, a.t, { id: a.id, 삭제: 'Y' }, a.ver, true)),
  saveDoc: (who, a) => saveDoc_(who, a.doc || {}, a.ver, !!a.issue),
  // 관리자
  users: (who) => { need_(who, '관리자'); return users_(); },
  saveUser: (who, a) => { need_(who, '관리자'); return saveUser_(who, a); },
  revoke: (who, a) => revoke_(who, a),
  myDevices: (who) => ({ devices: devices_().filter((d) => d.email === who.user.email).map((d) => Object.assign(d, { me: d.id === who.device.id })) }),
  importPreview: (who) => { need_(who, '관리자'); return importPreview_(who); },
  importCommit: (who, a) => { need_(who, '관리자'); return importCommit_(who, a.key); },
  audit: (who) => { need_(who, '관리자'); return audit_(); },
  backups: (who) => { need_(who, '관리자'); return { backups: backups_() }; },
  backupNow: (who) => { need_(who, '관리자'); return { backup: backup_('수동', who.user.email) }; },
  restorePreview: (who, a) => { need_(who, '관리자'); return restorePreview_(a.id); },
  restoreCommit: (who, a) => { need_(who, '관리자'); return restoreCommit_(who, a.id); },
  setSource: (who, a) => { need_(who, '관리자'); return setSource_(who, a.value); },
  saveSeal: (who, a) => { need_(who, '관리자'); return saveSeal_(who, a); },
  sourceList: (who) => { need_(who, '관리자'); return sourceList_(); },
  setSourceFile: (who, a) => { need_(who, '관리자'); return setSourceFile_(who, a.id); },
  exportInfo: (who) => { need_(who, '관리자'); return exportInfo_(); },
  updateInfo: (who) => { need_(who, '관리자'); return updateInfo_(); },
  updateNow: (who) => { need_(who, '관리자'); return updateNow_(); },
  exportNow: (who) => {
    need_(who, '관리자');
    if ((setting_('원본') || '엑셀') !== '웹') return { ok: false, msg: '원본이 웹일 때만 엑셀에 저장합니다.' };
    const d = P.getProperty('EXPORT_DIRTY') || '';
    try { const r = exportExcel_(who.user.email); if ((P.getProperty('EXPORT_DIRTY') || '') === d) P.setProperty('EXPORT_DIRTY', ''); return Object.assign({ result: r }, exportInfo_()); }
    catch (e) { setSetting_('엑셀 저장 오류', fmtT_(new Date()) + ' · ' + e.message); return { ok: false, msg: e.message }; }
  },
};
function need_(who, role) { if (!can_(who, role)) throw new Error(role + ' 권한이 필요합니다.'); }

// ===================== DB 읽기·쓰기 =====================
let DB_ = null;
function db_() { return DB_ || (DB_ = SpreadsheetApp.openById(P.getProperty('DB_ID'))); }
let TBL_ = {};
function rows_(t) {
  if (TBL_[t]) return TBL_[t];
  const sh = db_().getSheetByName(t);
  const v = sh ? sh.getDataRange().getValues() : [[]];
  const H = v[0].map(String), out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '' || v[i][0] == null) continue;
    const o = { _row: i + 1 };
    H.forEach((h, j) => { if (h) o[h] = norm_(h, v[i][j]); });
    out.push(o);
  }
  return (TBL_[t] = out);
}
function live_(t) { return rows_(t).filter((r) => r.삭제 !== 'Y'); }
/** DB 시트의 머리줄이 지금 SCHEMA 와 다르면(열 추가 등) 이름 기준으로 옮겨 다시 씀 */
function ensureCols_(t) {
  const want = SCHEMA[t].concat(META), ck = 'cols:' + t + ':' + want.join('|').length + ':' + SERVER_VERSION, C = CacheService.getScriptCache();
  if (C.get(ck)) return; // 머리줄 확인은 6시간에 한 번
  const sh = db_().getSheetByName(t) || ensureSheet_(db_(), t, want); // 새로 생긴 표(외주·단가 등)는 시트를 만듦
  C.put(ck, '1', 21600);
  const H = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  if (H.slice(0, want.length).join('|') === want.join('|') && H.slice(want.length).every((h) => !h)) return;
  TBL_ = {}; const rows = rows_(t).map(clean_); writeTable_(t, rows);
}
/** 모든 표의 머리줄을 지금 SCHEMA 에 맞춤 (서버 업데이트로 칸이 늘면 첫 요청 때 한 번). 칸 위치로 쓰는 곳(setSilent_ 등)이 어긋나지 않게 */
function ensureAll_() {
  const C = CacheService.getScriptCache(), key = (t) => 'cols:' + t + ':' + SCHEMA[t].concat(META).join('|').length + ':' + SERVER_VERSION;
  const ks = Object.keys(SCHEMA), got = C.getAll(ks.map(key)) || {}, miss = ks.filter((t) => !got[key(t)]);
  if (!miss.length || !P.getProperty('DB_ID')) return;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { miss.forEach((t) => ensureCols_(t)); TBL_ = {}; } finally { lock.releaseLock(); }
}
function norm_(h, v) {
  if (NUMCOLS.indexOf(h) >= 0) return num_(v);
  if (v instanceof Date) return h === '수정일' ? fmtT_(v) : iso_(v);
  return v == null ? '' : String(v);
}
function clean_(r) { const o = {}; for (const k in r) if (k !== '_row') o[k] = r[k]; return o; }
function byId_(t) { const m = {}; live_(t).forEach((r) => { m[r.id] = r; }); return m; }
function newId_(t) {
  const pre = { 현장: 'S', 거래처: 'P', 계약: 'C', 점검: 'I', 청구: 'B', 연락: 'T', 문서: 'D', 자료요청: 'Q', 메일: 'M', 검토: 'R', 일지: 'J', 택배: 'K', 종합: 'F', 외주: 'O', 단가: 'U' }[t] || 'X';
  return pre + 'W' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
}
let SET_ = null; // 한 번 실행하는 동안 설정 시트를 한 번만 읽음
function setting_(k) {
  if (!SET_) {
    SET_ = {}; const sh = db_().getSheetByName('설정'); if (!sh) return '';
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) { const kk = String(v[i][0]); if (!(kk in SET_)) SET_[kk] = String(v[i][1]); }
  }
  return k in SET_ ? SET_[k] : '';
}
function setSetting_(k, val, desc) {
  SET_ = null;
  const sh = db_().getSheetByName('설정'), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === k) { sh.getRange(i + 1, 2).setValue(String(val)); return; }
  sh.appendRow([k, String(val), desc || '']);
}
function log_(who, t, id, act, before, after) {
  bumpVer_();
  if (t !== '설정') markDirty_();
  const pick = (r) => (r ? JSON.stringify(clean_(r)).slice(0, 45000) : '');
  db_().getSheetByName('변경이력').appendRow([fmtT_(new Date()), who ? who.user.email : '시스템', t, id, act, pick(before), pick(after)]);
}
function history_(t, id) {
  const v = db_().getSheetByName('변경이력').getDataRange().getValues();
  const out = [];
  for (let i = v.length - 1; i >= 1 && out.length < 50; i--) if (String(v[i][2]) === t && String(v[i][3]) === id) out.push({ 시각: norm_('수정일', v[i][0]), 사용자: v[i][1], 동작: v[i][4], 전: v[i][5], 후: v[i][6] });
  return out;
}

/** 한 줄 등록·수정·삭제: 권한 → 원본 잠금 → 버전 충돌 → 중복 → 기록 */
function save_(who, t, row, ver, isDelete) {
  if (!SCHEMA[t] || t === '문서' && !isDelete) throw new Error('이 표는 여기서 고칠 수 없습니다.');
  const needRole = isDelete && t !== '연락' ? '관리자' : WRITE_ROLE[t];
  need_(who, needRole);
  if (FROM_EXCEL.indexOf(t) >= 0 && (setting_('원본') || '엑셀') === '엑셀') throw new Error('지금은 엑셀이 원본이라 웹에서 고칠 수 없습니다. 엑셀에서 고친 뒤 「가져오기」를 하세요. (시스템관리 › 시스템 설정에서 원본을 웹으로 바꾸면 고칠 수 있습니다)');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    TBL_ = {};
    ensureCols_(t);
    const cols = SCHEMA[t].concat(META), sh = db_().getSheetByName(t) || ensureSheet_(db_(), t, SCHEMA[t].concat(META));
    const cur = row.id ? rows_(t).find((r) => r.id === row.id) : null;
    if (row.id && !cur) throw new Error(LABEL[t] + '을(를) 찾지 못했습니다. 새로고침해 보세요.');
    if (cur && Number(ver) !== cur.버전) return { ok: false, code: 'conflict', msg: '다른 사람이 먼저 수정했습니다. 바뀐 내용을 확인한 뒤 다시 저장하세요.', current: clean_(cur) };
    const next = Object.assign({}, cur ? clean_(cur) : {});
    if (!isDelete) SCHEMA[t].forEach((c) => { if (c !== 'id' && c in row) next[c] = NUMCOLS.indexOf(c) >= 0 ? num_(row[c]) : String(row[c] == null ? '' : row[c]).trim(); });
    if (!cur) { next.id = newId_(t); if (t === '연락' || t === '일지') next.작성자 = who.user.email; }
    if (t === '종합' && !isDelete) { // 상태 맞추기: 제출일이 있으면 제출 완료, 작성일만 있으면 최소 작성 중
      if (next.제출일) next.상태 = '제출 완료'; else if (next.상태 === '제출 완료') next.상태 = next.작성일 ? '작성 중' : '미작성';
      if (!next.상태) next.상태 = next.작성일 ? '작성 중' : '미작성';
      if (next.작성일 && next.상태 === '미작성') next.상태 = '작성 중';
      const fc = byId_('계약')[next.계약id] || {}, fp = byId_('거래처');
      if (next.계약id) next.현장id = fc.현장id || next.현장id || '';
      next.차수 = 1; // 종합보고서는 계약마다 1회
      next.제출처 = (fp[fc.시공사id] || fp[fc.발주처id] || {}).거래처명 || ''; // 제출처 = 그 계약의 시공사 (고정)
      if (next.계약id && live_('종합').some((y) => y.계약id === next.계약id && y.id !== next.id)) return { ok: false, code: 'invalid', msg: '이 계약은 이미 종합보고서가 등록되어 있습니다. 목록에서 그 건을 여세요.' };
      if (!cur && !next.요청자) next.요청자 = '';
    }
    if ((t === '외주' || t === '단가') && !isDelete) { // 엑셀에 적을 글자 칸: 고른 계약·거래처 이름으로 맞춤
      const oc = byId_('계약')[next.계약id];
      if (oc) { if (t === '외주') next.현장명 = oc.계약명; else next.계약명 = oc.계약명; }
      if (t === '외주' && next.업체id) { const op = byId_('거래처')[next.업체id]; if (op) next.업체명 = op.거래처명; }
      if (t === '단가') { next.단가 = String(next.단가 == null ? '' : next.단가).replace(/[,\s원]/g, '') === '' ? '' : String(Math.round(num_(next.단가))); next.횟수 = Math.max(0, Math.round(num_(next.횟수))); }
    }
    const err = validate_(t, next, cur);
    if (err) return { ok: false, code: 'invalid', msg: err };
    if (t === '계약' && !isDelete) {
      next.부가세 = next.부가세 || Math.round(next.공급가 * 0.1); next.계약금 = next.계약금 || next.공급가 + next.부가세;
      if (!next.연도) next.연도 = +(String(next.계약일 || today_()).slice(0, 4));
      if (!next.No) next.No = live_('계약').filter((c) => c.연도 === next.연도).reduce((mx, c) => Math.max(mx, c.No || 0), 0) + 1; // 엑셀 용역코드(연도-번호)용
      if (!next.상태) next.상태 = '진행중';
    }
    next.버전 = (cur ? cur.버전 : 0) + 1; next.수정일 = fmtT_(new Date()); next.수정자 = who.user.email; next.삭제 = isDelete ? 'Y' : (cur ? cur.삭제 : '');
    const vals = [cols.map((c) => (next[c] == null ? '' : next[c]))];
    if (cur) sh.getRange(cur._row, 1, 1, cols.length).setValues(vals); else sh.appendRow(vals[0]);
    log_(who, t, next.id, isDelete ? '삭제' : cur ? '수정' : '등록', cur, next);
    delete TBL_[t];
    if (['계약', '점검', '청구', '검토'].indexOf(t) >= 0) derive_(true);
    if (t === '계약' && !isDelete && next.상태지정 && next.상태 !== next.상태지정) { setSilent_('계약', next.id, { 상태: next.상태지정 }); next.상태 = next.상태지정; } // 웹이 원본일 때만 계산 (엑셀 수식과 같게)
    return { row: next };
  } finally { lock.releaseLock(); }
}
function validate_(t, r, cur) {
  const others = live_(t).filter((x) => !cur || x.id !== cur.id);
  if (t === '현장') { if (!r.현장명) return '현장명을 입력하세요.'; if (others.some((x) => key_(x.현장명) === key_(r.현장명))) return '같은 이름의 현장이 이미 있습니다.'; }
  if (t === '거래처') {
    if (!r.거래처명) return '거래처명을 입력하세요.';
    const biz = digits_(r.사업자번호);
    if (biz && others.some((x) => digits_(x.사업자번호) === biz)) return '같은 사업자번호의 거래처가 이미 있습니다.';
  }
  if (t === '계약') {
    if (!r.계약명 || !r.현장id) return '현장과 계약명을 입력하세요.';
    if (others.some((x) => x.현장id === r.현장id && key_(x.계약명) === key_(r.계약명) && x.연도 === r.연도)) return '같은 현장·계약명·연도의 계약이 이미 있습니다.';
  }
  if (t === '점검' && (!r.계약id || !r.항목)) return '계약과 점검항목을 입력하세요.';
  if (t === '점검' && r.작성완료일 && !r.점검일) return '보고서 작성 완료일은 점검일을 입력한 뒤에 넣을 수 있습니다.';
  if (t === '청구' && (!r.계약id || !(r.발행일 || r.입금일 || num_(r.계약외) || num_(r.할증)))) return '계약과 발행일(또는 입금일)을 입력하세요. (날짜 없는 줄은 계약 외 금액·할증분이 있을 때만 저장)';
  if (t === '외주' && !(r.업체명 || r.업체id)) return '외주업체를 입력하세요.';
  if (t === '외주' && !(r.계약id || r.현장명)) return '용역(현장)을 고르세요.';
  if (t === '외주' && r.지급금액 && !r.청구금액) return '청구금액을 먼저 입력하세요.';
  if (t === '단가' && (!r.계약id || !r.항목)) return '계약과 점검항목(공정)을 입력하세요.';
  if (t === '단가' && others.some((x) => x.계약id === r.계약id && key_(x.항목) === key_(r.항목))) return '이 계약에 같은 점검항목 단가가 이미 있습니다. 그 줄을 고치세요.';
  if (t === '단가' && r.단가 !== '' && num_(r.단가) < 0) return '단가는 0 이상으로 입력하세요.';
  if (t === '연락' && (!r.현장id || !r.연락일)) return '현장과 연락일을 입력하세요.';
  if (t === '검토' && (!r.계약id || !r.항목 || !r.요청일)) return '계약, 점검항목, 검토요청일을 입력하세요.';
  if (t === '종합' && (!r.계약id || !r.요청일)) return '대상 계약(현장)과 요청일을 입력하세요.';
  if (t === '종합' && r.제출일 && r.작성일 && r.제출일 < r.작성일) return '제출일이 작성일보다 빠릅니다.';
  if (t === '일지' && (!r.날짜 || !r.내용)) return '날짜와 업무 내용을 입력하세요.';
  if (t === '택배' && (!r.현장id || !r.발송일)) return '현장과 발송일을 입력하세요.';
  return '';
}

// ===================== 화면용 계산 =====================
function today_() { return iso_(new Date()); }
function dd_(a, b) { return Math.round((new Date(a + 'T00:00:00') - new Date(b + 'T00:00:00')) / 864e5); }
function contractMoney_(c, inv) {
  const mine = inv.filter((b) => b.계약id === c.id);
  const paid = c.이월수금 + mine.reduce((s, b) => s + b.수금액, 0);
  const wait = mine.filter((b) => !b.입금일 && b.발행일).reduce((s, b) => s + Math.max(0, b.청구액 - b.수금액), 0);
  return { 수금: paid, 대기: wait, 잔여: c.계약금 - paid, 수금률: c.계약금 ? paid / c.계약금 : 0, 청구: mine.reduce((s, b) => s + b.청구액, 0) };
}
function stageOf_(x) {
  const s = x.단계 || (x.발송일 ? '완료' : x.점검일 ? '보고서 작성중' : x.예정일 ? '점검 예정' : '점검 전');
  // 보고서 작성 완료는 작성만 끝난 상태 — 검토·발송 단계로 넘기지 않음 (검토·발송은 실제 기록이 있을 때만)
  return s === '보고서 작성중' && x.작성완료일 ? '보고서 작성 완료' : s;
}
function isWriting_(st) { return st === '보고서 작성중' || st === '보고서 작성 완료'; } // 아직 제출 전 (보고서 기한이 살아 있는 상태)
function contractList_() {
  const S = byId_('현장'), Pn = byId_('거래처'), inv = live_('청구'), insp = live_('점검');
  return live_('계약').map((c) => {
    const m = contractMoney_(c, inv), items = insp.filter((x) => x.계약id === c.id);
    const done = c.상태 === '완료' ? items.length : items.filter((x) => x.발송일 || stageOf_(x) === '완료' || stageOf_(x) === '입금 대기').length;
    return { id: c.id, 연도: c.연도, 현장id: c.현장id, 현장명: (S[c.현장id] || {}).현장명 || c.계약명, 계약명: c.계약명, 계약일: c.계약일, 상태: c.상태,
      발주처: (Pn[c.발주처id] || {}).거래처명 || '', 시공사: (Pn[c.시공사id] || Pn[c.발주처id] || {}).거래처명 || '', 공급가: c.공급가, 부가세: c.부가세, 계약금: c.계약금,
      수금: m.수금, 대기: m.대기, 잔여: m.잔여, 수금률: m.수금률, 청구: m.청구, 점검: items.length, 제출: done, 공종: works_(items), 버전: c.버전,
      입금비고: c.입금비고 || '', 기존수금: c.기존수금 || 0, 기존미입금: c.기존미입금 || 0, 미수금: c.미수금 || 0, 종합보고서구분: c.종합보고서구분 || '', 종합보고서금액: c.종합보고서금액 || 0,
      이월수금: c.이월수금 || 0, 상태지정: c.상태지정 || '', VAT기준: c.VAT기준 || '', 완료일: c.완료일 || '' };
  }).sort((a, b) => b.연도 - a.연도 || String(b.id).localeCompare(String(a.id)));
}
function works_(items) {
  const w = [];
  items.forEach((x) => { const b = split_(x.항목)[0]; if (b && NON_TARGET.indexOf(b) < 0 && !/회차점검$/.test(b) && w.indexOf(b) < 0) w.push(b); });
  return w;
}
function sitesList_() {
  const T = today_(), contracts = contractList_(), insp = live_('점검'), cts = live_('연락');
  const byC = {}; insp.forEach((x) => { (byC[x.계약id] = byC[x.계약id] || []).push(x); });
  return live_('현장').map((s) => {
    const cs = contracts.filter((c) => c.현장id === s.id);
    const items = [].concat(...cs.map((c) => byC[c.id] || []));
    const active = cs.some((c) => c.상태 === '진행중');
    const writing = items.filter((x) => isWriting_(stageOf_(x)) && x.기한).map((x) => x.기한).sort();
    const next = items.find((x) => ['점검 예정', '점검 전'].indexOf(stageOf_(x)) >= 0);
    const last = cts.filter((x) => x.현장id === s.id).sort((a, b) => String(b.연락일).localeCompare(String(a.연락일)))[0];
    const money = cs.reduce((a, c) => ({ 계약금: a.계약금 + c.계약금, 수금: a.수금 + c.수금, 대기: a.대기 + c.대기 }), { 계약금: 0, 수금: 0, 대기: 0 });
    const latest = cs[0] || {};
    return { id: s.id, 현장명: s.현장명, 관할구: s.관할구, 주소: s.주소, 규모: s.규모, 진행: active, 계약수: cs.length, 시공사: latest.시공사 || '', 발주처: latest.발주처 || '',
      점검: items.length, 완료: items.filter((x) => stageOf_(x) === '완료').length + (active ? 0 : 0), 작성중: items.filter((x) => stageOf_(x) === '보고서 작성중').length, 작성완료: items.filter((x) => stageOf_(x) === '보고서 작성 완료').length,
      기한: writing[0] || '', 남은날: writing[0] ? dd_(writing[0], T) : null, 다음: next ? next.항목 : '', 다음시기: next ? next.시기 : '',
      계약금: money.계약금, 수금: money.수금, 대기: money.대기, 수금률: money.계약금 ? money.수금 / money.계약금 : 0, 최근연락: last ? last.연락일 : '', 버전: s.버전,
      단계목록: items.map((x) => [x.항목, stageOf_(x)]) };
  }).sort((a, b) => (b.진행 - a.진행) || ((a.남은날 == null ? 9999 : a.남은날) - (b.남은날 == null ? 9999 : b.남은날)) || a.현장명.localeCompare(b.현장명));
}
function siteDetail_(id) {
  const s = live_('현장').find((x) => x.id === id);
  if (!s) return { site: null };
  const cs = contractList_().filter((c) => c.현장id === id), ids = cs.map((c) => c.id);
  const T = today_();
  // 점검별 내부 검토 이력 (보고서 검토 관리에서): 마지막 요청일 · 검토 완료 회신일
  const RVI = {}; live_('검토').filter((v) => ids.indexOf(v.계약id) >= 0).forEach((v) => { const k = v.계약id + '|' + key_(v.항목), o = RVI[k] = RVI[k] || { 요청: '', 회신: '' }; if (v.요청일 && String(v.요청일) > o.요청) o.요청 = String(v.요청일); if (v.상태 === '검토완료' && v.회신일 && String(v.회신일) > o.회신) o.회신 = String(v.회신일); });
  const items = live_('점검').filter((x) => ids.indexOf(x.계약id) >= 0).map((x) => { const rv = RVI[x.계약id + '|' + key_(x.항목)] || {}; return Object.assign(clean_(x), { 단계: stageOf_(x), 남은날: x.기한 && isWriting_(stageOf_(x)) ? dd_(x.기한, T) : null, 검토요청일: rv.요청 || '', 검토회신일: rv.회신 || '' }); });
  return { site: clean_(s), contracts: cs, items, invoices: live_('청구').filter((b) => ids.indexOf(b.계약id) >= 0).map(clean_).sort((a, b) => String(b.발행일).localeCompare(String(a.발행일))),
    contacts: live_('연락').filter((x) => x.현장id === id).map(clean_).sort((a, b) => String(b.연락일).localeCompare(String(a.연락일))),
    docs: live_('문서').filter((d) => d.현장id === id || ids.indexOf(d.계약id) >= 0).map(clean_),
    requests: requestsList_().filter((q) => q.현장id === id), mails: mailsList_().filter((m) => m.현장id === id) };
}
function contractDetail_(id) {
  const c = live_('계약').find((x) => x.id === id);
  if (!c) return { contract: null };
  const Pn = byId_('거래처'), inv = live_('청구'), m = contractMoney_(c, inv);
  return { contract: Object.assign(clean_(c), m, { 현장명: (byId_('현장')[c.현장id] || {}).현장명 || '', 발주처: clean_(Pn[c.발주처id] || {}), 실제입금처: clean_(Pn[c.실제입금처id] || {}), 시공사: clean_(Pn[c.시공사id] || {}) }),
    items: live_('점검').filter((x) => x.계약id === id).map((x) => Object.assign(clean_(x), { 단계: stageOf_(x) })),
    invoices: inv.filter((b) => b.계약id === id).map(clean_).sort((a, b) => String(a.발행일).localeCompare(String(b.발행일))),
    docs: live_('문서').filter((d) => d.계약id === id).map(clean_), prices: live_('단가').filter((u) => u.계약id === id).map(clean_),
    final: (live_('종합').find((f) => f.계약id === id) || null) };
}
// ---- 회차별 단가 목록 (현장별 · 공정별 조회)
function pricesList_() {
  const C = byId_('계약'), S = byId_('현장'), insp = live_('점검');
  const cnt = {}; insp.forEach((x) => { const k = x.계약id + '|' + key_(procKey_(x.항목)); (cnt[k] = cnt[k] || []).push(x.항목); });
  const rows = live_('단가').map((u) => { const c = C[u.계약id] || {}, live = cnt[u.계약id + '|' + key_(procKey_(u.항목))] || [];
    return Object.assign(clean_(u), { 계약명: c.계약명 || u.계약명, 현장id: c.현장id || '', 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || u.계약명, 연도: c.연도 || '', 상태: c.상태 || '', VAT: c.VAT기준 === '별도' ? '별도 (공급가)' : '포함', 단가: u.단가 === '' ? '' : num_(u.단가), 금액: u.단가 !== '' && u.횟수 ? num_(u.단가) * u.횟수 : 0, 점검수: live.length, 점검회차: live.join(', ') }); });
  return { prices: rows };
}
// 점검항목 → 공정 이름 (뒤의 「N차」를 뗌: 지보공2차 → 지보공) — 엑셀 ProcKey 와 같음
function procKey_(it) { let s = String(it || '').trim(); if (/차$/.test(s)) { s = s.slice(0, -1).replace(/\d+$/, '').trim(); } return s || String(it || '').trim(); }
// ---- 외주관리 목록 (엑셀 L~P 계산과 같게)
function outsourceList_() {
  const C = byId_('계약'), S = byId_('현장'), Pn = byId_('거래처');
  const rows = live_('외주').map((o) => { const c = C[o.계약id] || {}, F = o.청구금액, G = o.지급금액, sup = F ? round_(F / 1.1) : 0;
    return Object.assign(clean_(o), { 업체: o.업체명 || (Pn[o.업체id] || {}).거래처명 || '', 계약명: c.계약명 || o.현장명, 현장id: c.현장id || '', 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || o.현장명,
      미지급: F ? Math.max(0, F - G) : 0, 공급가: sup, 부가세: F ? F - sup : 0, 상태: !F ? '금액 없음' : G >= F ? '지급 완료' : G > 0 ? '일부 지급' : '지급 대기', 연도: String(o.계산서일 || '').slice(0, 4) }); })
    .sort((a, b) => String(b.계산서일).localeCompare(String(a.계산서일)));
  return { outsource: rows };
}
function schedule_(from, to) {
  const C = byId_('계약'), S = byId_('현장'), ev = [];
  const nm = (cid) => { const c = C[cid] || {}; return { 현장: (S[c.현장id] || {}).현장명 || c.계약명 || '', 현장id: c.현장id || '', 종료: c.상태 && c.상태 !== '진행중' }; };
  live_('점검').forEach((x) => {
    const n = nm(x.계약id), st = stageOf_(x), done = st === '완료' || n.종료;
    if (x.점검일) ev.push(Object.assign({ d: x.점검일, k: '점검', t: x.항목, id: x.id, done }, n));
    else if (x.예정일) ev.push(Object.assign({ d: x.예정일, k: '예정', t: x.항목, id: x.id, done }, n));
    if (isWriting_(st) && x.기한) ev.push(Object.assign({ d: x.기한, k: '기한', t: x.항목, id: x.id }, n));
  });
  const last = {};
  live_('연락').forEach((c) => { if (!last[c.현장id] || c.연락일 > last[c.현장id].연락일) last[c.현장id] = c; });
  Object.values(last).forEach((c) => { if (c.다음연락) ev.push({ d: c.다음연락, k: '연락', 현장: (S[c.현장id] || {}).현장명 || '', 현장id: c.현장id, t: c.담당자 }); });
  return ev.filter((e) => (!from || e.d >= from) && (!to || e.d <= to)).sort((a, b) => a.d.localeCompare(b.d));
}
function requestsList_() {
  const S = byId_('현장'), T = today_();
  return live_('자료요청').map((q) => { const done = !!q.수신일 || /완료/.test(q.상태); return Object.assign(clean_(q), { 현장명: (S[q.현장id] || {}).현장명 || '', 받음: done, 경과: done ? (q.요청일 && q.수신일 ? dd_(q.수신일, q.요청일) : null) : (q.요청일 ? dd_(T, q.요청일) : null), 기한남음: !done && q.기한 ? dd_(q.기한, T) : null }); })
    .sort((a, b) => String(b.요청일).localeCompare(String(a.요청일)) || num_(b.NO) - num_(a.NO));
}
function mailsList_() {
  const S = byId_('현장'), T = today_();
  return live_('메일').map((m) => { const wait = /예/.test(m.회신필요) && !m.회신일 && !/받음|완료/.test(m.회신상태); return Object.assign(clean_(m), { 현장명: (S[m.현장id] || {}).현장명 || '', 회신대기: wait, 경과: wait && m.보낸날 ? dd_(T, m.보낸날) : null }); })
    .sort((a, b) => String(b.보낸날).localeCompare(String(a.보낸날)) || num_(b.NO) - num_(a.NO));
}
function contactsList_() {
  const S = byId_('현장');
  return live_('연락').map((c) => Object.assign(clean_(c), { 현장명: (S[c.현장id] || {}).현장명 || '' })).sort((a, b) => String(b.연락일).localeCompare(String(a.연락일)));
}
function money_() {
  const T = today_(), C = byId_('계약'), S = byId_('현장'), Pn = byId_('거래처');
  const contracts = contractList_();
  const invoices = live_('청구').map((b) => {
    const c = C[b.계약id] || {};
    const part = b.입금일 ? 0 : partsSum_(b.분할입금), left = Math.max(0, b.청구액 - b.수금액 - part);
    return Object.assign(clean_(b), { 부분입금: part, 분할: partsOf_(b.분할입금), 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || '', 현장id: c.현장id || '', 발주처: (Pn[c.발주처id] || {}).거래처명 || b.입금자, 미입금: b.입금일 || !b.발행일 ? 0 : left, 경과: b.발행일 ? dd_(T, b.발행일) : null, 연도: String(b.발행일).slice(0, 4) });
  }).sort((a, b) => String(b.발행일).localeCompare(String(a.발행일)));
  const monthly = {};
  invoices.forEach((b) => { const k = String(b.발행일).slice(0, 7); if (!k) return; const m = monthly[k] = monthly[k] || { 청구: 0, 수금: 0, 건수: 0 }; m.청구 += b.청구액; m.수금 += b.수금액; m.건수++; });
  // 월별 실적 (엑셀 진행현황보고 「월별 실적」과 같은 기준 · 완료된 용역 포함): 점검 건수(점검일) · 현장 방문(같은 날 같은 용역은 1회) · 보고서 발송(발송일) · 청구(발행일) · 입금(입금일 기준 수금액)
  const perf = {}, P0 = (k) => (perf[k] = perf[k] || { 점검: 0, 방문: 0, 일수: 0, 발송: 0, 청구: 0, 입금: 0, 계약외: 0 }), visit = {}, dayz = {};
  live_('점검').forEach((x) => {
    if (x.점검일) { const k = String(x.점검일).slice(0, 7); P0(k).점검++; const vk = x.점검일 + '|' + x.계약id; if (!visit[vk]) { visit[vk] = 1; perf[k].방문++; } const dk = String(x.점검일).slice(0, 10); if (!dayz[dk]) { dayz[dk] = 1; perf[k].일수++; } } // 일수: 점검한 날짜 수 (같은 날 여러 현장도 1일)
    if (/^\d{4}-\d\d-\d\d/.test(String(x.발송일 || ''))) P0(String(x.발송일).slice(0, 7)).발송++;
  });
  live_('청구').forEach((b) => { if (b.발행일) P0(String(b.발행일).slice(0, 7)).청구 += b.청구액; if (b.입금일) P0(String(b.입금일).slice(0, 7)).입금 += b.수금액; });
  const years = {};
  contracts.forEach((c) => { const y = years[c.연도] = years[c.연도] || { 건수: 0, 공급가: 0, 계약금: 0, 수금: 0, 진행: 0, 완료: 0 }; y.건수++; y.공급가 += c.공급가; y.계약금 += c.계약금; y.수금 += c.수금; if (c.상태 === '진행중') y.진행++; else y.완료++; });
  return { contracts, invoices, monthly, years, perf };
}
function partnersList_() {
  const cs = contractList_(), raw = live_('계약');
  return live_('거래처').map((p) => {
    const mine = raw.filter((c) => c.발주처id === p.id || c.시공사id === p.id || c.실제입금처id === p.id).map((c) => c.id);
    const list = cs.filter((c) => mine.indexOf(c.id) >= 0);
    return Object.assign(clean_(p), { 계약수: list.length, 계약금: list.reduce((s, c) => s + c.계약금, 0), 공급가: list.reduce((s, c) => s + c.공급가, 0), 대기: list.reduce((s, c) => s + c.대기, 0), 진행: list.filter((c) => c.상태 === '진행중').length });
  }).sort((a, b) => b.계약금 - a.계약금);
}
function partnerDetail_(id) {
  const p = live_('거래처').find((x) => x.id === id);
  if (!p) return { partner: null };
  const raw = live_('계약').filter((c) => c.발주처id === id || c.시공사id === id || c.실제입금처id === id).map((c) => c.id);
  const contracts = contractList_().filter((c) => raw.indexOf(c.id) >= 0);
  return { partner: clean_(p), contracts, invoices: live_('청구').filter((b) => raw.indexOf(b.계약id) >= 0).map(clean_).sort((a, b) => String(b.발행일).localeCompare(String(a.발행일))),
    docs: live_('문서').filter((d) => d.거래처id === id).map(clean_) };
}
function reviewsList_() {
  const T = today_(), C = byId_('계약'), S = byId_('현장'), Pn = byId_('거래처');
  return live_('검토').map((r) => {
    const c = C[r.계약id] || {}, wait = r.상태 === '검토중' || !r.상태, resend = r.상태 === '수정요청' && !r.재발송일;
    return Object.assign(clean_(r), { 계약명: c.계약명 || '', 현장id: c.현장id || '', 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || '', 시공사: (Pn[c.시공사id] || Pn[c.발주처id] || {}).거래처명 || '',
      소요: r.요청일 && r.회신일 ? dd_(r.회신일, r.요청일) : null, 경과: wait && r.요청일 ? dd_(T, r.요청일) : resend && r.회신일 ? dd_(T, r.회신일) : null, 재발송대기: resend, 대기: wait || resend });
  }).sort((a, b) => String(b.요청일).localeCompare(String(a.요청일)) || (+b.회차 || 0) - (+a.회차 || 0));
}
function parcelsList_() {
  const C = byId_('계약'), S = byId_('현장'), Pn = byId_('거래처');
  const list = live_('택배').map((k) => Object.assign(clean_(k), { 현장명: (S[k.현장id] || {}).현장명 || '' })).sort((a, b) => String(b.발송일).localeCompare(String(a.발송일)));
  const lastTo = {}; list.slice().reverse().forEach((k) => { if (k.받는사람 || k.주소) lastTo[k.현장id] = { 받는사람: k.받는사람, 주소: k.주소 }; });
  // 보낼 것: 점검 단계가 「발송 대기」(검토 완료 · 입금 확인, 거래처 발송 전)
  const ready = live_('점검').filter((x) => stageOf_(x) === '발송 대기').map((x) => { const c = C[x.계약id] || {}; return { 점검id: x.id, 계약id: x.계약id, 현장id: c.현장id, 현장명: (S[c.현장id] || {}).현장명 || c.계약명 || '', 시공사: (Pn[c.시공사id] || Pn[c.발주처id] || {}).거래처명 || '', 항목: x.항목, 점검일: x.점검일, 버전: x.버전, 예전: lastTo[c.현장id] || null }; });
  return { parcels: list, ready };
}
function report_() {
  const T = today_(), Y = +T.slice(0, 4), C = live_('계약'), bills = live_('청구'), Pn = byId_('거래처');
  const nk = (c) => key_(c.계약명), byName = {}; C.forEach((c) => { byName[c.id] = nk(c); });
  const agg = {}; bills.forEach((b) => { const k = byName[b.계약id]; if (!k) return; const a = (agg[k] = agg[k] || { 수금: 0, 올해: 0, 대기: 0, 계약외: 0, 할증: 0 }); a.수금 += b.수금액; if (b.입금일 && +String(b.입금일).slice(0, 4) === Y) a.올해 += b.수금액; if (b.청구액 > 0 && !b.입금일) a.대기 += b.청구액; a.계약외 += b.계약외; a.할증 += b.할증; });
  const rows = C.map((c) => {
    const a = agg[nk(c)] || { 수금: 0, 올해: 0, 대기: 0, 계약외: 0, 할증: 0 }, paid = c.이월수금 + a.수금, goal = c.계약금 + (c.종합보고서구분 === '별도' ? c.종합보고서금액 : 0);
    return { id: c.id, 현장id: c.현장id, 연도: c.연도, No: c.No, 계약명: c.계약명, 입금자: (Pn[c.발주처id] || {}).거래처명 || '', 시공사: (Pn[c.시공사id] || Pn[c.발주처id] || {}).거래처명 || '', 공급가: c.계약금 ? Math.round(c.계약금 / 1.1) : 0, 부가세: c.계약금 ? c.계약금 - Math.round(c.계약금 / 1.1) : 0, 계약금: c.계약금,
      대기: a.대기, 수금: paid, 올해수금: a.올해, 이전수금: paid - a.올해, 상태: c.상태 || '진행중', 잔여: c.계약금 ? goal - paid : 0, 종합보고서: c.종합보고서금액, 기성율: goal ? paid / goal : 0, 미수금: c.미수금, 비고: c.비고, 계약외: a.계약외 };
  });
  return { year: Y, rows };
}
function dash_() {
  const T = today_(), sites = sitesList_(), m = money_();
  const insp = live_('점검'), C = byId_('계약'), S = byId_('현장');
  const Pn = byId_('거래처'), bname = (c) => ((Pn[c.시공사id] || Pn[c.발주처id] || {}).거래처명 || '');
  const builder = {}; live_('계약').slice().sort((a, b) => a.연도 - b.연도).forEach((c) => { const n = bname(c); if (n) builder[c.현장id] = n; }); // 현장 → 시공사 (가장 최근 계약)
  const writing = insp.filter((x) => isWriting_(stageOf_(x))).map((x) => { const c = C[x.계약id] || {}; return { 작성완료일: x.작성완료일 || '', 현장: (S[c.현장id] || {}).현장명 || '', 현장id: c.현장id, 시공사: bname(c), 항목: x.항목, 점검일: x.점검일, 기한: x.기한, 남은날: x.기한 ? dd_(x.기한, T) : null, 비고: x.비고 }; })
    .sort((a, b) => (a.남은날 == null ? 9999 : a.남은날) - (b.남은날 == null ? 9999 : b.남은날));
  const mon = new Date(T + 'T00:00:00'); mon.setDate(mon.getDate() - ((mon.getDay() + 6) % 7));
  const sun = new Date(mon.getTime() + 6 * 864e5);
  const lastC = {};
  live_('연락').forEach((c) => { if (!lastC[c.현장id] || c.연락일 > lastC[c.현장id].연락일) lastC[c.현장id] = c; });
  const todo = Object.values(lastC).filter((c) => c.다음연락 && c.다음연락 <= T && sites.some((s) => s.id === c.현장id && s.진행)).map((c) => ({ k: '연락', 현장: (S[c.현장id] || {}).현장명, 현장id: c.현장id, t: `다음 연락 ${c.다음연락}${c.담당자 ? ' · ' + c.담당자 : ''}` }));
  const waitInv = m.invoices.filter((b) => b.미입금 > 0);
  const active = sites.filter((s) => s.진행);
  return { today: T, week: { from: iso_(mon), to: iso_(sun), events: schedule_(iso_(mon), iso_(sun)) },
    kpi: { 진행현장: active.length, 완료현장: sites.length - active.length, 작성중: writing.filter((w) => !w.작성완료일).length, 작성완료: writing.filter((w) => w.작성완료일).length, 기한14: writing.filter((w) => w.남은날 != null && w.남은날 <= 14).length, 기한초과: writing.filter((w) => w.남은날 != null && w.남은날 < 0).length,
      입금대기: waitInv.reduce((s, b) => s + b.미입금, 0), 입금대기건: waitInv.length, 계약진행: m.contracts.filter((c) => c.상태 === '진행중').length, 계약완료: m.contracts.filter((c) => c.상태 !== '진행중').length },
    final: (() => { try { const f = finals_(); return Object.assign({}, f.kpi, { open: f.list.filter((r) => r.상태 !== '제출 완료').slice(0, 40).map((r) => Object.assign({}, r, { 검토: undefined, 파일: undefined })) }); } catch (e) { return null; } })(),
    writing: writing.slice(0, 12), waitInv: waitInv.sort((a, b) => b.경과 - a.경과).slice(0, 10), contacts: contactsList_().slice(0, 6), todo, builder,
    reviews: reviewsList_().filter((r) => r.대기), jobs: live_('일지').filter((j) => (j.상태 === '진행중' || j.상태 === '이월') && j.날짜 < T).sort((a, b) => String(b.날짜).localeCompare(String(a.날짜))).slice(0, 10).map(clean_),
    years: m.years, recent: recentChanges_(8) };
}
function recentChanges_(n) {
  const sh = db_().getSheetByName('변경이력'), last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(Math.max(2, last - 60), 1, Math.min(61, last - 1), 5).getValues().reverse();
  return v.filter((r) => r[4] !== '가져오기' && r[4] !== '복구').slice(0, n).map((r) => ({ 시각: norm_('수정일', r[0]), 사용자: r[1], 표: r[2], id: r[3], 동작: r[4] }));
}
function counts_() { const o = {}; FROM_EXCEL.concat(['문서']).forEach((t) => { o[t] = live_(t).length; }); const cs = live_('계약'); o.진행계약 = cs.filter((c) => c.상태 === '진행중').length; o.완료계약 = cs.length - o.진행계약; return o; }
function lookups_() {
  return { sites: live_('현장').map((s) => ({ id: s.id, n: s.현장명 })).sort((a, b) => a.n.localeCompare(b.n)),
    partners: live_('거래처').map((p) => ({ id: p.id, n: p.거래처명, biz: p.사업자번호, 대표자: p.대표자, 주소: p.주소, 업태: p.업태, 종목: p.종목 })).sort((a, b) => a.n.localeCompare(b.n)),
    contracts: contractList_().map((c) => ({ id: c.id, n: c.계약명, 연도: c.연도, 현장id: c.현장id, 상태: c.상태, 시공사: c.시공사 })),
    items: live_('점검').reduce((o, x) => { (o[x.계약id] = o[x.계약id] || []).push(x.항목); return o; }, {}),
    company: company_() };
}
function company_() {
  const get = (k, d) => setting_(k) || d;
  let seal = '';
  const sid = setting_('회사_직인');
  if (sid) { try { const b = DriveApp.getFileById(sid).getBlob(); seal = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) {} }
  return { 상호: get('회사_상호', '㈜유토엔지니어링건축사사무소'), 대표: get('회사_대표', ''), 사업자번호: get('회사_사업자번호', ''), 주소: get('회사_주소', ''), 업태: get('회사_업태', ''), 종목: get('회사_종목', ''), 전화: get('회사_전화', ''), 팩스: get('회사_팩스', ''), 기타사항: get('회사_거래명세서_기타사항', ''), 직인: seal };
}
/** 거래명세서 직인 이미지: 공개 저장소(GitHub)에 두지 않고 회사 Drive 에 보관, 로그인한 사람에게만 보냄 */
function saveSeal_(who, a) {
  const m = /^data:(image\/(png|jpeg|gif|webp));base64,(.+)$/.exec(String(a.data || ''));
  const old = setting_('회사_직인');
  if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
  if (!m) { setSetting_('회사_직인', ''); log_(who, '설정', '회사_직인', '삭제', null, null); return {}; }
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 400000) return { ok: false, msg: '이미지가 너무 큽니다(400KB 이하).' };
  const f = folder_(CFG.BACKUP_FOLDER).createFile(Utilities.newBlob(bytes, m[1], '직인.' + (m[2] === 'jpeg' ? 'jpg' : m[2])));
  setSetting_('회사_직인', f.getId(), '거래명세서 직인 이미지 파일 ID (시스템 설정에서 등록)');
  log_(who, '설정', '회사_직인', '수정', null, { 파일: f.getId() });
  return {};
}

// ===================== 거래명세서 (엑셀 「거래명세서」와 같은 계산) =====================
function calcStatement_(lines) {
  let sup = 0, tax = 0;
  const out = (lines || []).filter((l) => String(l.품명 || '').trim()).slice(0, 26).map((l) => {
    const q = num_(l.수량), p = num_(l.단가), vat = l.VAT === '포함' ? '포함' : '별도';
    let s = 0, t = 0;
    if (q && p) {
      if (vat === '별도') { s = q * p; t = round_(s * 0.1); } else { s = round_(q * p / 1.1); t = q * p - s; }
    }
    sup += s; tax += t;
    return { 품명: String(l.품명).trim(), 규격: String(l.규격 || ''), 단위: String(l.단위 || '회'), 수량: q, 단가: p, VAT: vat, 공급가액: s, 세액: t };
  });
  return { lines: out, 공급가: sup, 세액: tax, 합계: sup + tax };
}
function round_(x) { return Math.sign(x) * Math.round(Math.abs(x)); } // 엑셀 ROUND 와 같은 반올림(0.5 는 0 에서 먼 쪽)
function saveDoc_(who, d, ver, issue) {
  need_(who, '일반');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    TBL_ = {};
    const cur = d.id ? rows_('문서').find((x) => x.id === d.id) : null;
    if (d.id && !cur) throw new Error('문서를 찾지 못했습니다.');
    if (cur && Number(ver) !== cur.버전) return { ok: false, code: 'conflict', msg: '다른 사람이 먼저 수정했습니다. 다시 열어 확인하세요.', current: clean_(cur) };
    if (cur && cur.상태 === '발행' && !issue) throw new Error('발행된 문서는 고칠 수 없습니다. 「복사해서 새로 작성」을 쓰세요.');
    const kind = d.종류 || '거래명세서';
    const calc = kind === '거래명세서' ? calcStatement_(d.lines) : { lines: [], 공급가: num_(d.공급가), 세액: num_(d.세액), 합계: num_(d.합계) };
    const body = { lines: calc.lines, 거래처: d.거래처 || {}, 현장명: d.현장명 || '', 기타사항: d.기타사항 || '', 메모: d.메모 || '' };
    const next = Object.assign({}, cur ? clean_(cur) : {}, { 종류: kind, 계약id: d.계약id || '', 현장id: d.현장id || '', 거래처id: d.거래처id || '', 일자: d.일자 || today_(), 제목: d.제목 || '',
      내용: JSON.stringify(body), 공급가: calc.공급가, 세액: calc.세액, 합계: calc.합계 });
    if (!next.거래처id && !(body.거래처 && body.거래처.거래처명)) return { ok: false, code: 'invalid', msg: '거래처를 고르세요.' };
    if (!calc.lines.length && kind === '거래명세서') return { ok: false, code: 'invalid', msg: '품목을 한 줄 이상 입력하세요.' };
    if (!cur) { next.id = newId_('문서'); next.작성자 = who.user.email; next.상태 = '임시저장'; }
    if (issue && next.상태 !== '발행') { next.상태 = '발행'; next.번호 = nextNo_(kind, String(next.일자).slice(0, 4)); }
    next.버전 = (cur ? cur.버전 : 0) + 1; next.수정일 = fmtT_(new Date()); next.수정자 = who.user.email; next.삭제 = '';
    const cols = SCHEMA['문서'].concat(META), sh = db_().getSheetByName('문서');
    const vals = [cols.map((c) => (next[c] == null ? '' : next[c]))];
    if (cur) sh.getRange(cur._row, 1, 1, cols.length).setValues(vals); else sh.appendRow(vals[0]);
    log_(who, '문서', next.id, issue ? '발행' : cur ? '수정' : '등록', cur, next);
    TBL_ = {};
    return { doc: next };
  } finally { lock.releaseLock(); }
}
function nextNo_(kind, year) {
  const pre = { 거래명세서: 'TS', 계약서: 'CT', 견적서: 'ES', 공문: 'UTO' }[kind] || 'DOC';
  const k = '번호:' + kind + ':' + year, n = num_(setting_(k)) + 1;
  setSetting_(k, n, kind + ' ' + year + '년 마지막 번호 (자동)');
  return `${pre}-${year}-${String(n).padStart(3, '0')}`;
}

// ===================== 사용자·기기 =====================
function devices_() {
  const v = db_().getSheetByName('기기').getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ id: r[0], email: String(r[1]).toLowerCase(), name: r[2], since: norm_('수정일', r[4]), last: norm_('수정일', r[5]), exp: norm_('수정일', r[6]), status: r[7], row: i + 2, h: r[3] }));
}
function users_() {
  const v = db_().getSheetByName('사용자').getDataRange().getValues();
  const devs = devices_();
  return { users: v.slice(1).filter((r) => r[1]).map((r) => ({ 이름: r[0], 이메일: String(r[1]).toLowerCase(), 권한: r[2], 상태: r[3], 메모: r[4], 기기: devs.filter((d) => d.email === String(r[1]).toLowerCase() && d.status === '사용').length })),
    devices: devs.map((d) => ({ id: d.id, email: d.email, name: d.name, since: d.since, last: d.last, exp: d.exp, status: d.status })).reverse() };
}
function saveUser_(who, a) {
  const email = String(a.이메일 || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, msg: '이메일을 확인하세요.' };
  const role = ROLE_RANK[a.권한] == null ? '조회' : a.권한, st = a.상태 === '차단' ? '차단' : '승인';
  const u = findUser_(email), sh = db_().getSheetByName('사용자');
  if (u && u.email === who.user.email && (role !== '관리자' || st !== '승인')) return { ok: false, msg: '자기 자신의 관리자 권한은 내릴 수 없습니다.' };
  const row = [String(a.이름 || '').trim() || email, email, role, st, String(a.메모 || '')];
  if (u) sh.getRange(u.row, 1, 1, 5).setValues([row]); else sh.appendRow(row);
  devices_().filter((d) => d.email === email).forEach((d) => CacheService.getScriptCache().remove('auth:' + d.h)); // 권한·상태 바뀌면 바로 적용
  log_(who, '사용자', email, u ? '수정' : '등록', null, { 이메일: email, 권한: role, 상태: st });
  return users_();
}
function setDevice_(row, status, h) { db_().getSheetByName('기기').getRange(row, 8).setValue(status); if (h) CacheService.getScriptCache().remove('auth:' + h); }
function revoke_(who, a) {
  const d = devices_().find((x) => x.id === a.id);
  if (!d) return { ok: false, msg: '기기를 찾지 못했습니다.' };
  if (d.email !== who.user.email && !can_(who, '관리자')) return { ok: false, msg: '내 기기만 해제할 수 있습니다.' };
  setDevice_(d.row, a.block && can_(who, '관리자') ? '차단' : '해제', d.h);
  return {};
}

// ===================== 원본 전환 =====================
function setSource_(who, v) {
  if (v !== '웹' && v !== '엑셀') return { ok: false, msg: '값이 올바르지 않습니다.' };
  const prev = setting_('원본') || '엑셀';
  if (prev === v) return { source: v };
  if (v === '웹') {
    const sync = readSync_();
    if (!sync) return { ok: false, msg: '엑셀 자동 저장에 필요한 연결 정보가 없습니다. 먼저 「엑셀 가져오기」를 한 번 한 뒤 전환하세요.' };
    let f; try { f = DriveApp.getFileById(sync.fileId); } catch (e) { return { ok: false, msg: '가져왔던 엑셀 파일을 찾지 못했습니다. 「엑셀 가져오기」를 다시 하세요.' }; }
    if (f.getLastUpdated().getTime() > sync.fileMod + 60000) return { ok: false, msg: '마지막 가져오기 뒤에 엑셀이 바뀌었습니다. 「엑셀 가져오기」를 한 번 더 해서 웹과 엑셀을 맞춘 뒤 전환하세요.' };
  } else if (+(P.getProperty('EXPORT_DIRTY') || 0)) {
    try { exportExcel_('원본 되돌리기 직전'); P.setProperty('EXPORT_DIRTY', ''); } catch (e) { return { ok: false, msg: '엑셀에 아직 저장하지 못한 웹 내용이 있어 되돌리지 않았습니다: ' + e.message }; }
  }
  backup_('원본 ' + prev + '→' + v + ' 전환 직전', who.user.email);
  setSetting_('원본', v);
  P.setProperty('EXPORT_DIRTY', '');
  exportTrigger_(v === '웹');
  log_(who, '설정', '원본', '수정', { 값: prev }, { 값: v });
  return { source: v };
}

// ===================== 백업·복구 =====================
function dailyBackup() { backup_('자동', '시스템'); }
function backup_(reason, by) {
  const f = folder_(CFG.BACKUP_FOLDER);
  const name = '업무관리DB_백업_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd-HHmm') + '_' + String(reason).replace(/[\\/:*?"<>|]/g, '');
  const copy = DriveApp.getFileById(P.getProperty('DB_ID')).makeCopy(name, f);
  copy.setDescription('백업 사유: ' + reason + ' · ' + (by || ''));
  if (reason === '자동') { // 오래된 자동 백업 정리
    const autos = backups_().filter((b) => /_자동$/.test(b.name));
    autos.slice(CFG.BACKUP_KEEP).forEach((b) => DriveApp.getFileById(b.id).setTrashed(true));
  }
  return { id: copy.getId(), name };
}
function backups_() {
  const it = folder_(CFG.BACKUP_FOLDER).getFiles(), out = [];
  while (it.hasNext()) { const f = it.next(); out.push({ id: f.getId(), name: f.getName(), at: fmtT_(f.getDateCreated()), url: 'https://docs.google.com/spreadsheets/d/' + f.getId(), xlsx: 'https://docs.google.com/spreadsheets/d/' + f.getId() + '/export?format=xlsx' }); }
  return out.sort((a, b) => b.at.localeCompare(a.at));
}
function readTables_(ss) {
  const out = {};
  Object.keys(SCHEMA).forEach((t) => {
    const sh = ss.getSheetByName(t); if (!sh) { out[t] = []; return; }
    const v = sh.getDataRange().getValues(), H = v[0].map(String);
    out[t] = v.slice(1).filter((r) => r[0] !== '' && r[0] != null).map((r) => { const o = {}; H.forEach((h, j) => { if (h) o[h] = norm_(h, r[j]); }); return o; });
  });
  return out;
}
function diff_(cur, next, cols) {
  const A = {}, B = {}; cur.forEach((r) => { A[r.id] = r; }); next.forEach((r) => { B[r.id] = r; });
  const res = { 추가: 0, 변경: 0, 삭제: 0, 같음: 0, 예: [] };
  next.forEach((r) => {
    const o = A[r.id];
    if (!o) { res.추가++; if (res.예.length < 6) res.예.push({ 동작: '추가', id: r.id, 이름: label_(r) }); return; }
    const ch = cols.filter((c) => String(o[c] == null ? '' : o[c]) !== String(r[c] == null ? '' : r[c]));
    if (ch.length) { res.변경++; if (res.예.length < 6) res.예.push({ 동작: '변경', id: r.id, 이름: label_(r), 칸: ch.slice(0, 4).map((c) => `${c}: ${o[c]} → ${r[c]}`) }); } else res.같음++;
  });
  cur.forEach((r) => { if (!B[r.id] && r.삭제 !== 'Y') { res.삭제++; if (res.예.length < 6) res.예.push({ 동작: '삭제', id: r.id, 이름: label_(r) }); } });
  return res;
}
function label_(r) { return r.현장명 || r.계약명 || r.거래처명 || [r.항목, r.품명, r.연락일, r.발행일].filter(Boolean).join(' ') || r.id; }
function writeTable_(t, rows) {
  const cols = SCHEMA[t].concat(META), sh = db_().getSheetByName(t) || ensureSheet_(db_(), t, SCHEMA[t].concat(META));
  sh.clearContents();
  const data = [cols].concat(rows.map((r) => cols.map((c) => (r[c] == null ? '' : r[c]))));
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), data.length), cols.length).setNumberFormat('@');
  sh.getRange(1, 1, data.length, cols.length).setValues(data);
  TBL_ = {};
}
function restorePreview_(id) {
  const ss = SpreadsheetApp.openById(id), from = readTables_(ss), cur = readTables_(db_());
  const tables = {};
  Object.keys(SCHEMA).forEach((t) => { tables[t] = Object.assign({ 지금: cur[t].filter((r) => r.삭제 !== 'Y').length, 백업: from[t].filter((r) => r.삭제 !== 'Y').length }, diff_(cur[t], from[t], SCHEMA[t])); });
  return { name: DriveApp.getFileById(id).getName(), tables };
}
function restoreCommit_(who, id) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const name = DriveApp.getFileById(id).getName();
    if (!/^업무관리DB_백업_/.test(name)) throw new Error('백업 파일이 아닙니다.');
    const from = readTables_(SpreadsheetApp.openById(id));
    const before = backup_('복구 직전', who.user.email);
    Object.keys(SCHEMA).forEach((t) => writeTable_(t, from[t]));
    log_(who, '전체', name, '복구', null, { 백업: name, 복구직전백업: before.name });
    return { restored: name, safety: before.name };
  } finally { lock.releaseLock(); }
}

// ===================== 엑셀 가져오기 =====================
/** Drive 에서 이름에 「통합관리」가 들어간 엑셀 후보 (최근 수정 순) */
function sourceCandidates_() {
  const it = DriveApp.searchFiles(`title contains '${CFG.SOURCE_NAME}' and trashed = false`), out = [];
  while (it.hasNext()) { const f = it.next(); if (/\.xls[mx]$/i.test(f.getName()) && !/^~\$/.test(f.getName())) out.push(f); }
  return out.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
}
function sourceFile_() {
  const id = P.getProperty('XLSM_FILE_ID');
  if (id) {
    let f = null;
    try { f = DriveApp.getFileById(id); if (f.isTrashed && f.isTrashed()) f = null; } catch (e) { f = null; }
    if (!f) throw new Error('고정해 둔 엑셀 파일을 찾지 못했습니다(삭제·이동·권한 변경). 시스템 설정 › 가져올 엑셀 파일에서 다시 고르세요.');
    return f;
  }
  const best = sourceCandidates_()[0];
  if (!best) throw new Error('Drive 에서 이름에 「통합관리」가 들어간 엑셀(.xlsm)을 찾지 못했습니다.');
  return best;
}
function fileInfo_(f) { return { id: f.getId(), name: f.getName(), modified: fmtT_(f.getLastUpdated()), url: 'https://drive.google.com/file/d/' + f.getId() + '/view' }; }
function sourceList_() {
  const pinned = P.getProperty('XLSM_FILE_ID') || '';
  const list = sourceCandidates_().map(fileInfo_);
  let current = null, error = '';
  try { current = fileInfo_(sourceFile_()); } catch (e) { error = e.message; }
  if (pinned && current && !list.some((x) => x.id === current.id)) list.unshift(current); // 이름이 바뀐 고정 파일도 보이게
  return { files: list, pinned, current, error, last: setting_('마지막 가져오기') || '' };
}
function setSourceFile_(who, id) {
  id = String(id || '').trim();
  if (id) {
    let f;
    try { f = DriveApp.getFileById(id); } catch (e) { return { ok: false, msg: '파일을 열 수 없습니다. 이 계정이 볼 수 있는 파일인지 확인하세요.' }; }
    if (!/\.xls[mx]$/i.test(f.getName())) return { ok: false, msg: '엑셀 파일(.xlsm/.xlsx)만 고를 수 있습니다.' };
  }
  const prev = P.getProperty('XLSM_FILE_ID') || '';
  P.setProperty('XLSM_FILE_ID', id);
  log_(who, '설정', '가져올 엑셀', '수정', { 파일: prev || '자동' }, { 파일: id || '자동' });
  return sourceList_();
}
function readWorkbook_(file) {
  const want = ['용역관리', '용역대장', '점검이력', '수금대장', '현장연락', '외주관리', '회차별 단가', '완료보관', '거래명세서', '자료요청관리', '메일관리', '입금관리', '보고서검토관리', '업무일지', '택배발송'];
  let tmpId;
  try { tmpId = Drive.Files.copy({ name: '_업무관리_가져오기_임시', mimeType: MimeType.GOOGLE_SHEETS }, file.getId()).id; }
  catch (e) { tmpId = Drive.Files.copy({ title: '_업무관리_가져오기_임시' }, file.getId(), { convert: true }).id; }
  try {
    const ss = SpreadsheetApp.openById(tmpId), out = {};
    want.forEach((n) => { const sh = ss.getSheetByName(n); if (sh) out[n] = sh.getDataRange().getValues(); });
    return out;
  } finally { DriveApp.getFileById(tmpId).setTrashed(true); }
}
function importPreview_(who) {
  if ((setting_('원본') || '엑셀') === '웹') return { ok: false, msg: '원본이 이미 웹으로 전환되어 엑셀 가져오기는 쓸 수 없습니다. 필요하면 시스템 설정에서 원본을 엑셀로 되돌리세요.' };
  const file = sourceFile_(), V = readWorkbook_(file);
  const mapped = mapWorkbook_(V);
  const cur = readTables_(db_());
  const tables = {};
  FROM_EXCEL.forEach((t) => { tables[t] = Object.assign({ 엑셀: (mapped.tables[t] || []).length, 지금: cur[t].filter((r) => r.삭제 !== 'Y').length }, diff_(cur[t], mapped.tables[t], SCHEMA[t])); });
  const key = Utilities.getUuid();
  const f = folder_(CFG.TMP_FOLDER).createFile('import-' + key + '.json', JSON.stringify({ tables: mapped.tables, company: mapped.company || {}, anchors: mapped.anchors || {}, rules: mapped.rules || {}, file: { id: file.getId(), mod: file.getLastUpdated().getTime() } }), 'application/json');
  CacheService.getScriptCache().put('imp:' + key, JSON.stringify({ file: f.getId(), by: who.user.email }), 1800);
  return { key, source: fileInfo_(file), tables, checks: mapped.checks, warnings: mapped.warnings };
}
function importCommit_(who, key) {
  const raw = CacheService.getScriptCache().get('imp:' + key);
  if (!raw) return { ok: false, msg: '미리보기가 만료되었습니다(30분). 다시 미리보기를 하세요.' };
  if ((setting_('원본') || '엑셀') === '웹') return { ok: false, msg: '원본이 웹으로 전환되어 가져올 수 없습니다.' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const info = JSON.parse(raw), file = DriveApp.getFileById(info.file);
    const saved = JSON.parse(file.getBlob().getDataAsString('UTF-8')), tables = saved.tables || saved;
    const safety = backup_('가져오기 직전', who.user.email);
    const cur = readTables_(db_()), now = fmtT_(new Date());
    FROM_EXCEL.forEach((t) => {
      if (!tables[t]) return; // 예전 미리보기에 없던 표는 건드리지 않음
      const old = {}; cur[t].forEach((r) => { old[r.id] = r; });
      const wo = WEB_ONLY[t] || [];
      const rows = tables[t].map((r0) => {
        const o = old[r0.id], r = Object.assign({}, r0);
        if (o) wo.forEach((c) => { if (r[c] == null || r[c] === '') r[c] = o[c]; }); // 엑셀에 칸이 없는 웹 전용 값은 가져오기 때 지우지 않음
        const same = o && SCHEMA[t].every((c) => String(o[c] == null ? '' : o[c]) === String(r[c] == null ? '' : r[c]));
        return Object.assign({}, r, { 버전: o ? (same ? o.버전 : o.버전 + 1) : 1, 수정일: o && same ? o.수정일 : now, 수정자: o && same ? o.수정자 : '가져오기', 삭제: '' });
      });
      writeTable_(t, rows);
    });
    file.setTrashed(true);
    CacheService.getScriptCache().remove('imp:' + key);
    Object.entries(saved.company || {}).forEach(([k, v]) => { if (v) setSetting_('회사_' + k, v, '거래명세서에 들어가는 회사 정보 (엑셀 「거래명세서」 시트에서 가져옴)'); });
    Object.entries(saved.rules || {}).forEach(([k, v]) => setSetting_(k, v, '엑셀 「점검이력」 시트 기준값 (웹이 원본일 때 같은 계산에 씀)'));
    if (saved.anchors && saved.file) writeSync_({ fileId: saved.file.id, fileMod: saved.file.mod, at: now, anchors: saved.anchors, base: syncBase_() });
    setSetting_('마지막 가져오기', now + ' · ' + who.user.email);
    log_(who, '전체', '엑셀', '가져오기', null, { 백업: safety.name });
    return { done: true, safety: safety.name, counts: counts_() };
  } finally { lock.releaseLock(); }
}

/** 1회 이관 (원본이 웹으로 전환된 뒤 추가된 표): 엑셀 「외주관리」·「회차별 단가」, 날짜 없는 계약 외 입금 줄
 *  - 이미 있는 자료는 고치지 않고, 빠진 줄만 넣음 · 넣기 전에 DB 백업 · 엑셀 줄 위치(연결 정보)도 함께 기억해 다음 엑셀 저장 때 줄이 겹치지 않게 함
 *  - 한 번 끝나면 스크립트 속성 MIG_G 에 결과를 남기고 다시 하지 않음 */
function migrateG_(by) {
  const done = P.getProperty('MIG_G') || '';
  if (/^done/.test(done)) return null;
  if ((setting_('원본') || '엑셀') !== '웹') return null; // 엑셀이 원본이면 「가져오기」가 함께 채움
  const lock = LockService.getScriptLock(); if (!lock.tryLock(2000)) return null;
  try {
    if (/^done/.test(P.getProperty('MIG_G') || '')) return null;
    const run = +((P.getProperty('MIG_G') || '').split(':')[1] || 0); if (run && Date.now() - run < 10 * 60000) return null;
    const tries = +(P.getProperty('MIG_G_N') || 0); if (tries >= 3) return null; // 3번 실패하면 멈춤 (시스템 설정에 오류 표시)
    P.setProperty('MIG_G_N', String(tries + 1));
    P.setProperty('MIG_G', 'run:' + Date.now());
    const sync = readSync_(); if (!sync) { P.setProperty('MIG_G', 'done:연결 정보 없음 (이관 안 함)'); return null; }
    const file = DriveApp.getFileById(sync.fileId);
    const safety = backup_('외주·단가 이관 직전', by || '시스템');
    const mp = mapWorkbook_(readWorkbook_(file)), T = mp.tables, now = fmtT_(new Date());
    TBL_ = {};
    // 엑셀로 다시 읽은 계약 → 지금 DB 계약 (이름+연도, 없으면 이름만)
    const dbC = live_('계약'), mC = {}; (T.계약 || []).forEach((c) => { mC[c.id] = c; });
    const cid = (id) => { const m = mC[id]; if (!m) return ''; const k = key_(m.계약명); const same = dbC.filter((c) => key_(c.계약명) === k); return ((same.find((c) => +c.연도 === +m.연도) || same.sort((a, b) => b.연도 - a.연도)[0]) || {}).id || ''; };
    const pid = (() => { const P2 = live_('거래처'); return (name) => { const k = key_(corp_(name)); return (P2.find((p) => key_(corp_(p.거래처명)) === k) || {}).id || ''; }; })();
    const base = sync.base = sync.base || {}, A = sync.anchors = sync.anchors || {}, res = {};
    const put = (t, rows) => {
      const have = {}; rows_(t).forEach((r) => { have[r.id] = 1; });
      const add = rows.filter((r) => !have[r.id]); if (!add.length) return 0;
      const cols = SCHEMA[t].concat(META), sh = db_().getSheetByName(t) || ensureSheet_(db_(), t, cols);
      const vals = add.map((r) => cols.map((c) => (c === '버전' ? 1 : c === '수정일' ? now : c === '수정자' ? '이관' : c === '삭제' ? '' : r[c] == null ? '' : r[c])));
      sh.getRange(sh.getLastRow() + 1, 1, vals.length, cols.length).setNumberFormat('@').setValues(vals);
      (base[t] = base[t] || []).push(...add.map((r) => { const o = {}; SCHEMA[t].forEach((c) => { o[c] = r[c] == null ? '' : r[c]; }); o.삭제 = ''; return o; }));
      add.forEach((r) => { const an = ((mp.anchors || {})[t] || {})[r.id]; if (an) (A[t] = A[t] || {})[r.id] = an; });
      delete TBL_[t]; return add.length;
    };
    const outs = (T.외주 || []).map((o) => Object.assign({}, o, { 계약id: cid(o.계약id), 업체id: pid(o.업체명) || '' }));
    const prices = (T.단가 || []).map((u) => Object.assign({}, u, { 계약id: cid(u.계약id) }));
    const bills = (T.청구 || []).filter((b) => !b.발행일 && !b.입금일).map((b) => Object.assign({}, b, { 계약id: cid(b.계약id) })).filter((b) => b.계약id);
    res.외주 = put('외주', outs); res.단가 = put('단가', prices); res.청구 = put('청구', bills);
    writeSync_(sync);
    log_({ user: { email: by || '시스템' } }, '전체', '엑셀', '이관', null, { 외주: res.외주, 단가: res.단가, 청구: res.청구, 백업: safety.name });
    const msg = `${now} · 외주 ${res.외주}건 (청구금액 ${outs.reduce((a, o) => a + o.청구금액, 0).toLocaleString('en-US')}원) · 회차별 단가 ${res.단가}건 · 날짜 없는 계약 외 줄 ${res.청구}건 (${bills.reduce((a, b) => a + b.계약외, 0).toLocaleString('en-US')}원) · 백업 ${safety.name}`;
    setSetting_('엑셀 이관(외주·단가)', msg, '1회 이관 결과 (자동)');
    P.setProperty('MIG_G', 'done:' + msg);
    bumpVer_();
    return res;
  } catch (e) {
    P.setProperty('MIG_G', 'err:' + Date.now() + ':' + e.message); setSetting_('엑셀 이관(외주·단가)', fmtT_(new Date()) + ' · 실패: ' + e.message + ' (다음 접속 때 다시 시도)', '1회 이관 결과 (자동)');
    return null;
  } finally { lock.releaseLock(); }
}

/** 엑셀 값(V) → DB 표 (순수 함수: 테스트 가능) */
function mapWorkbook_(V) {
  const warnings = [];
  const AN = {}; // 엑셀 위치: AN[표][id] = { 시트: 행 }
  const anchor = (t, id, sheet, row) => { if (!row) return; const a = (AN[t] = AN[t] || {}); (a[id] = a[id] || {})[sheet] = row; };
  const svc = {};
  table_(V['용역관리'], 6, 7, '용역명').forEach((r) => {
    const name = one_(g_(r, '용역명')); if (!name || /^#/.test(name)) return;
    svc[key_(name)] = { _r: r._r, name, gu: one_(g_(r, '관할구')), 시공사: one_(g_(r, '시공사')), 규모: one_(g_(r, '규모')), 상태: one_(g_(r, '상태')), 준공: iso_(g_(r, '준공')), 계약일: iso_(g_(r, '계약일')), 완료일: iso_(g_(r, '완료일')), code: one_(g_(r, '용역코드')), 비고: one_(g_(r, '비고')) };
  });
  // 거래처
  const PT = {}, pk = {};
  const partner = (name, info, role) => {
    name = one_(name); if (!name) return '';
    const biz = digits_(info && info.사업자번호);
    let k = biz ? 'b' + biz : 'n' + key_(corp_(name));
    if (!biz && pk['n' + key_(corp_(name))]) k = pk['n' + key_(corp_(name))];
    let p = PT[k];
    if (!p) { p = PT[k] = { id: 'P' + hashS_(k), 거래처명: name, 사업자번호: '', 대표자: '', 주소: '', 업태: '', 종목: '', 구분: '', 전화: '', 이메일: '', 비고: '' }; }
    pk['n' + key_(corp_(name))] = k;
    if (info) ['사업자번호', '대표자', '주소', '업태', '종목', '이메일'].forEach((f) => { if (!p[f] && info[f]) p[f] = info[f]; });
    if (role && p.구분.split(',').indexOf(role) < 0) p.구분 = p.구분 ? p.구분 + ',' + role : role;
    return p.id;
  };
  // 현장·계약 (용역대장 = 모든 계약)
  const S = {}, contracts = [], cByName = {};
  const ledger = table_(V['용역대장'], 6, 7, '용역명');
  // 사업자번호 있는 줄을 먼저 등록해 이름만 있는 줄이 같은 거래처로 묶이게
  ledger.slice().sort((a, b) => (digits_(g_(b, '사업자')) ? 1 : 0) - (digits_(g_(a, '사업자')) ? 1 : 0)).forEach((r) => {
    partner(g_(r, '입금자'), { 사업자번호: one_(g_(r, '사업자')), 대표자: one_(g_(r, '대표자')), 주소: one_(g_(r, '사업장 소재지')), 업태: one_(g_(r, '업태')), 종목: one_(g_(r, '종목')), 이메일: one_(g_(r, '세금계산서 메일')) }, '발주처');
  });
  const seen = {};
  ledger.forEach((r) => {
    const name = one_(g_(r, '용역명')); if (!name || /^#/.test(name)) return;
    const k = key_(name), sv = svc[k] || {};
    const m = one_(g_(r, '기준연도')).match(/(20\d\d)/), year = m ? +m[1] : (sv.계약일 ? +sv.계약일.slice(0, 4) : 0);
    const no = num_(g_(r, 'No'));
    let id = 'C' + (year || '0000') + '-' + String(no || contracts.length + 1).padStart(2, '0');
    if (seen[id]) { warnings.push(`용역대장 번호 중복: ${id} (${name}) → 뒤에 -2 를 붙여 따로 보관`); id += '-2'; }
    seen[id] = 1;
    if (!S[k]) S[k] = { id: 'S' + hashS_(k), 현장명: name, 주소: one_(g_(r, '현장 주소')), 관할구: sv.gu || '', 규모: sv.규모 || '', 비고: '' };
    else if (!S[k].주소) S[k].주소 = one_(g_(r, '현장 주소'));
    const payer = partner(g_(r, '입금자'), null, '발주처');
    const real = one_(g_(r, '실제 입금처')) ? partner(g_(r, '실제 입금처'), null, '입금처') : '';
    const builder = sv.시공사 ? partner(sv.시공사, null, '시공사') : '';
    const c = { id, 현장id: S[k].id, 계약명: name, 연도: year, No: no, 계약일: sv.계약일 || '', 공급가: num_(g_(r, '공급가')), 부가세: num_(g_(r, '부가세')), 계약금: num_(g_(r, '계약금')),
      VAT기준: one_(g_(r, 'VAT 기준')), 종합보고서구분: one_(g_(r, '종합보고서 구분')), 종합보고서금액: num_(g_(r, '종합보고서')), 계약외: num_(g_(r, '계약 외')), 할증: num_(g_(r, '할증')),
      이월수금: num_(g_(r, '이월 수금액')), 미수금: num_(g_(r, '미수금')), 상태: one_(g_(r, '상태')) || '진행중', 발주처id: payer, 실제입금처id: real, 시공사id: builder,
      용역코드: one_(g_(r, '용역코드')) || sv.code || '', 준공예정일: sv.준공 || '', 완료일: sv.완료일 || '', 비고: s_(g_(r, '비고')),
      _엑셀수금: num_(g_(r, '수금 합계')), _엑셀대기: num_(g_(r, '청구 후 입금 대기')) };
    anchor('계약', id, '용역대장', r._r);
    contracts.push(c);
    (cByName[k] = cByName[k] || []).push(c);
  });
  Object.keys(svc).forEach((k) => { if (!cByName[k]) warnings.push(`용역관리에는 있지만 용역대장에 없는 용역: ${svc[k].name} (가져오지 않음)`); });
  const pickC = (name, year) => {
    const cs = cByName[key_(one_(name))]; if (!cs) return null;
    if (cs.length === 1 || !year) return cs.sort((a, b) => b.연도 - a.연도)[0];
    return cs.filter((c) => c.연도 <= year).sort((a, b) => b.연도 - a.연도)[0] || cs[0];
  };
  Object.keys(svc).forEach((k) => { const c = cByName[k] && pickC(svc[k].name); if (c) anchor('계약', c.id, '용역관리', svc[k]._r); });
  contracts.forEach((c) => { c.입금비고 = ''; c.기존수금 = 0; c.기존미입금 = 0; });
  table_(V['입금관리'], 6, 7, '용역명').forEach((r) => {
    const c = pickC(g_(r, '용역명')); if (!c) return;
    anchor('계약', c.id, '입금관리', r._r);
    c.입금비고 = s_(g_(r, '비고')); c.기존수금 = num_(g_(r, '기존 수금액')); c.기존미입금 = num_(g_(r, '기존 청구 미입금'));
  });
  // 점검: 진행 중(점검이력) + 끝난 용역(완료보관)
  const insp = [], iseen = {};
  const addInsp = (c, x, row) => {
    let id = 'I' + c.id.slice(1) + '-' + key_(x.항목); while (iseen[id]) id += '+'; iseen[id] = 1;
    insp.push(Object.assign({ id, 계약id: c.id }, x));
    if (row) anchor('점검', id, '점검이력', row);
  };
  const activeNames = {};
  table_(V['점검이력'], 6, 7, '용역명').forEach((r) => {
    const item = one_(g_(r, '점검항목')); if (!item) return;
    const c = pickC(g_(r, '용역명')); if (!c) { warnings.push(`점검이력의 용역을 용역대장에서 찾지 못함: ${one_(g_(r, '용역명'))} ${item}`); return; }
    activeNames[key_(c.계약명)] = 1;
    addInsp(c, { 항목: item, 단계: one_(g_(r, '지금 단계')), 예정일: iso_(g_(r, '점검 예정일')), 점검일: iso_(g_(r, '점검일')), 기한: iso_(g_(r, '보고서 기한')), 발송일: iso_(g_(r, '거래처 발송일')),
      시기: one_(g_(r, '점검시기')), 위치: one_(g_(r, '점검 위치')), 내부검토: one_(g_(r, '내부검토')), 비고: one_(g_(r, '비고')) }, r._r);
  });
  const arc = archive_(V['완료보관']);
  Object.keys(arc).forEach((n) => {
    if (activeNames[key_(n)]) return;
    const c = pickC(n); if (!c) { warnings.push(`완료보관의 용역을 용역대장에서 찾지 못함: ${n}`); return; }
    arc[n].forEach((x) => addInsp(c, { 항목: x.항목, 단계: '완료', 예정일: '', 점검일: x.점검일, 기한: '', 발송일: x.발송일, 시기: '', 위치: '', 내부검토: '', 비고: '완료보관에서 가져옴' }));
  });
  // 청구(수금대장)
  const bills = [], bseen = {};
  table_(V['수금대장'], 6, 7, '용역명').forEach((r) => {
    const d = iso_(g_(r, '세/계 발행일')), paidOn = iso_(g_(r, '입금일')), anyD = d || paidOn;
    // 발행일도 입금일도 없는 줄: 계약 외 금액·할증분·금액이 적혀 있으면 가져옴 (예: 계약 외 현금 990,000원), 아무것도 없으면 빈 줄
    if (!anyD && !(num_(g_(r, '계약 외')) || num_(g_(r, '할증')) || num_(g_(r, '청구액')))) return;
    const name = one_(g_(r, '용역명')), c = pickC(name, anyD ? +anyD.slice(0, 4) : 0);
    if (!c) { warnings.push(`수금대장의 용역을 용역대장에서 찾지 못함: ${name} (${d})`); return; }
    let id = 'B' + hashS_([anyD, key_(name), key_(one_(g_(r, '품명'))), num_(g_(r, '청구액'))].join('|')); while (bseen[id]) id += '+'; bseen[id] = 1;
    const links = [1, 2, 3, 4].map((i) => one_(g_(r, '연결 점검 ' + i))).filter((x) => x && !/금액/.test(x));
    anchor('청구', id, '수금대장', r._r);
    const [memo, parts] = splitParts_(s_(g_(r, '비고')));
    bills.push({ id, 계약id: c.id, 발행일: d, 품명: one_(g_(r, '품명')), 청구액: num_(g_(r, '청구액')), 공급가: num_(g_(r, '공급가')), 부가세: num_(g_(r, '부가세')), 입금일: iso_(g_(r, '입금일')),
      수금액: num_(g_(r, '수금액')), 입금자: one_(g_(r, '입금자')), 상태: one_(g_(r, '상태')), 연결점검: links.join(', '), 계약외: num_(g_(r, '계약 외')), 할증: num_(g_(r, '할증')), 비고: memo || (d || !paidOn ? '' : '계산서 발행 없이 입금'), 분할입금: parts });
  });
  // 연락(현장연락)
  const calls = [], tseen = {};
  table_(V['현장연락'], 6, 7, '용역명').forEach((r) => {
    const name = one_(g_(r, '용역명')), s = S[key_(name)];
    if (!s) { warnings.push(`현장연락의 용역을 찾지 못함: ${name}`); return; }
    const d = iso_(g_(r, '연락일'));
    let id = 'T' + hashS_([d, key_(name), one_(g_(r, '방법')), s_(g_(r, '통화 메모'))].join('|')); while (tseen[id]) id += '+'; tseen[id] = 1;
    anchor('연락', id, '현장연락', r._r);
    calls.push({ id, 현장id: s.id, 연락일: d, 방법: one_(g_(r, '방법')), 담당자: one_(g_(r, '현장 담당자')), 공정: one_(g_(r, '현재 공정')), 메모: s_(g_(r, '통화 메모')), 예상점검: one_(g_(r, '다음 점검 예상')),
      예상시기: iso_(g_(r, '예상 시기')), 다음연락: iso_(g_(r, '다음 연락 예정일')), 비고: s_(g_(r, '비고')), 작성자: '엑셀' });
  });
  // 자료요청(자료요청관리) · 메일(메일관리)
  const reqs = [], qseen = {};
  table_(V['자료요청관리'], 6, 7, '용역명').forEach((r) => {
    const name = one_(g_(r, '용역명')), s = S[key_(name)];
    if (!s) { warnings.push(`자료요청관리의 용역을 찾지 못함: ${name}`); return; }
    const no = one_(g_(r, 'NO')), d = iso_(g_(r, '요청일'));
    let id = 'Q' + hashS_([no, d, key_(name), s_(g_(r, '요청 자료'))].join('|')); while (qseen[id]) id += '+'; qseen[id] = 1;
    anchor('자료요청', id, '자료요청관리', r._r);
    reqs.push({ id, NO: no, 현장id: s.id, 요청일: d, 방법: one_(g_(r, '요청방법')), 담당자: one_(g_(r, '현장 담당자')), 항목: one_(g_(r, '관련 점검항목')), 점검일: iso_(g_(r, '점검일')), 구분: one_(g_(r, '요청구분')),
      원요청: one_(g_(r, '원요청')), 자료: s_(g_(r, '요청 자료')), 기한: iso_(g_(r, '회신 기한')), 수신일: iso_(g_(r, '수신일')), 상태: one_(g_(r, '처리상태')), 검토: s_(g_(r, '수신 자료 검토')), 비고: s_(g_(r, '비고')) });
  });
  const mails = [], mseen = {};
  table_(V['메일관리'], 5, 6, '제목').forEach((r) => {
    const name = one_(g_(r, '용역명')), s = S[key_(name)];
    if (!s) { warnings.push(`메일관리의 용역을 찾지 못함: ${name}`); return; }
    const no = one_(g_(r, 'NO')), d = iso_(g_(r, '보낸 날짜'));
    let id = 'M' + hashS_([no, d, key_(name), one_(g_(r, '제목'))].join('|')); while (mseen[id]) id += '+'; mseen[id] = 1;
    anchor('메일', id, '메일관리', r._r);
    mails.push({ id, NO: no, 현장id: s.id, 구분: one_(g_(r, '구분')), 항목: one_(g_(r, '점검항목')), 받는사람: one_(g_(r, '받는 사람')), 이메일: one_(g_(r, '이메일')), 보낸날: d, 제목: one_(g_(r, '제목')), 내용: s_(g_(r, '메일 내용')),
      첨부: one_(g_(r, '첨부파일')), 회신필요: one_(g_(r, '회신 필요')), 회신상태: one_(g_(r, '회신 상태')), 회신일: iso_(g_(r, '회신일')), 회신내용: s_(g_(r, '회신 내용')), 회신첨부: one_(g_(r, '회신 첨부')), 비고: s_(g_(r, '비고')) });
  });
  // 보고서 검토(보고서검토관리 · 1줄 = 1회 발송)
  const revs = [], rseen = {};
  table_(V['보고서검토관리'], 6, 7, '용역명').forEach((r) => {
    const name = one_(g_(r, '용역명')), c = pickC(name);
    if (!c) { warnings.push(`보고서검토관리의 용역을 찾지 못함: ${name}`); return; }
    const no = one_(g_(r, 'NO')), d = iso_(g_(r, '검토요청일'));
    let id = 'R' + hashS_([no, d, key_(name), key_(one_(g_(r, '점검항목'))), one_(g_(r, '발송 회차'))].join('|')); while (rseen[id]) id += '+'; rseen[id] = 1;
    anchor('검토', id, '보고서검토관리', r._r);
    revs.push({ id, NO: no, 계약id: c.id, 항목: one_(g_(r, '점검항목')), 회차: one_(g_(r, '발송 회차')), 요청일: d, 검토자: one_(g_(r, '검토자')), 제목: one_(g_(r, '메일 제목')), 상태: one_(g_(r, '검토상태')),
      회신일: iso_(g_(r, '검토 회신일')), 의견: s_(g_(r, '검토내용')), 조치: s_(g_(r, '수정사항')), 재발송일: iso_(g_(r, '수정 재발송일')), 비고: s_(g_(r, '비고')) });
  });
  // 업무일지 (날짜 줄 아래에 같은 날 업무가 더 있으면 날짜 칸이 빈 줄로 이어짐)
  const jobs = [], jseen = {}, JV = V['업무일지'] || [];
  let jd = '', jn = 0;
  for (let r = 6; r < JV.length; r++) {
    const row = JV[r] || [], c3 = row[2];
    if (dt_(c3)) { jd = iso_(c3); jn = 0; } else if (String(c3 == null ? '' : c3).trim()) { jd = ''; continue; } // 월 제목줄
    const 현장 = one_(row[4]), 구분 = one_(row[5]), 내용 = s_(row[6]);
    if (!jd || !(현장 || 구분 || 내용)) continue;
    jn++;
    let id = 'J' + hashS_([jd, jn, 내용].join('|')); while (jseen[id]) id += '+'; jseen[id] = 1;
    jobs.push({ id, 날짜: jd, 현장, 구분, 내용, 상태: one_(row[7]), 후속: one_(row[8]), 비고: s_(row[9]), 작성자: '엑셀' });
  }
  // 택배발송
  const parcels = [], kseen = {};
  table_(V['택배발송'], 6, 7, '용역명').forEach((r) => {
    const name = one_(g_(r, '용역명')), s = S[key_(name)];
    if (!s) { warnings.push(`택배발송의 용역을 찾지 못함: ${name}`); return; }
    const d = iso_(g_(r, '발송일'));
    let id = 'K' + hashS_([d, key_(name), one_(g_(r, '점검항목')), one_(g_(r, '받는 사람'))].join('|')); while (kseen[id]) id += '+'; kseen[id] = 1;
    anchor('택배', id, '택배발송', r._r);
    parcels.push({ id, 현장id: s.id, 발송일: d, 항목: one_(g_(r, '점검항목')), 받는사람: one_(g_(r, '받는 사람')), 주소: one_(g_(r, '주소')), 비고: one_(g_(r, '비고')) });
  });
  // 외주관리 (외주업체도 거래처로). 엑셀 글자(업체·용역명)는 그대로 보관
  const outs = [], oseen = {};
  table_(V['외주관리'], 6, 7, '외주업체').forEach((r) => {
    const vn = one_(g_(r, '외주업체')), pid = partner(vn, null, '외주'), name = one_(g_(r, '용역명')), d = iso_(g_(r, '계산서일자'));
    const c = name ? pickC(name, d ? +d.slice(0, 4) : 0) : null;
    if (name && !c) warnings.push(`외주관리의 용역을 용역대장에서 찾지 못함: ${name} (글자 그대로 보관)`);
    let id = 'O' + hashS_([d, key_(vn), key_(name), key_(one_(g_(r, '작업내용'))), num_(g_(r, '청구금액'))].join('|')); while (oseen[id]) id += '+'; oseen[id] = 1;
    anchor('외주', id, '외주관리', r._r);
    const pd = g_(r, '점검일');
    outs.push({ id, 계산서일: d, 업체id: pid, 업체명: vn, 계약id: c ? c.id : '', 현장명: name, 점검일: iso_(pd) || one_(pd), 내용: s_(g_(r, '작업내용')), 청구금액: num_(g_(r, '청구금액')), 지급금액: num_(g_(r, '지급금액')),
      지급일: iso_(g_(r, '지급일')), 비고: s_(g_(r, '비고')), 연결점검: [1, 2].map((i) => one_(g_(r, '연결 점검 ' + i))).filter(Boolean).join(', ') });
  });
  // 회차별 단가 (1줄 = 용역 × 점검항목)
  const prices = [], useen = {};
  table_(V['회차별 단가'], 6, 7, '용역명').forEach((r) => {
    const name = one_(g_(r, '용역명')), item = one_(g_(r, '점검항목')); if (!name || !item) return;
    const c = pickC(name);
    if (!c) warnings.push(`회차별 단가의 용역을 용역대장에서 찾지 못함: ${name} ${item} (글자 그대로 보관)`);
    let id = 'U' + hashS_([key_(name), key_(item)].join('|')); while (useen[id]) id += '+'; useen[id] = 1;
    anchor('단가', id, '회차별 단가', r._r);
    const up = g_(r, '회당 단가');
    prices.push({ id, 계약id: c ? c.id : '', 계약명: name, 항목: item, 단가: up === '' || up == null || /^#/.test(String(up)) ? '' : String(num_(up)), 횟수: num_(g_(r, '횟수')), 회차: one_(g_(r, '해당 회차')), 메모: s_(g_(r, '메모')) });
  });
  // 검증: 엑셀이 계산한 값과 DB 로 다시 계산한 값 비교
  const checks = [];
  const byYear = {};
  contracts.forEach((c) => { const y = byYear[c.연도] = byYear[c.연도] || { 건수: 0, 공급가: 0, 계약금: 0 }; y.건수++; y.공급가 += c.공급가; y.계약금 += c.계약금; });
  let paidMismatch = 0, waitMismatch = 0, paidX = 0, paidD = 0;
  contracts.forEach((c) => {
    const mine = bills.filter((b) => b.계약id === c.id);
    const paid = c.이월수금 + mine.reduce((s, b) => s + b.수금액, 0);
    const wait = mine.filter((b) => !b.입금일 && b.발행일).reduce((s, b) => s + Math.max(0, b.청구액 - b.수금액), 0);
    paidX += c._엑셀수금; paidD += paid;
    if (Math.abs(paid - c._엑셀수금) > 1) { paidMismatch++; if (paidMismatch <= 5) warnings.push(`수금 합계 차이: ${c.계약명} 엑셀 ${c._엑셀수금} / 다시 계산 ${paid}`); }
    if (Math.abs(wait - c._엑셀대기) > 1) waitMismatch++;
    delete c._엑셀수금; delete c._엑셀대기;
  });
  checks.push({ 항목: '계약 건수', 엑셀: ledger.filter((r) => one_(g_(r, '용역명'))).length, DB: contracts.length });
  Object.keys(byYear).sort().forEach((y) => checks.push({ 항목: `${y}년 계약금(VAT 포함)`, 엑셀: byYear[y].계약금, DB: byYear[y].계약금 }));
  checks.push({ 항목: '수금 합계 (엑셀 계산 vs 이월+수금대장)', 엑셀: paidX, DB: paidD, 차이건수: paidMismatch });
  checks.push({ 항목: '청구 후 입금 대기가 다른 계약 수', 엑셀: 0, DB: waitMismatch });
  checks.push({ 항목: '외주 내역 건수 · 청구금액 합계', 엑셀: table_(V['외주관리'], 6, 7, '외주업체').length + '건 · ' + table_(V['외주관리'], 6, 7, '외주업체').reduce((a, r) => a + num_(g_(r, '청구금액')), 0), DB: outs.length + '건 · ' + outs.reduce((a, o) => a + o.청구금액, 0) });
  checks.push({ 항목: '회차별 단가 줄 수', 엑셀: table_(V['회차별 단가'], 6, 7, '용역명').length, DB: prices.length });
  const partners = Object.values(PT);
  // 엑셀 「거래명세서」 시트의 공급자 칸 (H3 사업자번호, H4 업체명, J4 대표, H5 주소, H6 업태, J6 종목, H7 전화, J7 팩스, A39 기타사항)
  const TS = V['거래명세서'] || [], cell = (r, c) => String(((TS[r - 1] || [])[c - 1]) == null ? '' : TS[r - 1][c - 1]).trim();
  const company = TS.length ? { 사업자번호: cell(3, 8), 상호: cell(4, 8), 대표: cell(4, 10), 주소: cell(5, 8), 업태: cell(6, 8), 종목: cell(6, 10), 전화: cell(7, 8), 팩스: cell(7, 10), 거래명세서_기타사항: cell(39, 1) } : {};
  const PI = V['점검이력'] || [], numAt = (r, c, d) => { const v = (PI[r - 1] || [])[c - 1]; return typeof v === 'number' && v > 0 ? v : d; };
  const rules = { 보고서기한일수: numAt(4, 4, 30), 입금지연일수: numAt(4, 14, 30), 일지연도: (typeof (JV[1] || [])[3] === 'number' && JV[1][3]) || new Date().getFullYear() }; // 점검이력 D4 · N4, 업무일지 D2
  return { tables: { 현장: Object.values(S), 거래처: partners, 계약: contracts, 점검: insp, 청구: bills, 연락: calls, 자료요청: reqs, 메일: mails, 검토: revs, 일지: jobs, 택배: parcels, 외주: outs, 단가: prices }, checks, warnings, company, anchors: AN, rules };
}
function archive_(v) {
  const out = {};
  if (!v || !v.length) return out;
  let hr = -1, hc = {};
  for (let r = 0; r < Math.min(v.length, 12) && hr < 0; r++) {
    const row = v[r].map(cleanHead_);
    if (row.indexOf('용역코드') >= 0 && row.indexOf('점검항목') >= 0) { hr = r; row.forEach((h, i) => { if (h) hc[h] = i; }); }
  }
  if (hr < 0) return out;
  const code2name = {};
  v.forEach((row) => { const m = String(row[0] == null ? '' : row[0]).match(/^■\s*\[(.+?)\]\s*(.+?)(\s{2,}\(|\s*\(보관일|$)/); if (m) code2name[m[1].trim()] = m[2].trim(); });
  const col = (row, k) => { for (const h in hc) if (h.indexOf(k) === 0) return row[hc[h]]; return null; };
  for (let r = hr + 1; r < v.length; r++) {
    const row = v[r], code = one_(col(row, '용역코드')), item = one_(col(row, '점검항목'));
    if (!code || !item) continue;
    const n = code2name[code] || code;
    (out[n] = out[n] || []).push({ 항목: item, 점검일: iso_(col(row, '점검일')), 발송일: iso_(col(row, '거래처 발송일')) });
  }
  return out;
}

// ===================== 웹 → 엑셀 저장 (원본이 웹일 때) =====================
// 통합관리 엑셀(.xlsm)을 틀로 두고, 웹에서 바뀐 「입력 칸」만 고쳐 같은 파일에 덮어씁니다.
// 수식 칸·서식·매크로·다른 시트는 건드리지 않습니다. 덮어쓰기 직전 원본을 「업무관리 엑셀백업」 폴더에 복사합니다.
const XL = { // 시트: 데이터 첫 행, 수식이 미리 깔린 마지막 행, 행을 알아보는 열
  용역대장: { first: 7, last: 506, key: 'C' },
  용역관리: { first: 7, last: 106, key: 'B' },
  입금관리: { first: 7, last: 106, key: 'A' },
  점검이력: { first: 7, last: 606, key: 'C' },
  보고서검토관리: { first: 7, last: 206, key: 'C' },
  수금대장: { first: 7, last: 1006, key: 'B' },
  현장연락: { first: 7, last: 506, key: 'B' },
  자료요청관리: { first: 7, last: 306, key: 'C' },
  메일관리: { first: 6, last: 2000, key: 'H' },
  택배발송: { first: 7, last: 506, key: 'B' },
  외주관리: { first: 7, last: 506, key: 'B' },
  '회차별 단가': { first: 7, last: 1006, key: 'C' },
};
const SYNC_TABLES = ['계약', '점검', '청구', '연락', '자료요청', '메일', '검토', '택배', '외주', '단가'];
const XL_FOLDER = '업무관리 엑셀백업';
const XL_KEEP = 30;

function syncBase_() {
  const all = readTables_(db_()), out = {};
  FROM_EXCEL.forEach((t) => { out[t] = (all[t] || []).map((r) => { const o = {}; SCHEMA[t].forEach((c) => { o[c] = r[c] == null ? '' : r[c]; }); o.삭제 = r.삭제 || ''; return o; }); });
  return out;
}
function syncFile_(create) {
  const id = P.getProperty('SYNC_FILE_ID');
  if (id) { try { const f = DriveApp.getFileById(id); if (!(f.isTrashed && f.isTrashed())) return f; } catch (e) {} }
  if (!create) return null;
  const f = folder_(XL_FOLDER).createFile('엑셀연결.json', '{}', 'application/json');
  P.setProperty('SYNC_FILE_ID', f.getId());
  return f;
}
function readSync_() { const f = syncFile_(false); if (!f) return null; try { const o = JSON.parse(f.getBlob().getDataAsString('UTF-8')); return o && o.base ? o : null; } catch (e) { return null; } }
function writeSync_(o) { syncFile_(true).setContent(JSON.stringify(o)); }
function markDirty_() { if ((setting_('원본') || '엑셀') === '웹') P.setProperty('EXPORT_DIRTY', String(Date.now())); }

// 표 한 줄 → 엑셀 칸 { 시트: { 열: [종류, 값] } }  (종류: s 글자, n 숫자, d 날짜)
function syncCtx_(T) {
  const live = (t) => (T[t] || []).filter((r) => r.삭제 !== 'Y');
  const by = (t) => { const m = {}; live(t).forEach((r) => { m[r.id] = r; }); return m; };
  const S = by('현장'), Pn = by('거래처'), C = by('계약'), siteName = {};
  live('계약').slice().sort((a, b) => num_(a.연도) - num_(b.연도)).forEach((c) => { siteName[c.현장id] = c.계약명; });
  return { S, Pn, C, pn: (id) => (Pn[id] || {}).거래처명 || '', site: (id) => siteName[id] || (S[id] || {}).현장명 || '' };
}
function renderXl_(t, r, X, isNew) {
  const s = (v) => ['s', v == null ? '' : String(v)], n = (v) => ['n', num_(v)], d = (v) => ['d', iso_(v) || (/^\d{4}-\d\d-\d\d/.test(String(v || '')) ? String(v).slice(0, 10) : '')];
  if (t === '계약') {
    const st = X.S[r.현장id] || {}, p = X.Pn[r.발주처id] || {};
    const o = {
      용역대장: { A: n(r.No), B: s(r.연도 ? r.연도 + '년' : ''), C: s(r.계약명), D: s(p.거래처명), E: n(r.계약금), H: n(r.이월수금), P: n(r.종합보고서금액), Q: n(r.미수금), R: s(r.비고),
        T: s(p.사업자번호), U: s(p.이메일), V: s(p.대표자), W: s(p.주소), X: s(p.업태), Y: s(p.종목), Z: s(r.VAT기준), AA: s(st.주소), AB: s(r.종합보고서구분), AC: s(X.pn(r.실제입금처id)) },
      용역관리: { A: n(r.No), B: s(r.계약명), C: n(r.연도), D: s(st.규모), E: s(st.관할구), F: s(X.pn(r.시공사id)), G: d(r.준공예정일), AK: d(r.계약일) },
      입금관리: { A: s(r.계약명), B: n(r.계약금), C: s(p.거래처명), D: n(r.기존수금), E: n(r.기존미입금), H: s(r.입금비고) },
    };
    return o;
  }
  if (t === '점검') {
    const c = X.C[r.계약id] || {};
    return { 점검이력: { C: s(c.계약명), D: s(r.항목), F: d(r.예정일), H: s(r.위치), I: d(r.점검일), P: d(r.발송일), Q: s(r.비고) } };
  }
  if (t === '청구') {
    const c = X.C[r.계약id] || {}, L = String(r.연결점검 || '').split(/\s*,\s*/).filter(Boolean);
    return { 수금대장: { A: d(r.발행일), B: s(c.계약명), C: s(r.품명), D: n(r.청구액), E: d(r.입금일), F: s(r.입금자), I: s([r.비고 === '계산서 발행 없이 입금' ? '' : r.비고, partsText_(r.분할입금, !!r.입금일)].filter(Boolean).join(' · ')),
      J: s(L[0]), K: s(L[1]), L: s(L[2]), M: s(L[3]), AH: n(r.계약외), AI: n(r.할증) } };
  }
  if (t === '연락') return { 현장연락: { A: d(r.연락일), B: s(X.site(r.현장id)), C: s(r.방법), D: s(r.담당자), E: s(r.공정), F: s(r.예상점검), G: d(r.예상시기), H: d(r.다음연락), I: s(r.메모), J: s(r.비고) } };
  if (t === '자료요청') return { 자료요청관리: { C: s(X.site(r.현장id)), D: d(r.요청일), E: s(r.방법), F: s(r.담당자), G: s(r.항목), I: s(r.구분), J: s(r.원요청), K: s(r.자료), L: d(r.기한), M: d(r.수신일), N: s(r.상태), Q: s(r.검토), R: s(r.비고) } };
  if (t === '검토') { const c = X.C[r.계약id] || {}; return { 보고서검토관리: { C: s(c.계약명), D: s(r.항목), F: /^\d+$/.test(String(r.회차)) ? n(r.회차) : s(r.회차), G: d(r.요청일), H: s(r.검토자), I: s(r.제목), J: s(r.상태), K: d(r.회신일), N: s(r.의견), O: s(r.조치), P: d(r.재발송일), Q: s(r.비고) } }; }
  if (t === '외주') { const c = X.C[r.계약id] || {}, L = String(r.연결점검 || '').split(/\s*,\s*/).filter(Boolean);
    return { 외주관리: { A: d(r.계산서일), B: s(r.업체명 || X.pn(r.업체id)), C: s(c.계약명 || r.현장명), D: /^\d{4}-\d\d-\d\d$/.test(String(r.점검일 || '')) ? d(r.점검일) : s(r.점검일), E: s(r.내용), F: n(r.청구금액), G: n(r.지급금액), H: d(r.지급일), I: s(r.비고), J: s(L[0]), K: s(L[1]) } }; }
  if (t === '단가') { const c = X.C[r.계약id] || {}; return { '회차별 단가': { C: s(c.계약명 || r.계약명), D: s(r.항목), E: String(r.단가 == null ? '' : r.단가) === '' ? ['s', ''] : ['z', num_(r.단가)], F: n(r.횟수), G: s(r.회차), J: s(r.메모) } }; }
  if (t === '택배') return { 택배발송: { A: d(r.발송일), B: s(X.site(r.현장id)), C: s(r.항목), D: s(r.받는사람), E: s(r.주소), F: s(r.비고) } };
  if (t === '메일') return { 메일관리: { A: /^\d+$/.test(String(r.NO)) ? n(r.NO) : s(r.NO), B: s(X.site(r.현장id)), C: s(r.구분), D: s(r.항목), E: s(r.받는사람), F: s(r.이메일), G: d(r.보낸날), H: s(r.제목), I: s(r.내용), J: s(r.첨부), K: s(r.회신필요), L: s(r.회신상태), M: d(r.회신일), N: s(r.회신내용), O: s(r.회신첨부), Q: s(r.비고) } };
  return {};
}
function xlKey_(v) { return (v[0] === 'z' && String(v[1]) === '' ? 's' : v[0]) + '|' + (v[0] === 'n' ? (num_(v[1]) || '') : String(v[1] == null ? '' : v[1]).trim()); }
// 엑셀에 아직 없는 줄을 새로 만들지
function wantNew_(t, sheet, r, b) {
  if (t === '계약') return !b && (sheet === '용역대장' || r.상태 !== '완료');
  return !b;
}

/** 웹 DB 와 마지막 동기화 때의 차이만 엑셀에 씁니다. 반환: { 칸, 줄 } */
function exportExcel_(by) {
  // 엑셀 내려받기·고치기·올리기는 오래 걸리므로 그동안 웹 저장이 기다리지 않게 스크립트 잠금을 잡지 않음.
  // 대신 엑셀 저장끼리 겹치지 않도록 짧은 잠금으로 「저장 중」 표시만 확인.
  const lk = LockService.getScriptLock(); lk.waitLock(30000);
  try {
    const run = +(P.getProperty('EXPORT_RUN') || 0);
    if (run && Date.now() - run < 6 * 60000) throw new Error('엑셀 저장이 이미 진행 중입니다. 잠시 뒤 다시 시도하세요.');
    P.setProperty('EXPORT_RUN', String(Date.now()));
  } finally { lk.releaseLock(); }
  try {
    const sync = readSync_();
    if (!sync) throw new Error('엑셀 연결 정보가 없습니다. 원본을 엑셀로 둔 상태에서 「엑셀 가져오기」를 한 번 해 주세요.');
    const cur = syncBase_(), base = sync.base, A = sync.anchors || {};
    const Xc = syncCtx_(cur), Xb = syncCtx_(base), ops = [];
    SYNC_TABLES.forEach((t) => {
      const B = {}, seen = {}; (base[t] || []).forEach((r) => { if (r.삭제 !== 'Y') B[r.id] = r; });
      const one = (id, r, b, del) => {
        const an = (A[t] || {})[id] || {};
        const now = del ? {} : renderXl_(t, r, Xc, !b), was = b ? renderXl_(t, b, Xb, false) : {};
        Object.keys(Object.assign({}, now, was)).forEach((sheet) => {
          const nw = now[sheet] || {}, ws = was[sheet] || {}, row = an[sheet];
          if (!row) {
            if (del || !wantNew_(t, sheet, r, b) || !XL[sheet]) return;
            const cells = {}; Object.keys(nw).forEach((c) => { if (xlKey_(nw[c]) !== xlKey_([nw[c][0], ''])) cells[c] = nw[c]; });
            ops.push({ t, id, sheet, row: 0, cells, keyWas: '' }); return;
          }
          const cells = {};
          Object.keys(Object.assign({}, nw, ws)).forEach((c) => { const a = nw[c] || [ws[c][0], ''], z = ws[c] || [a[0], '']; if (xlKey_(a) !== xlKey_(z)) cells[c] = a; });
          if (Object.keys(cells).length || del) ops.push({ t, id, sheet, row, cells, del, keyWas: ws[XL[sheet].key] ? String(ws[XL[sheet].key][1]) : '' });
        });
      };
      (cur[t] || []).forEach((r) => { seen[r.id] = 1; const b = B[r.id], del = r.삭제 === 'Y'; if (!b && del) return; if (del && !Object.keys((A[t] || {})[r.id] || {}).length) return; one(r.id, r, b, del); });
      (base[t] || []).forEach((b) => { if (!seen[b.id] && B[b.id]) one(b.id, null, b, true); }); // 복구 등으로 표에서 사라진 줄
    });
    const jy = +(setting_('일지연도') || new Date().getFullYear());
    const jsig = (T) => JSON.stringify((T.일지 || []).filter((r) => r.삭제 !== 'Y' && String(r.날짜).slice(0, 4) === String(jy)).map((r) => SCHEMA.일지.filter((c) => c !== '작성자').map((c) => r[c])));
    const journal = jsig(cur) !== jsig(base);
    if (!ops.length && !journal) { sync.base = cur; sync.at = fmtT_(new Date()); writeSync_(sync); return { 칸: 0, 줄: 0 }; }

    const file = DriveApp.getFileById(sync.fileId);
    const book = xlOpen_(file.getBlob()), sheets = {};
    const sh = (name) => sheets[name] || (sheets[name] = xlSheet_(book.sheetXml(name), book.strings));
    // 1) 기존 줄이 그 자리에 그대로 있는지 확인 (엑셀을 직접 고쳐 줄이 밀렸으면 멈춤)
    const bad = [];
    ops.forEach((o) => { if (!o.row || !o.keyWas) return; const v = sh(o.sheet).value(o.row, colNum_(XL[o.sheet].key)); if (key_(one_(v)) !== key_(one_(o.keyWas))) bad.push(`${o.sheet} ${o.row}행 (기대: ${o.keyWas} / 엑셀: ${one_(v) || '빈칸'})`); });
    if (bad.length) throw new Error('엑셀 줄 위치가 웹이 기억하는 것과 다릅니다. 엑셀을 직접 고친 것 같습니다: ' + bad.slice(0, 3).join(', ') + (bad.length > 3 ? ` 외 ${bad.length - 3}건` : '') + ' — 시스템 설정에서 원본을 엑셀로 되돌려 「가져오기」를 다시 한 뒤 웹으로 전환하세요.');
    // 2) 새 줄 자리 잡기
    const next = {};
    const freeRow = (name) => {
      const cfg = XL[name], s = sh(name), kc = colNum_(cfg.key);
      if (!next[name]) { let last = cfg.first - 1; for (let r = cfg.first; r <= cfg.last; r++) if (one_(s.value(r, kc))) last = r; next[name] = last + 1; }
      const r = next[name]++;
      if (r > cfg.last) throw new Error(`엑셀 「${name}」 시트의 입력 칸(${cfg.last}행까지)이 가득 찼습니다. 엑셀에서 수식 줄을 늘린 뒤 다시 가져오기·전환하세요.`);
      return r;
    };
    let cells = 0, rows = 0; const skipped = [], dbFix = [];
    ops.forEach((o) => {
      const s = sh(o.sheet);
      if (!o.row) {
        o.row = freeRow(o.sheet);
        ((A[o.t] = A[o.t] || {})[o.id] = A[o.t][o.id] || {})[o.sheet] = o.row;
        if (o.t === '자료요청' || o.t === '검토') { const no = s.value(o.row, 1); if (no !== '' && no != null) dbFix.push([o.t, o.id, 'NO', String(no)]); }
        if (o.t === '메일' && !String((o.cells.A || [])[1] || '')) { let mx = 0; for (let r = XL.메일관리.first; r <= XL.메일관리.last; r++) { const v = s.value(r, 1); if (typeof v === 'number' && v > mx) mx = v; } o.cells.A = ['n', mx + 1]; dbFix.push(['메일', o.id, 'NO', String(mx + 1)]); }
      }
      rows++;
      if (o.del) delete ((A[o.t] || {})[o.id] || {})[o.sheet]; // 지운 줄은 더 이상 기억하지 않음
      Object.keys(o.cells).forEach((c) => { const [k, v] = o.cells[c]; if (s.set(o.row, colNum_(c), k, xlVal_(k, v))) cells++; else skipped.push(o.sheet + ' ' + c + o.row); });
    });
    Object.keys(sheets).forEach((name) => { if (sheets[name].dirty()) book.putSheet(name, sheets[name].toXml()); });
    if (journal) { const n = xlJournal_(book, (cur.일지 || []).filter((r) => r.삭제 !== 'Y'), jy); rows += n; cells += n; }
    // 3) 백업 → 덮어쓰기
    const bak = file.makeCopy('엑셀백업_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd-HHmmss') + '.xlsm', folder_(XL_FOLDER));
    bak.setDescription('웹→엑셀 저장 직전 원본: ' + file.getName());
    pruneXlBackups_();
    const blob = book.toBlob(file.getName(), 'application/vnd.ms-excel.sheet.macroEnabled.12');
    try { Drive.Files.update({}, sync.fileId, blob); } catch (e) { throw new Error('엑셀 파일을 덮어쓰지 못했습니다: ' + e.message); }
    // 4) 웹 쪽 번호(NO) 맞추기 · 기준 갱신 (기준은 이번에 엑셀에 쓴 내용 그대로 — 그 사이 웹에서 고친 것은 다음 저장 때 반영)
    const fileMod = DriveApp.getFileById(sync.fileId).getLastUpdated().getTime(), now = fmtT_(new Date());
    dbFix.forEach(([t, id, col, v]) => { const r = (cur[t] || []).find((x) => x.id === id); if (r) r[col] = v; });
    if (dbFix.length) {
      const lk2 = LockService.getScriptLock(); lk2.waitLock(30000);
      try { TBL_ = {}; dbFix.forEach(([t, id, col, v]) => setSilent_(t, id, { [col]: v })); bumpVer_(); } finally { lk2.releaseLock(); }
    }
    Object.assign(sync, { base: cur, anchors: A, at: now, fileMod });
    writeSync_(sync);
    setSetting_('마지막 엑셀 저장', `${now} · ${by || ''} · ${rows}줄 ${cells}칸` + (skipped.length ? ` (수식 칸이라 건너뜀: ${skipped.slice(0, 3).join(', ')})` : ''));
    setSetting_('엑셀 저장 오류', '');
    return { 칸: cells, 줄: rows, 건너뜀: skipped };
  } finally { P.deleteProperty('EXPORT_RUN'); }
}
function pruneXlBackups_() {
  const it = folder_(XL_FOLDER).getFiles(), list = [];
  while (it.hasNext()) { const f = it.next(); if (/^엑셀백업_/.test(f.getName())) list.push(f); }
  list.sort((a, b) => b.getName().localeCompare(a.getName())).slice(XL_KEEP).forEach((f) => f.setTrashed(true));
}
function xlVal_(k, v) {
  if (v === '' || v == null) return '';
  if (k === 'n') { const x = num_(v); return x ? x : ''; }
  if (k === 'z') return num_(v); // 0 도 숫자로 씀 (회차별 단가: 0원과 빈 칸을 구분)
  if (k === 'd') { const m = String(v).match(/^(\d{4})-(\d\d)-(\d\d)/); return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 864e5 : ''; }
  return String(v);
}
/** 자동 저장: 5분마다 돌며, 마지막 입력 뒤 2분이 지났으면 저장 */
function exportTick() {
  if ((setting_('원본') || '엑셀') !== '웹') return;
  try { ensureAll_(); migrateG_('자동'); } catch (e) {}
  const d = +(P.getProperty('EXPORT_DIRTY') || 0);
  if (!d || Date.now() - d < 2 * 60000) return;
  try {
    exportExcel_('자동');
    if (+(P.getProperty('EXPORT_DIRTY') || 0) === d) P.setProperty('EXPORT_DIRTY', '');
  } catch (e) {
    const msg = fmtT_(new Date()) + ' · ' + e.message, prev = setting_('엑셀 저장 오류');
    setSetting_('엑셀 저장 오류', msg);
    if (!prev || prev.slice(19) !== msg.slice(19)) { try { MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: '[업무관리] 엑셀 자동 저장 실패', body: e.message + '\n\n웹 자료는 안전하게 남아 있습니다. 시스템관리 › 시스템 설정에서 「지금 엑셀에 저장」으로 다시 시도할 수 있습니다.' }); } catch (x) {} }
  }
}
function exportTrigger_(on) {
  ScriptApp.getProjectTriggers().forEach((t) => { if (t.getHandlerFunction && t.getHandlerFunction() === 'exportTick') ScriptApp.deleteTrigger(t); });
  if (on) ScriptApp.newTrigger('exportTick').timeBased().everyMinutes(5).create();
}
function exportInfo_() {
  const sync = readSync_(); let file = null;
  try { if (sync && sync.fileId) file = fileInfo_(DriveApp.getFileById(sync.fileId)); } catch (e) {}
  const d = +(P.getProperty('EXPORT_DIRTY') || 0);
  return { ready: !!sync, file, synced: sync ? sync.at : '', last: setting_('마지막 엑셀 저장'), error: setting_('엑셀 저장 오류'), pending: d ? fmtT_(new Date(d)) : '' };
}
/** 화면에 안 보이는 계산 칸만 조용히 고침 (버전·변경이력 그대로) */
function setSilent_(t, id, ch) {
  const r = rows_(t).find((x) => x.id === id); if (!r) return;
  const cols = SCHEMA[t].concat(META), sh = db_().getSheetByName(t);
  const next = Object.assign(clean_(r), ch);
  sh.getRange(r._row, 1, 1, cols.length).setValues([cols.map((c) => (next[c] == null ? '' : next[c]))]);
  Object.assign(r, ch); // 읽어 둔 표를 그 자리에서 고침 (표 전체를 다시 읽지 않음)
}
/** 웹이 원본일 때 엑셀 수식과 같은 계산: 계약 공급가·부가세·상태, 청구 공급가·부가세·수금액, 점검 기한·단계 */
function derive_(fresh) {
  if ((setting_('원본') || '엑셀') !== '웹') return;
  if (!fresh) TBL_ = {};
  const days = +setting_('보고서기한일수') || 30;
  const bills = live_('청구'), cons = live_('계약'), C = {}; cons.forEach((c) => { C[c.id] = c; });
  const nk = (cid) => key_((C[cid] || {}).계약명);
  const ch = (t, r, o) => { const diff = {}; Object.keys(o).forEach((k) => { if (String(r[k]) !== String(o[k])) diff[k] = o[k]; }); if (Object.keys(diff).length) { setSilent_(t, r.id, diff); Object.assign(r, diff); changed++; } };
  let changed = 0;
  bills.forEach((b) => { const sup = b.청구액 ? Math.round(b.청구액 / 1.1) : 0; ch('청구', b, { 공급가: sup, 부가세: b.청구액 ? b.청구액 - sup : 0, 수금액: b.청구액 && b.입금일 ? b.청구액 : 0 }); });
  const paidByName = {}; bills.forEach((b) => { const k = nk(b.계약id); paidByName[k] = (paidByName[k] || 0) + b.수금액; });
  cons.forEach((c) => {
    const sup = c.계약금 ? Math.round(c.계약금 / 1.1) : 0, paid = c.이월수금 + (paidByName[key_(c.계약명)] || 0);
    const goal = c.계약금 + (c.종합보고서구분 === '별도' ? c.종합보고서금액 : 0);
    const auto = !c.계약금 ? (c.상태 || '진행중') : paid >= goal ? '완료' : '진행중';
    ch('계약', c, { 공급가: sup || c.공급가, 부가세: c.계약금 ? c.계약금 - sup : c.부가세, 상태: c.상태지정 === '완료' || c.상태지정 === '진행중' ? c.상태지정 : auto }); // 완료 처리·진행으로 복원한 계약은 그 상태 유지
  });
  const RV = {}; // 보고서검토관리 → 점검이력 「내부검토」 (엑셀 L열과 같은 규칙)
  live_('검토').forEach((r) => { const k = nk(r.계약id) + '|' + key_(r.항목); const a = (RV[k] = RV[k] || { n: 0, done: 0, ing: 0 }); a.n++; if (r.상태 === '검토완료') a.done++; if (r.상태 === '검토중') a.ing++; });
  live_('점검').forEach((x) => {
    const g = RV[nk(x.계약id) + '|' + key_(x.항목)];
    const iv = !g ? '' : g.done ? '검토완료' : g.ing ? '검토중' : '수정요청';
    if (x.비고 === '완료보관에서 가져옴' || (x.단계 === '완료' && !x.발송일)) return; // 엑셀 발송일 칸에 날짜 대신 글자(예: 인상1)가 있어 완료인 줄은 그대로
    const k = nk(x.계약id), item = key_(x.항목);
    const link = bills.find((b) => b.발행일 && nk(b.계약id) === k && String(b.연결점검 || '').split(/\s*,\s*/).some((v) => key_(v) === item));
    const rv = iv, issued = link && link.발행일, paid = link && link.입금일;
    const st = x.발송일 ? '완료' : rv === '검토완료' && paid ? '발송 대기' : rv === '검토완료' && issued ? '입금 대기' : rv === '검토완료' ? '청구 대기' : (rv === '검토중' || rv === '수정요청') ? '내부검토중' : x.점검일 ? '보고서 작성중' : x.예정일 ? '점검 예정' : '점검 전';
    const due = x.점검일 ? iso_(new Date(toDate_(x.점검일).getTime() + days * 864e5)) : '';
    ch('점검', x, { 단계: st, 기한: due, 내부검토: iv });
  });
  if (changed) bumpVer_();
}

// ---- xlsx 파일 다루기 (압축 안의 XML 을 직접 고침) ----
function attr_(tag, name) { const m = tag.match(new RegExp('\\s' + name.replace(':', '\\:') + '="([^"]*)"')); return m ? m[1] : null; }
function xmlDec_(s) { return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCharCode(parseInt(n, 16))).replace(/&amp;/g, '&'); }
function xmlEnc_(s) { return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function colNum_(L) { let n = 0; for (let i = 0; i < L.length; i++) n = n * 26 + L.charCodeAt(i) - 64; return n; }
function colName_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function xlOpen_(blob) {
  const z = blob.copyBlob ? blob.copyBlob() : blob; if (z.setContentType) z.setContentType('application/zip');
  const parts = Utilities.unzip(z), map = {}, txt = {}, dirty = {};
  parts.forEach((p) => { map[p.getName()] = p; });
  const get = (n) => (n in txt ? txt[n] : (txt[n] = map[n].getDataAsString('UTF-8')));
  const rid = {}, sheets = {};
  get('xl/_rels/workbook.xml.rels').replace(/<Relationship\b[^>]*>/g, (m) => { const tg = attr_(m, 'Target') || ''; rid[attr_(m, 'Id')] = tg.charAt(0) === '/' ? tg.slice(1) : 'xl/' + tg; return m; });
  get('xl/workbook.xml').replace(/<sheet\b[^>]*>/g, (m) => { sheets[xmlDec_(attr_(m, 'name'))] = rid[attr_(m, 'r:id')]; return m; });
  let sst = null;
  return {
    sheetXml(name) { const p = sheets[name]; if (!p || !map[p]) throw new Error('엑셀에 「' + name + '」 시트가 없습니다.'); return get(p); },
    putSheet(name, xml) { txt[sheets[name]] = xml; dirty[sheets[name]] = 1; },
    names() { return Object.keys(sheets); },
    part(path) { return map[path] ? get(path) : null; },
    putPart(path, xml) { txt[path] = xml; dirty[path] = 1; },
    drop(path) { const i = parts.findIndex((p) => p.getName() === path); if (i >= 0) parts.splice(i, 1); delete map[path]; },
    strings() {
      if (sst) return sst; sst = [];
      if (map['xl/sharedStrings.xml']) get('xl/sharedStrings.xml').replace(/<si>([\s\S]*?)<\/si>|<si\/>/g, (m, inner) => { let t = ''; String(inner || '').replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (mm, x) => { t += x; return mm; }); sst.push(xmlDec_(t)); return m; });
      return sst;
    },
    toBlob(name, type) {
      const out = parts.map((p) => { const n = p.getName(); if (!dirty[n]) return p; const nb = Utilities.newBlob('', 'application/xml', n); nb.setDataFromString(txt[n], 'UTF-8'); return nb; });
      const zz = Utilities.zip(out, name); zz.setContentType(type); return zz;
    },
  };
}
/** 업무일지 시트를 웹 자료로 다시 그림 (그 해 달력: 월 제목줄 + 날짜줄, 같은 날 업무가 여러 건이면 아래에 빈 날짜 줄 추가) */
function xlJournal_(book, rows, year) {
  let xml = book.sheetXml('업무일지');
  const m = xml.match(/<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/);
  if (!m) throw new Error('업무일지 시트 형식을 읽지 못했습니다.');
  const rowXml = {}; m[1].replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (rx) => { rowXml[+attr_(rx.match(/^<row\b[^>]*>/)[0], 'r')] = rx; return rx; });
  const st = (rx, col) => { const c = (rx || '').match(new RegExp('<c r="' + col + '\\d+"[^>]*?s="(\\d+)"')); return c ? c[1] : null; };
  const H = rowXml[7] || '', Dr = rowXml[8] || '';
  const hf = ((H.match(/<f>([\s\S]*?)<\/f>/) || [])[1]) || '"▶ 1월      업무 "&amp;COUNTA(G8:G38)&amp;"건"';
  const S = {}; 'CDEFGHIJ'.split('').forEach((c) => { S['h' + c] = st(H, c); S['d' + c] = st(Dr, c); });
  const sa = (k) => (S[k] != null ? ' s="' + S[k] + '"' : '');
  const byDay = {}; rows.forEach((r) => { (byDay[r.날짜] = byDay[r.날짜] || []).push(r); });
  const DAYS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  const tx = (c, r, v, k) => (v === '' || v == null ? `<c r="${c}${r}"${sa(k)}/>` : `<c r="${c}${r}"${sa(k)} t="inlineStr"><is><t xml:space="preserve">${xmlEnc_(v)}</t></is></c>`);
  const out = [], merges = []; let r = 7, n = 0;
  for (let mo = 1; mo <= 12; mo++) {
    const h = r++, a = r, b = r; let rowsInMonth = 0;
    const body = [];
    for (let d = 1; d <= DAYS[mo - 1]; d++) {
      const iso = year + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      const list = byDay[iso] || [null];
      list.forEach((e, i) => {
        const rr = r++; rowsInMonth++;
        const lines = e ? Math.max(1, Math.ceil(String(e.내용 || '').length / 55), String(e.내용 || '').split('\n').length) : 1;
        const C = i === 0 ? `<c r="C${rr}"${sa('dC')}><f>IF(MONTH(DATE($D$2,${mo},${d}))&lt;&gt;${mo},"",DATE($D$2,${mo},${d}))</f></c><c r="D${rr}"${sa('dD')} t="str"><f>IF(C${rr}="","",CHOOSE(WEEKDAY(C${rr}),"일","월","화","수","목","금","토"))</f></c>` : `<c r="C${rr}"${sa('dC')}/><c r="D${rr}"${sa('dD')}/>`;
        const E = e || {};
        body.push(`<row r="${rr}" ht="${(15.6 * lines).toFixed(1)}"${lines > 1 ? ' customHeight="1"' : ''} outlineLevel="1">${C}${tx('E', rr, E.현장, 'dE')}${tx('F', rr, E.구분, 'dF')}${tx('G', rr, E.내용, 'dG')}${tx('H', rr, E.상태, 'dH')}${tx('I', rr, E.후속, 'dI')}${tx('J', rr, E.비고, 'dJ')}</row>`);
        if (e) n++;
      });
    }
    const last = a + rowsInMonth - 1;
    const f = hf.replace(/\d+월/, mo + '월').replace(/([GH])\d+:([GH])\d+/g, (x, c1, c2) => `${c1}${a}:${c2}${last}`);
    out.push(`<row r="${h}" ht="21.75" customHeight="1"><c r="C${h}"${sa('hC')} t="str"><f>${f}</f></c>${'DEFGHIJ'.split('').map((c) => `<c r="${c}${h}"${sa('h' + c)}/>`).join('')}</row>`);
    out.push(body.join(''));
    merges.push(`<mergeCell ref="C${h}:J${h}"/>`);
    void b;
  }
  const lastRow = r - 1, L = Math.max(lastRow, 547);
  const keep = Object.keys(rowXml).map(Number).filter((x) => x < 7).sort((a, b) => a - b).map((x) => rowXml[x]).join('');
  const head = m[0].match(/^<sheetData\b[^>]*>/)[0];
  xml = xml.replace(m[0], () => head + keep + out.join('') + '</sheetData>');
  // 월 제목 병합칸 · 범위(조건부 서식, 목록 상자) · 크기
  xml = xml.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, (all, inner) => {
    const old = (inner.match(/<mergeCell\b[^>]*\/>/g) || []).filter((mc) => { const r1 = +(attr_(mc, 'ref') || '').replace(/^[A-Z]+(\d+).*$/, '$1'); return r1 < 7; });
    const list = old.concat(merges); return `<mergeCells count="${list.length}">${list.join('')}</mergeCells>`;
  });
  const i = xml.indexOf('</sheetData>');
  xml = xml.slice(0, i) + xml.slice(i).replace(/(sqref="|<xm:sqref>)([^"<]*)/g, (all, p1, ref) => p1 + ref.replace(/([A-Z]+)7:([A-Z]+)\d+/g, '$17:$2' + L));
  xml = xml.replace(/<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"\/>/, `<dimension ref="$1:$2${L}"/>`);
  book.putSheet('업무일지', xml);
  // 다른 시트(대시보드·오늘할일)가 보는 업무일지 범위
  book.names().forEach((nm) => {
    if (nm === '업무일지') return;
    const x = book.sheetXml(nm); if (x.indexOf('업무일지!') < 0) return;
    const y = x.replace(/업무일지!\$([A-Z]+)\$7:\$([A-Z]+)\$\d+/g, (all, c1, c2) => `업무일지!$${c1}$7:$${c2}$${L}`);
    if (y !== x) book.putSheet(nm, y);
  });
  // 수식 칸 위치가 바뀌었으므로 계산 순서 파일은 지우고 엑셀이 다시 만들게 함
  if (book.part('xl/calcChain.xml') != null) {
    book.drop('xl/calcChain.xml');
    book.putPart('[Content_Types].xml', book.part('[Content_Types].xml').replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
    book.putPart('xl/_rels/workbook.xml.rels', book.part('xl/_rels/workbook.xml.rels').replace(/<Relationship[^>]*Target="(?:\/xl\/)?calcChain\.xml"[^>]*\/>/, ''));
  }
  return n;
}
function xlSheet_(xml, strings) {
  const m = xml.match(/<sheetData\b[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/);
  if (!m) throw new Error('시트 형식을 읽지 못했습니다.');
  const rows = {}, order = []; let changed = false;
  (m[1] || '').replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (rx) => { const n = +attr_(rx.match(/^<row\b[^>]*>/)[0], 'r'); rows[n] = { xml: rx }; order.push(n); return rx; });
  const parse = (n) => {
    const R = rows[n]; if (R.cells) return R;
    const open = R.xml.match(/^<row\b[^>]*?(\/?)>/);
    R.attrs = open[0].replace(/^<row/, '').replace(/\/?>$/, ''); R.cells = {};
    if (!open[1]) R.xml.slice(open[0].length).replace(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, (cx) => { R.cells[colNum_(attr_(cx.match(/^<c\b[^>]*>/)[0], 'r').replace(/\d+/g, ''))] = cx; return cx; });
    return R;
  };
  const cellAt = (n, c) => (rows[n] ? parse(n).cells[c] : null);
  return {
    value(n, c) {
      const cx = cellAt(n, c); if (!cx) return '';
      const t = attr_(cx.match(/^<c\b[^>]*>/)[0], 't'), v = (cx.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      if (t === 's') return v == null ? '' : strings()[+v] || '';
      if (t === 'inlineStr') { let s = ''; cx.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (mm, x) => { s += x; return mm; }); return xmlDec_(s); }
      if (v == null) return '';
      return t === 'str' || t === 'e' || t === 'b' ? xmlDec_(v) : Number(v);
    },
    set(n, c, kind, val) {
      let R = rows[n];
      if (!R) { R = rows[n] = { xml: '', attrs: ' r="' + n + '"', cells: {} }; order.push(n); order.sort((a, b) => a - b); } else parse(n);
      const old = R.cells[c];
      if (old && /<f[\s>\/]/.test(old)) return false; // 수식 칸은 절대 안 고침
      let st = old ? attr_(old.match(/^<c\b[^>]*>/)[0], 's') : null;
      for (let up = n - 1; st == null && up >= n - 5 && up > 0; up--) { const a = cellAt(up, c); if (a) st = attr_(a.match(/^<c\b[^>]*>/)[0], 's'); }
      const ref = colName_(c) + n, sa = st != null ? ' s="' + st + '"' : '';
      R.cells[c] = val === '' || val == null ? '<c r="' + ref + '"' + sa + '/>'
        : kind === 's' ? '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEnc_(val) + '</t></is></c>'
        : '<c r="' + ref + '"' + sa + '><v>' + val + '</v></c>';
      R.dirty = true; changed = true; return true;
    },
    dirty() { return changed; },
    toXml() {
      const body = order.map((n) => {
        const R = rows[n]; if (!R.dirty) return R.xml;
        const cs = Object.keys(R.cells).map(Number).sort((a, b) => a - b).map((c) => R.cells[c]).join('');
        const at = R.attrs.replace(/\sspans="[^"]*"/, '');
        return cs ? '<row' + at + '>' + cs + '</row>' : '<row' + at + '/>';
      }).join('');
      const head = m[0].match(/^<sheetData\b[^>]*?(?=\/?>)/)[0];
      return xml.replace(m[0], () => head + '>' + body + '</sheetData>');
    },
  };
}

// ---- 분할(부분) 입금: 엑셀 비고 칸에 「분할 입금: 금액원 날짜 / …」로 적던 방식 그대로
function partsOf_(s) { return String(s || '').split(/\s*,\s*/).map((x) => { const m = x.match(/^(\d{4}-\d\d-\d\d):(-?\d+)$/); return m ? { d: m[1], a: +m[2] } : null; }).filter(Boolean); }
function partsSum_(s) { return partsOf_(s).reduce((t, p) => t + p.a, 0); }
function partsText_(s, done) { const ps = partsOf_(s); return ps.length ? '분할 입금: ' + ps.map((p) => p.a.toLocaleString('en-US') + '원 ' + p.d).join(' / ') + (done ? ' (완납)' : '') : ''; }
/** 엑셀 비고에서 「분할 입금: …」 부분을 떼어 냄 → [나머지 비고, DB 형식] */
function splitParts_(memo) {
  const m = String(memo || '').match(/분할\s*입금\s*:\s*((?:[\d,]+\s*원\s*\d{4}-\d\d-\d\d\s*(?:\/\s*)?)+)(?:\(완납\))?/);
  if (!m) return [String(memo || ''), ''];
  const parts = []; m[1].replace(/([\d,]+)\s*원\s*(\d{4}-\d\d-\d\d)/g, (x, a, d) => { parts.push(d + ':' + a.replace(/,/g, '')); return x; });
  const rest = String(memo).replace(m[0], '').replace(/^[\s·]+|[\s·]+$/g, '').replace(/\s*·\s*·\s*/g, ' · ');
  return [rest, parts.join(', ')];
}
// ===================== 서버 자동 업데이트 =====================
// Apps Script 에는 「불러오기 코드」만 있고, 이 파일(GitHub server/Code.gs)을 5분마다 받아 실행합니다. (불러오기/Code.gs 참고)
function updateInfo_() {
  const L = typeof loaderInfo_ === 'function' ? loaderInfo_() : { loader: false };
  return { code: SERVER_VERSION, loader: !!L.loader, deployed: L.version ? L.version + (L.at ? ' · ' + fmtT_(new Date(L.at)) : '') : '', checked: L.checked ? fmtT_(new Date(L.checked)) : '', error: L.error || '', auto: !!L.auto };
}
function updateNow_() {
  if (typeof loaderCheck_ !== 'function') return { ok: false, msg: '아직 불러오기 코드가 설치되지 않았습니다. Apps Script 에 불러오기 Code.gs 를 넣고 installLoader 를 실행하세요.' };
  const r = loaderCheck_(false); return Object.assign({ result: { changed: r.changed, version: r.version, waiting: !!r.waiting } }, updateInfo_());
}
// ===================== 공통 도우미 =====================
const NON_TARGET = ['초기점검', '종합보고서', '정밀안전점검', '공사재개전점검', '착수전점검'];
function cleanHead_(h) { return String(h == null ? '' : h).replace(/\s+/g, ' ').replace(/▼/g, '').replace(/선택/g, '').replace(/\s*\((자동|직접.*?|진행 중)\)\s*/g, '').trim(); }
function table_(vals, headRow, startRow, key) {
  if (!vals || vals.length < headRow) return [];
  const heads = vals[headRow - 1].map(cleanHead_);
  const ki = heads.findIndex((h) => h.indexOf(key) === 0);
  const out = [];
  for (let r = startRow - 1; r < vals.length; r++) {
    const row = vals[r];
    if (ki < 0 || row[ki] === '' || row[ki] == null) continue;
    const o = {};
    heads.forEach((h, i) => { if (h && !(h in o)) o[h] = row[i]; });
    Object.defineProperty(o, '_r', { value: r + 1, enumerable: false }); // 엑셀 행 번호 (웹→엑셀 저장 때 같은 줄을 고치려고)
    out.push(o);
  }
  return out;
}
function g_(o, ...names) { for (const n of names) for (const k in o) if (k.indexOf(n) === 0) return o[k]; return null; }
function num_(v) { if (typeof v === 'number') return isNaN(v) ? 0 : v; const n = parseFloat(String(v == null ? '' : v).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
function dt_(v) { return v instanceof Date && !isNaN(v) ? v : null; }
function toDate_(v) { if (dt_(v)) return v; const m = String(v || '').match(/^(\d{4})-(\d\d)-(\d\d)(?: (\d\d):(\d\d))?/); return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : null; }
function iso_(v) { const d = dt_(v) || (typeof v === 'string' && /^\d{4}-\d\d-\d\d/.test(v) ? toDate_(v) : null); return d ? Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd') : ''; }
function fmtT_(d) { return dt_(d) ? Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(d || ''); }
function s_(v) { const t = v instanceof Date ? iso_(v) : String(v == null ? '' : v).trim(); return /^#/.test(t) ? '' : t.replace(/\s*\n\s*/g, '\n'); }
function one_(v) { return s_(v).replace(/\n/g, ' '); }
function split_(item) { const m = String(item || '').replace(/\s/g, '').match(/^(.+?)(\d+)차$/); return m ? [m[1], +m[2]] : [String(item || '').replace(/\s/g, ''), null]; }
function key_(s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); }
function corp_(s) { return String(s || '').replace(/\(주\)|㈜|주식회사|\(유\)|유한회사/g, ''); }
function digits_(s) { const d = String(s || '').replace(/\D/g, ''); return d.length >= 10 ? d : ''; }
function hashS_(s) { let h = 5381; s = String(s); for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36).toUpperCase(); }
