/**
 * 업무관리 서버 · 불러오기 코드 (@uto-erp-loader)
 * ─────────────────────────────────────────────
 * 이 파일은 직접 고치지 않습니다. 실제 서버 코드는 GitHub(inseon10342-alt/uto-erp)의 server/Code.gs 이고,
 * 여기서는 그것을 받아 실행만 합니다.
 *  · 5분마다 GitHub 의 server/VERSION 을 확인해 바뀌었으면 새 코드를 받습니다.
 *  · 새 코드는 먼저 시험 실행해 보고, 문제가 없을 때만 바꿉니다. 문제가 있으면 마지막으로 잘 되던 코드를 계속 씁니다.
 *  · 받은 코드는 Drive 「업무관리 서버코드」 폴더에 보관합니다(마지막으로 잘 되던 코드).
 * 처음 한 번: 위 함수 목록에서 installLoader 를 실행해 권한을 허용합니다.
 */
const LOADER = { RAW: 'https://raw.githubusercontent.com/inseon10342-alt/uto-erp/main/server/', FOLDER: '업무관리 서버코드', CHUNK: 90000 };
const LP_ = PropertiesService.getScriptProperties();
let APP_ = null;

// ---- 바깥에서 부르는 입구 (웹 앱 · 트리거 · 편집기)
function doGet(e) { return app_().doGet(e); }
function doPost(e) { return app_().doPost(e); }
function dailyBackup() { return app_().dailyBackup(); }
function exportTick() { return app_().exportTick(); }
function setup() { return app_().setup(); }
function loaderTick() { try { loaderCheck_(false); } catch (e) { loaderFail_(e); } }

/** 처음 한 번 실행: GitHub 에서 서버 코드를 받아 보관하고, 5분마다 확인하는 트리거를 겁니다 */
function installLoader() {
  ScriptApp.getProjectTriggers().forEach((t) => { if (['loaderTick', 'updateTick'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('loaderTick').timeBased().everyMinutes(5).create();
  const r = loaderCheck_(true);
  Logger.log('불러오기 연결 완료 · 서버 코드 ' + r.version + ' · 이제 「배포 → 배포 관리 → 새 버전」을 한 번만 하면 끝입니다.');
}

// ---- 서버 코드 꺼내 오기: 캐시 → Drive 보관본 → GitHub 순서
function app_() {
  if (APP_) return APP_;
  const C = CacheService.getScriptCache(), meta = JSON.parse(C.get('ldr:meta') || 'null');
  let src = null;
  if (meta && meta.n) { const got = C.getAll(Array.from({ length: meta.n }, (_, i) => 'ldr:' + meta.v + ':' + i)); const parts = Array.from({ length: meta.n }, (_, i) => got['ldr:' + meta.v + ':' + i]); if (parts.every((x) => x != null)) src = parts.join(''); }
  if (src == null) { const f = keptFile_(); if (f) { src = f.getBlob().getDataAsString('UTF-8'); cacheSrc_(src, LP_.getProperty('LDR_VER') || 'drive'); } }
  if (src == null) { const r = loaderCheck_(true); return r.app; }
  APP_ = build_(src);
  return APP_;
}
/** 서버 코드를 함수로 만들어 입구들을 꺼냄 (문법 오류면 여기서 실패) */
function build_(src) {
  const make = new Function(src + '\n;return { doGet: doGet, doPost: doPost, dailyBackup: dailyBackup, exportTick: exportTick, setup: setup, version: (typeof SERVER_VERSION === "string" ? SERVER_VERSION : "") };');
  const app = make();
  // 서버 코드에 해당 함수가 없으면 이 불러오기 파일의 같은 이름 함수가 잡히므로, 그런 경우는 받아들이지 않음
  const mine = { doGet: doGet, doPost: doPost, dailyBackup: dailyBackup, exportTick: exportTick, setup: setup };
  Object.keys(mine).forEach((k) => { if (typeof app[k] !== 'function' || app[k] === mine[k]) throw new Error('서버 코드에 ' + k + ' 함수가 없습니다.'); });
  return app;
}
function cacheSrc_(src, v) {
  const C = CacheService.getScriptCache(), o = {}, n = Math.ceil(src.length / LOADER.CHUNK);
  for (let i = 0; i < n; i++) o['ldr:' + v + ':' + i] = src.slice(i * LOADER.CHUNK, (i + 1) * LOADER.CHUNK);
  C.putAll(o, 21600); C.put('ldr:meta', JSON.stringify({ v, n }), 21600);
}
function keptFile_() { const id = LP_.getProperty('LDR_FILE'); if (!id) return null; try { const f = DriveApp.getFileById(id); return f.isTrashed() ? null : f; } catch (e) { return null; } }
function raw_(path) {
  const r = UrlFetchApp.fetch(LOADER.RAW + path + '?t=' + Date.now(), { muteHttpExceptions: true, headers: { 'Cache-Control': 'no-cache' } });
  if (r.getResponseCode() !== 200) throw new Error('GitHub 에서 ' + path + ' 를 읽지 못했습니다 (' + r.getResponseCode() + ')');
  return r.getContentText();
}

/** GitHub 버전 확인 → 새 코드면 시험 실행 후 바꿈. 서버 코드의 「지금 업데이트 확인」 버튼도 이것을 부릅니다 */
function loaderCheck_(force) {
  const v = raw_('VERSION').trim(), cur = LP_.getProperty('LDR_VER') || '';
  LP_.setProperty('LDR_CHECKED', new Date().toISOString());
  if (!force && v === cur && keptFile_()) return { changed: false, version: v };
  const lock = LockService.getScriptLock(); if (!lock.tryLock(10000)) return { changed: false, version: cur, busy: true };
  try {
    const src = raw_('Code.gs');
    if (src.indexOf('@uto-erp-server') < 0) throw new Error('받은 코드에 서버 표시(@uto-erp-server)가 없어 바꾸지 않았습니다.');
    if (src.indexOf("SERVER_VERSION = '" + v + "'") < 0) { if (force && !keptFile_()) throw new Error('GitHub 가 아직 옛 코드를 주고 있습니다. 1~2분 뒤 다시 실행하세요.'); return { changed: false, version: cur, waiting: true }; }
    const app = build_(src); // 시험 실행: 문법 오류·필수 함수 없음이면 여기서 멈추고 지금 코드를 계속 씀
    let f = keptFile_();
    if (f) f.setContent(src);
    else { const it = DriveApp.getFoldersByName(LOADER.FOLDER), dir = it.hasNext() ? it.next() : DriveApp.createFolder(LOADER.FOLDER); f = dir.createFile('server-Code.js', src, 'text/plain'); LP_.setProperty('LDR_FILE', f.getId()); }
    f.setDescription('GitHub server/Code.gs ' + v + ' · ' + new Date().toISOString());
    cacheSrc_(src, v);
    LP_.setProperty('LDR_VER', v); LP_.setProperty('LDR_AT', new Date().toISOString()); LP_.deleteProperty('LDR_ERR');
    APP_ = app;
    return { changed: v !== cur, version: v, app };
  } finally { lock.releaseLock(); }
}
function loaderFail_(e) {
  const prev = LP_.getProperty('LDR_ERR') || '', msg = e.message;
  LP_.setProperty('LDR_ERR', new Date().toISOString() + ' · ' + msg);
  if (prev.slice(25) !== msg) { try { MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: '[업무관리] 서버 코드 업데이트 실패', body: msg + '\n\n서버는 마지막으로 잘 되던 코드로 계속 동작합니다.' }); } catch (x) {} }
}
/** 서버 화면(시스템 설정)에 보여 줄 상태 */
function loaderInfo_() { return { loader: true, version: LP_.getProperty('LDR_VER') || '', at: LP_.getProperty('LDR_AT') || '', checked: LP_.getProperty('LDR_CHECKED') || '', error: LP_.getProperty('LDR_ERR') || '', auto: ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === 'loaderTick') }; }
