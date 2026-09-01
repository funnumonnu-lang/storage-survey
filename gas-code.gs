/**
 * 収納量調査 回答収集 — バックエンド（Google Apps Script／API専用）
 *
 * 案件ごとにスプレッドシートを分け、1つのスクリプトから切り替えて使う。
 *
 *   マスター（このスクリプトを紐づけたスプレッドシート）
 *     案件         … キー／案件名／スプレッドシートID／備考
 *     マスター設定  … 管理用あいことば
 *
 *   案件ごとのスプレッドシート
 *     設定 ／ ②什器サイズ ／ 部門マスタ ／ 提出 ／ 回答ログ ／ 採用 ／ ①回答
 *
 * フロントは GitHub Pages に置き、?p=案件キー を付けて呼ぶ。
 * デプロイ：ウェブアプリ／実行=自分／アクセス=全員
 */

var MAX_TIERS = 7;

var S_CONF   = '設定';
var S_UNIT   = '②什器サイズ';
var S_DEPT   = '部門マスタ';
var S_SUB    = '提出';
var S_LOG    = '回答ログ';
var S_ADOPT  = '採用';
var S_OUT    = '①回答';

var PRESENCE_YES   = 'ある';
var PRESENCE_OTHER = '収納物はあるが、自分の部署以外の物品である';
var DISPOSAL_50    = '50%程度廃棄もしくは外部倉庫へ移管可能（50%は今後も保管）';
var DISPOSAL_KEEP  = '殆どすべて今後も保管';
var TIMING_NONE    = '使用タイミングは決まっていない';

var ST_OPEN = '入力中';
var ST_DONE = '提出済み';

var S_PROJ   = '案件';
var S_MCONF  = 'マスター設定';
var PROJ_COLS  = ['キー', '案件名', 'スプレッドシートID', '備考'];
var MCONF_COLS = ['項目', '値'];

var UNIT_COLS  = ['収納番号', '什器幅(mm)', 'FM', '什器奥行(mm)', '什器高さ(mm)', '什器仕様', '段数', '部門'];
var DEPT_COLS  = ['部門', '課', '出力先'];
var SUB_COLS   = ['提出ID', '部門', '課', '回答者', '状態', '開始日時', '提出日時'];
var LOG_COLS   = ['提出ID', '回答者', '収納番号', '段', '部門', '課', '収納物有無', '品目',
                  '整理・移動', '使用タイミング', '保管理由', '残50%タイミング', '残50%保管理由', '更新日時'];
var ADOPT_COLS = ['収納番号', '採用する提出ID', '決定日時'];

/* ================================================================
 * セットアップ
 * ================================================================ */

/** 最初に1回だけ実行する。マスター側のシートを作る */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートに紐づいたスクリプトとして作成してください（拡張機能 > Apps Script）。');

  sheetWithHeader_(ss, S_PROJ, PROJ_COLS);

  var mc = ss.getSheetByName(S_MCONF);
  if (!mc) {
    mc = ss.insertSheet(S_MCONF);
    mc.getRange(1, 1, 1, 2).setValues([MCONF_COLS]).setFontWeight('bold').setBackground('#F2F2F2');
    mc.getRange(2, 1, 1, 2).setValues([['管理用あいことば', 'change-me']]);
    mc.setColumnWidth(1, 200);
    mc.setColumnWidth(2, 400);
  }

  var first = ss.getSheets()[0];
  if (first.getName() === 'シート1' && first.getLastRow() === 0) ss.deleteSheet(first);
  dropCache_();

  SpreadsheetApp.getUi().alert(
    'マスターを作成しました。\n\n' +
    '1. 「マスター設定」の管理用あいことばを変更\n' +
    '2. 案件ごとに新しいスプレッドシートを作り、そのIDを控える\n' +
    '3. デプロイ > 新しいデプロイ > ウェブアプリ\n' +
    '   実行=自分 ／ アクセス=全員\n' +
    '4. 発行された /exec を index.html の API_URL に貼る\n' +
    '5. Pages を ?admin=1 で開き、案件を登録して「初期化」を押す'
  );
}

/** 案件のスプレッドシートにデータ用シートを作る（管理画面から呼ぶ） */
function setupProject(pass, key) {
  assertPass_(pass);
  var ss = projectSs_(key);

  sheetWithHeader_(ss, S_UNIT,  UNIT_COLS);
  sheetWithHeader_(ss, S_DEPT,  DEPT_COLS);
  sheetWithHeader_(ss, S_SUB,   SUB_COLS);
  sheetWithHeader_(ss, S_LOG,   LOG_COLS);
  sheetWithHeader_(ss, S_ADOPT, ADOPT_COLS);

  var conf = ss.getSheetByName(S_CONF);
  if (!conf) {
    conf = ss.insertSheet(S_CONF);
    conf.getRange(1, 1, 1, 2).setValues([['項目', '値']]).setFontWeight('bold');
    conf.getRange(2, 1, DEFAULT_CONF_.length, 2).setValues(DEFAULT_CONF_);
    conf.setColumnWidth(1, 200);
    conf.setColumnWidth(2, 700);
    conf.getRange(2, 2, DEFAULT_CONF_.length, 1).setWrap(true);
  }

  var first = ss.getSheets()[0];
  if (first.getName() === 'シート1' && first.getLastRow() === 0) ss.deleteSheet(first);
  dropCache_();
  return { name: ss.getName(), url: ss.getUrl() };
}

var DEFAULT_CONF_ = [
  ['案件名', '収納量調査'],
  ['収納物品', [
    '書類（図面など）',
    '支払帳票（見積書・注文請書・納品書・請求書など）',
    '契約書（賃貸借契約・業務委託契約など）',
    'カタログ・営業資料（サンプル資料、商品パンフレット、メーカー・ベンダーカタログなど）',
    '一時保管書類',
    '個人情報の入った書類',
    'その他書類',
    '文房具・オフィス備品',
    '書籍',
    '防災備品'
  ].join('\n')],
  ['その他（自由入力）を出す', 'はい'],
  ['整理・移動', [
    '100%廃棄もしくは外部倉庫へ移管可能', DISPOSAL_50, DISPOSAL_KEEP
  ].join('\n')],
  ['使用タイミング', [
    '3ヶ月以内に使用する', '3～6ヶ月以内に使用する', '6ヶ月～1年以上先に使用する', TIMING_NONE
  ].join('\n')],
  ['保管理由', [
    'いつか使用する可能性があるため（急な顧客対応等）',
    'アーカイブ（参考資料）として残したいため',
    '社内ルールで保管が定められているため',
    '電子化が必要な書類のため（電子化後であれば移管可能）',
    '廃棄していいのか不明のため'
  ].join('\n')],
  ['他の人が登録中の収納を知らせる', 'はい']
];

function sheetWithHeader_(ss, name, cols) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols])
      .setFontWeight('bold').setBackground('#F2F2F2');
    sh.setFrozenRows(1);
  }
  return sh;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('収納量調査')
    .addItem('マスターを初期化する', 'setup')
    .addToUi();
}

/* ================================================================
 * API（GitHub Pages から呼ばれる）
 *
 * 読み取りは JSONP（?callback=）、書き込みは POST。
 * GAS は他オリジンからの通常の fetch を受けにくいので、この2本立てにしている。
 * ================================================================ */

/** 外から呼べる関数だけをここに並べる */
function api_() {
  return {
    getBootstrap: getBootstrap,
    startSubmission: startSubmission,
    saveTier: saveTier,
    submitAll: submitAll,
    reopenSubmission: reopenSubmission,
    getReview: getReview,
    setAdopt: setAdopt,
    adminReopen: adminReopen,
    getSetupInfo: getSetupInfo,
    importUnits: importUnits,
    importDepts: importDepts,
    clearCache: clearCache,
    adminBuild: adminBuild,
    listProjects: listProjects,
    saveProject: saveProject,
    removeProject: removeProject,
    setupProject: setupProject
  };
}

function dispatch_(fn, args) {
  var map = api_();
  if (!map[fn]) throw new Error('unknown function: ' + fn);
  return map[fn].apply(null, args || []);
}

/** 案件を必要としない関数 */
var MASTER_FNS_ = ['listProjects', 'saveProject', 'removeProject', 'setupProject'];

function run_(fn, argsJson, projectKey) {
  try {
    useProject_(MASTER_FNS_.indexOf(fn) >= 0 ? '' : projectKey);
    return { ok: true, v: dispatch_(fn, argsJson ? JSON.parse(argsJson) : []) };
  } catch (err) {
    return { ok: false, e: String((err && err.message) || err) };
  }
}

function jsonp_(callback, payload) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (!p.fn) {
    return ContentService.createTextOutput(
      '収納量調査 回答収集API。回答画面は GitHub Pages 側のURLを開いてください。'
    ).setMimeType(ContentService.MimeType.TEXT);
  }
  var res = run_(p.fn, p.args, p.p);
  return p.callback ? jsonp_(p.callback, res) : json_(res);
}

function doPost(e) {
  var body = {};
  try {
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, e: 'リクエストを読めませんでした。' });
  }
  var res = run_(body.fn, body.args ? JSON.stringify(body.args) : '[]', body.p);
  return body.callback ? jsonp_(body.callback, res) : json_(res);
}

/* ================================================================
 * 共通
 * ================================================================ */

/* ---- 案件の切り替え ---------------------------------------------- *
 * リクエストごとに「どの案件か」を決め、ss_() がその案件のスプレッドシートを返す。
 * ------------------------------------------------------------------ */
var CUR_KEY_ = '';        /* 現在の案件キー */
var CUR_SS_  = null;      /* 解決済みのスプレッドシート */

function master_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function masterRows_(name) {
  var sh = master_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

/** 案件マスタ [{key,name,sheetId,note}] */
function projects_() {
  var out = [];
  masterRows_(S_PROJ).forEach(function(r){
    var k = String(r[0] || '').trim();
    if (!k) return;
    out.push({ key: k, name: String(r[1] || ''), sheetId: String(r[2] || '').trim(),
               note: String(r[3] || '') });
  });
  return out;
}

function project_(key) {
  var ps = projects_();
  for (var i = 0; i < ps.length; i++) if (ps[i].key === String(key)) return ps[i];
  return null;
}

/** 指定した案件のスプレッドシート */
function projectSs_(key) {
  var p = project_(key);
  if (!p) throw new Error('案件「' + key + '」が登録されていません。');
  if (!p.sheetId) throw new Error('案件「' + key + '」にスプレッドシートIDが入っていません。');
  try {
    return SpreadsheetApp.openById(p.sheetId);
  } catch (e) {
    throw new Error('案件「' + key + '」のスプレッドシートを開けません。IDを確認してください。');
  }
}

/** リクエストの入口で呼ぶ */
function useProject_(key) {
  key = String(key || '').trim();
  if (CUR_KEY_ === key && CUR_SS_) return;
  CUR_KEY_ = key;
  CUR_SS_ = null;
  MEMO_ = {};
}

function ss_() {
  if (!CUR_KEY_) throw new Error('案件が指定されていません（URLに ?p=案件キー が必要です）。');
  if (!CUR_SS_) CUR_SS_ = projectSs_(CUR_KEY_);
  return CUR_SS_;
}

/* ---- キャッシュ（設定・什器・部門マスタ。シート読み込みが起動の重さの大半） ---- */
var CACHE_TTL = 600;                 /* 秒。シートを直接編集した場合はこの時間内に反映される */
var MEMO_ = {};                      /* 同一実行内の重複読み込みも防ぐ */

function cached_(key, fn) {
  key = CUR_KEY_ + ':' + key;
  if (MEMO_[key] !== undefined) return MEMO_[key];
  var c = null;
  try { c = CacheService.getScriptCache(); } catch (e) {}
  if (c) {
    var hit = c.get(key);
    if (hit) {
      try { MEMO_[key] = JSON.parse(hit); return MEMO_[key]; } catch (e) {}
    }
  }
  var v = fn();
  MEMO_[key] = v;
  if (c) { try { c.put(key, JSON.stringify(v), CACHE_TTL); } catch (e) {} }
  return v;
}

function dropCache_() {
  MEMO_ = {};
  try {
    CacheService.getScriptCache().removeAll(
      [CUR_KEY_ + ':cfg', CUR_KEY_ + ':units', CUR_KEY_ + ':depts']);
  } catch (e) {}
}

/** 管理画面から即時反映したいとき */
function clearCache(pass) {
  assertPass_(pass);
  dropCache_();
  return true;
}

function rows_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function now_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'); }

function confAll_() {
  return cached_('cfg', function(){
    var m = {};
    rows_(S_CONF).forEach(function(r){
      var k = String(r[0]).trim();
      if (k) m[k] = String(r[1]);
    });
    return m;
  });
}
function confValue_(key) {
  var m = confAll_();
  return m[key] === undefined ? '' : m[key];
}
function confList_(key) {
  return confValue_(key).split('\n').map(function(s){ return s.trim(); }).filter(function(s){ return s; });
}
function confYes_(key) { return /^(はい|yes|true|1|○)$/i.test(confValue_(key).trim()); }

function toHalf_(v) {
  return String(v == null ? '' : v).replace(/[０-９]/g, function(c){
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
}
function pad3_(v) {
  var s = toHalf_(v).replace(/[^0-9]/g, '');
  while (s.length < 3) s = '0' + s;
  return s;
}
function tiersOf_(tierCell, spec) {
  var m = toHalf_(tierCell).match(/(\d+)/);
  var n = m ? parseInt(m[1], 10) : 0;
  if (!n) {
    var m2 = toHalf_(spec).match(/(\d+)\s*段/);
    n = m2 ? parseInt(m2[1], 10) : 0;
  }
  if (!n || n < 1) return 0;
  return Math.min(n, MAX_TIERS);
}

function units_() { return cached_('units', unitsRaw_); }

function unitsRaw_() {
  var out = [];
  rows_(S_UNIT).forEach(function(r){
    var raw = String(r[0] || '').trim();
    if (!raw) return;
    var m = raw.match(/^([A-Za-z]+)\s*(\d+)$/);
    if (!m) return;
    var alpha = m[1].toUpperCase(), num = pad3_(m[2]);
    out.push({
      no: alpha + num, alpha: alpha, num: num,
      w: r[1], d: r[3], h: r[4],
      spec: String(r[5] || ''),
      tiers: tiersOf_(r[6], r[5]),
      dept: String(r[7] || '')
    });
  });
  return out;
}

function depts_() { return cached_('depts', deptsRaw_); }

function deptsRaw_() {
  var map = {}, order = [], last = '';
  rows_(S_DEPT).forEach(function(r){
    var dn = String(r[0] || '').trim();
    var sn = String(r[1] || '').trim();
    var col = String(r[2] || '').trim().toUpperCase();
    if (dn) last = dn; else dn = last;
    if (!dn) return;
    if (!map[dn]) { map[dn] = { name: dn, sections: [], col: col === 'H' ? 'H' : 'G' }; order.push(dn); }
    if (col === 'H') map[dn].col = 'H';
    if (sn && map[dn].sections.indexOf(sn) < 0) map[dn].sections.push(sn);
  });
  return order.map(function(n){ return map[n]; });
}

/* ================================================================
 * 提出
 * ================================================================ */

function label_(dept, section, who) {
  var base = (dept || '') + (section ? ' ' + section : '');
  if (!base) base = who || '(不明)';
  else if (who) base += '（' + who + '）';
  return base;
}

function subs_() {
  var m = {}, order = [];
  rows_(S_SUB).forEach(function(r){
    var id = String(r[0] || '').trim();
    if (!id) return;
    var dept = String(r[1] || ''), section = String(r[2] || ''), who = String(r[3] || '');
    m[id] = { id: id, dept: dept, section: section, who: who,
              label: label_(dept, section, who),
              state: String(r[4] || ST_OPEN),
              startedAt: String(r[5] || ''), submittedAt: String(r[6] || '') };
    order.push(id);
  });
  return { map: m, order: order };
}

/** 全回答ログ {提出ID: {収納番号: {段: rec}}} */
function logBy_() {
  var out = {};
  rows_(S_LOG).forEach(function(r){
    var sid = String(r[0] || '').trim();
    var no = String(r[2] || '').trim().toUpperCase();
    var tier = parseInt(toHalf_(r[3]), 10);
    if (!sid || !no || !tier) return;
    if (!out[sid]) out[sid] = {};
    if (!out[sid][no]) out[sid][no] = {};
    out[sid][no][tier] = {
      dept: String(r[4] || ''), section: String(r[5] || ''),
      presence: String(r[6] || ''), item: String(r[7] || ''),
      disposal: String(r[8] || ''), timing: String(r[9] || ''),
      reason: String(r[10] || ''), timing50: String(r[11] || ''),
      reason50: String(r[12] || ''), at: String(r[13] || '')
    };
  });
  return out;
}

function answered_(x) {
  if (!x || !x.presence) return false;
  if (x.presence === PRESENCE_OTHER) return true;
  return !!(x.dept && x.section);
}

/** 1つの提出の中の進捗 */
function progressOf_(us, byUnit) {
  var out = {};
  us.forEach(function(u){
    var t = (byUnit || {})[u.no] || {};
    var filled = 0, pending = false;
    for (var n = 1; n <= u.tiers; n++) {
      if (!answered_(t[n])) continue;
      filled++;
      if (t[n].presence === PRESENCE_OTHER) pending = true;
    }
    if (filled) out[u.no] = { filled: filled, done: filled === u.tiers && u.tiers > 0, pending: pending };
  });
  return out;
}

/** 他の提出が触っている収納 {収納番号: [回答者,…]} */
function othersOf_(myId, lb, sm) {
  var out = {};
  for (var sid in lb) {
    if (sid === myId) continue;
    var who = (sm.map[sid] && sm.map[sid].label) || '(不明)';
    for (var no in lb[sid]) {
      if (!out[no]) out[no] = [];
      if (out[no].indexOf(who) < 0) out[no].push(who);
    }
  }
  return out;
}

function newId_() {
  return 'S' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') +
         '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
}

/** 新しい提出を始める。所属は必須、氏名は任意 */
function startSubmission(dept, section, who) {
  dept = String(dept || '').trim();
  section = String(section || '').trim();
  who = String(who || '').trim();
  if (!dept || !section) throw new Error('部門と課を選んでください。');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id = newId_();
    ss_().getSheetByName(S_SUB).appendRow([id, dept, section, who, ST_OPEN, now_(), '']);
    SpreadsheetApp.flush();
    return id;
  } finally { lock.releaseLock(); }
}

/** 起動時。submitId を渡すと自分の進捗も返す */
function getBootstrap(submitId) {
  var all = units_();
  var us = all.filter(function(u){ return u.tiers > 0; });
  var sm = subs_(), lb = logBy_();
  var me = submitId && sm.map[submitId] ? sm.map[submitId] : null;

  return {
    project: confValue_('案件名') || '収納量調査',
    units: us,
    noTiers: all.filter(function(u){ return !u.tiers; }).map(function(u){ return u.no; }),
    depts: depts_(),
    options: {
      items:    confList_('収納物品'),
      disposal: confList_('整理・移動'),
      timing:   confList_('使用タイミング'),
      reason:   confList_('保管理由')
    },
    allowOther: confYes_('その他（自由入力）を出す'),
    me: me,
    /* 自分の回答をまるごと返し、収納を開くたびの往復をなくす */
    myAnswers: me ? (lb[me.id] || {}) : {},
    others: (me && confYes_('他の人が登録中の収納を知らせる')) ? othersOf_(me.id, lb, sm) : {}
  };
}

/** 自分の進捗だけ取り直す */
function getMyProgress(submitId) {
  var us = units_().filter(function(u){ return u.tiers > 0; });
  return progressOf_(us, logBy_()[submitId]);
}

/** 自分の、その収納の回答 */
function getMyAnswers(submitId, no) {
  var t = ((logBy_()[submitId]) || {})[String(no).toUpperCase()] || {};
  var out = [];
  for (var n = 1; n <= MAX_TIERS; n++) out.push(t[n] || null);
  return out;
}

/** 1段ぶん保存。自分の提出の中だけを上書きする */
function saveTier(submitId, no, tierNo, tier) {
  submitId = String(submitId || '').trim();
  no = String(no).toUpperCase();
  tierNo = parseInt(tierNo, 10);
  if (!submitId || !no || !tierNo) throw new Error('提出ID・収納番号・段が必要です。');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sm = subs_();
    var me = sm.map[submitId];
    if (!me) throw new Error('提出が見つかりません。お名前を入れ直してください。');
    if (me.state === ST_DONE) throw new Error('提出済みです。編集するには「提出を取り消す」を押してください。');

    var sh = ss_().getSheetByName(S_LOG);
    var mine = tier.presence !== PRESENCE_OTHER;
    var yes  = tier.presence === PRESENCE_YES;
    var keep = yes && tier.disposal === DISPOSAL_KEEP;
    var half = yes && tier.disposal === DISPOSAL_50;

    var row = [
      submitId, me.label, no, tierNo,
      mine ? (tier.dept || '') : '',
      mine ? (tier.section || '') : '',
      tier.presence || '',
      yes ? (tier.item || '') : '',
      yes ? (tier.disposal || '') : '',
      keep ? (tier.timing || '') : '',
      (keep && tier.timing === TIMING_NONE) ? (tier.reason || '') : '',
      half ? (tier.timing50 || '') : '',
      (half && tier.timing50 === TIMING_NONE) ? (tier.reason50 || '') : '',
      now_()
    ];

    var at = findLogRow_(sh, submitId, no, tierNo);
    if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else    sh.appendRow(row);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  /* 進捗は画面側で持っているので返さない。返す値を減らすほど保存が速い */
  return true;
}

function findLogRow_(sh, sid, no, tierNo) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var v = sh.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === sid &&
        String(v[i][2]).trim().toUpperCase() === no &&
        parseInt(toHalf_(v[i][3]), 10) === tierNo) return i + 2;
  }
  return 0;
}

function setState_(submitId, state) {
  var sh = ss_().getSheetByName(S_SUB);
  var last = sh.getLastRow();
  for (var i = 2; i <= last; i++) {
    if (String(sh.getRange(i, 1).getValue()).trim() === submitId) {
      sh.getRange(i, 5).setValue(state);
      sh.getRange(i, 7).setValue(state === ST_DONE ? now_() : '');
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}

/** 提出する。未完了の収納が残っていれば拒否 */
function submitAll(submitId) {
  var us = units_().filter(function(u){ return u.tiers > 0; });
  var pg = progressOf_(us, logBy_()[submitId]);
  var partial = [], count = 0;
  us.forEach(function(u){
    var p = pg[u.no];
    if (!p) return;
    if (p.done) count++; else partial.push(u.no);
  });
  if (partial.length) return { ok: false, partial: partial };
  if (!count) return { ok: false, empty: true };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { setState_(submitId, ST_DONE); } finally { lock.releaseLock(); }
  return { ok: true, count: count };
}

/** 提出を取り消して編集に戻す */
function reopenSubmission(submitId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { setState_(submitId, ST_OPEN); } finally { lock.releaseLock(); }
  return true;
}

/* ================================================================
 * 管理：突き合わせ
 * ================================================================ */

function assertPass_(pass) {
  var want = '';
  masterRows_(S_MCONF).forEach(function(r){
    if (String(r[0]).trim() === '管理用あいことば') want = String(r[1]);
  });
  if (!want) throw new Error('マスター設定に管理用あいことばがありません。setup を実行してください。');
  if (!pass || String(pass) !== want) throw new Error('あいことばが違います。');
}

function masterPassIsDefault_() {
  var v = '';
  masterRows_(S_MCONF).forEach(function(r){
    if (String(r[0]).trim() === '管理用あいことば') v = String(r[1]);
  });
  return v === 'change-me';
}

/** 案件一覧（管理画面用。準備状況も返す） */
function listProjects(pass) {
  assertPass_(pass);
  var out = projects_().map(function(p){
    var o = { key: p.key, name: p.name, sheetId: p.sheetId, note: p.note,
              ready: false, units: 0, depts: 0, answers: 0, error: '' };
    try {
      useProject_(p.key);
      var ss = ss_();
      o.sheetName = ss.getName();
      o.sheetUrl = ss.getUrl();
      var us = units_().filter(function(u){ return u.tiers > 0; });
      o.units = us.length;
      o.depts = depts_().length;
      o.answers = rows_(S_LOG).length;
      o.ready = !!(o.units && o.depts);
    } catch (e) {
      o.error = String((e && e.message) || e);
    }
    return o;
  });
  useProject_('');
  return { projects: out, passIsDefault: masterPassIsDefault_() };
}

/** 案件を追加・更新する（キーで突き合わせ） */
function saveProject(pass, key, name, sheetId, note) {
  assertPass_(pass);
  key = String(key || '').trim();
  sheetId = String(sheetId || '').trim();
  if (!key) throw new Error('キーを入力してください。');
  if (!/^[A-Za-z0-9_-]+$/.test(key)) throw new Error('キーは半角英数字とハイフン・アンダースコアだけにしてください（URLに入るため）。');
  if (!sheetId) throw new Error('スプレッドシートIDを入力してください。');

  /* URLを貼られてもIDを取り出す */
  var m = sheetId.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (m) sheetId = m[1];
  try { SpreadsheetApp.openById(sheetId); }
  catch (e) { throw new Error('そのIDのスプレッドシートを開けません。IDと共有設定を確認してください。'); }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = master_().getSheetByName(S_PROJ);
    if (!sh) throw new Error('「案件」シートがありません。setup を実行してください。');
    var last = sh.getLastRow(), at = 0;
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < v.length; i++) {
        if (String(v[i][0]).trim() === key) { at = i + 2; break; }
      }
    }
    var row = [key, String(name || ''), sheetId, String(note || '')];
    if (at) sh.getRange(at, 1, 1, 4).setValues([row]);
    else    sh.appendRow(row);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  useProject_(key);
  dropCache_();
  useProject_('');
  return true;
}

/** 案件を一覧から外す（スプレッドシート自体は消さない） */
function removeProject(pass, key) {
  assertPass_(pass);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = master_().getSheetByName(S_PROJ);
    var last = sh.getLastRow();
    for (var i = 2; i <= last; i++) {
      if (String(sh.getRange(i, 1).getValue()).trim() === String(key)) {
        sh.deleteRow(i);
        SpreadsheetApp.flush();
        return true;
      }
    }
  } finally { lock.releaseLock(); }
  return false;
}

function adoptMap_() {
  var m = {};
  rows_(S_ADOPT).forEach(function(r){
    var no = String(r[0] || '').trim().toUpperCase();
    var sid = String(r[1] || '').trim();
    if (no && sid) m[no] = sid;
  });
  return m;
}

function sameTier_(a, b) {
  if (!a && !b) return true;
  if (!a || !b) return false;
  var keys = ['dept','section','presence','item','disposal','timing','reason','timing50','reason50'];
  for (var i = 0; i < keys.length; i++) {
    if (String(a[keys[i]] || '') !== String(b[keys[i]] || '')) return false;
  }
  return true;
}

function lastAt_(byTier) {
  var at = '';
  for (var n in byTier) if (byTier[n].at > at) at = byTier[n].at;
  return at;
}

/** 完了データを持つ提出IDを収納番号ごとに集める */
function holders_(us, sm, lb) {
  var h = {};
  sm.order.forEach(function(sid){
    var pg = progressOf_(us, lb[sid]);
    us.forEach(function(u){
      if (!pg[u.no] || !pg[u.no].done) return;
      if (!h[u.no]) h[u.no] = [];
      h[u.no].push(sid);
    });
  });
  return h;
}

/** 提出状況・重複・不足 */
function getReview(pass) {
  assertPass_(pass);
  dropCache_();
  var us = units_().filter(function(u){ return u.tiers > 0; });
  var sm = subs_(), lb = logBy_(), ad = adoptMap_();
  var hd = holders_(us, sm, lb);

  var submissions = sm.order.map(function(sid){
    var s = sm.map[sid];
    var pg = progressOf_(us, lb[sid]);
    var done = 0, part = 0;
    us.forEach(function(u){
      var p = pg[u.no];
      if (!p) return;
      if (p.done) done++; else part++;
    });
    return { id: sid, who: s.label, state: s.state, submittedAt: s.submittedAt,
             done: done, partial: part };
  });

  var duplicates = [];
  us.forEach(function(u){
    var ids = hd[u.no];
    if (!ids || ids.length < 2) return;
    var base = lb[ids[0]][u.no];
    var identical = true;
    for (var i = 1; i < ids.length && identical; i++) {
      for (var n = 1; n <= u.tiers; n++) {
        if (!sameTier_(base[n], lb[ids[i]][u.no][n])) { identical = false; break; }
      }
    }
    duplicates.push({
      no: u.no, identical: identical,
      adopted: (ad[u.no] && ids.indexOf(ad[u.no]) >= 0) ? ad[u.no] : ids[0],
      entries: ids.map(function(sid){
        return { id: sid, who: sm.map[sid].label, state: sm.map[sid].state, at: lastAt_(lb[sid][u.no]) };
      })
    });
  });

  var missing = [], partialUnits = [];
  us.forEach(function(u){
    if (hd[u.no] && hd[u.no].length) return;
    var who = [];
    sm.order.forEach(function(sid){
      if (lb[sid] && lb[sid][u.no]) who.push(sm.map[sid].label);
    });
    if (who.length) partialUnits.push({ no: u.no, tiers: u.tiers, who: who });
    else missing.push({ no: u.no, tiers: u.tiers, spec: u.spec });
  });

  var waiting = [];
  us.forEach(function(u){
    var ids = hd[u.no];
    if (!ids || !ids.length) return;
    var sid = (ad[u.no] && ids.indexOf(ad[u.no]) >= 0) ? ad[u.no] : ids[0];
    var t = lb[sid][u.no], ns = [];
    for (var n = 1; n <= u.tiers; n++) if (t[n] && t[n].presence === PRESENCE_OTHER) ns.push(n);
    if (ns.length) waiting.push({ no: u.no, tiers: ns });
  });

  return {
    units: us.length,
    submissions: submissions,
    duplicates: duplicates,
    missing: missing,
    partialUnits: partialUnits,
    waiting: waiting,
    ready: us.length - missing.length - partialUnits.length
  };
}

/** 重複したときにどの提出を採るか決める */
function setAdopt(pass, no, submitId) {
  assertPass_(pass);
  no = String(no).toUpperCase();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName(S_ADOPT);
    var last = sh.getLastRow(), at = 0;
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < v.length; i++) {
        if (String(v[i][0]).trim().toUpperCase() === no) { at = i + 2; break; }
      }
    }
    var row = [no, submitId, now_()];
    if (at) sh.getRange(at, 1, 1, 3).setValues([row]);
    else    sh.appendRow(row);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return true;
}

/** 管理者が提出を差し戻す */
function adminReopen(pass, submitId) {
  assertPass_(pass);
  return reopenSubmission(submitId);
}

/* ================================================================
 * 管理：設定の確認と取り込み
 * ================================================================ */

/** アプリが②と部門マスタをどう解釈したかを返す */
function getSetupInfo(pass) {
  assertPass_(pass);
  dropCache_();
  var all = units_();
  var byTier = {}, noTiers = [];
  all.forEach(function(u){
    if (!u.tiers) { noTiers.push({ no: u.no, spec: u.spec }); return; }
    if (!byTier[u.tiers]) byTier[u.tiers] = { tiers: u.tiers, count: 0, sample: [] };
    byTier[u.tiers].count++;
    if (byTier[u.tiers].sample.length < 3) byTier[u.tiers].sample.push(u.no + '（' + (u.spec || '仕様なし') + '）');
  });
  var tiers = [];
  Object.keys(byTier).sort(function(a, b){ return a - b; }).forEach(function(k){ tiers.push(byTier[k]); });

  var seen = {}, dup = [];
  all.forEach(function(u){ if (seen[u.no]) dup.push(u.no); seen[u.no] = 1; });

  var ds = depts_();
  return {
    unitCount: all.length,
    tierBreakdown: tiers,
    noTiers: noTiers,
    duplicated: dup,
    totalTiers: all.reduce(function(a, u){ return a + u.tiers; }, 0),
    depts: ds.map(function(d){ return { name: d.name, col: d.col, sections: d.sections }; }),
    options: {
      items:    confList_('収納物品'),
      disposal: confList_('整理・移動'),
      timing:   confList_('使用タイミング'),
      reason:   confList_('保管理由')
    },
    allowOther: confYes_('その他（自由入力）を出す'),
    warnOthers: confYes_('他の人が登録中の収納を知らせる'),
    project: confValue_('案件名'),
    passIsDefault: masterPassIsDefault_(),
    hasAnswers: rows_(S_LOG).length > 0
  };
}

function parseTsv_(text) {
  return String(text || '').replace(/\r\n?/g, '\n').split('\n')
    .map(function(line){ return line.split('\t'); })
    .filter(function(cells){
      return cells.some(function(c){ return String(c).trim() !== ''; });
    });
}

/** ②什器サイズを貼り付けから入れ替える */
function importUnits(pass, text) {
  assertPass_(pass);
  var rows = parseTsv_(text);
  if (!rows.length) throw new Error('貼り付けが空です。');

  /* 1行目が見出しなら落とす */
  if (String(rows[0][0]).replace(/\s/g, '').indexOf('収納番号') >= 0) rows.shift();

  var out = [], bad = [];
  rows.forEach(function(c){
    var raw = String(c[0] || '').trim();
    if (!raw) return;
    if (!raw.match(/^([A-Za-z]+)\s*(\d+)$/)) { bad.push(raw); return; }
    out.push([ raw, c[1] || '', c[2] || '', c[3] || '', c[4] || '', c[5] || '', c[6] || '', c[7] || '' ]);
  });
  if (!out.length) throw new Error('収納番号として読める行がありませんでした。1列目が収納番号になっているか確認してください。');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = ss_();
    var old = ss.getSheetByName(S_UNIT);
    if (old) ss.deleteSheet(old);
    var sh = sheetWithHeader_(ss, S_UNIT, UNIT_COLS);
    sh.getRange(2, 1, out.length, UNIT_COLS.length).setValues(out);
    SpreadsheetApp.flush();
    dropCache_();
  } finally { lock.releaseLock(); }

  return { imported: out.length, skipped: bad };
}

/** 部門マスタを貼り付けから入れ替える */
function importDepts(pass, text) {
  assertPass_(pass);
  var rows = parseTsv_(text);
  if (!rows.length) throw new Error('貼り付けが空です。');

  var head = -1;
  for (var i = 0; i < Math.min(rows.length, 10); i++) {
    var j = rows[i].map(function(c){ return String(c).replace(/\s/g, ''); });
    if (j.indexOf('部門') >= 0 && j.indexOf('課') >= 0) { head = i; break; }
  }
  var body = rows.slice(head + 1);

  var out = [], last = '', count = 0;
  body.forEach(function(c){
    var dn = String(c[0] || '').trim();
    var sn = String(c[1] || '').trim();
    var col = String(c[2] || '').trim().toUpperCase();
    if (dn) last = dn; else dn = last;
    if (!dn) return;
    out.push([dn, sn, col === 'H' ? 'H' : '']);
    if (sn) count++;
  });
  if (!out.length) throw new Error('部門として読める行がありませんでした。');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = ss_();
    var old = ss.getSheetByName(S_DEPT);
    if (old) ss.deleteSheet(old);
    var sh = sheetWithHeader_(ss, S_DEPT, DEPT_COLS);
    sh.getRange(2, 1, out.length, DEPT_COLS.length).setValues(out);
    SpreadsheetApp.flush();
    dropCache_();
  } finally { lock.releaseLock(); }

  var ds = depts_();
  return { depts: ds.length, sections: count };
}

/* ================================================================
 * ①回答シートの生成
 * ================================================================ */

function headerRow_() {
  var h = ['タイムスタンプ', '回答者', '収納番号',
           '収納番号の頭にあるアルファベットを選択してください',
           '収納番号を選択してください'];
  for (var n = 1; n <= MAX_TIERS; n++) {
    h.push(n + '段目を使用している部門');
    h.push(n + '段目を使用している課');
    h.push('上から' + n + '段目に収納物はありますか？');
    h.push(n + '段目の収納で過半数を占めているものは何ですか？');
    h.push(n + '段目の収納物に関して、今後、執務環境の見直しを行う場合に、どの程度整理・移動が可能ですか？');
    h.push(n + '段目の今後使用するタイミングを教えてください。');
    h.push(n + '段目の保管が必要な理由を教えてください。');
    h.push(n + '段目の残す50%の、今後使用するタイミングを教えてください。');
    h.push(n + '段目の残す50%の、保管が必要な理由を教えてください。');
  }
  return h;
}

function buildAnswerSheet() {
  var us = units_().filter(function(u){ return u.tiers > 0; });
  var sm = subs_(), lb = logBy_(), ad = adoptMap_();
  var hd = holders_(us, sm, lb);
  var aoa = [headerRow_()];
  var skipped = [];

  us.forEach(function(u){
    var ids = hd[u.no];
    if (!ids || !ids.length) { skipped.push(u.no); return; }
    var use = (ad[u.no] && ids.indexOf(ad[u.no]) >= 0) ? ad[u.no] : ids[0];

    var t = lb[use][u.no];
    var row = [lastAt_(t), sm.map[use].label, u.no, u.alpha, u.num];
    for (var n = 1; n <= MAX_TIERS; n++) {
      var x = t[n];
      if (!x || n > u.tiers) {
        if (n === u.tiers + 1) row.push('', '', n + '段目以降は存在しない', '', '', '', '', '', '');
        else row.push('', '', '', '', '', '', '', '', '');
        continue;
      }
      row.push(x.dept, x.section, x.presence, x.item, x.disposal,
               x.timing, x.reason, x.timing50, x.reason50);
    }
    aoa.push(row);
  });

  var ss = ss_();
  var sh = ss.getSheetByName(S_OUT);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(S_OUT);
  sh.getRange(1, 1, aoa.length, aoa[0].length).setValues(aoa);
  sh.getRange(1, 1, 1, aoa[0].length).setFontWeight('bold').setBackground('#F2F2F2');
  sh.setFrozenRows(1);

  return { rows: aoa.length - 1, skipped: skipped, url: ss.getUrl() };
}

function adminBuild(pass) {
  assertPass_(pass);
  return buildAnswerSheet();
}
