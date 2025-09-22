/************************************************************
 * TimeTrigger 一式（メニュー＋定期＋ワンショット台帳＋一覧）
 * 方式B：月末→月初ウォッチャは “検知したその場で即実行”
 * ！onOpenは定義しません（中央onOpenから buildMenu_TimeTrigger_ を呼ぶ前提）
 ************************************************************/

/***** ========= 予約台帳の定義 ========= *****/
const TTRES = {
  SHEET : 'time_trigger_reserve',
  HEADER: ['Created(JST)','Scheduled(JST)','Handler','Label','Status','Updated(JST)'] // PENDING | FIRED | CANCELED
};

/***** ========= 共通ユーティリティ ========= *****/
function _ensureSheet_(name, header){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,header.length).setValues([header]);
    sh.setFrozenRows(1);
    try { for (let c=1;c<=header.length;c++) sh.autoResizeColumn(c); } catch(_){}
  }
  return sh;
}
function _tz_(){ return Session.getScriptTimeZone() || 'Asia/Tokyo'; }
function _fmtNow_(){ return Utilities.formatDate(new Date(), _tz_(), 'yyyy-MM-dd HH:mm'); }
function _fmtAt_(d){ return Utilities.formatDate(d, _tz_(), 'yyyy-MM-dd HH:mm'); }

/** どんなセルでも Date に直す（Date型/シリアル/英語長文/ISO/カスタム） */
function _parseWhenFlexible_(v){
  if (v && typeof v.getTime === 'function' && !isNaN(v.getTime())) return v;           // Date
  if (typeof v === 'number' && isFinite(v)){                                           // serial
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d  = new Date(ms);
    if (!isNaN(d)) return d;
  }
  const s = String(v||'').trim();
  if (!s) return null;
  const d1 = new Date(s); if (!isNaN(d1)) return d1;                                   // ISO/英語長文
  const m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})$/.exec(s);                    // yyyy-MM-dd HH:mm
  if (m) return new Date(+m[1], +m[2]-1, +m[3], +m[4], +m[5], 0, 0);
  return null;
}

function _appendReservation_(whenDate, handler, label){
  const sh = _ensureSheet_(TTRES.SHEET, TTRES.HEADER);
  sh.appendRow([_fmtNow_(), _fmtAt_(whenDate), handler, String(label||''), 'PENDING', '']);
}

function _markReservationAsFired_(handler){
  const sh = SpreadsheetApp.getActive().getSheetByName(TTRES.SHEET);
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  const now = new Date();
  let targetRow = -1, bestDiff = Infinity;

  for (let r=2; r<=values.length; r++){
    const row = values[r-1];
    if (row[4] !== 'PENDING' || row[2] !== handler) continue;
    const at = _parseWhenFlexible_(row[1]);
    if (!at) continue;
    const diff = now.getTime() - at.getTime(); // 正: 予定 <= 今
    if (diff >= -5*60*1000 && diff <= 3*60*1000*60 && diff < bestDiff) {
      bestDiff = diff;
      targetRow = r;
    }
  }
  if (targetRow > 0){
    sh.getRange(targetRow, 5).setValue('FIRED');
    sh.getRange(targetRow, 6).setValue(_fmtNow_());
  }
}

function _callIfExists_(fname, ...args){
  const fn = (typeof globalThis !== 'undefined' ? globalThis : this)[fname];
  if (typeof fn === 'function') return fn.apply(null, args);
  SpreadsheetApp.getActive().toast(`未実装: ${fname}（スキップ）`);
  return null;
}

/** 汎用：ハンドラー名配列でトリガー削除し件数を返す */
function _deleteTriggersByHandlers_(handlers){
  let cnt = 0;
  (ScriptApp.getProjectTriggers()||[]).forEach(t=>{
    const h = t.getHandlerFunction && t.getHandlerFunction();
    if (handlers.includes(h)) { ScriptApp.deleteTrigger(t); cnt++; }
  });
  return cnt;
}

/***** ========= メニュー（中央onOpenから呼ばれる） ========= *****/
function buildMenu_TimeTrigger_() {
  SpreadsheetApp.getUi()
    .createMenu('TimeTrigger')
    .addItem('任意時刻スケジュールを設定（フル実行）', 'openOneShotFullDialog_')
    .addItem('任意時刻スケジュールを解除（フル実行）', 'cancelOneShotFullTriggers_')
    .addSeparator()
    .addItem('定期スケジュールを設定（10:00 フル実行／月・水・土）', 'installWeeklySchedules_Full_v1')
    .addItem('定期スケジュールを解除（週次フル）', 'deleteWeeklySchedules_Full_v1')
    .addSeparator()
    .addItem('月末23:55〜月初00:05を設定（フル｜ウォッチャ）', 'installMonthlyBoundaryWatcher_Full_')
    .addItem('月末23:55〜月初00:05を解除（フル｜ウォッチャ）', 'deleteMonthlyBoundaryWatcher_Full_')
    .addSeparator()
    .addItem('予約状況を表示', 'showTriggerReservations_')
    .addToUi();
}

/***** ========= 定期：週次（フル） ========= *****/
function installWeeklySchedules_Full_v1(){
  // 旧仕様の掃除（重複防止）
  _deleteTriggersByHandlers_(['jobWeekly7_CSV_', 'jobWeekly10_UploadPivot_', 'jobWeekly10_Full_']);

  [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.SATURDAY].forEach(d=>{
    ScriptApp.newTrigger('jobWeekly10_Full_')
      .timeBased().onWeekDay(d).atHour(10).create();
  });
  SpreadsheetApp.getActive().toast('週次（フル）トリガー：月・水・土 10:00 を設定しました。');
}
function deleteWeeklySchedules_Full_v1(){
  // 新旧どちらの週次も一括削除
  const removed = _deleteTriggersByHandlers_(['jobWeekly10_Full_', 'jobWeekly7_CSV_', 'jobWeekly10_UploadPivot_']);
  SpreadsheetApp.getActive().toast(`週次トリガーを ${removed} 件削除しました。`);
  try{ showTriggerReservations_(); }catch(_){}
}

/***** ========= 定期：月またぎ（23:55〜00:05｜フル・短間隔フィルタ） =========
 * 毎分ウォッチャで「JSTの月末23:55〜23:59 または 月初00:00〜00:05」を検知。
 * 再入防止に ScriptProperties を使用します。
 **********************************************************************/
const TT_KEYS = {
  MONTH_BOUNDARY_LAST: 'MONTH_BOUNDARY_LAST_YYYYMM' // この境界で直近実行した“翌月”を記録
};

function installMonthlyBoundaryWatcher_Full_(){
  // 旧 23:30/23:59/0:00±5 を掃除
  _deleteTriggersByHandlers_([
    'jobMonthlyBoundary_Watcher_',
    'jobMonthlyBOM_0000_Watcher_',
    'jobMonthlyEOM_2359_Watcher_',
    'jobMonthlyEOM_2330_Full_', 'jobMonthlyEOM_2330_'
  ]);

  // 新ウォッチャ作成（重複防止のため先に掃除）
  ScriptApp.newTrigger('jobMonthlyBoundary_Watcher_')
    .timeBased()
    .everyMinutes(1)   // 毎分起動 → 内部で「23:55〜00:05」のみ実行
    .create();

  SpreadsheetApp.getActive().toast('月末23:55〜月初00:05（フル）ウォッチャを設定しました（毎分起動→内部判定）。');
}
function deleteMonthlyBoundaryWatcher_Full_(){
  const removed = _deleteTriggersByHandlers_([
    'jobMonthlyBoundary_Watcher_',
    'jobMonthlyBOM_0000_Watcher_',
    'jobMonthlyEOM_2359_Watcher_',
    'jobMonthlyEOM_2330_Full_', 'jobMonthlyEOM_2330_'
  ]);
  SpreadsheetApp.getActive().toast(`月末23:55〜月初00:05ウォッチャ／旧23:30/23:59/0:00±5 を ${removed} 件削除しました。`);
  try{ showTriggerReservations_(); }catch(_){}
}

/** 毎分起動のウォッチャ本体（JSTで23:55〜00:05の 1回のみ実行）—方式B（即実行） */
function jobMonthlyBoundary_Watcher_(){
  const tz = _tz_();
  const now = new Date();
  const y  = +Utilities.formatDate(now, tz, 'yyyy');
  const m  = +Utilities.formatDate(now, tz, 'MM');
  const d  = +Utilities.formatDate(now, tz, 'dd');
  const hh = +Utilities.formatDate(now, tz, 'HH');
  const mm = +Utilities.formatDate(now, tz, 'mm');

  // 当月の最終日（JST）
  const lastDay = new Date(y, m, 0).getDate();

  // ウィンドウ判定：月末23:55〜23:59、または月初00:00〜00:05
  const inTail = (d === lastDay && hh === 23 && mm >= 55); // 23:55〜23:59
  const inHead = (d === 1       && hh === 0  && mm <= 5);  // 00:00〜00:05
  if (!(inTail || inHead)) return;

  // “境界ID”を「この境界で1回だけ」にするための年月で表現
  // 月末側で動いたら「翌月YYYY-MM」、月初側で動いたら「当月YYYY-MM」
  let boundaryYm;
  if (inHead) {
    boundaryYm = Utilities.formatDate(now, tz, 'yyyy-MM');       // 月初→当月
  } else {
    const next = new Date(y, m, 1);                              // 翌月1日
    boundaryYm = Utilities.formatDate(next, tz, 'yyyy-MM');      // 月末→翌月
  }

  // 再実行防止
  const props = PropertiesService.getScriptProperties();
  const key = TT_KEYS.MONTH_BOUNDARY_LAST;
  if (props.getProperty(key) === boundaryYm) return;

  // 実行（フルパイプライン）— 即時
  runFullPipeline_('jobMonthlyBoundary_Full_');
  props.setProperty(key, boundaryYm);

  // ログ
  _ttLog('MONTH_BOUNDARY', 'FIRED', boundaryYm);
}

/***** ========= フル実行パイプライン（CSV→Upload→CSV→PivotA→PivotB） ========= *****/
function runFullPipeline_(caller){
  const lock = LockService.getScriptLock();
  lock.tryLock(30 * 1000);

  // ① CSVインポート（前処理）
  _callIfExists_('runCsvImportPipeline_', `${caller}#step1_csv_pre`);

  // ② アップロード→③ 再CSV → ④ PivotA → ⑤ PivotB
  _callIfExists_('runUploadAndPivotPipeline_', `${caller}#step2_upload_and_pivot`);

  lock.releaseLock();
  SpreadsheetApp.getActive().toast('フル実行（CSV→Upload→CSV→PivotA→PivotB）完了');
  return true;
}

/***** ========= トリガーの入口 ========= *****/
// 週次（フル）
function jobWeekly10_Full_(){ return runFullPipeline_('jobWeekly10_Full_'); }

// ワンショット（フル）
function jobOneShot_Full_(){
  _markReservationAsFired_('jobOneShot_Full_');
  return runFullPipeline_('jobOneShot_Full_');
}

/***** ========= ワンショット：設定／解除 ========= *****/
// ダイアログで日時入力 → jobOneShot_Full_ を at() で予約（※B本筋とは独立、従来どおり残置）
function openOneShotFullDialog_(){
  const ui = SpreadsheetApp.getUi();
  const labelResp = ui.prompt('ラベルを入力', '例: フル実行', ui.ButtonSet.OK_CANCEL);
  if (labelResp.getSelectedButton() !== ui.Button.OK) return;
  const label = labelResp.getResponseText().trim() || 'フル実行';

  const dtResp = ui.prompt('実行日時を入力',
    '形式: YYYY-MM-DD HH:mm（例: 2025-09-06 13:00）', ui.ButtonSet.OK_CANCEL);
  if (dtResp.getSelectedButton() !== ui.Button.OK) return;
  const when = _parseWhenFlexible_(dtResp.getResponseText().trim());
  if (!when) { ui.alert('日時の形式が不正です。YYYY-MM-DD HH:mm などで入力してください。'); return; }
  if (when.getTime() - Date.now() < 60*1000) { ui.alert('現在時刻より1分以上先の日時を指定してください。'); return; }

  createOneOffTriggerFor_('jobOneShot_Full_', {
    y: when.getFullYear(), m: when.getMonth()+1, d: when.getDate(),
    h: when.getHours(), mi: when.getMinutes(), label
  });
}

// 予約済みワンショット（新旧すべて）を削除し、台帳PENDING→CANCELED
function cancelOneShotFullTriggers_(){
  // 新旧ハンドラーをまとめて削除
  const targets = ['jobOneShot_Full_', 'jobOneShot_CSV_', 'jobOneShot_UploadPivot_'];
  const removed = _deleteTriggersByHandlers_(targets);

  // 台帳更新
  const sh = SpreadsheetApp.getActive().getSheetByName(TTRES.SHEET);
  if (sh){
    const v = sh.getDataRange().getValues();
    for (let r=2; r<=v.length; r++){
      if (targets.includes(v[r-1][2]) && v[r-1][4] === 'PENDING'){
        sh.getRange(r, 5).setValue('CANCELED');
        sh.getRange(r, 6).setValue(_fmtNow_());
      }
    }
  }
  SpreadsheetApp.getActive().toast(`任意時刻ワンショットを ${removed} 件解除しました。`);
  try{ showTriggerReservations_(); }catch(_){}
}

/***** ========= ワンショット作成 (共通) ========= *****/
function createOneOffTriggerFor_(handlerName, p){
  const y = Number(p && p.y), m = Number(p && p.m), d = Number(p && p.d);
  const h = Number(p && p.h), mi = Number(p && p.mi);
  if (!(y && m && d) || isNaN(h) || isNaN(mi)) throw new Error('日時の指定が不正です。');

  const when = new Date(y, m-1, d, h, mi, 0, 0);
  if (when.getTime() - Date.now() < 60*1000) throw new Error('現在時刻より1分以上先を指定してください。');

  ScriptApp.newTrigger(handlerName).timeBased().at(when).create();
  _appendReservation_(when, handlerName, p && p.label);

  SpreadsheetApp.getActive().toast(`${p.label}（${_fmtAt_(when)}）に実行を予約しました`);
  return `「${p.label}」に実行を予約しました（${_fmtAt_(when)}・${handlerName}）。`;
}

/***** ========= 予約状況ビューア（サイドバー） ========= *****/
function showTriggerReservations_(){
  const tz = _tz_();

  // A) 現在インストール済みトリガー
  const triggers = (ScriptApp.getProjectTriggers()||[])
    .map(t=>t.getHandlerFunction && t.getHandlerFunction())
    .filter(Boolean);
  const rowsA = triggers.map(h=>{
    let note = '';
    if (h === 'jobWeekly10_Full_')                 note = '毎週 月・水・土 の 10:00（フル実行）';
    else if (h === 'jobMonthlyBoundary_Watcher_')  note = '毎分ウォッチャ → 月末23:55〜月初00:05のどこかで1回（フル）';
    else if (h === 'jobOneShot_Full_')             note = '任意時刻（フル実行・下の台帳参照）';
    else if (h === 'jobWeekly7_CSV_')              note = '（旧）週次 07:00 CSV';
    else if (h === 'jobWeekly10_UploadPivot_')     note = '（旧）週次 10:00 Upload&Pivot';
    else if (['jobMonthlyEOM_2359_Watcher_','jobMonthlyEOM_2330_Full_','jobMonthlyEOM_2330_','jobMonthlyBOM_0000_Watcher_'].includes(h))
      note = '（旧）月末/月初ウォッチャ';
    return [h, note];
  });

  // B) ワンショット予約台帳
  let rowsB = [];
  const sh = SpreadsheetApp.getActive().getSheetByName(TTRES.SHEET);
  if (sh) {
    const v = sh.getDataRange().getValues();
    const idx = {}; TTRES.HEADER.forEach((h,i)=>idx[h]=i);
    rowsB = v.slice(1).map(r=>({
      created: r[idx['Created(JST)']],
      scheduled: r[idx['Scheduled(JST)']],
      handler: r[idx['Handler']],
      label: r[idx['Label']],
      status: r[idx['Status']],
      updated: r[idx['Updated(JST)']]
    })).sort((a,b)=>{
      const da=_parseWhenFlexible_(a.scheduled)?.getTime()||0;
      const db=_parseWhenFlexible_(b.scheduled)?.getTime()||0;
      return da-db;
    });
  }

  const esc = s => String(s||'').replace(/[&<>"']/g, c=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
  const tableA = `<table class="tbl"><tr><th>Handler</th><th>Note</th></tr>${
    rowsA.map(r=>`<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td></tr>`).join('')
  }</table>`;
  const tableB = `<table class="tbl"><tr><th>Created</th><th>Scheduled</th><th>Handler</th><th>Label</th><th>Status</th><th>Updated</th></tr>${
    rowsB.map(r=>`<tr><td>${esc(r.created)}</td><td>${esc(r.scheduled)}</td><td>${esc(r.handler)}</td><td>${esc(r.label)}</td><td>${esc(r.status)}</td><td>${esc(r.updated)}</td></tr>`).join('')
  }</table>`;

  const html = HtmlService.createHtmlOutput(
`<style>
body{font-family:system-ui, sans-serif; font-size:13px; line-height:1.5;}
.tbl{border-collapse:collapse; width:100%;}
.tbl th,.tbl td{border:1px solid #ddd; padding:6px; font-size:12px;}
.tbl th{background:#f6f6f6; text-align:left;}
.note{font-size:12px; color:#444; margin:6px 0 12px;}
.muted{color:#777; font-size:12px;}
</style>
<h1>予約状況</h1>
<div class="note">ワンショットの日時は予約時に台帳へ保存しています（APIでは取得不可）。</div>
<h2>現在インストール済みトリガー（ScriptApp）</h2>
${tableA}
<h2>ワンショット予約台帳（シート: ${esc(TTRES.SHEET)}）</h2>
${tableB}
<div class="muted" style="margin-top:10px;">タイムゾーン: ${esc(tz)}</div>`
  ).setTitle('予約状況');
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** ========= （任意）古いPENDINGを掃除 ========= *****/
function cleanupOldPendingReservations_(days=60){
  const sh = SpreadsheetApp.getActive().getSheetByName(TTRES.SHEET);
  if (!sh) return 0;
  const v = sh.getDataRange().getValues();
  const idx = {}; TTRES.HEADER.forEach((h,i)=>idx[h]=i);
  const now = Date.now(), limit = days*24*60*60*1000;
  let removed = 0;
  for (let r=v.length; r>=2; r--){
    const row = v[r-1];
    const status = row[idx['Status']];
    const at = _parseWhenFlexible_(row[idx['Scheduled(JST)']]);
    if (status==='PENDING' && at && (now - at.getTime() > limit)){
      sh.deleteRow(r); removed++;
    }
  }
  if (removed) SpreadsheetApp.getActive().toast(`古いPENDINGを ${removed} 行削除しました。`);
  return removed;
}

/*** ▼ TimeTrigger から呼ばれる実体ラッパー（ログ付き） ***/
// CSVだけ（前処理）
function runCsvImportPipeline_(tag){
  _ttLog('CSV_PRE', 'START', tag);
  appendDiffUsers_toCopySheetA();   // CSVインポート（copy→ユーザーへ差分追記）
  _ttLog('CSV_PRE', 'DONE', tag);
}

// アップロード→CSV→PivotA→PivotB（後処理まとめ）
function runUploadAndPivotPipeline_(tag){
  _ttLog('PIPE_AFTER', 'START', tag);
  // ② Bubbleアップロード（ユーザー末尾100件・ログつき）
  menuUploadUserTail100_LogSheet();

  // ③ 再CSV（Bubble更新を取り込むため）
  appendDiffUsers_toCopySheetA();

  // ④ PivotA（最新月：書式→クロス→転記）
  buildFormatCrossAndTransferLatest();

  // ⑤ PivotB（最新月：全部入り）
  menuPivotB_RunAllLatest();

  _ttLog('PIPE_AFTER', 'DONE', tag);
}

/*** シンプルなログ（time_trigger_log シートへ） ***/
function _ttLog(step, status, tag, note){
  const sh = (function(){
    const ss = SpreadsheetApp.getActive();
    return ss.getSheetByName('time_trigger_log') || ss.insertSheet('time_trigger_log');
  })();
  sh.getRange(1,1,1,6).setValues([['Timestamp','Step','Status','Caller','User','Note']]).setFontWeight('bold');
  sh.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    step, status, String(tag||''), Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : '', String(note||'')
  ]);
}
