/***** ===== Bubble Upload（ユーザー末尾100件｜別シートにログを自動作成） ===== *****
 * ENVで Live / Dev（/version-test） を切替。トークンもENVに応じて使い分け。
 * Script Properties 推奨キー:
 *   - BUBBLE_TOKEN        （Live）
 *   - BUBBLE_TOKEN_DEV    （Dev）
 *****************************************************************************************/

//////////////////// ENV ////////////////////
const ENV = 'dev'; // ← 'live' | 'dev' で切替
const BASES = {
  live: 'https://manepon00.com',
  dev : 'https://manepon00.com/version-test'
};

//////////////////// 設定 ////////////////////
const FLOW_CFG = {
  BASE : BASES[ENV],
  TOKEN: (function(){
    const sp = PropertiesService.getScriptProperties();
    return sp.getProperty(ENV==='live' ? 'BUBBLE_TOKEN' : 'BUBBLE_TOKEN_DEV') || '';
  })(),
  TYPE : 'user',          // Bubble データタイプ
  FIELD: '流入経路'        // 上書きするフィールド名
};

// HTTPヘッダ
const FLOW_HGET   = { Authorization: 'Bearer ' + String(FLOW_CFG.TOKEN).replace(/^Bearer\s+/i,'').trim() };
const FLOW_HPATCH = { ...FLOW_HGET, 'Content-Type':'application/json' };

/** ログ設定（別シートに追記） */
const FLOW_LOG = {
  SHEET : 'アップロード履歴',
  HEADER: ['日時','RUN','行','ID','流入経路（手入力後）','結果','GET(1)コード','PATCHコード','GET(2)コード','流入経路（上書き後）','メモ']
};

/** ENVチェック（Liveのみ厳格にチェックしたい場合に使用） */
function flow_assertEnv_(){
  // 旧コードの「/version-test を拒否」ガードは Live 前提だった（参考：元実装）:
  // flow_assertLive_() が /version-test を検出すると例外にしていた。 :contentReference[oaicite:5]{index=5}
  // 今回はENVで制御するので、Liveで /version-test が含まれていたらだけ落とす。
  if (ENV==='live' && /\/version-test\/?$/.test(FLOW_CFG.BASE)) {
    throw new Error('BASE が DEV です（Live運用時は /version-test を外してください）');
  }
}

/** 動作確認（200ならOK） */
function pingFlowUpload(){
  const url = `${FLOW_CFG.BASE}/api/1.1/obj/${FLOW_CFG.TYPE}?limit=1`;
  const r = UrlFetchApp.fetch(url, { headers: FLOW_HGET, muteHttpExceptions: true });
  Logger.log(`GET ${r.getResponseCode()}`);
}

/** メニュー */
function buildMenu_CsvUpload_() {
  SpreadsheetApp.getUi()
    .createMenu('Bubble Upload')
    .addItem('Upload（ユーザー末尾100件｜別シートログ）', 'menuUploadUserTail100_LogSheet')
    .addToUi();
}

/** ========= 本体：ユーザーの末尾100件 + 別シートにログ =========
 * 対象：ユーザーの最終行から上へ向かって unique id が入っている行を最大100件
 * 入力：同行の「流入経路」セル（'CLEAR'/'-' は null、空はスキップ）
 * 出力：Bubbleへ上書き（差分時のみPATCH）し、別シートにログ
 * ログ：ID / 流入経路（手入力後） / 流入経路（上書き後） を含む
 */
function menuUploadUserTail100_LogSheet() {
  const runId = Utilities.getUuid().slice(0,8);
  const tz    = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const shLog = flow_getOrCreateLogSheet_();

  try {
    flow_assertEnv_();

    const ss = SpreadsheetApp.getActive();
    const shUser = ss.getSheetByName('ユーザー');
    if (!shUser) throw new Error('シート「ユーザー」が見つかりません');

    // ヘッダ/列
    const header = flow_getHeaderRow_(shUser);
    const colUserId   = flow_findHeader_(header, 'unique id');
    const colUserFlow = flow_findHeader_(header, '流入経路');
    if (colUserId <= 0 || colUserFlow <= 0) {
      throw new Error('「ユーザー」シートに「unique id」または「流入経路」列がありません');
    }

    const lastRow = shUser.getLastRow();
    if (lastRow < 2) { SpreadsheetApp.getActive().toast('ユーザーにデータがありません'); return; }

    const nRows  = Math.max(lastRow - 1, 0);
    const ids    = shUser.getRange(2, colUserId,   nRows, 1).getDisplayValues().map(r=>String(r[0]||'').trim());
    const flows  = shUser.getRange(2, colUserFlow, nRows, 1).getDisplayValues().map(r=>String(r[0]||'').trim());

    const targets = []; // [{row, id, input}]
    for (let idx = ids.length - 1; idx >= 0 && targets.length < 100; idx--) {
      const id = ids[idx]; if (!id) continue;
      targets.push({ row: idx + 2, id, input: flows[idx] || '' }); // 行番号は+2
    }
    if (!targets.length) { SpreadsheetApp.getActive().toast('対象IDが見つかりません'); return; }

    // STARTログ
    flow_appendLog_(shLog, tz, runId, '-', '-', 'START', '-', '-', '-', '-', `対象 ${targets.length} 件`);

    let patched = 0, same = 0, skipped = 0, errors = 0;

    for (const t of targets) {
      const id = t.id;
      const input = t.input;

      // 1) 現在値（GET1）
      const g1 = UrlFetchApp.fetch(`${FLOW_CFG.BASE}/api/1.1/obj/${FLOW_CFG.TYPE}/${encodeURIComponent(id)}`, {
        headers: FLOW_HGET, muteHttpExceptions: true
      });
      const code1 = g1.getResponseCode();
      if (code1 !== 200) {
        errors++;
        flow_appendLog_(shLog, tz, runId, String(t.row), id, input, 'ERROR(GET1)', `GET:${code1}`, '-', '-', '');
        Utilities.sleep(80);
        continue;
      }
      const cur = flow_extractDisplay_(JSON.parse(g1.getContentText()).response[FLOW_CFG.FIELD]);

      // 2) 次値決定
      let next; // string | null | undefined
      if (/^(-|CLEAR)$/i.test(input)) {
        next = null;
      } else if (input === '') {
        next = undefined; // スキップ
      } else {
        next = input;
      }

      // 3) PATCH（差分時）
      let patchCode = '-', result = '';
      if (typeof next === 'undefined') {
        result = 'SKIP(入力空)'; skipped++;
      } else {
        const isSame = (next === null ? (cur === '') : (String(cur) === String(next)));
        if (!isSame) {
          const body = {}; body[FLOW_CFG.FIELD] = next;
          const p = UrlFetchApp.fetch(`${FLOW_CFG.BASE}/api/1.1/obj/${FLOW_CFG.TYPE}/${encodeURIComponent(id)}`, {
            method:'patch', headers: FLOW_HPATCH, payload: JSON.stringify(body), muteHttpExceptions:true
          });
          patchCode = String(p.getResponseCode());
          if (p.getResponseCode() < 200 || p.getResponseCode() >= 300) {
            errors++;
            flow_appendLog_(shLog, tz, runId, String(t.row), id, input, 'ERROR(PATCH)', `GET:${code1}`, `PATCH:${patchCode}`, '-', '');
            Utilities.sleep(80);
            continue;
          }
          patched++; result = 'PATCH';
        } else {
          same++; result = 'SAME';
        }
      }

      // 4) 最終値（GET2）→ ログ「流入経路（上書き後）」に記録
      const g2 = UrlFetchApp.fetch(`${FLOW_CFG.BASE}/api/1.1/obj/${FLOW_CFG.TYPE}/${encodeURIComponent(id)}`, {
        headers: FLOW_HGET, muteHttpExceptions: true
      });
      const code2 = g2.getResponseCode();
      if (code2 !== 200) {
        errors++;
        flow_appendLog_(shLog, tz, runId, String(t.row), id, input, 'ERROR(GET2)', `GET:${code1}`, patchCode==='-'?'-':`PATCH:${patchCode}`, `GET2:${code2}`, '');
      } else {
        const after = flow_extractDisplay_(JSON.parse(g2.getContentText()).response[FLOW_CFG.FIELD]);
        flow_appendLog_(shLog, tz, runId, String(t.row), id, input, result, `GET:${code1}`, patchCode==='-'?'-':`PATCH:${patchCode}`, `GET2:${code2}`, after);
      }

      Utilities.sleep(80);
    }

    flow_appendLog_(shLog, tz, runId, '-', '-', 'END', '-', '-', '-', '-', `PATCH:${patched} / SAME:${same} / SKIP:${skipped} / ERR:${errors}`);
    SpreadsheetApp.getActive().toast(`完了：対象 ${targets.length} 件（PATCH:${patched} / SAME:${same} / SKIP:${skipped} / ERR:${errors}）`);

  } catch (e) {
    // 例外もログ（ログシートが取れないケースも考慮）
    try {
      const shLog2 = flow_getOrCreateLogSheet_();
      const tz2 = Session.getScriptTimeZone() || 'Asia/Tokyo';
      const runId2 = Utilities.getUuid().slice(0,8);
      flow_appendLog_(shLog2, tz2, runId2, '-', '-', 'FATAL', '-', '-', '-', '-', String(e));
    } catch (_) {}
    SpreadsheetApp.getActive().toast('アップロードでエラー: ' + e);
    throw e;
  }
}

/** ========= ヘルパ ========= */
function flow_getHeaderRow_(sh){
  const lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
}
function flow_findHeader_(headerArray, name){
  const idx = headerArray.indexOf(String(name).trim());
  return (idx >= 0) ? (idx + 1) : -1; // 1-based
}
function flow_extractDisplay_(v){
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'object') {
    if (v.display) return String(v.display);
    if (v._id) return String(v._id);
    try { return JSON.stringify(v); } catch(_) { return String(v); }
  }
  return String(v);
}

/** ログシート取得（なければ作成、非表示でも再表示、ヘッダ整備） */
function flow_getOrCreateLogSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FLOW_LOG.SHEET);
  if (!sh) { sh = ss.insertSheet(FLOW_LOG.SHEET); sh.appendRow(FLOW_LOG.HEADER); }
  sh.setTabColor('#1a73e8'); // 青マーキング
  // ヘッダが一致しない場合は合わせる（不足列を右に追加）
  const header = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), FLOW_LOG.HEADER.length)).getValues()[0];
  if (header.join('｜') !== FLOW_LOG.HEADER.join('｜')){
    sh.getRange(1,1,1,FLOW_LOG.HEADER.length).setValues([FLOW_LOG.HEADER]);
  }
  return sh;
}

/** ログ1行追記 */
function flow_appendLog_(shLog, tz, runId, rowStr, id, inputValue, result, http1, patchCode, http2, afterValue){
  const ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const ordered = [ ts, runId, rowStr, id, inputValue, result, http1, patchCode, http2, afterValue, '' ];
  shLog.appendRow(ordered);
}

/** ==== Backward-compat: 旧関数名に対応 ==== */
function menuUploadOverwriteManual() { return menuUploadUserTail100_LogSheet(); }
function menuUploadFlowFromGamma100(){ return menuUploadUserTail100_LogSheet(); }
