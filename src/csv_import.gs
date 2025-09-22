/******************** CSVインポート メニュー ********************/
function buildMenu_CsvImport_() {
  SpreadsheetApp.getUi()
    .createMenu('CSV import')
    .addItem('user', 'appendDiffUsers_toCopySheetA')
    .addToUi();
}

/***** ENV & 設定 *****/
const ENV_CSV = 'dev'; // ← 'live' | 'dev'
const BASE_CSV = (ENV_CSV === 'live')
  ? 'https://manepon00.com'
  : 'https://manepon00.com/version-test'; // 元は LIVE_BASE 固定だった部分を切替に変更（参考）。 :contentReference[oaicite:9]{index=9}

const TYPE_USER     = 'user';
const ORIG_SHEET    = 'testA_user';      // 原本
const COPY_SHEET    = 'testA_user_copy'; // 取り込み先（複製）
const USER_SHEET    = 'ユーザー';         // ユーザー本体へも追記
const HEADER_ROW    = 1;                 // ヘッダ行
const MAX_ROWS      = 1000;              // 取得上限
const PAGE_LIMIT    = 100;               // Bubble 1リクエスト上限

// 表示ラベル→API実キーのエイリアス（必要に応じて追加）
const ALIASES = {
  'unique id': ['_id','Unique id'],
  'Creation Date': ['Created Date'],
  'email': ['email','Email','authentication.email.email'],
  '00_お名前': ['00_お名前','お名前','User_ユーザー情報.お名前'],
  '流入経路': [
    '00_R_ユーザー情報.流入経路','User_ユーザー情報.流入経路','流入経路',
    '00_R_ユーザー情報.流入経路.display','流入経路.display',
    '00_R_ユーザー情報.OS_流入経路','OS_流入経路','OS_流入経路.display'
  ],
  'LINE ID': ['LINE ID','00_R_ユーザー情報.LINE ID','User_ユーザー情報.LINE ID']
};
// 「月/年月」を作る元の日時候補
const YM_SOURCE_CANDIDATES = ['Created Date','Creation Date','登録日時'];

/***** エントリ（差分取り込み→copyへ追記→ユーザーへ転送） *****/
function appendDiffUsers_toCopySheetA(){
  const token = PropertiesService.getScriptProperties().getProperty(ENV_CSV==='live' ? 'BUBBLE_TOKEN' : 'BUBBLE_TOKEN_DEV');
  if (!token) throw new Error('BUBBLE_TOKEN（ENV別）が未設定です');

  const ss   = SpreadsheetApp.getActive();
  const copy = getOrCreateCopySheet_(ss, ORIG_SHEET, COPY_SHEET); // 初回だけ複製
  ensureLineIdColumnInCopy_(copy);                                // ← LINE ID を所定位置に

  const header = getHeader_(copy, HEADER_ROW);
  const monthHeader = header.includes('月') ? '月' : (header.includes('年月') ? '年月' : null);

  // 直近取り込みの基準ID（copyの末尾 unique id）
  const baseId = getLastNonEmptyInColumn_(copy, 1, HEADER_ROW);

  // 基準の Created Date
  let baseCreatedISO = null;
  if (baseId) {
    const obj = fetchOneById_(token, TYPE_USER, baseId);
    if (obj) {
      const created = obj['Created Date'] ?? obj['Creation Date'];
      const d = normalizeToDate_(created);
      if (d) baseCreatedISO = d.toISOString();
    }
  }

  // 「基準日時より新しい」データを昇順で取得（元実装と同様のロジック） :contentReference[oaicite:10]{index=10}
  const newer = fetchUsersNewerThan_(token, baseCreatedISO, MAX_ROWS);
  if (newer.length) {
    // フラット化して copy へ追記
    const flat = newer.map(o => flatten_(o));
    const rows = flat.map(r => header.map(col => {
      if (monthHeader && col === monthHeader) {
        const dval = pickSmart_(r, YM_SOURCE_CANDIDATES);
        return toYearMonthString_(dval); // "YYYY年M月"
      }
      return pickSmart_(r, [col, ...(ALIASES[col] || [])]);
    }));

    const startRow = copy.getLastRow() + 1;
    const cols = Math.max(copy.getLastColumn(), header.length);
    if (copy.getLastColumn() < cols) copy.insertColumnsAfter(copy.getLastColumn(), cols - copy.getLastColumn());

    const padded = rows.map(row => {
      const out = new Array(cols).fill('');
      for (let i=0;i<Math.min(row.length, cols);i++) out[i]=row[i];
      return out;
    });

    // “月/年月”列がある場合は先にプレーンテキストに（自動日付化対策）
    if (monthHeader) {
      const mcol = header.indexOf(monthHeader) + 1; // 1-based
      copy.getRange(startRow, mcol, padded.length, 1).setNumberFormat('@');
    }

    copy.getRange(startRow, 1, padded.length, cols).setValues(padded);
  }

  // --- ユーザーシートへ“差分のみ”追記（元実装の手順と同じ） :contentReference[oaicite:11]{index=11}
  const userSh = ss.getSheetByName(USER_SHEET) || ss.insertSheet(USER_SHEET);
  let userHeader = getHeader_(userSh, HEADER_ROW);
  if (userHeader.length === 0) {
    userSh.getRange(1,1,1,header.length).setValues([header]); // 初期ヘッダは copy と揃える
    userHeader = header.slice();
  }

  // “ユーザー”の unique id 列の最終データ行（=ID_αのある行）を基準に、連続して追記する
  const lastIdRow = getLastNonEmptyRowByHeader_(userSh, HEADER_ROW, 'unique id'); // 行番号（ヘッダ行なら=1）
  const idAlpha   = (lastIdRow > HEADER_ROW)
    ? String(userSh.getRange(lastIdRow, userHeader.indexOf('unique id')+1).getValue() || '').trim()
    : '';

  const copyVals   = copy.getDataRange().getValues(); // A:unique id 前提
  const copyHeader = copyVals[0].map(v=>String(v||'').trim());
  const idxId      = copyHeader.indexOf('unique id');
  if (idxId < 0) throw new Error('copyシートに「unique id」ヘッダがありません');

  // ID_α が見つかる位置の“次の行”から末尾までを転送
  let startIdx = 1; // 2行目(配列index=1)から
  if (idAlpha) {
    for (let r=1; r<copyVals.length; r++) {
      if (String(copyVals[r][idxId]||'').trim() === idAlpha) { startIdx = r + 1; break; }
    }
  }

  if (startIdx < copyVals.length) {
    const toAppend = copyVals.slice(startIdx);
    if (toAppend.length) {
      // ヘッダ合わせ（ユーザー側に存在する列のみコピー）
      const mapIdx = userHeader.map(h => copyHeader.indexOf(h));
      const out = toAppend.map(row => mapIdx.map(i => (i>=0 ? row[i] : '')));

      // 追記開始位置＝「unique id 列の最終データ行の次の行」
      const start = Math.max(lastIdRow, HEADER_ROW) + 1;

      const width = Math.max(userSh.getLastColumn(), userHeader.length);
      if (userSh.getLastColumn() < width) userSh.insertColumnsAfter(userSh.getLastColumn(), width - userSh.getLastColumn());

      // “月/年月”列がある場合は先にプレーンテキストに（自動日付化対策）
      const userMonthHeader = userHeader.includes('月') ? '月' : (userHeader.includes('年月') ? '年月' : null);
      if (userMonthHeader) {
        const umcol = userHeader.indexOf(userMonthHeader) + 1;
        userSh.getRange(start, umcol, out.length, 1).setNumberFormat('@');
      }

      userSh.getRange(start, 1, out.length, width).setValues(out);

      // ID_β と “月” を保存（PivotA 側が参照）
      const lastRow2 = out[out.length-1];
      const idBeta = String(lastRow2[userHeader.indexOf('unique id')] || '').trim();
      let ym = '';
      if (userMonthHeader) {
        ym = String(lastRow2[userHeader.indexOf(userMonthHeader)] || '').trim();
      } else if (userHeader.includes('Creation Date')) {
        ym = toYearMonthString_(lastRow2[userHeader.indexOf('Creation Date')]);
      }
      const props = PropertiesService.getScriptProperties();
      if (idBeta) props.setProperty('LAST_IMPORTED_ID', idBeta);
      if (ym)     props.setProperty('LAST_IMPORTED_MONTH', ym); // 例: "2025年9月"
    }
  }

  SpreadsheetApp.getActive().toast('CSV import 完了（copy & ユーザーを更新）');
}

/***** API *****/
function fetchOneById_(token, typeName, uniqueId){
  const url = `${BASE_CSV}/api/1.1/obj/${encodeURIComponent(typeName)}/${encodeURIComponent(uniqueId)}`; // 旧: LIVE_BASE を置換 :contentReference[oaicite:12]{index=12}
  const res = UrlFetchApp.fetch(url, { headers:{ Authorization:`Bearer ${token}` }, muteHttpExceptions:true });
  if (res.getResponseCode() === 200) {
    const j = JSON.parse(res.getContentText());
    return j.response?.results ? j.response.results[0] : j.response;
  }
  if (res.getResponseCode() === 404) return null;
  throw new Error(`fetchOneById_ HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
}
// 基準日時より後のみ（Created Date > base）昇順で最大n件
function fetchUsersNewerThan_(token, baseCreatedISO, maxRows){
  let out = [], cursor = 0;
  while (out.length < maxRows) {
    const remaining = maxRows - out.length;
    const perPage   = Math.min(PAGE_LIMIT, remaining);
    const params = { cursor, limit: perPage, sort_field: 'Created Date', descending: 'false' };
    if (baseCreatedISO) {
      params.constraints = JSON.stringify([{ key:'Created Date', constraint_type:'greater than', value: baseCreatedISO }]);
    }
    const url = `${BASE_CSV}/api/1.1/obj/${TYPE_USER}?${buildQuery_(params)}`; // 旧: LIVE_BASE を置換 :contentReference[oaicite:13]{index=13}
    const res = UrlFetchApp.fetch(url, { headers:{ Authorization:`Bearer ${token}` }, muteHttpExceptions:true });
    if (res.getResponseCode() !== 200) throw new Error(`HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    const chunk = (JSON.parse(res.getContentText()).response?.results) || [];
    if (chunk.length === 0) break;
    out = out.concat(chunk);
    cursor += chunk.length;
  }
  return out;
}

/***** ユーティリティ *****/
function buildQuery_(params){
  return Object.entries(params)
    .filter(([,v]) => v !== undefined && v !== null && v !== '')
    .map(([k,v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&');
}
function flatten_(obj, pref=''){
  const out = {};
  for (const [k,v] of Object.entries(obj || {})){
    const key = pref ? `${pref}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) Object.assign(out, flatten_(v, key));
    else out[key] = v;
  }
  return out;
}
function pickSmart_(row, candidates){
  const r = row, norm = s => String(s).toLowerCase().replace(/[\s_]/g,'');
  for (const cand of candidates){
    if (r.hasOwnProperty(cand)) return r[cand];
    if (cand.includes('.')) {
      const v = cand.split('.').reduce((a,k)=>(a && a[k]!==undefined ? a[k] : undefined), r);
      if (v !== undefined) return v ?? '';
    }
    const t = norm(cand);
    for (const k of Object.keys(r)){ if (norm(k) === t) return r[k]; }
  }
  return '';
}
function normalizeToDate_(value){
  if (value == null || value === '') return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return isNaN(value) ? null : value;
  if (typeof value === 'string'){ const d1 = new Date(value); if (!isNaN(d1)) return d1; }
  return null;
}
function toYearMonthString_(value){
  const d = normalizeToDate_(value); if (!d) return '';
  return `${d.getFullYear()}年${d.getMonth()+1}月`;
}

/***** 補助 *****/
// 原本があり複製が無ければ1回だけ複製して返す
function getOrCreateCopySheet_(ss, originalName, copyName){
  const orig = ss.getSheetByName(originalName);
  let copy = ss.getSheetByName(copyName);
  if (!orig) throw new Error(`原本シートが見つかりません: ${originalName}`);
  if (!copy) {
    copy = orig.copyTo(ss).setName(copyName);
    ss.setActiveSheet(copy);
    // コピー直後にA1にヘッダが入る前提のため、余計な行列を削る等は必要に応じて
  }
  return copy;
}
// copyに LINE ID 列が無ければ追加（位置は既存仕様に合わせて適宜）
function ensureLineIdColumnInCopy_(copy){
  const header = getHeader_(copy, HEADER_ROW);
  if (!header.includes('LINE ID')) {
    copy.insertColumnAfter(header.length); // 末尾に追加
    copy.getRange(HEADER_ROW, header.length+1).setValue('LINE ID');
  }
}
function getHeader_(sheet, headerRow){
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(headerRow,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
}
// 指定列の末尾非空セルの値（ID）を返す（無ければ空）
function getLastNonEmptyInColumn_(sheet, colIndex1based, headerRow){
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return '';
  const vals = sheet.getRange(headerRow+1, colIndex1based, lastRow-headerRow, 1).getValues();
  for (let i = vals.length-1; i>=0; i--) {
    const s = String(vals[i][0] ?? '').trim();
    if (s) return s;
  }
  return '';
}
// 指定ヘッダ列の「最後に値が入っている行番号」を返す（無ければ headerRow を返す）
function getLastNonEmptyRowByHeader_(sheet, headerRow, headerName){
  const header = getHeader_(sheet, headerRow);
  const idx = header.indexOf(headerName);
  if (idx < 0) return headerRow;
  const col = idx + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return headerRow;
  const vals = sheet.getRange(headerRow+1, col, lastRow-headerRow, 1).getValues();
  for (let i = vals.length-1; i>=0; i--) {
    const s = String(vals[i][0] ?? '').trim();
    if (s) return headerRow + 1 + i;
  }
  return headerRow;
}

/***** Backward-compat: 旧名に対応 *****/
function menuCsvImportManual() { return appendDiffUsers_toCopySheetA(); }
function runCsvImport()        { return appendDiffUsers_toCopySheetA(); }
