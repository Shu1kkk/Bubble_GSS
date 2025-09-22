/*******************************************************
 * PivotB 一式（最新月自動／0列除外＋空白統合＋空白列削除＋KPI再計算）
 *  - クロス表2種：合計0の列は出力しない
 *  - レイアウト：ピボット見出しは0列のみ採用
 *  - 後処理①："(空白)"→「経路不明」化 & 総計の直前へ移動
 *  - 後処理②：経路不明←→総計の間の“ヘッダー空白列”を削除して総計を左詰め
 *  - 転記＋KPI：行4/7 と C12〜K12 を再計算して書き戻し
 *  - 共有ブック対応：シート取得は「名前→ID自動再バインド（DocumentProperties）」で統一
 *******************************************************/

/* ========= 共有安全版シート取得（gidに依存しない） ========= */
// 例: getSheetSafe_('面談報告') / getSheetSafe_('PivotB_counta', {create:true})
function getSheetSafe_(sheetName, opts) {
  const create = !!(opts && opts.create);
  const ss    = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties(); // 共有（ユーザー差が出ない）
  const key   = `PIVOTB_SHEET_ID__${sheetName}`;

  // 1) 保存gidがあればまず試す
  const saved = props.getProperty(key);
  if (saved) {
    try {
      const sh = ss.getSheetById(Number(saved));
      if (sh) return sh;
    } catch (e) { /* fallthrough */ }
  }
  // 2) 名前で取得（必要なら作成）
  let sh = ss.getSheetByName(sheetName);
  if (!sh && create) sh = ss.insertSheet(sheetName);
  if (!sh) throw new Error(`シート "${sheetName}" が見つかりません`);

  // 3) 新gidを保存（以後はIDで安定参照）
  props.setProperty(key, String(sh.getSheetId()));
  return sh;
}

/* ========= メニュー ========= */
function buildMenu_PivotB_() {
  SpreadsheetApp.getUi()
    .createMenu('Pivot_面談報告')
    .addItem('面談報告AB列オートフィル(最新月)', 'fillAB_fromUsers_values')
    .addSeparator()
    .addItem('クロス作成(最新月)', 'menuPivotB_BuildCrossLatest')
    .addItem('書式(最新月)', 'renderPivotBLayoutA2_v8')  // ← 書式後にKPI再計算まで自動
    .addSeparator()
    .addItem('後処理①: 空白→経路不明 & 総計直前へ', 'fixBlankToUnknownAndPlaceBeforeTotal_')
    .addItem('後処理②: 経路不明←→総計の間の空白列を削除', 'compressBlankColumnsBetweenUnknownAndTotal_')
    .addSeparator()
    .addItem('転記＋KPI(最新月)', 'fillPivotBModifiedValues_v3') // KPI も含む
    .addSeparator()
    .addItem('全部入り(最新月)', 'menuPivotB_RunAllLatest')
    .addToUi();
}

/** クロス作成(最新月) 2種まとめて */
function menuPivotB_BuildCrossLatest() { runPivotBCountA(); runPivotBUniqueCountA(); }

/** 全部入り(最新月)：A/B補完→クロス2種→書式（後処理①②内包）→KPI再計算 */
function menuPivotB_RunAllLatest() {
  fillAB_fromUsers_values();
  runPivotBCountA();
  runPivotBUniqueCountA();
  renderPivotBLayoutA2_v8(); // 内部で後処理①②→KPI再計算まで実施
}

/* ========= A/B補完（ユーザー参照） ========= */
const PIVOTB_ABCFG = { targetSheetName: '面談報告', userSheetName: 'ユーザー' };

function getLastDataRowByColC_(sh) {
  const last = Math.max(2, sh.getLastRow());
  const colC = sh.getRange(2, 3, last - 1, 1).getValues();
  for (let i = colC.length - 1; i >= 0; i--) if (String(colC[i][0]).trim() !== '') return i + 2;
  return 1;
}

function fillAB_fromUsers_values() {
  const ss = SpreadsheetApp.getActive();
  const sh = getSheetSafe_(PIVOTB_ABCFG.targetSheetName);     // 面談報告（既存必須）
  const us = getSheetSafe_(PIVOTB_ABCFG.userSheetName);       // ユーザー（既存必須）

  const lastRow = getLastDataRowByColC_(sh);
  if (lastRow < 2) return;

  const uLast = us.getLastRow(); if (uLast < 2) return;
  const uVals = us.getRange(1, 1, uLast, 4).getValues();
  const dict = new Map();
  for (let r = 1; r < uVals.length; r++) {
    const key = String(uVals[r][0] || '').trim(); if (!key) continue;
    dict.set(key, { b: uVals[r][1], d: uVals[r][3] });
  }

  const n = lastRow - 1;
  const keys = sh.getRange(2, 3, n, 1).getValues();
  const outA = new Array(n), outB = new Array(n);
  for (let i = 0; i < n; i++) {
    const k = String(keys[i][0] || '').trim();
    if (k && dict.has(k)) { const v = dict.get(k); outA[i] = [v.d ?? '']; outB[i] = [v.b ?? '']; }
    else { outA[i] = ['']; outB[i] = ['']; }
  }
  sh.getRange(2, 1, n, 1).setValues(outA);
  sh.getRange(2, 2, n, 1).setValues(outB);
}

/* ========= ログユーティリティ ========= */
const __PIVOTB_LOG__ = [];
function logLine_(msg) {
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const line = `[${stamp}] ${msg}`;
  __PIVOTB_LOG__.push([line]); console.log(line); Logger.log(line);
}
function logFlush_(logSheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = getSheetSafe_(logSheetName, { create: true });
  if (!__PIVOTB_LOG__.length) return;
  const start = (sh.getLastRow() || 0) + 1;
  sh.getRange(start, 1, __PIVOTB_LOG__.length, 1).setValues(__PIVOTB_LOG__);
  __PIVOTB_LOG__.length = 0;
}

/* ========= 共通ユーティリティ ========= */
function toInt_(x) { return Number(x) | 0; }
function normalizeText_(s, opts) {
  let t = String(s ?? ''); if (opts && opts.trim) t = t.trim(); if (opts && opts.caseInsensitive) t = t.toLowerCase(); return t;
}
function parseYearMonth_(v) {
  if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1, ok: true, src: 'date' };
  const s = String(v ?? '').trim(); if (!s) return { y: null, m: null, ok: false, src: 'empty' };
  let m; if ((m = s.match(/^(\d{4})年(\d{1,2})月$/))) return { y: toInt_(m[1]), m: toInt_(m[2]), ok: true, src: 'YYYY年M月' };
  if ((m = s.match(/^(\d{4})[\/\-](\d{1,2})$/))) return { y: toInt_(m[1]), m: toInt_(m[2]), ok: true, src: 'YYYY/MM' };
  if ((m = s.match(/^(\d{1,2})月$/)))          return { y: null, m: toInt_(m[1]), ok: true, src: 'M月' };
  const t = Date.parse(s); if (!isNaN(t)) { const d = new Date(t); return { y: d.getFullYear(), m: d.getMonth() + 1, ok: true, src: 'parsedDate' }; }
  return { y: null, m: null, ok: false, src: 'unknown:' + s };
}
function resolveYearMonth_(row, monthFieldIdx, header, opts, filterY) {
  const pm = parseYearMonth_(row[monthFieldIdx]);
  if (pm.ok && pm.y && pm.m) return { y: pm.y, m: pm.m, src: pm.src };
  if (pm.ok && !pm.y && pm.m) {
    if (opts && opts.useCreationDateYearFallback) {
      for (const cand of (opts.creationDateCandidates || [])) {
        const idx = header.indexOf(cand);
        if (idx >= 0) { const candParsed = parseYearMonth_(row[idx]); if (candParsed.ok && candParsed.y) return { y: candParsed.y, m: pm.m, src: pm.src + '+fallback:' + cand }; }
      }
    }
    if (opts && opts.assumeFilterYearIfMissing && filterY) return { y: filterY, m: pm.m, src: pm.src + '+assumeFilterYear' };
    return { y: null, m: pm.m, src: pm.src + '+noYear' };
  }
  return { y: null, m: null, src: pm.src };
}
function ymKey_(y, m) { return `${y}年${m}月`; }

/** 面談報告!月 の最新年月を検出（年欠落は Creation Date/申請日/確定日で補完） */
function detectLatestYearMonthFromMendan_() {
  const sh = getSheetSafe_('面談報告');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('面談報告にデータがありません');
  const header = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  const idxMonth = header.indexOf('月');
  if (idxMonth === -1) throw new Error('面談報告に「月」列が見つかりません');
  const fallbackCols = ['Creation Date','申請日','確定日'].map(n => header.indexOf(n)).filter(i => i >= 0);

  let best = null;
  for (const r of rows) {
    let ym = parseYearMonth_(r[idxMonth]);
    if (ym.ok && !ym.y && ym.m && fallbackCols.length) {
      for (const fi of fallbackCols) { const f = parseYearMonth_(r[fi]); if (f.ok && f.y) { ym = { y: f.y, m: ym.m, ok: true }; break; } }
    }
    if (ym && ym.y && ym.m) {
      const key = ym.y * 100 + ym.m;
      if (!best || key > best.key) best = { y: ym.y, m: ym.m, key };
    }
  }
  if (!best) { const now = new Date(); return { y: now.getFullYear(), m: now.getMonth() + 1, label: ymKey_(now.getFullYear(), now.getMonth() + 1), monthOnly: `${now.getMonth() + 1}月` }; }
  return { y: best.y, m: best.m, label: ymKey_(best.y, best.m), monthOnly: `${best.m}月` };
}

/* ========= クロス（COUNTA：0列除外） ========= */
function runPivotBCountA() {
  const latest = detectLatestYearMonthFromMendan_();
  const CONFIG = {
    sourceSheetName: '面談報告',
    targetSheetName: 'PivotB_counta',
    logSheetName: 'PivotB_log',
    rowField: null,
    colField: '経路',
    valueField: 'R_ユーザー情報',
    filters: { '月': latest.label },

    includeBlankBucket: true,
    blankLabel: '(空白)',
    showTotals: true,
    dropZeroColumns: true,

    monthOptions: {
      useCreationDateYearFallback: true,
      creationDateCandidates: ['Creation Date', '申請日', '確定日'],
      assumeFilterYearIfMissing: true
    }
  };

  try { buildCrossTabWithLogs_v3_drop0_(CONFIG); logFlush_(CONFIG.logSheetName); }
  catch (err) { logLine_('❌ 例外発生: ' + (err && err.stack ? err.stack : err)); logFlush_(CONFIG.logSheetName); throw err; }
}

function buildCrossTabWithLogs_v3_drop0_(params) {
  logLine_('▶ 開始: クロス表作成 v3（0列除外）');

  const src = getSheetSafe_(params.sourceSheetName);
  const values = src.getDataRange().getValues();
  if (values.length < 2) throw new Error('データがありません（ヘッダーのみ）');
  const header = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  const idx = (name) => { const i = header.indexOf(String(name).trim()); if (i === -1) throw new Error(`ヘッダーが見つかりません: "${name}"`); return i; };
  const rowIdx = params.rowField ? idx(params.rowField) : -1;
  const colIdx = idx(params.colField);
  const valIdx = idx(params.valueField);
  const filterIdxMap = {}; for (const k in (params.filters || {})) filterIdxMap[k] = idx(k);

  // 月フィルタ
  let filterY = null, filterM = null, filterKey = null;
  if ('月' in (params.filters || {})) {
    const parsed = parseYearMonth_(params.filters['月']); if (!parsed.ok) throw new Error(`フィルタ "月" を解釈できません: ${params.filters['月']}`);
    filterM = parsed.m; filterY = parsed.y || null;
    if (!filterY) {
      const mi = filterIdxMap['月']; let bestY = null;
      for (const r of rows) { const resolved = resolveYearMonth_(r, mi, header, params.monthOptions || {}, null); if (resolved.m === filterM && resolved.y) bestY = Math.max(bestY ?? 0, resolved.y); }
      if (!bestY) throw new Error(`フィルタ"月=${params.filters['月']}"の年が特定できません`);
      filterY = bestY;
    }
    filterKey = ymKey_(filterY, filterM);
  }
  logLine_(`適用フィルタ: 月="${filterKey}"`);

  const filtered = rows.filter(r => {
    for (const f in filterIdxMap) {
      if (f === '月') {
        const resolved = resolveYearMonth_(r, filterIdxMap[f], header, params.monthOptions || {}, filterY);
        if (!(resolved.y && resolved.m)) return false;
        if (resolved.y !== filterY || resolved.m !== filterM) return false;
      } else if (String(r[filterIdxMap[f]]).trim() !== String(params.filters[f]).trim()) return false;
    }
    return true;
  });

  // 列候補（出現順）
  const allColKeys = []; const seenCols = new Set();
  for (const r of filtered) {
    let key = String(r[colIdx]).trim();
    if (!key && params.includeBlankBucket) key = params.blankLabel;
    if (!key && !params.includeBlankBucket) continue;
    if (!seenCols.has(key)) { seenCols.add(key); allColKeys.push(key); }
  }
  const target = getSheetSafe_(params.targetSheetName, { create: true });
  if (!allColKeys.length) { target.clear().getRange(1,1).setValue('データなし（該当する経路がありません）'); return; }

  // 行キー
  const rowKeys = [];
  if (rowIdx >= 0) {
    const seenRows = new Set();
    for (const r of filtered) { const label = String(r[rowIdx]).trim(); if (!seenRows.has(label)) { seenRows.add(label); rowKeys.push(label); } }
    if (!rowKeys.length) rowKeys.push('（該当なし）');
  } else rowKeys.push('件数');

  // 集計
  const mat = {}; for (const rk of rowKeys) { mat[rk] = {}; for (const ck of allColKeys) mat[rk][ck] = 0; }
  for (const r of filtered) {
    const colRaw = String(r[colIdx]).trim();
    const cKey = colRaw || (params.includeBlankBucket ? params.blankLabel : ''); if (!cKey) continue;
    const v = String(r[valIdx]).trim(); if (!v) continue;
    const rKey = (rowIdx >= 0) ? String(r[rowIdx]).trim() : '件数'; if (!(rKey in mat) || !(cKey in mat[rKey])) continue;
    mat[rKey][cKey] += 1;
  }

  // 列合計 → 0列を除外
  const colTotals = {}; for (const ck of allColKeys) colTotals[ck] = 0;
  for (const rk of rowKeys) for (const ck of allColKeys) colTotals[ck] += (mat[rk][ck] || 0);
  const colKeys = (params.dropZeroColumns === false) ? allColKeys : allColKeys.filter(ck => (colTotals[ck] || 0) > 0);
  if (!colKeys.length) { target.clear().getRange(1,1).setValue('データはありますが、すべての列合計が 0 でした。'); return; }

  // 出力
  const showTotals = !!params.showTotals;
  const headerRow = [''].concat(colKeys); if (showTotals) headerRow.push('合計');
  const out = [headerRow];

  for (const rk of rowKeys) {
    const line = [rk]; let rowSum = 0;
    for (const ck of colKeys) { const n = mat[rk][ck] || 0; line.push(n); rowSum += n; }
    if (showTotals) line.push(rowSum);
    out.push(line);
  }
  if (showTotals) {
    const tot = ['合計'];
    for (const ck of colKeys) tot.push(colTotals[ck] || 0);
    tot.push((tot.length>1) ? tot.slice(1).reduce((a,b)=>a+(+b||0),0) : 0);
    out.push(tot);
  }

  target.clear();
  target.getRange(1, 1, out.length, out[0].length).setValues(out);
  target.setFrozenRows(1);
  target.getRange(1,1,1,out[0].length).setFontWeight('bold');
  target.getRange(2,2,out.length-1,out[0].length-1).setNumberFormat('0');
  target.autoResizeColumns(1, out[0].length);
  target.getRange(1,1,out.length,out[0].length).setBorder(true,true,true,true,true,true);

  logLine_(`出力: ${out.length}行 x ${out[0].length}列 / 採用列=${colKeys.length}（0列除外）`);
}

/* ========= クロス（COUNTA UNIQUE：0列除外） ========= */
function runPivotBUniqueCountA() {
  const latest = detectLatestYearMonthFromMendan_();
  const CONFIG = {
    sourceSheetName: '面談報告',
    targetSheetName: 'PivotB_countaunique',
    logSheetName: 'PivotB_log',
    rowField: null,
    colField: '経路',
    valueField: 'R_ユーザー情報',
    uniqueKeyField: 'R_ユーザー情報',
    filters: { '月': latest.label },
    includeBlankBucket: true, blankLabel: '(空白)', showTotals: true, dropZeroColumns: true,
    monthOptions: { useCreationDateYearFallback: true, creationDateCandidates: ['Creation Date','申請日','確定日'], assumeFilterYearIfMissing: true },
    uniqueNormalize: { trim: true, caseInsensitive: true }
  };

  try { buildCrossTabUniqueWithLogs_v3_drop0_(CONFIG); logFlush_(CONFIG.logSheetName); }
  catch (err) { logLine_('❌ 例外発生: ' + (err && err.stack ? err.stack : err)); logFlush_(CONFIG.logSheetName); throw err; }
}

function buildCrossTabUniqueWithLogs_v3_drop0_(params) {
  logLine_('▶ 開始: クロス表作成 v3（unique・0列除外）');

  const src = getSheetSafe_(params.sourceSheetName);
  const values = src.getDataRange().getValues();
  if (values.length < 2) throw new Error('データがありません（ヘッダーのみ）');
  const header = values[0].map(h => String(h).trim());
  const rows = values.slice(1);
  const idx = (name) => { const i = header.indexOf(String(name).trim()); if (i === -1) throw new Error(`ヘッダーが見つかりません: "${name}"`); return i; };

  const rowIdx = params.rowField ? idx(params.rowField) : -1;
  const colIdx = idx(params.colField);
  const valIdx = idx(params.valueField);
  const uniqueIdx = params.uniqueKeyField ? idx(params.uniqueKeyField) : valIdx;

  const filterIdxMap = {}; for (const k in (params.filters || {})) filterIdxMap[k] = idx(k);

  let filterY = null, filterM = null, filterKey = null;
  if ('月' in (params.filters || {})) {
    const parsed = parseYearMonth_(params.filters['月']); if (!parsed.ok) throw new Error(`フィルタ "月" を解釈できません: ${params.filters['月']}`);
    filterM = parsed.m; filterY = parsed.y || null;
    if (!filterY) {
      const mi = filterIdxMap['月']; let bestY = null;
      for (const r of rows) { const resolved = resolveYearMonth_(r, mi, header, params.monthOptions || {}, null); if (resolved.m === filterM && resolved.y) bestY = Math.max(bestY ?? 0, resolved.y); }
      if (!bestY) throw new Error(`フィルタ"月=${params.filters['月']}"の年が特定できません`);
      filterY = bestY;
    }
    filterKey = ymKey_(filterY, filterM);
  }
  logLine_(`適用フィルタ: 月="${filterKey}"`);

  const filtered = rows.filter(r => {
    for (const f in filterIdxMap) {
      if (f === '月') {
        const resolved = resolveYearMonth_(r, filterIdxMap[f], header, params.monthOptions || {}, filterY);
        if (!(resolved.y && resolved.m)) return false;
        if (resolved.y !== filterY || resolved.m !== filterM) return false;
      } else if (String(r[filterIdxMap[f]]).trim() !== String(params.filters[f]).trim()) return false;
    }
    return true;
  });

  const allColKeys = []; const seenCols = new Set();
  for (const r of filtered) {
    let key = String(r[colIdx]).trim();
    if (!key && params.includeBlankBucket) key = params.blankLabel;
    if (!key && !params.includeBlankBucket) continue;
    if (!seenCols.has(key)) { seenCols.add(key); allColKeys.push(key); }
  }
  const target = getSheetSafe_(params.targetSheetName, { create: true });
  if (!allColKeys.length) { target.clear().getRange(1,1).setValue('データなし（該当する経路がありません）'); return; }

  const rowKeys = [];
  if (rowIdx >= 0) {
    const seenRows = new Set();
    for (const r of filtered) { const label = String(r[rowIdx]).trim(); if (!seenRows.has(label)) { seenRows.add(label); rowKeys.push(label); } }
    if (!rowKeys.length) rowKeys.push('（該当なし）');
  } else rowKeys.push('件数');

  // 集計（ユニーク）
  const matSets = {}; for (const rk of rowKeys) { matSets[rk] = {}; for (const ck of allColKeys) matSets[rk][ck] = new Set(); }
  for (const r of filtered) {
    const cKey = String(r[colIdx]).trim() || (params.includeBlankBucket ? params.blankLabel : ''); if (!cKey) continue;
    const rKey = (rowIdx >= 0) ? String(r[rowIdx]).trim() : '件数'; if (!(rKey in matSets) || !(cKey in matSets[rKey])) continue;
    let uKey = String(r[uniqueIdx] != null ? r[uniqueIdx] : ''); uKey = normalizeText_(uKey, params.uniqueNormalize || { trim: true, caseInsensitive: true });
    if (!uKey) continue;
    matSets[rKey][cKey].add(uKey);
  }

  // 列合計サイズ → 0列を除外
  const colTotals = {}; for (const ck of allColKeys) colTotals[ck] = 0;
  for (const rk of rowKeys) for (const ck of allColKeys) colTotals[ck] += matSets[rk][ck].size;

  const colKeys = (params.dropZeroColumns === false) ? allColKeys : allColKeys.filter(ck => (colTotals[ck] || 0) > 0);
  if (!colKeys.length) { target.clear().getRange(1,1).setValue('データはありますが、すべての列合計が 0 でした。'); return; }

  const showTotals = !!params.showTotals;
  const headerRow = [''].concat(colKeys); if (showTotals) headerRow.push('合計');
  const out = [headerRow];

  let grandTotal = 0;
  for (const rk of rowKeys) {
    const line = [rk]; let rowSum = 0;
    for (const ck of colKeys) { const n = matSets[rk][ck].size; line.push(n); rowSum += n; }
    if (showTotals) { line.push(rowSum); grandTotal += rowSum; }
    out.push(line);
  }
  if (showTotals) { const tot = ['合計']; for (const ck of colKeys) tot.push(colTotals[ck] || 0); tot.push(grandTotal); out.push(tot); }

  target.clear();
  target.getRange(1, 1, out.length, out[0].length).setValues(out);
  target.setFrozenRows(1);
  target.getRange(1,1,1,out[0].length).setFontWeight('bold');
  target.getRange(2,2,out.length-1,out[0].length-1).setNumberFormat('0');
  target.autoResizeColumns(1, out[0].length);
  target.getRange(1,1,out.length,out[0].length).setBorder(true,true,true,true,true,true);

  logLine_(`出力: ${out.length}行 x ${out[0].length}列 / 採用列=${colKeys.length}（0列除外）`);
}

/* ========= レイアウト v8（書式→後処理①②→KPI再計算まで自動） ========= */
function renderPivotBLayoutA2_v8() {
  const SS  = SpreadsheetApp.getActiveSpreadsheet();
  const PIV = getSheetSafe_('PivotB_counta');
  if (!PIV || PIV.getLastRow() < 2 || PIV.getLastColumn() < 2) {
    throw new Error('先に「クロス作成(最新月)」を実行してください。');
  }

  // 1) アクティブ見出し収集
  const lc = PIV.getLastColumn();
  const h1 = PIV.getRange(1, 1, 1, lc).getValues()[0].map(s => String(s || '').trim());
  const r2 = PIV.getRange(2, 1, 1, lc).getValues()[0];
  const isPos = (v) => (typeof v === 'number' ? v : Number(String(v||'').replace(/,/g,''))) > 0;

  const active = [];
  for (let c = 1; c <= lc; c++) {
    const label = h1[c - 1];
    if (!label || label === '合計') continue;
    if (isPos(r2[c - 1])) active.push(label);
  }

  // 2) 見出し構成
  const FIXED_FP_AP = ['FP_一條','FP_三田','FP_枝川','FP_松下','FP_西','FP_西田','FP_青木','FP_洗','FP_大山','FP_大澤','FP_滝澤','FP_鳥山','FP_辻村','FP_白岩','FP_北野','FP_廣瀬'];
  const set       = new Set(active);
  const fpFixed   = FIXED_FP_AP.filter(x => set.has(x));
  const fpExtra   = active.filter(x => /^FP_/.test(x) && !fpFixed.includes(x));
  const nonFP     = active.filter(x => !/^FP_/.test(x) && x !== '経路不明');
  const hasAnyFP  = fpFixed.length + fpExtra.length > 0;
  const hasUnknown= set.has('経路不明') || set.has('空白') || set.has('(空白)');

  const headers = [
    ...fpFixed,
    ...fpExtra,
    ...(hasAnyFP ? ['FP'] : []),
    ...nonFP,
    ...(hasUnknown ? ['経路不明'] : []),
    '総計'
  ];

  // 3) 出力シートの用意（PivotB_auto に統一。旧 pivotB_auto があればリネーム）
  const ALT = SpreadsheetApp.getActive().getSheetByName('pivotB_auto');
  if (ALT && !SpreadsheetApp.getActive().getSheetByName('PivotB_auto')) ALT.setName('PivotB_auto');

  const OUT = getSheetSafe_('PivotB_auto', { create: true }); // 共有安全
  OUT.setTabColor('#1a73e8'); // 青タブ

  // 4) レイアウト初期化
  OUT.clear();
  OUT.setHiddenGridlines(true);
  const R = (r, c, rr = 1, cc = 1) => OUT.getRange(r, c, rr, cc);

  // 5) 見出し描画
  const latest = detectLatestYearMonthFromMendan_();
  R(2, 1).setValue(`${latest.m}月`);
  R(3, 1, 1, headers.length).setValues([headers]);
  R(6, 1, 1, headers.length).setValues([headers]);

  // バンド（装飾）
  const bandW = Math.max(headers.length, 20);
  OUT.getRange(1, 1, 1, bandW).setBackground('#cfcfcf');
  OUT.getRange(27,1, 1, bandW).setBackground('#cfcfcf');
  R(3, 1, 2, headers.length).setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');
  R(6, 1, 2, headers.length).setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');

  // FP列の目印
  const fpCol = headers.indexOf('FP') + 1;
  if (fpCol > 0) {
    R(3, fpCol).setBackground('#f7dce8');
    R(6, fpCol).setBackground('#f7dce8');
  }

  // KPI枠の見出し
  R(10, 3).setValue('FP');
  R(10, 8).setValue('その他経路');
  R(11, 3, 1, 4).setValues([['登録数（まねぽんのみ）','面談人数','面談移行率','1人当たりの面談件数']]).setBackground('#fde2ea').setHorizontalAlignment('center');
  R(11, 8, 1, 4).setValues([['登録数','面談人数','面談移行率','1人当たりの面談件数']]).setBackground('#fde2ea').setHorizontalAlignment('center');

  R(10, 2).setValue(`${latest.m}月`);
  R(12, 2).setValue('投資コンシェル');
  R(13, 2).setValue('みらいマップ');

  // 罫線
  R(11, 2, 3, 2).setBorder(true,true,true,true,true,true);
  R(11, 4, 2, 3).setBorder(true,true,true,true,true,true);
  R(11, 8, 2, 4).setBorder(true,true,true,true,true,true);
  R(11, 2, 3, 1).setBorder(true,true,true,true,true,true);

  // 6) 後処理①②
  fixBlankToUnknownAndPlaceBeforeTotal_('PivotB_auto');
  compressBlankColumnsBetweenUnknownAndTotal_('PivotB_auto');

  // 7) KPI再計算して反映
  fillPivotBModifiedValues_v3();
}

/* ========= 転記＋KPI（列移動/削除後のヘッダーを見て安全に書く） ========= */
function fillPivotBModifiedValues_v3() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const out = getSheetSafe_('PivotB_auto');

  const LOG = [];
  const log = (m) => { const t = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); const line = `[${t}] ${m}`; LOG.push([line]); console.log(line); Logger.log(line); };
  const flushLog = () => { const ls = getSheetSafe_('PivotB_log', { create: true }); const start = (ls.getLastRow() || 0) + 1; if (LOG.length) ls.getRange(start, 1, LOG.length, 1).setValues(LOG); };
  const num = (v) => (typeof v === 'number') ? v : Number(String(v ?? '').replace(/,/g,'')) || 0;

  const lastCol = out.getLastColumn();
  if (lastCol < 2) return;

  const hdr3 = out.getRange(3, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());
  const hdr6 = out.getRange(6, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());

  // ピボット -> dict
  function dictFromPivot(sheetName){
    const sh = getSheetSafe_(sheetName);
    if(!sh || sh.getLastRow()<2 || sh.getLastColumn()<1) return {};
    const lc=sh.getLastColumn();
    const h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(s=>String(s||'').trim());
    const v=sh.getRange(2,1,1,lc).getValues()[0]; const d={}; for(let i=0;i<h.length;i++) d[h[i]]=v[i]; return d;
  }
  const dictUnique = dictFromPivot('PivotB_countaunique');
  const dictCountA = dictFromPivot('PivotB_counta');

  // 行4 / 行7 の書き込み
  function buildRow(hdr, dict) {
    const row = hdr.map(h => num(dict[h]));
    const idxTotal = hdr.indexOf('総計');
    const totalFromDict = ('合計' in dict) ? num(dict['合計']) : NaN;
    if (idxTotal >= 0) row[idxTotal] = !isNaN(totalFromDict) ? totalFromDict : row.reduce((s, v, i) => (i === idxTotal ? s : s + num(v)), 0);
    const idxFP = hdr.indexOf('FP');
    if (idxFP >= 0) row[idxFP] = hdr.reduce((s,h,i)=> s + (/^FP_/.test(h) ? num(row[i]) : 0), 0);
    const idxUnknown = hdr.indexOf('経路不明');
    if (idxUnknown >= 0) {
      const total = (idxTotal >= 0) ? num(row[idxTotal]) : 0;
      const fpSum = (idxFP >= 0) ? num(row[idxFP]) : 0;
      const others = hdr.reduce((s,h,i)=> { if(/^FP_/.test(h)||h==='FP'||h==='経路不明'||h==='総計') return s; return s + num(row[i]); }, 0);
      row[idxUnknown] = total - fpSum - others;
    }
    return row;
  }
  const row4 = buildRow(hdr3, dictUnique); if (row4.length) out.getRange(4, 1, 1, row4.length).setValues([row4]);
  const row7 = buildRow(hdr6, dictCountA); if (row7.length) out.getRange(7, 1, 1, row7.length).setValues([row7]);

  // ====== KPI: C12..K12 を再計算・反映 ======
  function valueBelowHeader(hdrRowIndex, label) {
    const hdr = (hdrRowIndex === 3) ? hdr3 : hdr6;
    const bodyRow = (hdrRowIndex === 3) ? 4 : 7;
    const idx = hdr.indexOf(label); if (idx === -1) return 0;
    return num(out.getRange(bodyRow, idx + 1).getValue());
  }

  const latest = detectLatestYearMonthFromMendan_();
  const monthLabel = `${latest.m}月`;
  const rowKeyAll      = `${monthLabel}全体`;   // C12 用
  const rowKeyAllCount = `${monthLabel}全体数`; // C13 用

  function getManepOnFixedHeader_v3(url, tag, rowKey, headerRows) {
    const norm = (s) => String(s ?? '').trim().normalize('NFKC');
    try {
      const book = SpreadsheetApp.openByUrl(url);
      const sh = book.getSheetByName('マッチング状況');
      if (!sh) return 0;
      const lastR = sh.getLastRow(), lastC = sh.getLastColumn();
      if (lastR < 2 || lastC < 2) return 0;

      const hdrRows = Array.isArray(headerRows) ? headerRows : [headerRows];
      let manepCol = -1, usedHeaderRow = -1;
      for (const hr of hdrRows) {
        if (hr < 1 || hr > lastR) continue;
        const headerRow = sh.getRange(hr, 1, 1, lastC).getDisplayValues()[0];
        const headerNorm = headerRow.map(norm);
        let idx = headerNorm.findIndex(h => h === 'まねぽん'); if (idx === -1) idx = headerNorm.findIndex(h => h.includes('まねぽん'));
        if (idx !== -1) { manepCol = idx + 1; usedHeaderRow = hr; break; }
      }
      if (manepCol === -1) return 0;

      const dataStart = usedHeaderRow + 1;
      const bDisp = sh.getRange(dataStart, 2, lastR - dataStart + 1, 1).getDisplayValues().map(r => r[0]);
      const normB = bDisp.map(v => norm(v));
      const key = norm(rowKey);
      const idxes = []; const strip = (s) => s.replace(/\s/g,'');
      for (let i = 0; i < normB.length; i++) if (normB[i] === key || strip(normB[i]) === strip(key)) idxes.push(i);
      if (!idxes.length) return 0;

      let chosenRow = dataStart + idxes[idxes.length - 1];
      let v = sh.getRange(chosenRow, manepCol).getValue();
      let n = (typeof v === 'number') ? v : (String(v).trim() ? Number(String(v).replace(/,/g,'')) : 0);
      if (n === 0 && String(v).trim() === '') {
        for (let k = idxes.length - 2; k >= 0; k--) {
          const r = dataStart + idxes[k];
          v = sh.getRange(r, manepCol).getValue();
          n = (typeof v === 'number') ? v : (String(v).trim() ? Number(String(v).replace(/,/g,'')) : 0);
          if (n !== 0 || String(v).trim() !== '') { chosenRow = r; break; }
        }
      }
      return n;
    } catch (e) { return 0; }
  }

  // ① C12/C13：外部シート（URLは運用に合わせて調整）
  const C12_val = getManepOnFixedHeader_v3(
    'https://docs.google.com/spreadsheets/d/1ug-8pYA4sx3BrRQBnHuns0V_xxGQsYNFT2ipyjKpDQU/edit?gid=607304295#gid=607304295',
    'C12', `${monthLabel}全体`, 2
  );
  const C13_val = getManepOnFixedHeader_v3(
    'https://docs.google.com/spreadsheets/d/1AyhnVHOO20B83HZ0scNcHOOL5ABzp_Jazu-wAjvHcvc/edit?gid=205978315#gid=205978315',
    'C13', `${monthLabel}全体数`, 3
  );

  // ② D12〜F12：本シートの FP/総計 を使って算出
  const D12_val = valueBelowHeader(3, 'FP');                      // FP_*総計（unique）
  const E12_val = (C12_val + C13_val) ? D12_val / (C12_val + C13_val) : '';
  const F12_val = D12_val ? valueBelowHeader(6, 'FP') / D12_val : '';

  // ③ H12〜K12：その他系
  const shA = (function(){ try { return getSheetSafe_('PivotA_auto'); } catch(e){ return null; } })();
  function firstOrLastNumberBelowHeader(sheet, headerRow, headerLabel, which) {
    if (!sheet) return 0;
    const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
    const hdr = sheet.getRange(headerRow, 1, 1, lc).getDisplayValues()[0].map(s => String(s || '').trim());
    const idx = hdr.indexOf(headerLabel); if (idx === -1) return 0;
    const col = idx + 1; const vals = sheet.getRange(headerRow + 1, col, lr - headerRow, 1).getValues().map(r => r[0]);
    const toNum = (x) => (typeof x === 'number') ? x : (String(x).trim() ? Number(String(x).replace(/,/g,'')) : null);
    if (which === 'first') { for (let i = 0; i < vals.length; i++) { const n = toNum(vals[i]); if (n !== null && !isNaN(n)) return n; } }
    else { for (let i = vals.length - 1; i >= 0; i--) { const n = toNum(vals[i]); if (n !== null && !isNaN(n)) return n; } }
    return 0;
  }
  const fpFirst = firstOrLastNumberBelowHeader(shA, 6, 'FP', 'first');
  const fpLast  = firstOrLastNumberBelowHeader(shA, 6, 'FP', 'last');
  const H12_val = fpLast - fpFirst;

  const I12_val = valueBelowHeader(3, '総計') - valueBelowHeader(3, 'FP');
  const J12_val = H12_val ? I12_val / H12_val : '';
  const K12_num = valueBelowHeader(6, '総計') - valueBelowHeader(6, 'FP');
  const K12_val = I12_val ? (K12_num / I12_val) : '';

  // ===== 反映 =====
  out.getRange('C12').setValue(C12_val);
  out.getRange('D12').setValue(D12_val);
  out.getRange('E12').setValue(E12_val).setNumberFormat('0.00%');
  out.getRange('F12').setValue(F12_val).setNumberFormat('0.000');
  out.getRange('C13').setValue(C13_val);

  out.getRange('H12').setValue(H12_val);
  out.getRange('I12').setValue(I12_val);
  out.getRange('J12').setValue(J12_val).setNumberFormat('0.0%');
  out.getRange('K12').setValue(K12_val).setNumberFormat('0.000');

  log('KPI(C12〜K12) を再計算して反映しました。');
  flushLog();
}

/* ========= 後処理①：空白→経路不明 & 総計直前へ（列数可変対応） ========= */
function fixBlankToUnknownAndPlaceBeforeTotal_(sheetName) {
  const name = sheetName || 'PivotB_auto';
  const sh = getSheetSafe_(name);

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 2 || lastRow < 3) { SpreadsheetApp.getActive().toast('データが足りません（列/行不足）'); return; }

  const hdr3 = sh.getRange(3, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());
  const hdr6 = sh.getRange(6, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());

  const findColByLabel = (labels) => {
    for (let c = 1; c <= lastCol; c++) {
      const h3 = hdr3[c - 1], h6 = hdr6[c - 1];
      if (labels.includes(h3) || labels.includes(h6)) return c;
    }
    return -1;
  };

  const cBlank = findColByLabel(['空白', '(空白)']);
  if (cBlank === -1) { SpreadsheetApp.getActive().toast('「空白」列が見つかりませんでした（処理不要）'); return; }

  let cTotal = findColByLabel(['総計']);
  if (cTotal === -1) throw new Error('「総計」列が見つかりません。先にピボット生成を実行してください。');

  if (cBlank === cTotal - 1) {
    sh.getRange(3, cBlank).setValue('経路不明');
    sh.getRange(6, cBlank).setValue('経路不明');
    SpreadsheetApp.getActive().toast('改名のみ実施（位置は既に総計の直前でした）');
    return;
  }

  sh.insertColumnBefore(cTotal);
  sh.getRange(1, cBlank, lastRow, 1).copyTo(sh.getRange(1, cTotal - 1, lastRow, 1), { contentsOnly: true });
  sh.getRange(3, cTotal - 1).setValue('経路不明');
  sh.getRange(6, cTotal - 1).setValue('経路不明');

  const delIndex = (cBlank >= cTotal) ? (cBlank + 1) : cBlank;
  sh.deleteColumn(delIndex);

  SpreadsheetApp.getActive().toast('「空白」→「経路不明」へ改名し、総計の直前へ移動しました');
}

/* ========= 後処理②：経路不明←→総計の間の“ヘッダー空白列”を削除 ========= */
function compressBlankColumnsBetweenUnknownAndTotal_(sheetName) {
  const name = sheetName || 'PivotB_auto';
  const sh = getSheetSafe_(name);
  const lastCol = sh.getLastColumn(), lastRow = sh.getLastRow(); if (lastCol < 2 || lastRow < 1) return;

  // ヘッダー行を検出（3/6 行を優先、無ければ上位 80 行スキャン）
  const candidateRows = [3,6]; const maxScan = Math.min(80, lastRow);
  for (let r = 1; r <= maxScan; r++) if (!candidateRows.includes(r)) candidateRows.push(r);

  let headerRow = -1, idxUnknown = -1, idxTotal = -1, hdr = null;
  for (const r of candidateRows) {
    if (r > lastRow) continue;
    const row = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s || '').trim());
    const iu = row.indexOf('経路不明'); const it = row.indexOf('総計');
    if (iu >= 0 && it >= 0){ headerRow = r; idxUnknown = iu; idxTotal = it; hdr = row; break; }
  }
  if (headerRow === -1){ SpreadsheetApp.getActive().toast('ヘッダー行が特定できなかったため、空白列削除はスキップ'); return; }

  if (idxUnknown > idxTotal){ const t = idxUnknown; idxUnknown = idxTotal; idxTotal = t; }

  const deleteCols = []; // 1-based col index
  for (let i = idxUnknown + 1; i <= idxTotal - 1; i++) {
    const label = hdr[i];
    if (label === '' || label === null || typeof label === 'undefined') deleteCols.push(i + 1);
  }
  if (!deleteCols.length){ SpreadsheetApp.getActive().toast('削除対象となる「空白ヘッダー列」はありません'); return; }

  deleteCols.sort((a,b)=>b-a).forEach(c=>sh.deleteColumn(c));
  SpreadsheetApp.getActive().toast(`空白ヘッダー列を ${deleteCols.length} 列削除し、「総計」を左詰めしました`);
}
