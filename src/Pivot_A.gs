/*********************** PivotA（ユーザー｜最新月→バケット＋クロス→転記）— 完全版 ************************
 * 1) 書式（pibotA_modified_auto）
 *    - 9行: A〜=アフィリエイター / FP列=「アフィリエイター」
 *    - 12行: A〜=ASP            / FP列=「ASP」
 *    - 15行: A〜=ポイントサイト / FP列=「ポイントサイト」
 *    - 未分類なし：A18&FP18=経路不明, FP21=総計, 4行/24行=グレー帯
 *      未分類あり：A18〜=未分類, FP18=未分類項目, A21&FP21=経路不明, FP24=総計, 4行/27行=グレー帯
 *    - 6行=FP_* 見出し, 7行=A〜に FP_* ベクタ、FP列にその合計
 * 2) クロス（testA_user_copy_pivot）
 *    - unique id があれば「月内の最新 非FP」×「月内の最新 FP_*」で1カウント
 * 3) 転記（今回のルール）
 *    - pivot の「経路不明」行の FP_* 値 → Aシート 7行 A〜（FP列左まで）
 *    - 同ベクタ合計 → 7行 FP 列
 *    - pivot の「経路不明, FP」値 → A/FP の “経路不明ラベル直下” に入れる
 *    - 各バケット行（9/12/15/必要なら18）の “ラベル直下” に pivot の「FP列」の値（=各流入経路の合計）を反映
 *    - 各値行の合計を FP 列に入れる
 *    - 総計 = FP_*合計 + アフィリエイター合計 + ASP合計 + ポイントサイト合計 +（未分類があれば）未分類合計 + 経路不明合計
 **********************************************************************************************************/

//////////////////// 設定 ////////////////////
const PVA_CFG = {
  SRC_SHEET_NAME : 'ユーザー',
  HEADER_ROW_SRC : 1,
  COL_CHANNEL    : '流入経路',
  COL_MONTH      : '月',
  COL_CREATED    : 'Creation Date',
  COL_UID        : 'unique id',

  OUT_SHEET_NAME   : 'PivotA_auto',
  CROSS_SHEET_NAME : 'testA_user_copy_pivot',

  ROW_AFF     : 9,
  ROW_ASP     : 12,
  ROW_POINT   : 15,
  ROW_OTHERS  : 18,  // 未分類 or（未分類なし時は）経路不明のラベル行
  ROW_UNKNOWN2: 21,  // 未分類あり時の経路不明ラベル行
  ROW_TOTAL   : 24,  // 未分類ありは総計ラベル行 / 未分類なしはグレー帯行
  ROW_GRAY_HAS_OTHERS: 27, // 未分類ありのグレー帯

  COLOR_GRAY  : '#e6e6e6',

  FP_BASE: [
    'FP_一條','FP_三田','FP_枝川','FP_松下','FP_西','FP_西田','FP_青木',
    'FP_洗','FP_大山','FP_鳥山','FP_辻村','FP_庭山','FP_白岩','FP_北野','FP_廣瀬'
  ],
  FP_TAIL_LABEL: 'FP'
};

// バケット（複数ヒット→未分類）
const PAT = {
  FP:        /^FP_/i,
  AFF:       /(?:Twitter|\bX\b|ブログ|Youtuve|You.?Tube?|\bYouTube\b|アフィリエイター)/i,
  ASP:       /(レントラックス|\bASP\b)/i,
  POINT:     /(ハピタス|Zucks|ポイントインカム|poikey|ポイントサイト)/i
};

//////////////////// メニュー ////////////////////
function buildMenu_PivotA_() {
  SpreadsheetApp.getUi()
    .createMenu('Pivot_user')
    .addItem('書式（最新月）', 'buildFormatOnlyLatest')
    .addItem('クロス作成（最新月）', 'buildCrossLatest')
    .addItem('転記（最新月）', 'transferFromCrossLatest')
    .addSeparator()
    .addItem('全部入り（最新月）', 'buildFormatCrossAndTransferLatest')
    .addToUi();
}

//////////////////// エントリ ////////////////////
function buildFormatOnlyLatest(){ const t=_findLatestMonth(); if(!t) _err('最新月が判定できません'); _buildFormatForTarget(t); }
function buildCrossLatest(){       const t=_findLatestMonth(); if(!t) _err('最新月が判定できません'); _buildCrossForTarget(t); }
function transferFromCrossLatest(){const t=_findLatestMonth(); if(!t) _err('最新月が判定できません'); _transferFromCrossForTarget(t); }
function buildFormatCrossAndTransferLatest(){
  const t=_findLatestMonth(); if(!t) _err('最新月が判定できません');
  _buildFormatForTarget(t);
  _buildCrossForTarget(t);
  _transferFromCrossForTarget(t);
}

//////////////////// 書式作成 ////////////////////
function _buildFormatForTarget(target){
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(PVA_CFG.SRC_SHEET_NAME);
  if (!src) _err(`元データが見つかりません: ${PVA_CFG.SRC_SHEET_NAME}`);

  const { headers, rows } = _readAll(src);
  const iCh = headers.indexOf(PVA_CFG.COL_CHANNEL);
  if (iCh < 0) _err(`ヘッダーに「${PVA_CFG.COL_CHANNEL}」がありません`);
  const iMo = headers.indexOf(PVA_CFG.COL_MONTH);
  const iCr = headers.indexOf(PVA_CFG.COL_CREATED);

  // 最新月の「流入経路」を収集
  const list = [];
  for (const r of rows){
    const ym = (iMo>=0 ? _parseYYYYM(r[iMo]) : null) || (iCr>=0 ? _ymFromDate(r[iCr]) : null);
    if (!ym || ym.y!==target.y || ym.m!==target.m) continue;
    const s = String(r[iCh]||'').trim();
    if (s) list.push(s);
  }

  // 分類（複数ヒット→未分類）
  const aff=[]; const asp=[]; const point=[]; const others=[]; const fpExtra=new Set();
  for (const s of list){
    if (s === '経路不明') continue;
    const hits=[];
    if (PAT.FP.test(s)) { hits.push('FP'); fpExtra.add(s); }
    if (PAT.AFF.test(s)) hits.push('アフィリエイター');
    if (PAT.ASP.test(s)) hits.push('ASP');
    if (PAT.POINT.test(s)) hits.push('ポイントサイト');

    let bucket='未分類項目';
    if (hits.length===1) bucket=hits[0];
    switch (bucket){
      case 'FP': break;
      case 'アフィリエイター': _pushUnique(aff, s); break;
      case 'ASP':               _pushUnique(asp, s); break;
      case 'ポイントサイト':     _pushUnique(point, s); break;
      default:                  _pushUnique(others, s); break;
    }
  }

    // 出力シート作成（PivotA）
    const name = PVA_CFG.OUT_SHEET_NAME;           // 'PivotA_auto' を想定
    const old  = ss.getSheetByName(name); if (old) ss.deleteSheet(old);
    const out  = ss.insertSheet(name);
    out.setTabColor('#1a73e8');                    // ★ 青マーキング



  // FP見出し（6行目）
  const fpHeaders = Array.from(new Set([...PVA_CFG.FP_BASE, ...fpExtra]));
  const headersRow = [...fpHeaders, PVA_CFG.FP_TAIL_LABEL];
  out.getRange(6,1,1,headersRow.length).setValues([headersRow]).setWrap(false);
  const pCol = headersRow.length; // 列「FP」

  // 4行目のグレー帯
  out.getRange(4, 1, 1, headersRow.length).setBackground(PVA_CFG.COLOR_GRAY);

  // バケットのラベルと値置き場
  if (aff.length)   out.getRange(PVA_CFG.ROW_AFF,   1, 1, aff.length).setValues([aff]);
  if (asp.length)   out.getRange(PVA_CFG.ROW_ASP,   1, 1, asp.length).setValues([asp]);
  if (point.length) out.getRange(PVA_CFG.ROW_POINT, 1, 1, point.length).setValues([point]);
  out.getRange(PVA_CFG.ROW_AFF,   pCol).setValue('アフィリエイター');
  out.getRange(PVA_CFG.ROW_ASP,   pCol).setValue('ASP');
  out.getRange(PVA_CFG.ROW_POINT, pCol).setValue('ポイントサイト');

  // 未分類の有無で出し分け
  if (others.length){
    out.getRange(PVA_CFG.ROW_OTHERS, 1, 1, others.length).setValues([others]); // A18〜
    out.getRange(PVA_CFG.ROW_OTHERS, pCol).setValue('未分類項目');             // FP18
    out.getRange(PVA_CFG.ROW_UNKNOWN2, 1).setValue('経路不明');                // A21
    out.getRange(PVA_CFG.ROW_UNKNOWN2, pCol).setValue('経路不明');             // FP21
    out.getRange(PVA_CFG.ROW_TOTAL, pCol).setValue('総計');                    // FP24
    out.getRange(PVA_CFG.ROW_GRAY_HAS_OTHERS, 1, 1, headersRow.length).setBackground(PVA_CFG.COLOR_GRAY); // 27
  } else {
    out.getRange(PVA_CFG.ROW_OTHERS, 1).setValue('経路不明');                  // A18
    out.getRange(PVA_CFG.ROW_OTHERS, pCol).setValue('経路不明');               // FP18
    out.getRange(21, pCol).setValue('総計');                                   // FP21
    out.getRange(24, 1, 1, headersRow.length).setBackground(PVA_CFG.COLOR_GRAY); // 24
  }

  // 折返しオフ→幅フィット
  const usedRows = Math.max(24, out.getLastRow());
  out.getRange(1, 1, usedRows, headersRow.length).setWrap(false);
  _autoFitColumns(out, headersRow.length);
  out.setFrozenRows(0);

  SpreadsheetApp.getActive().toast(`書式作成：${target.y}年${target.m}月`);
}

//////////////////// クロス作成 ////////////////////
function _buildCrossForTarget(target){
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(PVA_CFG.SRC_SHEET_NAME);
  if (!src) _err(`元データが見つかりません: ${PVA_CFG.SRC_SHEET_NAME}`);

  const { headers, rows } = _readAll(src);
  const iCh  = headers.indexOf(PVA_CFG.COL_CHANNEL);
  const iMo  = headers.indexOf(PVA_CFG.COL_MONTH);
  const iCr  = headers.indexOf(PVA_CFG.COL_CREATED);
  const iUid = headers.indexOf(PVA_CFG.COL_UID);
  if (iCh < 0) _err(`ヘッダーに「${PVA_CFG.COL_CHANNEL}」がありません`);

  const recs = rows
    .filter(r => {
      const ym = (iMo>=0 ? _parseYYYYM(r[iMo]) : null) || (iCr>=0 ? _ymFromDate(r[iCr]) : null);
      return ym && ym.y===target.y && ym.m===target.m;
    })
    .map(r => {
      const ch = String(r[iCh]||'').trim();
      const ts = (iCr>=0 && r[iCr] instanceof Date) ? r[iCr].getTime()
               : (iCr>=0 && r[iCr]) ? new Date(r[iCr]).getTime() : 0;
      const uid= iUid>=0 ? String(r[iUid]||'').trim() : '';
      return { uid, ch, ts };
    });

  const pairs = [];
  if (iUid >= 0){
    const byId = new Map();
    const sorted = recs.filter(x=>x.uid).sort((a,b)=>a.ts-b.ts);
    for (const r of sorted){
      if (!byId.has(r.uid)) byId.set(r.uid, { src:null, fp:null });
      const o = byId.get(r.uid);
      if (PAT.FP.test(r.ch)) o.fp = _normalizeFP(r.ch) || PVA_CFG.FP_TAIL_LABEL;
      else                   o.src = r.ch || '経路不明';
    }
    for (const {src,fp} of byId.values()){
      pairs.push({ ch: src || '経路不明', fp: fp || PVA_CFG.FP_TAIL_LABEL });
    }
  } else {
    for (const r of recs){
      const ch = PAT.FP.test(r.ch) ? '経路不明' : (r.ch || '経路不明');
      const fp = PAT.FP.test(r.ch) ? _normalizeFP(r.ch) || PVA_CFG.FP_TAIL_LABEL : PVA_CFG.FP_TAIL_LABEL;
      pairs.push({ ch, fp });
    }
  }

  // FP列（ヘッダ）
  const fpSet = new Set(PVA_CFG.FP_BASE);
  pairs.forEach(p=>{ if (p.fp && p.fp!==PVA_CFG.FP_TAIL_LABEL) fpSet.add(p.fp); });
  const fpHeaders = [...PVA_CFG.FP_BASE.filter(x=>fpSet.has(x))];
  for (const x of fpSet) if (!fpHeaders.includes(x)) fpHeaders.push(x);
  if (!fpHeaders.includes(PVA_CFG.FP_TAIL_LABEL)) fpHeaders.push(PVA_CFG.FP_TAIL_LABEL);

  // チャンネル
  const chSet = new Set(['経路不明']);
  pairs.forEach(p=>chSet.add(p.ch));
  const channels = [...chSet].sort((a,b)=>a.localeCompare(b,'ja'));

  // 集計
  const cross = new Map();
  const bump = (ch, fp) => {
    if (!cross.has(ch)) cross.set(ch, new Map());
    const m = cross.get(ch);
    m.set(fp, (m.get(fp)||0) + 1);
  };
  pairs.forEach(p=> bump(p.ch, p.fp));

  // 出力
  const old = ss.getSheetByName(PVA_CFG.CROSS_SHEET_NAME); if (old) ss.deleteSheet(old);
  const pv  = ss.insertSheet(PVA_CFG.CROSS_SHEET_NAME);

  pv.getRange(1,2,1,fpHeaders.length).setValues([fpHeaders]);
  pv.getRange(2,1,channels.length,1).setValues(channels.map(x=>[x]));
  const body = channels.map(ch=>{
    const m = cross.get(ch) || new Map();
    return fpHeaders.map(fp => m.get(fp)||0);
  });
  if (body.length) pv.getRange(2,2,body.length,body[0].length).setValues(body);

  _autoFitColumns(pv, fpHeaders.length+1);
  pv.setFrozenRows(1); pv.setFrozenColumns(1);

  SpreadsheetApp.getActive().toast(`クロス作成：${PVA_CFG.CROSS_SHEET_NAME}`);
}

//////////////////// 転記（ご指定ルール） ////////////////////
function _transferFromCrossForTarget(_target){
  const ss  = SpreadsheetApp.getActive();
  const out = ss.getSheetByName(PVA_CFG.OUT_SHEET_NAME);
  const pv  = ss.getSheetByName(PVA_CFG.CROSS_SHEET_NAME);
  if (!out || !pv) _err('出力シートまたはクロスシートが見つかりません');

  const pvLastCol = pv.getLastColumn(), pvLastRow = pv.getLastRow();
  if (pvLastCol < 2 || pvLastRow < 2) _err('クロス表にデータがありません');

  // pivot 読み込み
  const pvFPs   = pv.getRange(1, 2, 1, pvLastCol-1).getDisplayValues()[0].map(s=>String(s||'').trim()); // B1〜
  const pvChans = pv.getRange(2, 1, pvLastRow-1, 1).getDisplayValues().map(r=>String(r[0]||'').trim());
  const pvBody  = pv.getRange(2, 2, pvLastRow-1, pvLastCol-1).getValues();

  const idxFP = pvFPs.indexOf(PVA_CFG.FP_TAIL_LABEL);
  if (idxFP < 0) _err('クロス表に「FP」列が見つかりません');

  // 経路不明の行
  const rowUnknown = pvChans.indexOf('経路不明');

  // 出力側ヘッダ
  const outLastCol = out.getLastColumn();
  const headerRow  = out.getRange(6,1,1,outLastCol).getDisplayValues()[0].map(s=>String(s||'').trim());
  const pCol = headerRow.lastIndexOf(PVA_CFG.FP_TAIL_LABEL) + 1;
  const fpHeaders = headerRow.slice(0, pCol-1);

  // === 7行目：FP_* ベクタ（A〜FP左） & 合計（FP列） ===
  const colIndexByFP = new Map(fpHeaders.map((fp, i)=>[fp, i]));
  const pvIndexByFP  = new Map(pvFPs.map((fp, i)=>[fp, i]));
  const row7 = new Array(fpHeaders.length).fill(0);
  if (rowUnknown >= 0){
    for (let j=0;j<fpHeaders.length;j++){
      const fpName = fpHeaders[j];
      const pvIdx  = pvIndexByFP.get(fpName);
      row7[j] = Number((pvIdx!=null ? pvBody[rowUnknown][pvIdx] : 0) || 0);
    }
  }
  out.getRange(7, 1, 1, fpHeaders.length).clearContent();
  if (row7.length) out.getRange(7, 1, 1, row7.length).setValues([row7]);
  const fpStarSum = row7.reduce((a,b)=>a+b, 0);
  out.getRange(7, pCol).setValue(fpStarSum);

  // === 経路不明（A/FP の“ラベル直下”） ===
  const unknownFPVal = (rowUnknown >= 0) ? Number(pvBody[rowUnknown][idxFP]||0) : 0;
  const hasOthers = String(out.getRange(PVA_CFG.ROW_OTHERS, pCol).getDisplayValue()||'') === '未分類項目';
  const rowUnknownVal = (hasOthers ? PVA_CFG.ROW_UNKNOWN2 : PVA_CFG.ROW_OTHERS) + 1; // 22 or 19
  out.getRange(rowUnknownVal, 1).setValue(unknownFPVal);    // A
  out.getRange(rowUnknownVal, pCol).setValue(unknownFPVal); // FP

  // === 各バケットの“ラベル直下”に FP 列値を反映（アフィ/ASP/ポイント/（未分類）） ===
  const channelTotals = new Map(); // ch -> FP列値
  pvChans.forEach((ch, i)=> channelTotals.set(ch, Number(pvBody[i][idxFP]||0)));

  const rowMap = new Map([[PVA_CFG.ROW_AFF,10],[PVA_CFG.ROW_ASP,13],[PVA_CFG.ROW_POINT,16]]);
  if (hasOthers) rowMap.set(PVA_CFG.ROW_OTHERS, 19); // 未分類値行
  for (const [rLabel, rVal] of rowMap){
    const labels = _readRowLabels(out, rLabel, pCol-1);
    const vals   = labels.map(ch => Number(channelTotals.get(ch)||0));
    out.getRange(rVal, 1, 1, pCol-1).clearContent();
    if (vals.length) out.getRange(rVal, 1, 1, vals.length).setValues([vals]);
  }

  // === 各行の合計を FP 列へ ===
  _sumRowToFP(out, 10, pCol);  // アフィリエイター
  _sumRowToFP(out, 13, pCol);  // ASP
  _sumRowToFP(out, 16, pCol);  // ポイントサイト
  if (hasOthers) _sumRowToFP(out, 19, pCol); // 未分類

  // === 総計を「総計」ラベル直下（FP列）へ ===
  // 総計 = FP_*合計（7行FP） + アフィ合計 + ASP合計 + ポイント合計 +（未分類があれば）未分類合計 + 経路不明合計
  const sumAff    = Number(out.getRange(10, pCol).getValue()||0);
  const sumASP    = Number(out.getRange(13, pCol).getValue()||0);
  const sumPoint  = Number(out.getRange(16, pCol).getValue()||0);
  const sumOthers = hasOthers ? Number(out.getRange(19, pCol).getValue()||0) : 0;
  const sumUnknown= unknownFPVal;
  const grand     = fpStarSum + sumAff + sumASP + sumPoint + sumOthers + sumUnknown;

  // 「総計」ラベル行の検出（未分類なし=21 / あり=24）
  let totalLabelRow = 21;
  const lblAt21 = String(out.getRange(21, pCol).getDisplayValue()||'');
  if (lblAt21 !== '総計') totalLabelRow = 24;
  out.getRange(totalLabelRow+1, pCol).setValue(grand);

  SpreadsheetApp.getActive().toast('転記完了（指定ルール）');
}

//////////////////// ヘルパ ////////////////////
function _readAll(sh){
  const v = sh.getDataRange().getValues();
  const headers = (v[0]||[]).map(x=>String(x||'').trim());
  const rows = v.slice(1);
  return { headers, rows };
}
function _parseYYYYM(s){
  const t = String(s||'').trim();
  const m = t.match(/^(\d{4})年\s*(\d{1,2})月$/);
  return m ? { y:+m[1], m:+m[2] } : null;
}
function _ymFromDate(v){
  if (!v && v!==0) return null;
  let d = v;
  if (Object.prototype.toString.call(v) !== '[object Date]') d = new Date(v);
  if (isNaN(d)) return null;
  return { y: d.getFullYear(), m: d.getMonth()+1 };
}
function _findLatestMonth(){
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(PVA_CFG.SRC_SHEET_NAME);
  if (!src) return null;
  const { headers, rows } = _readAll(src);
  const iMo = headers.indexOf(PVA_CFG.COL_MONTH);
  const iCr = headers.indexOf(PVA_CFG.COL_CREATED);
  let best = null;
  for (const r of rows){
    const ym = (iMo>=0 ? _parseYYYYM(r[iMo]) : null) || (iCr>=0 ? _ymFromDate(r[iCr]) : null);
    if (!ym) continue;
    const v = ym.y*100 + ym.m;
    if (best===null || v>best) best=v;
  }
  return best ? { y:Math.floor(best/100), m:best%100 } : null;
}
function _pushUnique(arr, s){ if (!arr.includes(s)) arr.push(s); }
function _err(msg){ SpreadsheetApp.getActive().toast(String(msg)); throw new Error(msg); }

// FP名を正規化
function _normalizeFP(v){
  const s = String(v||'').trim();
  const m = s.match(/^([ＦF][ＰP])[\s_＿:：\-]*(.+)$/i);
  if (!m) return null;
  let name = m[2].trim().replace(/[？?（）\(\)\[\]【】・･\s]/g,'');
  if (!name) return null;
  return 'FP_' + name;
}

// 行の A〜（空まで）のラベル名を取得
function _readRowLabels(sheet, row, maxCol){
  const arr = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0].map(s=>String(s||'').trim());
  const out=[]; for (let i=0;i<arr.length;i++){ if (!arr[i]) break; out.push(arr[i]); }
  return out;
}

// 行の A〜（FP手前）合計を FP 列へ
function _sumRowToFP(sheet, row, pCol){
  const nums = sheet.getRange(row, 1, 1, pCol-1).getValues()[0].map(n=>Number(n||0));
  const sum  = nums.reduce((a,b)=>a+b, 0);
  sheet.getRange(row, pCol).setValue(sum);
}

// 文字幅から列幅を直接設定（全角×2換算）
function _autoFitColumns(sheet, colCount) {
  SpreadsheetApp.flush();
  const lastRow = Math.max(sheet.getLastRow(), 24);
  const values  = sheet.getRange(1, 1, lastRow, colCount).getDisplayValues();
  const displayLen = s => String(s || '').replace(/[^\x00-\x7F]/g, 'xx').length;
  for (let c = 0; c < colCount; c++) {
    let maxLen = 0;
    for (let r = 0; r < values.length; r++) {
      const len = displayLen(values[r][c]);
      if (len > maxLen) maxLen = len;
    }
    const px = Math.min(640, Math.max(64, Math.round(maxLen * 7 + 24)));
    sheet.setColumnWidth(c + 1, px);
  }
}
