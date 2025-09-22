/************ すべてのメニューをここで作る ************/
function onOpen(e) {
  if (typeof buildMenu_CsvImport_     === 'function') buildMenu_CsvImport_();     // CSV import（既存想定）
  if (typeof buildMenu_CsvUpload_     === 'function') buildMenu_CsvUpload_();     // Bubble Upload（既存想定）
  if (typeof buildMenu_PivotA_        === 'function') buildMenu_PivotA_();        // Pivot_user（既存想定）
  if (typeof buildMenu_PivotB_        === 'function') buildMenu_PivotB_();        // Pivot_面談報告（既存想定）
  if (typeof buildMenu_TimeTrigger_   === 'function') buildMenu_TimeTrigger_();   // ★今回追加：TimeTrigger
  if (typeof buildMenu_SheetTabTools_ === 'function') buildMenu_SheetTabTools_();
  if (typeof buildMenu_SheetDelete_   === 'function') buildMenu_SheetDelete_(); // ← クリア版の代わり
}
function onInstall(e){ onOpen(e); }
