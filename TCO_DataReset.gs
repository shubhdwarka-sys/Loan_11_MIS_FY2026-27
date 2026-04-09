// ============================================================
// TCO_DataReset.gs
// Data Reset — Headers safe, sirf data clear hoga
// Run: resetAllData() — sab sheets
// Run: resetDmOnly() / resetAccOnly() / resetRtoOnly()
// ============================================================

function resetAllData() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '⚠️ CONFIRM FULL RESET',
    'DM + ACCOUNT + RTO + MASTER DATA — sab clear ho jayega!\nHeaders safe rahenge.\n\nKya aap sure hain?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) { ui.alert('❌ Reset cancelled.'); return; }

  // Double confirm
  var resp2 = ui.alert(
    '🔴 LAST WARNING',
    'Yeh action UNDO nahi hoga!\nSabse pehle Backup lena recommended hai.\n\nFir bhi proceed karna hai?',
    ui.ButtonSet.YES_NO
  );
  if (resp2 !== ui.Button.YES) { ui.alert('❌ Reset cancelled.'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var count = 0;

  count += clearSheetData_(ss, 'DM_DISBURSEMENT MEMO',    4, 57);
  count += clearSheetData_(ss, 'ACCOUNT_PAYMENT_TRACKER', 4, 45);
  count += clearSheetData_(ss, 'RTO_TRACKER',             4, 30);
  count += clearSheetData_(ss, 'MASTER_DATA',             3, 99);
  count += clearSheetData_(ss, '_BACKUP_DM',              2, 57);
  count += clearSheetData_(ss, '_BACKUP_ACCOUNT',         2, 45);
  count += clearSheetData_(ss, '_BACKUP_RTO',             2, 30);
  count += clearSheetData_(ss, 'AUTH_AUDIT_LOG',          3,  6);

  ui.alert('✅ Reset Complete!\n\n' + count + ' sheets cleared.\nHeaders safe hain.');
}


// ── INDIVIDUAL SHEET RESETS ───────────────────────────────────
function resetDmOnly() {
  if (!confirmReset_('DM_DISBURSEMENT MEMO')) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetData_(ss, 'DM_DISBURSEMENT MEMO', 4, 57);
  clearSheetData_(ss, '_BACKUP_DM', 2, 57);
  SpreadsheetApp.getUi().alert('✅ DM Sheet reset done!');
}

function resetAccOnly() {
  if (!confirmReset_('ACCOUNT_PAYMENT_TRACKER')) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetData_(ss, 'ACCOUNT_PAYMENT_TRACKER', 4, 45);
  clearSheetData_(ss, '_BACKUP_ACCOUNT', 2, 45);
  SpreadsheetApp.getUi().alert('✅ Account Sheet reset done!');
}

function resetRtoOnly() {
  if (!confirmReset_('RTO_TRACKER')) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetData_(ss, 'RTO_TRACKER', 4, 30);
  clearSheetData_(ss, '_BACKUP_RTO', 2, 30);
  SpreadsheetApp.getUi().alert('✅ RTO Sheet reset done!');
}

function resetMasterOnly() {
  if (!confirmReset_('MASTER_DATA')) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetData_(ss, 'MASTER_DATA', 3, 99);
  SpreadsheetApp.getUi().alert('✅ Master Data reset done!');
}


// ── HELPER: CLEAR DATA (headers safe) ────────────────────────
function clearSheetData_(ss, sheetName, dataStartRow, numCols) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return 0; // No data to clear

  var numRows = lastRow - dataStartRow + 1;
  sheet.getRange(dataStartRow, 1, numRows, numCols)
    .clearContent()
    .clearFormat();

  // Restore base formatting after clear
  sheet.getRange(dataStartRow, 1, numRows, numCols)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setFontColor('#000000')
    .setBorder(true, true, true, true, true, true,
      '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);

  Logger.log(sheetName + ' cleared from row ' + dataStartRow);
  return 1;
}

// ── HELPER: CONFIRM ───────────────────────────────────────────
function confirmReset_(sheetName) {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '⚠️ Reset: ' + sheetName,
    sheetName + ' ka data clear ho jayega.\nHeaders safe rahenge.\n\nConfirm?',
    ui.ButtonSet.YES_NO
  );
  return resp === ui.Button.YES;
}


// ── ADD TO MENU (onOpen mein call karo) ──────────────────────
// TCO_Operations_System.gs ke onOpen() mein yeh add karo:
//
// .addSeparator()
// .addSubMenu(SpreadsheetApp.getUi().createMenu('🗑️ Data Reset')
//   .addItem('Reset ALL Sheets',  'resetAllData')
//   .addItem('Reset DM Only',     'resetDmOnly')
//   .addItem('Reset Account Only','resetAccOnly')
//   .addItem('Reset RTO Only',    'resetRtoOnly')
//   .addItem('Reset Master Data', 'resetMasterOnly'))