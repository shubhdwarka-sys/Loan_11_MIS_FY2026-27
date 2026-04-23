/**
 * ============================================================
 *  TCO_UserControl.gs — v4.0 | FY 2026-27
 *  Admin Menu | Backup | Restore | Reset | Protections
 * ============================================================
 *  ADMIN   : admin.loan11@gmail.com
 *  SHEET   : 16j_0Szvo3WAQrEMgTbJx8oqtkiWbBd3IvaoZ2hdftgI
 * ============================================================
 */




// ── CONFIG ───────────────────────────────────────────────────
var UC = {
  ADMIN_EMAIL : 'admin.loan11@gmail.com',
  SHEET_ID    : '16j_0Szvo3WAQrEMgTbJx8oqtkiWbBd3IvaoZ2hdftgI',
  DATA_ROW    : 4,

  VISIBLE : [
    'DM_DISBURSEMENT MEMO',
    'ACCOUNT_PAYMENT_TRACKER',
    'RTO_TRACKER'
  ],

  HIDDEN : [
    'MASTER_DATA',
    'DEALER_MASTER',
    'TCO_EMPLOYEE_MASTER'
  ],

  BACKUP : [
    '_BACKUP_DM',
    '_BACKUP_ACCOUNT',
    '_BACKUP_RTO'
  ],

  // AUTO + GEN cols per sheet (1-based) — matches PDF v3 mapping
  AUTO_COLS : {
    'DM_DISBURSEMENT MEMO'    : [3, 4, 38, 40, 41, 42, 43, 44, 46],
    'ACCOUNT_PAYMENT_TRACKER' : [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 15, 16, 37, 38, 41, 42],
    'RTO_TRACKER'             : [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 15, 17, 18, 19, 26]
  }
};


// ─────────────────────────────────────────────────────────────
// ── ADMIN MENU (runs on open) ─────────────────────────────────
// ─────────────────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ TCO Admin')
    .addItem('📦 Manual Backup',           'manualBackup')
    .addItem('🔁 Restore from Backup',     'restoreFromBackup')
    .addSeparator()
    .addItem('🔒 Re-Apply Protections',    'reApplyProtections')
    .addItem('🎨 Apply Borders',           'applyBordersAll')
    .addItem('🖌️ Restore BG Colors',      'applyBackgroundToAllRows')
    .addItem('📋 Update Master Data',      'updateMasterDataManual')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('🗑️ Reset Data')
        .addItem('Reset ALL Sheets',       'resetAllData')
        .addItem('Reset DM Only',          'resetDmOnly')
        .addItem('Reset Account Only',     'resetAccOnly')
        .addItem('Reset RTO Only',         'resetRtoOnly')
    )
    .addToUi();
}


// ─────────────────────────────────────────────────────────────
// ── BACKUP ───────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function manualBackup() {
  var ss  = SpreadsheetApp.openById(UC.SHEET_ID);
  var ui  = SpreadsheetApp.getUi();
  var now = new Date().toLocaleString('en-IN');

  var pairs = [
    ['DM_DISBURSEMENT MEMO',    '_BACKUP_DM'     ],
    ['ACCOUNT_PAYMENT_TRACKER', '_BACKUP_ACCOUNT'],
    ['RTO_TRACKER',             '_BACKUP_RTO'    ]
  ];

  var log = [];

  pairs.forEach(function(p) {
    var src = ss.getSheetByName(p[0]);
    var bk  = ss.getSheetByName(p[1]);

    if (!src) { log.push('❌ Not found: ' + p[0]); return; }

    // Create backup sheet if missing
    
    if (!bk) {
  bk = ss.insertSheet(p[1]);
  bk.hideSheet();
} else {
  // Remove protection if any
  bk.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(function(prot) { prot.remove(); });
}

    bk.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .forEach(function(prot) { prot.remove(); });
    try { bk.showSheet(); } catch(e) {}
    bk.clearContents();
    var d = src.getDataRange().getValues();
    if (d.length > 0) bk.getRange(1, 1, d.length, d[0].length).setValues(d);
    log.push('✅ ' + p[0] + ' → ' + p[1]);
  });

  ui.alert(
    '📦  BACKUP COMPLETE\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    log.join('\n') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    now
  );
}


// ─────────────────────────────────────────────────────────────
// ── RESTORE ──────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function restoreFromBackup() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '⚠️  RESTORE FROM BACKUP',
    'DM, Account, RTO sheets will be overwritten with backup data.\n\nThis cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss   = SpreadsheetApp.openById(UC.SHEET_ID);
  var pairs = [
    ['_BACKUP_DM',      'DM_DISBURSEMENT MEMO'   ],
    ['_BACKUP_ACCOUNT', 'ACCOUNT_PAYMENT_TRACKER'],
    ['_BACKUP_RTO',     'RTO_TRACKER'            ]
  ];

  var log = [];

  pairs.forEach(function(p) {
    var src = ss.getSheetByName(p[0]);
    var tgt = ss.getSheetByName(p[1]);

    if (!src) { log.push('❌ Backup not found: ' + p[0]); return; }
    if (!tgt) { log.push('❌ Sheet not found: '  + p[1]); return; }

    var d = src.getDataRange().getValues();
    if (d.length === 0) { log.push('⚠️ Empty backup: ' + p[0]); return; }

    tgt.clearContents();
    tgt.getRange(1, 1, d.length, d[0].length).setValues(d);
    log.push('✅ Restored: ' + p[1]);
  });

  ui.alert(
    '🔁  RESTORE COMPLETE\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    log.join('\n') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    new Date().toLocaleString('en-IN')
  );
}


// ─────────────────────────────────────────────────────────────
// ── RE-APPLY PROTECTIONS ─────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function reApplyProtections() {
  var ss  = SpreadsheetApp.openById(UC.SHEET_ID);
  var ui  = SpreadsheetApp.getUi();
  var total = 0;

  // ── Step 1: Remove + re-lock AUTO cols on visible sheets ──
  UC.VISIBLE.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;

    // Remove existing range protections
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
      p.remove();
    });

    var autoCols = UC.AUTO_COLS[name] || [];
    var maxRow   = sh.getMaxRows();

    autoCols.forEach(function(col) {
      var p = sh.getRange(1, col, maxRow, 1).protect();
      p.setDescription('🔒 AUTO | ' + name + ' | Col-' + col);
      p.setWarningOnly(false);
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
      p.addEditor(UC.ADMIN_EMAIL);
      total++;
    });
  });

  // ── Step 2: Re-lock + hide master/ref sheets ──
  UC.HIDDEN.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;

    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) {
      p.remove();
    });

    sh.hideSheet();
    var p = sh.protect();
    p.setDescription('🔒 SYSTEM | ' + name);
    p.setWarningOnly(false);
    p.removeEditors(p.getEditors());
    if (p.canDomainEdit()) p.setDomainEdit(false);
    p.addEditor(UC.ADMIN_EMAIL);
  });

  ui.alert(
    '✅  PROTECTIONS APPLIED\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '🔒 AUTO Cols Locked  : ' + total + '\n' +
    '🔒 System Sheets     : ' + UC.HIDDEN.length + ' (Hidden)\n' +
    '👤 Admin             : ' + UC.ADMIN_EMAIL + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    new Date().toLocaleString('en-IN')
  );
}


// ─────────────────────────────────────────────────────────────
// ── BORDERS ──────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function applyBordersAll() {
  var ss = SpreadsheetApp.openById(UC.SHEET_ID);
  UC.VISIBLE.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var lr = Math.max(sh.getLastRow(), UC.DATA_ROW + 5);
    var lc = sh.getLastColumn();
    if (lc < 1) return;
    sh.getRange(UC.DATA_ROW, 1, lr - UC.DATA_ROW + 1, lc)
      .setBorder(
        true, true, true, true, true, true,
        '#CBD5E1',
        SpreadsheetApp.BorderStyle.SOLID
      );
  });
  SpreadsheetApp.getUi().alert('✅  Borders applied to DM, Account, RTO!');
}


// ─────────────────────────────────────────────────────────────
// ── RESTORE BG COLORS ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function applyBackgroundToAllRows() {
  var ss = SpreadsheetApp.openById(UC.SHEET_ID);
  var sheetDefs = [
    { name: 'DM_DISBURSEMENT MEMO',    key: 'DM'  },
    { name: 'ACCOUNT_PAYMENT_TRACKER', key: 'ACC' },
    { name: 'RTO_TRACKER',             key: 'RTO' }
  ];

  var AUTO_BG = '#EBF2FF';   // Light blue — AUTO
  var GEN_BG  = '#FEFCE8';   // Light yellow — GEN

  sheetDefs.forEach(function(sd) {
    var sheet   = ss.getSheetByName(sd.name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < UC.DATA_ROW) return;
    var autoCols = UC.AUTO_COLS[sd.name] || [];
    for (var row = UC.DATA_ROW; row <= lastRow; row++) {
      autoCols.forEach(function(c) {
        sheet.getRange(row, c).setBackground(AUTO_BG);
      });
    }
  });

  SpreadsheetApp.getUi().alert('✅  Background colors restored!');
}


// ─────────────────────────────────────────────────────────────
// ── UPDATE MASTER (Manual) ────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function updateMasterDataManual() {
  // updateMasterData_ is defined in TCO_Operations.gs
  updateMasterData_(SpreadsheetApp.openById(UC.SHEET_ID));
  SpreadsheetApp.getUi().alert('✅  Master Data updated!\n' + new Date().toLocaleString('en-IN'));
}


// ─────────────────────────────────────────────────────────────
// ── RESET FUNCTIONS ───────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function resetAllData() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '⚠️  RESET ALL DATA',
    'DM, Account, RTO — all data from Row 4 will be permanently deleted.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  clearSheetDataOnly_('DM_DISBURSEMENT MEMO');
  clearSheetDataOnly_('ACCOUNT_PAYMENT_TRACKER');
  clearSheetDataOnly_('RTO_TRACKER');
  clearMasterData_();

  SpreadsheetApp.getUi().alert('✅  All data reset!\n' + new Date().toLocaleString('en-IN'));
}

function resetDmOnly() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ Reset DM Data?', 'Row 4+ will be deleted.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  clearSheetDataOnly_('DM_DISBURSEMENT MEMO');
  ui.alert('✅  DM data reset!');
}

function resetAccOnly() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ Reset Account Data?', 'Row 4+ will be deleted.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  clearSheetDataOnly_('ACCOUNT_PAYMENT_TRACKER');
  ui.alert('✅  Account data reset!');
}

function resetRtoOnly() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ Reset RTO Data?', 'Row 4+ will be deleted.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  clearSheetDataOnly_('RTO_TRACKER');
  ui.alert('✅  RTO data reset!');
}

// ── HELPER: Clear single sheet data only (MASTER nahi) ──
function clearSheetDataOnly_(name) {
  var ss = SpreadsheetApp.openById(UC.SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) return;
  var lr = sh.getLastRow();
  if (lr < UC.DATA_ROW) return;
  sh.getRange(UC.DATA_ROW, 1, lr - UC.DATA_ROW + 1, sh.getLastColumn())
    .clearContent()
    .setBackground(null);
}

// ── HELPER: Clear MASTER sheet only ──
function clearMasterData_() {
  var ss    = SpreadsheetApp.openById(UC.SHEET_ID);
  var ms    = ss.getSheetByName('MASTER_DATA');
  if (!ms) return;

  // Protection temporarily remove
  var prots = ms.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  prots.forEach(function(p) { p.remove(); });

  // Clear data rows
  var mlr = ms.getLastRow();
  if (mlr >= UC.DATA_ROW) {
    ms.getRange(UC.DATA_ROW, 1, mlr - UC.DATA_ROW + 1, ms.getLastColumn() || 97)
      .clearContent();
  }

  // Re-protect
  var mp = ms.protect();
  mp.setWarningOnly(false);
  mp.removeEditors(mp.getEditors());
  if (mp.canDomainEdit()) mp.setDomainEdit(false);
  mp.addEditor(UC.ADMIN_EMAIL);
}

// यह स्क्रिप्ट Row 3 (Labels) से लेकर एंड तक फ़िल्टर सेट कर देगी
function setupMasterFilters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var visibleSheets = [
    'DM_DISBURSEMENT MEMO',
    'ACCOUNT_PAYMENT_TRACKER',
    'RTO_TRACKER'
  ];

  visibleSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      // अगर पहले से कोई फ़िल्टर है तो उसे हटा दें
      if (sheet.getFilter() !== null) {
        sheet.getFilter().remove();
      }
      
      // Row 3 से लेकर शीट के एंड तक नया फ़िल्टर लगा दें
      var maxRows = sheet.getMaxRows();
      var maxCols = sheet.getMaxColumns();
      // getRange(row, col, numRows, numColumns)
      sheet.getRange(3, 1, maxRows - 2, maxCols).createFilter();
    }
  });
  
  SpreadsheetApp.getUi().alert('✅ सभी वर्किंग शीट्स की Row 3 पर फ़िल्टर सेट हो गया है!');
}

// ─────────────────────────────────────────────────────────────
// ── BULK REFRESH ALL COLORS (For Copy-Paste & Old Data) ──────
// ─────────────────────────────────────────────────────────────
function refreshAllColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TCO.SHEET.ACC);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < TCO.DATA_ROW) return;

  // यह लूप हर रो पर जाकर कलर का रूल चेक करेगा और लगा देगा
  for (var row = TCO.DATA_ROW; row <= lastRow; row++) {
    applyPaymentColors_(sheet, row);
  }
  
  SpreadsheetApp.getUi().alert('✅ Account शीट के सभी कलर्स अपडेट हो गए हैं!');
}

