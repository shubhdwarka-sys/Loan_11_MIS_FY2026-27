// ========================================================
// TCO_Operations_System.gs — v4.0 | FY 2026-27
// Fixed: All 14 points
// ============================================================

// ── GLOBAL CONFIG ─────────────────────────────────────────────
var TCO = {
  SHEET_ID    : '1yGs3OfRX9UdRUEMYzKHr1yt0rUaDQ5rBjoxg6aeNBgM',
  DATA_ROW    : 4,
  ADMIN_EMAIL : 'admin.loan11@gmail.com',

  SHEET : {
    DM     : 'DM_DISBURSEMENT MEMO',
    ACC    : 'ACCOUNT_PAYMENT_TRACKER',
    RTO    : 'RTO_TRACKER',
    MASTER : 'MASTER_DATA',
    DEALER : 'DEALER_MASTER',
    EMP    : 'TCO_EMPLOYEE_MASTER',
    AUDIT  : 'AUTH_AUDIT_LOG',
    USER   : 'USER_MASTER',
    BK_DM  : '_BACKUP_DM',
    BK_ACC : '_BACKUP_ACCOUNT',
    BK_RTO : '_BACKUP_RTO'
  },

  STYLE : {
    AUTO_GEN  : { color: '#166534', style: 'italic', size: 10 },
    DM_NO     : { color: '#5D4037', style: 'italic', size: 12 },
    AUTO_PULL : { color: '#1A4D8F', style: 'italic', size: 10 },
    MANUAL    : { color: '#000000', style: 'normal', size: 10 },
    REG_NO    : { color: '#c2410c', style: 'normal', size: 12 },
    NOT_FOUND : { color: '#ef4444', style: 'italic', size: 10 },
    PENDING   : { color: '#ef4444', style: 'normal', size: 12 }, // RTO PENDING DAYS
  },

  // DM columns (1-indexed)
  DM : {
    DM_NO:1, DM_DATE:2, MONTH:3, DM_RECV_DATE:4, NO_OF_DAYS:5,
    BUYER:6, PHONE:7, PHONE2:8, EMAIL:9, REG_NO:10,
    SELLER:11, MODEL:12, MFG_YEAR:13, BANK:14, PRODUCT:15,
    LOAN_APPLIED:16, LOAN_APP1:17, TENURE1:18, INS1:19, SURAKHSA1:20,
    TOTAL_LOAN1:21, SURAKHSA_TEN1:22, FILE_CHG1:23, ROI1:24, EMI1:25, DISB_AMT1:26,
    BANK2:27, LOAN_APP2:28, TENURE2:29, INS2:30, SURAKHSA2:31,
    TOTAL_LOAN2:32, SURAKHSA_TEN2:33, FILE_CHG2:34, ROI2:35, EMI2:36, DISB_AMT2:37,
    CHALLAN_AMT:38,  // e-Challan AMT (If Any)
    DEALER_PAY_ST:39, EMP_ID:40,
    EXEC_NAME:41, BRANCH:42, TEAM_LEADER:43, DEALER_NAME:44,
    AUTH_PERSON:45, CONTACT_DLR:46, LOCATION:47, SELLER_AUTO:48,
    PAYOUT_TCO:49, PAYOUT_CD:50, PAYOUT_SCORE:51, LOAN_SCORE:52,
    RTO_CHARGES:53, RTO_SCORE:54, EXE_INCENTIVE:55, NET_SCORE:56, REMARKS:57
  },

  // ACC columns
  ACC : {
    DM_DATE:1, DM_NO:2, BUYER:3, PHONE:4, PHONE2:5, REG_NO:6,
    MODEL:7, MFG_YEAR:8, BANK:9, PRODUCT:10, DEALER:11, EXEC:12,
    DISB1:13, DISB2:14, DISB_RECV:15, DISB_RECV_DATE:16, UTR:17,
    P1_NAME:18, P1_DATE:19, P1_AMT:20, P1_ACC:21,
    P2_NAME:22, P2_DATE:23, P2_AMT:24, P2_ACC:25,
    P3_NAME:26, P3_DATE:27, P3_AMT:28, P3_ACC:29,
    DEALER_PAY_ST:30, DEALER_PAY_DT:31, PAYOUT_DEALER:32,
    HOLD_BANK:33, HOLD_TCO:34, EXE_INC:35, LOAN_SCORE:36,
    PAYOUT_BANK:37, NET_SCORE:38, RTO_CHARGES:39, RTO_VENDOR:40,
    RTO_PAID:41, RTO_PAY_DATE:42, RC_STATUS:43, RTO_PROFIT:44, REMARKS:45
  },

  // RTO columns
  RTO : {
    DM_DATE:1, DM_NO:2, DEALER_PAY_ST:3, PRODUCT:4, BRANCH:5,
    EXEC:6, DEALER:7, REG_NO:8, FILE_REC_DATE:9, RTO_CODE:10,
    MODEL:11, MFG_YEAR:12, CHASSIS:13, ENGINE:14, SELLER:15,
    SELLER_PHONE:16, BUYER:17, BUYER_PHONE:18, BANK:19, CASE_TYPE:20,
    SCAN_LINK:21, VENDOR:22, RECEIPT_NO:23, RC_STATUS:24,
    RC_DATE:25, PENDING_DAYS:26, ORIG_RC:27, RC_PROOF:28,
    SYS_REMARKS:29, REMARKS:30
  }
};

// DM→ACC field sync map
var DM_TO_ACC = {
  6:3,   // BUYER
  7:4,   // PHONE
  8:5,   // PHONE2
  10:6,  // REG NO
  12:7,  // MODEL
  13:8,  // MFG YEAR
  14:9,  // BANK
  15:10, // PRODUCT
  26:13, // DISB AMT1
  37:14, // DISB AMT2
  53:39  // RTO CHARGES
};

// DM→RTO field sync map
var DM_TO_RTO = {
  10:8,  // REG NO
  12:11, // MODEL
  13:12, // MFG YEAR
  14:19, // BANK
  15:4,  // PRODUCT
  11:15, // SELLER
  7:18   // PHONE → BUYER PHONE (Point 12)
};

// Columns where CAPS should NOT be applied
// EMAIL cols
var EMAIL_COLS  = { DM:[9], ACC:[], RTO:[] };
// URL/Link cols — lowercase only
var LINK_COLS   = { DM:[], ACC:[], RTO:[21,28] };
// Dropdown cols — skip caps
var DROPDOWN_COLS = { DM:[14,15,39], ACC:[30], RTO:[22,24,27] };
// Number/currency cols — skip caps
var NUMBER_COLS = {
  DM  : [7,8,13,16,17,18,19,20,21,22,23,24,25,26,28,29,30,31,32,33,34,35,36,37,38,46,49,50,51,52,53,54,55,56],
  ACC : [4,5,8,15,19,20,23,24,27,28,36,37,38,39,41],
  RTO : [12,16,26]
};

// Auto-gen / auto-pull column maps for styling
var AUTO_GEN_COLS  = { DM:[3,5], ACC:[44], RTO:[10,26] };
var AUTO_PULL_COLS = {
  DM  : [4,41,42,43,44,45,47,48],
  ACC : [1,2,3,4,5,6,7,8,9,10,11,12,13,14,39,40,43],
  RTO : [1,2,3,4,5,6,7,8,11,12,15,17,18,19]
};

var SPECIAL_COLS = {
  DM : {
    1 : {color:'#5D4037', style:'italic', size:12},
    10: {color:'#c2410c', style:'normal', size:12},
    16: {color:'#134e4a', style:'bold',   size:12},
    26: {color:'#134e4a', style:'bold',   size:12},
    37: {color:'#134e4a', style:'bold',   size:12},
    39: {color:'#134e4a', style:'bold',   size:12},
    55: {color:'#134e4a', style:'bold',   size:12},
    56: {color:'#134e4a', style:'bold',   size:12}
  },
  ACC: {
    2 : {color:'#5D4037', style:'italic', size:12},
    6 : {color:'#c2410c', style:'normal', size:12},
    30: {color:'#134e4a', style:'bold',   size:12}
  },
  RTO: {
    2 : {color:'#5D4037', style:'italic', size:12},
    8 : {color:'#c2410c', style:'normal', size:12},
    24: {color:'#134e4a', style:'bold',   size:12},
    26: {color:'#ef4444', style:'normal', size:12}
  }
};
// Currency columns (for ₹ watermark + Indian format)
var CURRENCY_COLS = {
  DM  : [16,17,19,20,21,23,25,26,28,30,31,32,34,36,37,38,49,50,51,52,53,54,55,56],  // Score cols removed
  ACC : [15,20,24,28,32,33,34,35,37,41]
};


// ─────────────────────────────────────────────────────────────
// ── MAIN TRIGGER ─────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;  // ← YEH ADD KARO
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sName = sheet.getName();
    var row   = range.getRow();
    var col   = range.getColumn();
    var value = range.getValue();

    if (row < TCO.DATA_ROW) return;
    if (sName !== TCO.SHEET.DM && sName !== TCO.SHEET.ACC && sName !== TCO.SHEET.RTO) return;

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var key  = sName === TCO.SHEET.DM  ? 'DM'  :
               sName === TCO.SHEET.ACC ? 'ACC' : 'RTO';
    var cell = sheet.getRange(row, col);
    var strVal = (value !== null && value !== undefined) ? String(value).trim() : '';

    // ── 1. PHONE VALIDATION — red bg, no block ──────────────
    var phoneCols = { DM:[7,8,46], ACC:[], RTO:[16] };
    if ((phoneCols[key] || []).indexOf(col) !== -1) {
      var digits = strVal.replace(/[^0-9]/g,'');
      cell.setBackground(strVal !== '' && digits.length !== 10 ? '#fca5a5' : null);
    }

    // ── 2. CURRENCY — red if text, clear if valid number ────
    var currCols = CURRENCY_COLS[key] || [];
    if (currCols.indexOf(col) !== -1) {
      if (strVal === '') {
        cell.setBackground(null);
      } else if (typeof value === 'string' && !/^\d+(\.\d+)?$/.test(strVal)) {
        cell.setBackground('#fca5a5'); // red — text in currency col
      } else {
        cell.setBackground(null);
        cell.setNumberFormat('₹ ##,##,##0');
      }
    }

    // ── 3. WORD ONLY — yellow if has numbers ────────────────
    var wordOnly = { DM:[6,11,27], ACC:[18,22,26], RTO:[] };
    if ((wordOnly[key] || []).indexOf(col) !== -1 && strVal !== '') {
      if (typeof value === 'string') {
  cell.setValue(strVal.toUpperCase());
  cell.setBackground(/[0-9]/.test(strVal) ? '#fef08a' : null);
}
      }
    

    // ── 4. EMAIL — lowercase, yellow if no @ ────────────────
    if (sName === TCO.SHEET.DM && col === TCO.DM.EMAIL && strVal !== '') {
      cell.setValue(strVal.toLowerCase());
      cell.setBackground(strVal.indexOf('@') === -1 ? '#fef08a' : null);
    }

    // ── 5. WORD + NO — caps, clear bg ───────────────────────
    var wordNum = { DM:[10,12,40], ACC:[17,21,25,29], RTO:[13,14,20,23] };
    if ((wordNum[key] || []).indexOf(col) !== -1 && typeof value === 'string' && strVal !== '') {
      cell.setValue(strVal.toUpperCase());
      cell.setBackground(null);
    }

    // ── 6. GENERAL CAPS (all other text cols) ───────────────
    var skipAll = [].concat(
      phoneCols[key] || [], currCols,
      wordOnly[key] || [], wordNum[key] || [],
      DROPDOWN_COLS[key] || [], NUMBER_COLS[key] || [],
      [TCO.DM.EMAIL]
    );
    if (skipAll.indexOf(col) === -1 && typeof value === 'string' && strVal !== '') {
      if (strVal.indexOf('http') !== 0 && strVal.indexOf('www') !== 0) {
        cell.setValue(strVal.toUpperCase());
      }
    }

    // ── 7. COLUMN STYLE ──────────────────────────────────────
    applyColumnStyle_(sheet, sName, row, col);

    // ── 8. ROUTE TO HANDLER — fresh read after caps ──────────
    var freshVal = range.getValue();
    if      (sName === TCO.SHEET.DM)  handleDmEdit_(ss, sheet, row, col, freshVal);
    else if (sName === TCO.SHEET.ACC) handleAccEdit_(ss, sheet, row, col, freshVal);
    else if (sName === TCO.SHEET.RTO) handleRtoEdit_(ss, sheet, row, col, freshVal);

  } catch(err) {
    Logger.log('onEdit ERROR: ' + err.message);
  }
}
// ─────────────────────────────────────────────────────────────
// ── DM SHEET HANDLER ─────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function handleDmEdit_(ss, sheet, row, col, value) {

  // 1. DM DATE entered → try generate
  if (col === TCO.DM.DM_DATE) {
    tryGenerateEntry_(ss, sheet, row);
    return;
  }

  // 2. BUYER NAME → try generate or sync
  if (col === TCO.DM.BUYER) {
    var exist = sheet.getRange(row, TCO.DM.DM_NO).getValue();
    var dmDate = sheet.getRange(row, TCO.DM.DM_DATE).getValue();
    if (dmDate && !String(exist).startsWith('TCO')) {
      tryGenerateEntry_(ss, sheet, row);
    } else if (String(exist).startsWith('TCO')) {
      syncFieldToAccRto_(ss, sheet, row, col, toUpper_(value));
    }
    return;
  }

  // 3. EMP ID → lookup (works on paste + delete)
  if (col === TCO.DM.EMP_ID) {
    var empVal = sheet.getRange(row, TCO.DM.EMP_ID).getValue();
    if (!empVal || String(empVal).trim() === '') {
      clearEmpFields_(sheet, row);
      clearEmpInAccRto_(ss, sheet, row);
    } else {
      fillExecDetails_(ss, sheet, row, String(empVal).trim());
    }
    return;
  }

  // 4. DEALER CONTACT → lookup (works on paste + delete)
  if (col === TCO.DM.CONTACT_DLR) {
    // Re-read value fresh
    var freshContact = sheet.getRange(row, TCO.DM.CONTACT_DLR).getValue();
    value = freshContact;
    if (!value || String(value).trim() === '') {
      clearDlrFields_(sheet, row);        // FIX Point 8: delete → clear
      clearDlrInAccRto_(ss, sheet, row);
    } else {
      fillDealerDetails_(ss, sheet, row, value);
    }
    return;
  }

  // 5. SELLER → sync to SELLER_AUTO + RTO
  if (col === TCO.DM.SELLER) {
    var up = toUpper_(value);
    var saCell = sheet.getRange(row, TCO.DM.SELLER_AUTO);
    saCell.setValue(up);
    styleCell_(saCell, TCO.STYLE.AUTO_PULL);
    syncFieldToAccRto_(ss, sheet, row, col, up);
    return;
  }

  // 6. PHONE → sync to RTO BUYER PHONE (Point 12)
  if (col === TCO.DM.PHONE) {
    validatePhone_(sheet, row, col, value);
    syncFieldToAccRto_(ss, sheet, row, col, value);
    return;
  }

  // 7. PHONE2 validation
  if (col === TCO.DM.PHONE2) {
    validatePhone_(sheet, row, col, value);
    syncFieldToAccRto_(ss, sheet, row, col, value);
    return;
  }

  // 8. CONTACT DLR phone validation
  if (col === TCO.DM.CONTACT_DLR) {
    validatePhone_(sheet, row, col, value);
    return;
  }

  // 9. Fields that sync to ACC/RTO
  var syncCols = [
    TCO.DM.REG_NO, TCO.DM.MODEL, TCO.DM.MFG_YEAR,
    TCO.DM.BANK, TCO.DM.PRODUCT, TCO.DM.DISB_AMT1,
    TCO.DM.DISB_AMT2, TCO.DM.RTO_CHARGES
  ];
  if (syncCols.indexOf(col) !== -1) {
    var val = isTextCol_(col, 'DM') ? toUpper_(value) : value;
    syncFieldToAccRto_(ss, sheet, row, col, val);
    // REG NO special style
    if (col === TCO.DM.REG_NO) {
      sheet.getRange(row, col)
        .setFontColor(TCO.STYLE.REG_NO.color)
        .setFontSize(TCO.STYLE.REG_NO.size)
        .setFontStyle('normal');
      // Update RTO CODE too
      var dmNo = getDmNo_(sheet, row);
      if (dmNo) {
        var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
        var rtoRow   = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);
        if (rtoRow) {
          var code = generateRtoCode_(value);
          if (code) {
            rtoSheet.getRange(rtoRow, TCO.RTO.RTO_CODE)
              .setValue(code)
              .setFontColor(TCO.STYLE.AUTO_GEN.color)
              .setFontStyle(TCO.STYLE.AUTO_GEN.style)
              .setFontSize(10);
          }
        }
      }
    }
    return;
  }

  // 10. ROI % format
  if (col === TCO.DM.ROI1 || col === TCO.DM.ROI2) {
    sheet.getRange(row, col).setNumberFormat('0.00" %"');
    return;
  }

  // 11. Currency columns — apply Indian format
  var currDM = CURRENCY_COLS.DM || [];
  if (currDM.indexOf(col) !== -1 && value !== '') {
    sheet.getRange(row, col).setNumberFormat('₹ ##,##,##0');
    return;
  }
}


// ─────────────────────────────────────────────────────────────
// ── ACC SHEET HANDLER ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function handleAccEdit_(ss, sheet, row, col, value) {

  // DEALER PAY STATUS → sync to DM + RTO
  if (col === TCO.ACC.DEALER_PAY_ST) {
    var dmNo = sheet.getRange(row, TCO.ACC.DM_NO).getValue();
    if (dmNo) syncDealerPayStatus_(ss, dmNo, value);
    return;
  }

  // DISB RECV DATE → sync to DM
  if (col === TCO.ACC.DISB_RECV_DATE) {
    var dmNo2 = sheet.getRange(row, TCO.ACC.DM_NO).getValue();
    if (dmNo2) syncDisbDateToDm_(ss, dmNo2, value);
    return;
  }

  // RTO PAID → recalc profit
  if (col === TCO.ACC.RTO_PAID) {
    var charges = parseFloat(sheet.getRange(row, TCO.ACC.RTO_CHARGES).getValue()) || 0;
    var profit  = charges - (parseFloat(value) || 0);
    var pc = sheet.getRange(row, TCO.ACC.RTO_PROFIT);
    pc.setValue(profit)
      .setFontColor(profit >= 0 ? '#10b981' : '#ef4444')
      .setFontWeight('bold')
      .setFontStyle('italic');
    return;
  }

  // Currency format
  var currACC = CURRENCY_COLS.ACC || [];
  if (currACC.indexOf(col) !== -1 && value !== '') {
    sheet.getRange(row, col).setNumberFormat('₹ ##,##,##0');
    return;
  }
}


// ─────────────────────────────────────────────────────────────
// ── RTO SHEET HANDLER ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function handleRtoEdit_(ss, sheet, row, col, value) {

  // SELLER PHONE — 10 digit validation (warning only)
  if (col === TCO.RTO.SELLER_PHONE) {
    validatePhone_(sheet, row, col, value);
    return;
  }

  // REG NO → RTO CODE
  if (col === TCO.RTO.REG_NO) {
    var code = generateRtoCode_(value);
    if (code) {
      sheet.getRange(row, TCO.RTO.RTO_CODE)
        .setValue(code)
        .setFontColor(TCO.STYLE.AUTO_GEN.color)
        .setFontStyle(TCO.STYLE.AUTO_GEN.style)
        .setFontSize(10);
    }
    sheet.getRange(row, col)
      .setFontColor(TCO.STYLE.REG_NO.color)
      .setFontSize(TCO.STYLE.REG_NO.size);
    return;
  }

  // RTO VENDOR → sync to ACC
  if (col === TCO.RTO.VENDOR) {
    var dmNo = sheet.getRange(row, TCO.RTO.DM_NO).getValue();
    if (dmNo) {
      var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
      var accRow   = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
      if (accRow) {
        styleCell_(accSheet.getRange(accRow, TCO.ACC.RTO_VENDOR).setValue(toUpper_(value)), TCO.STYLE.AUTO_PULL);
      }
    }
    return;
  }

  // RC STATUS → sync to ACC
  if (col === TCO.RTO.RC_STATUS) {
    var dmNo2 = sheet.getRange(row, TCO.RTO.DM_NO).getValue();
    if (dmNo2) {
      var accSheet2 = ss.getSheetByName(TCO.SHEET.ACC);
      var accRow2   = findRowByDmNo_(accSheet2, dmNo2, TCO.ACC.DM_NO);
      if (accRow2) {
        styleCell_(accSheet2.getRange(accRow2, TCO.ACC.RC_STATUS).setValue(toUpper_(value)), TCO.STYLE.AUTO_PULL);
      }
    }
    updatePendingDays_(sheet, row);
    return;
  }

  // RC DATE → freeze pending days
  if (col === TCO.RTO.RC_DATE) {
    updatePendingDays_(sheet, row);
    return;
  }
}


// ─────────────────────────────────────────────────────────────
// ── GENERATE & PROPAGATE ─────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function tryGenerateEntry_(ss, sheet, row) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
  var dmDate = sheet.getRange(row, TCO.DM.DM_DATE).getValue();
  var buyer  = sheet.getRange(row, TCO.DM.BUYER).getValue();
  if (!dmDate || !buyer || String(buyer).trim() === '') return;
  var currentDmNo = sheet.getRange(row, TCO.DM.DM_NO).getValue();
  if (currentDmNo && String(currentDmNo).startsWith('TCO')) return;
  var dmNo = generateDmNo_(sheet, row, dmDate);
  if (!dmNo) return;

  Utilities.sleep(2000);
  fillDmAutoFields_(sheet, row, dmNo);
  createAccRow_(ss, sheet, row, dmNo);
  createRtoRow_(ss, sheet, row, dmNo);
  updateMasterData_(ss);
  } finally {
    lock.releaseLock();
  }
}


// ── DM_NO GENERATOR ───────────────────────────────────────────
function generateDmNo_(sheet, row, dmDate) {
  var d      = new Date(dmDate);
  var months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  var prefix = 'TCO' + months[d.getMonth()] + String(d.getFullYear()).slice(-2);
  var data   = sheet.getDataRange().getValues();
  var maxSeq = 0;

  for (var i = TCO.DATA_ROW - 1; i < data.length; i++) {
    var v = String(data[i][0]).trim();
    if (v.startsWith(prefix)) {
      var seq = parseInt(v.replace(prefix, ''), 10) || 0;
      if (seq > maxSeq) maxSeq = seq;
    }
  }





  var dmNo = prefix + String(maxSeq + 1).padStart(3, '0');
  var cell = sheet.getRange(row, TCO.DM.DM_NO);
 var richText = SpreadsheetApp.newRichTextValue()
  .setText(dmNo)
  .setTextStyle(0, 3, SpreadsheetApp.newTextStyle().setForegroundColor('#78909C').setBold(true).build())
  .setTextStyle(3, 6, SpreadsheetApp.newTextStyle().setForegroundColor('#5D4037').setBold(true).build())
  .setTextStyle(6, 8, SpreadsheetApp.newTextStyle().setForegroundColor('#78909C').setBold(true).build())
  .setTextStyle(8, 11, SpreadsheetApp.newTextStyle().setForegroundColor('#134e4a').setBold(true).setItalic(true).build())
  .build();
cell.setRichTextValue(richText);
cell.setFontSize(12).setHorizontalAlignment('center');
  return dmNo;
}


// ── DM AUTO FIELDS ────────────────────────────────────────────
function fillDmAutoFields_(sheet, row, dmNo) {
  var dmDate = sheet.getRange(row, TCO.DM.DM_DATE).getValue();
  if (!dmDate) return;

  var d  = new Date(dmDate);
  var mm = String(d.getMonth() + 1).padStart(2, '0');

  // MONTH
  styleCell_(
    sheet.getRange(row, TCO.DM.MONTH).setValue(mm + '-' + d.getFullYear()),
    TCO.STYLE.AUTO_GEN
  );

  // NO OF DAYS formula
  styleCell_(
    sheet.getRange(row, TCO.DM.NO_OF_DAYS).setFormula(
      '=IF(B'+row+'="","",IF(D'+row+'="",TODAY()-B'+row+',MAX(0,D'+row+'-B'+row+')))'
    ),
    TCO.STYLE.AUTO_GEN
  );

  // SELLER AUTO
  var seller = sheet.getRange(row, TCO.DM.SELLER).getValue();
  if (seller) {
    styleCell_(
      sheet.getRange(row, TCO.DM.SELLER_AUTO).setValue(toUpper_(seller)),
      TCO.STYLE.AUTO_PULL
    );
  }

  // Default: DEALER PAY STATUS = PENDING (DM + ACC + RTO)
  sheet.getRange(row, TCO.DM.DEALER_PAY_ST).setValue('PENDING');
  sheet.getRange(row, TCO.DM.DEALER_PAY_ST)
  .setFontColor('#134e4a')
  .setFontStyle('italic')
  .setFontWeight('bold')
  .setFontSize(12);
}


// ── FILL EXEC DETAILS ─────────────────────────────────────────
function fillExecDetails_(ss, sheet, row, empId) {
  var execC   = sheet.getRange(row, TCO.DM.EXEC_NAME);
  var branchC = sheet.getRange(row, TCO.DM.BRANCH);
  var tlC     = sheet.getRange(row, TCO.DM.TEAM_LEADER);

  var empSheet = ss.getSheetByName(TCO.SHEET.EMP);
  if (!empSheet) return;
  var data  = empSheet.getDataRange().getValues();
  var found = false;

  for (var i = 2; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(empId).trim().toUpperCase()) {
      var name   = toUpper_(data[i][1]);
      var branch = toUpper_(data[i][7]); // col 8 confirmed
      var tl     = toUpper_(data[i][8]); // col 9

      styleCell_(execC.setValue(name),     TCO.STYLE.AUTO_PULL);
      styleCell_(branchC.setValue(branch), TCO.STYLE.AUTO_PULL);
      styleCell_(tlC.setValue(tl),         TCO.STYLE.AUTO_PULL);

      // Sync to ACC + RTO
      var dmNo = getDmNo_(sheet, row);
      if (dmNo) {
        var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
        var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
        var accRow   = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
        var rtoRow   = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);
        if (accRow) styleCell_(accSheet.getRange(accRow, TCO.ACC.EXEC).setValue(name),     TCO.STYLE.AUTO_PULL);
        if (rtoRow) {
          styleCell_(rtoSheet.getRange(rtoRow, TCO.RTO.EXEC).setValue(name),     TCO.STYLE.AUTO_PULL);
          styleCell_(rtoSheet.getRange(rtoRow, TCO.RTO.BRANCH).setValue(branch), TCO.STYLE.AUTO_PULL);
        }
      }
      found = true;
      break;
    }
  }

  if (!found) {
    // FIX Point 8: EMP NOT FOUND — red italic
    [execC, branchC, tlC].forEach(function(c) {
      styleCell_(c.setValue('EMP NOT FOUND'), TCO.STYLE.NOT_FOUND);
    });
  }
}


// ── FILL DEALER DETAILS ───────────────────────────────────────
function fillDealerDetails_(ss, sheet, row, contactNo) {
  var dlrC  = sheet.getRange(row, TCO.DM.DEALER_NAME);
  var authC = sheet.getRange(row, TCO.DM.AUTH_PERSON);
  var locC  = sheet.getRange(row, TCO.DM.LOCATION);

  var dlrSheet = ss.getSheetByName(TCO.SHEET.DEALER);
  if (!dlrSheet) return;
  var data  = dlrSheet.getDataRange().getValues();
  var found = false;

  for (var i = 2; i < data.length; i++) {
    if (String(data[i][5]).trim() === String(contactNo).trim()) {
      var dlr  = toUpper_(data[i][1]);
      var auth = toUpper_(data[i][4]);
      var loc  = toUpper_(data[i][2]);

      styleCell_(dlrC.setValue(dlr),   TCO.STYLE.AUTO_PULL);
      styleCell_(authC.setValue(auth), TCO.STYLE.AUTO_PULL);
      styleCell_(locC.setValue(loc),   TCO.STYLE.AUTO_PULL);

      // Sync to ACC + RTO
      var dmNo = getDmNo_(sheet, row);
      if (dmNo) {
        var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
        var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
        var accRow   = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
        var rtoRow   = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);
        if (accRow) styleCell_(accSheet.getRange(accRow, TCO.ACC.DEALER).setValue(dlr), TCO.STYLE.AUTO_PULL);
        if (rtoRow) styleCell_(rtoSheet.getRange(rtoRow, TCO.RTO.DEALER).setValue(dlr), TCO.STYLE.AUTO_PULL);
      }
      found = true;
      break;
    }
  }

  if (!found) {
    [dlrC, authC, locC].forEach(function(c) {
      styleCell_(c.setValue('DLR NOT FOUND'), TCO.STYLE.NOT_FOUND);
    });
  }
}


// ── CLEAR EMP FIELDS (Point 8 — delete) ──────────────────────
function clearEmpFields_(sheet, row) {
  [TCO.DM.EXEC_NAME, TCO.DM.BRANCH, TCO.DM.TEAM_LEADER].forEach(function(c) {
    sheet.getRange(row, c).setValue('').setFontColor('#000000').setFontStyle('normal');
  });
}

function clearEmpInAccRto_(ss, sheet, row) {
  var dmNo = getDmNo_(sheet, row);
  if (!dmNo) return;
  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  var accRow   = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
  var rtoRow   = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);
  if (accRow) accSheet.getRange(accRow, TCO.ACC.EXEC).setValue('').setFontStyle('normal');
  if (rtoRow) {
    rtoSheet.getRange(rtoRow, TCO.RTO.EXEC).setValue('').setFontStyle('normal');
    rtoSheet.getRange(rtoRow, TCO.RTO.BRANCH).setValue('').setFontStyle('normal');
  }
}


// ── CLEAR DEALER FIELDS (Point 8 — delete) ───────────────────
function clearDlrFields_(sheet, row) {
  [TCO.DM.DEALER_NAME, TCO.DM.AUTH_PERSON, TCO.DM.LOCATION].forEach(function(c) {
    sheet.getRange(row, c).setValue('').setFontColor('#000000').setFontStyle('normal');
  });
}

function clearDlrInAccRto_(ss, sheet, row) {
  var dmNo = getDmNo_(sheet, row);
  if (!dmNo) return;
  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  var accRow   = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
  var rtoRow   = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);
  if (accRow) accSheet.getRange(accRow, TCO.ACC.DEALER).setValue('').setFontStyle('normal');
  if (rtoRow) rtoSheet.getRange(rtoRow, TCO.RTO.DEALER).setValue('').setFontStyle('normal');
}


// ── CREATE ACC ROW ────────────────────────────────────────────
function createAccRow_(ss, dmSheet, row, dmNo) {
  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  if (!accSheet) return;
  if (findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO)) return;

  var d      = dmSheet.getRange(row, 1, 1, TCO.DM.REMARKS).getValues()[0];
  var newRow = Math.max(accSheet.getLastRow() + 1, TCO.DATA_ROW);
  var data   = new Array(TCO.ACC.REMARKS).fill('');

  // All auto-pull fields — CAPS applied
  data[TCO.ACC.DM_DATE-1]       = d[TCO.DM.DM_DATE-1];
  data[TCO.ACC.DM_NO-1]         = dmNo;
  data[TCO.ACC.BUYER-1]         = toUpper_(d[TCO.DM.BUYER-1]);
  data[TCO.ACC.PHONE-1]         = d[TCO.DM.PHONE-1];
  data[TCO.ACC.PHONE2-1]        = d[TCO.DM.PHONE2-1];
  data[TCO.ACC.REG_NO-1]        = toUpper_(d[TCO.DM.REG_NO-1]);
  data[TCO.ACC.MODEL-1]         = toUpper_(d[TCO.DM.MODEL-1]);
  data[TCO.ACC.MFG_YEAR-1]      = d[TCO.DM.MFG_YEAR-1];
  data[TCO.ACC.BANK-1]          = toUpper_(d[TCO.DM.BANK-1]);
  data[TCO.ACC.PRODUCT-1]       = toUpper_(d[TCO.DM.PRODUCT-1]);
  data[TCO.ACC.DEALER-1]        = toUpper_(d[TCO.DM.DEALER_NAME-1]);
  data[TCO.ACC.EXEC-1]          = toUpper_(d[TCO.DM.EXEC_NAME-1]);
  data[TCO.ACC.DISB1-1]         = d[TCO.DM.DISB_AMT1-1];
  data[TCO.ACC.DISB2-1]         = d[TCO.DM.DISB_AMT2-1];
  data[TCO.ACC.DEALER_PAY_ST-1] = 'PENDING';
  data[TCO.ACC.RTO_CHARGES-1]   = d[TCO.DM.RTO_CHARGES-1];

  accSheet.getRange(newRow, 1, 1, data.length).setValues([data]);
  applyRowFormat_(accSheet, newRow);
}


// ── CREATE RTO ROW ────────────────────────────────────────────
function createRtoRow_(ss, dmSheet, row, dmNo) {
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  if (!rtoSheet) return;
  if (findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO)) return;

  var d      = dmSheet.getRange(row, 1, 1, TCO.DM.REMARKS).getValues()[0];
  var newRow = Math.max(rtoSheet.getLastRow() + 1, TCO.DATA_ROW);
  var data   = new Array(TCO.RTO.REMARKS).fill('');
  var code   = generateRtoCode_(d[TCO.DM.REG_NO-1]);

  // All auto-pull fields — CAPS applied
  data[TCO.RTO.DM_DATE-1]       = d[TCO.DM.DM_DATE-1];
  data[TCO.RTO.DM_NO-1]         = dmNo;
  data[TCO.RTO.DEALER_PAY_ST-1] = 'PENDING';
  data[TCO.RTO.PRODUCT-1]       = toUpper_(d[TCO.DM.PRODUCT-1]);
  data[TCO.RTO.BRANCH-1]        = toUpper_(d[TCO.DM.BRANCH-1]);
  data[TCO.RTO.EXEC-1]          = toUpper_(d[TCO.DM.EXEC_NAME-1]);
  data[TCO.RTO.DEALER-1]        = toUpper_(d[TCO.DM.DEALER_NAME-1]);
  data[TCO.RTO.REG_NO-1]        = toUpper_(d[TCO.DM.REG_NO-1]);
  data[TCO.RTO.RTO_CODE-1]      = code;
  data[TCO.RTO.MODEL-1]         = toUpper_(d[TCO.DM.MODEL-1]);
  data[TCO.RTO.MFG_YEAR-1]      = d[TCO.DM.MFG_YEAR-1];
  data[TCO.RTO.SELLER-1]        = toUpper_(d[TCO.DM.SELLER-1]);
  data[TCO.RTO.BUYER-1]         = toUpper_(d[TCO.DM.BUYER-1]);
  data[TCO.RTO.BUYER_PHONE-1]   = d[TCO.DM.PHONE-1];   // FIX Point 12: DM col G
  data[TCO.RTO.BANK-1]          = toUpper_(d[TCO.DM.BANK-1]);
  data[TCO.RTO.RC_STATUS-1]     = 'PENDING';

  rtoSheet.getRange(newRow, 1, 1, data.length).setValues([data]);

  // PENDING DAYS formula — uses RTO col A (DM DATE)
  rtoSheet.getRange(newRow, TCO.RTO.PENDING_DAYS).setFormula(
    '=IF(A'+newRow+'="","",IF(Y'+newRow+'="",TODAY()-A'+newRow+',MAX(0,Y'+newRow+'-A'+newRow+')))'
  );

  applyRowFormat_(rtoSheet, newRow);
}


// ── SYNC FUNCTIONS ────────────────────────────────────────────
function syncFieldToAccRto_(ss, sheet, row, dmCol, value) {
  var dmNo = getDmNo_(sheet, row);
  if (!dmNo) return;

  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  var accRow   = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
  var rtoRow   = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);

  var accCol = DM_TO_ACC[dmCol];
  var rtoCol = DM_TO_RTO[dmCol];

  if (accCol && accRow) {
    var ac = accSheet.getRange(accRow, accCol);
    styleCell_(ac.setValue(value), TCO.STYLE.AUTO_PULL);
    if (accCol === TCO.ACC.REG_NO) {
      ac.setFontColor(TCO.STYLE.REG_NO.color).setFontSize(TCO.STYLE.REG_NO.size).setFontStyle('normal');
    }
  }
  if (rtoCol && rtoRow) {
    var rc = rtoSheet.getRange(rtoRow, rtoCol);
    styleCell_(rc.setValue(value), TCO.STYLE.AUTO_PULL);
    if (rtoCol === TCO.RTO.REG_NO) {
      rc.setFontColor(TCO.STYLE.REG_NO.color).setFontSize(TCO.STYLE.REG_NO.size).setFontStyle('normal');
    }
  }
}

function syncDealerPayStatus_(ss, dmNo, val) {
  var dm  = ss.getSheetByName(TCO.SHEET.DM);
  var rto = ss.getSheetByName(TCO.SHEET.RTO);
  var dr  = findRowByDmNo_(dm,  dmNo, TCO.DM.DM_NO);
  var rr  = findRowByDmNo_(rto, dmNo, TCO.RTO.DM_NO);
  if (dr) styleCell_(dm.getRange(dr,   TCO.DM.DEALER_PAY_ST).setValue(toUpper_(val)),  TCO.STYLE.AUTO_PULL);
  if (rr) styleCell_(rto.getRange(rr,  TCO.RTO.DEALER_PAY_ST).setValue(toUpper_(val)), TCO.STYLE.AUTO_PULL);
}

function syncDisbDateToDm_(ss, dmNo, val) {
  var dm  = ss.getSheetByName(TCO.SHEET.DM);
  var dr  = findRowByDmNo_(dm, dmNo, TCO.DM.DM_NO);
  if (dr) styleCell_(dm.getRange(dr, TCO.DM.DM_RECV_DATE).setValue(val), TCO.STYLE.AUTO_PULL);
}


// ─────────────────────────────────────────────────────────────
// ── STYLING FUNCTIONS ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────

// Apply style object to a cell
function styleCell_(cell, style) {
  return cell
    .setFontColor(style.color)
    .setFontStyle(style.style)
    .setFontSize(style.size || 10);
}

// Apply per-column style based on type
function applyColumnStyle_(sheet, sName, row, col) {
  var key = sName === TCO.SHEET.DM  ? 'DM'  :
            sName === TCO.SHEET.ACC ? 'ACC' : 'RTO';

  // SPECIAL COLS — size 12 + custom color (DM NO, REG NO, key cols)
  var specMap = SPECIAL_COLS[key] || {};
  if (specMap[col]) {
    var sp = specMap[col];
    sheet.getRange(row, col)
      .setFontColor(sp.color)
      .setFontStyle(sp.style)
      .setFontSize(sp.size);
    return;
  }

  // AUTO GEN
  if ((AUTO_GEN_COLS[key] || []).indexOf(col) !== -1) {
    styleCell_(sheet.getRange(row, col), TCO.STYLE.AUTO_GEN);
    return;
  }
  // AUTO PULL
  if ((AUTO_PULL_COLS[key] || []).indexOf(col) !== -1) {
    styleCell_(sheet.getRange(row, col), TCO.STYLE.AUTO_PULL);
    return;
  }
}

// Full row format on new row creation
function applyRowFormat_(sheet, row) {
  var sName   = sheet.getName();
  var maxCols = sName === TCO.SHEET.DM  ? TCO.DM.REMARKS  :
                sName === TCO.SHEET.ACC ? TCO.ACC.REMARKS  : TCO.RTO.REMARKS;
  var key     = sName === TCO.SHEET.DM  ? 'DM'  :
                sName === TCO.SHEET.ACC ? 'ACC' : 'RTO';

  // Base format
  sheet.getRange(row, 1, 1, maxCols)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setFontColor('#000000')
    .setFontStyle('normal')
    .setFontSize(10)
    .setFontFamily('Arial')
    .setBorder(true,true,true,true,true,true,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);

  // SPECIAL COLS — size 12 + custom color (all 3 sheets)
  var specMap = SPECIAL_COLS[key] || {};
  Object.keys(specMap).forEach(function(c) {
    var sp = specMap[Number(c)];
    sheet.getRange(row, Number(c))
      .setFontColor(sp.color)
      .setFontStyle(sp.style)
      .setFontSize(sp.size);
  });

  // AUTO GEN
  (AUTO_GEN_COLS[key] || []).forEach(function(c) {
    styleCell_(sheet.getRange(row, c), TCO.STYLE.AUTO_GEN);
    sheet.getRange(row, c).setBackground('#BBF7D0');
  });

  // AUTO PULL
  (AUTO_PULL_COLS[key] || []).forEach(function(c) {
    styleCell_(sheet.getRange(row, c), TCO.STYLE.AUTO_PULL);
    sheet.getRange(row, c).setBackground('#EEF2FF');
    // SPECIAL COLS override — size 12
    if (specMap[c]) {
      sheet.getRange(row, c).setFontSize(specMap[c].size);
    }
  });

  // Currency format on currency cols
  var curr = CURRENCY_COLS[key] || [];
  curr.forEach(function(c) {
    sheet.getRange(row, c).setNumberFormat('₹ ##,##,##0');
  });

  // ROI % format
  if (sName === TCO.SHEET.DM) {
    sheet.getRange(row, TCO.DM.ROI1).setNumberFormat('0.00" %"');
    sheet.getRange(row, TCO.DM.ROI2).setNumberFormat('0.00" %"');
  }
}


// ─────────────────────────────────────────────────────────────
// ── UTILITY FUNCTIONS ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────

// CAPS conversion + Warning colors (no popup, no block)
function applyCaps_(sheet, sName, row, col, value) {
  if (!value || typeof value !== 'string') return;
  var key = sName === TCO.SHEET.DM  ? 'DM'  :
            sName === TCO.SHEET.ACC ? 'ACC' : 'RTO';
  var cell = sheet.getRange(row, col);

  // Email → plain text, no validation, no hyperlink, no caps
  if ((EMAIL_COLS[key] || []).indexOf(col) !== -1) {
    cell.setValue(value.toLowerCase());
    cell.setFontLine('none');  // remove hyperlink underline
    // Yellow warning if not valid email format
    if (value && value.indexOf('@') === -1) {
      cell.setBackground('#fef08a'); // yellow
    } else {
      cell.setBackground(null);
    }
    return;
  }

  // Link/URL → lowercase, no caps
  if ((LINK_COLS[key] || []).indexOf(col) !== -1) {
    cell.setValue(value.toLowerCase()); return;
  }

  // URL value → skip
  if (value.indexOf('http') === 0 || value.indexOf('www') === 0) return;

  // Numbers → skip caps, handle currency warning
  if ((NUMBER_COLS[key] || []).indexOf(col) !== -1) {
    // Currency warning — red if not a number
    var currKey = CURRENCY_COLS[key] || [];
    if (currKey.indexOf(col) !== -1) {
      if (isNaN(parseFloat(value)) || !isFinite(value)) {
        cell.setBackground('#fca5a5'); // red
      } else {
        cell.setBackground(null);
        cell.setNumberFormat('₹ ##,##,##0');
      }
    }
    return;
  }

  // Dropdown → skip
  if ((DROPDOWN_COLS[key] || []).indexOf(col) !== -1) return;

  // All text cols → CAPS + word warning
  
  // Word only cols — yellow if contains numbers
  var wordOnlyCols = { DM:[6,11,27], ACC:[18,22,26], RTO:[] };
  if ((wordOnlyCols[key] || []).indexOf(col) !== -1) {
    if (/\d/.test(value)) {
      cell.setBackground('#fef08a'); // yellow warning
    } else {
      cell.setBackground(null);
    }
    return;
  }

  // Clear any previous warning
  cell.setBackground(null);
}

// Smart toUpperCase — skip null/numbers
function toUpper_(val) {
  if (!val) return '';
  if (typeof val === 'number') return val;
  return String(val).toUpperCase();
}

// Handle number entries — currency validation + clear red
function handleNumberEntry_(sheet, sName, row, col, value) {
  var key = sName === TCO.SHEET.DM  ? 'DM'  :
            sName === TCO.SHEET.ACC ? 'ACC' : 'RTO';
  var cell = sheet.getRange(row, col);
  var curr = CURRENCY_COLS[key] || [];
  if (curr.indexOf(col) !== -1) {
    if (isNaN(value) || value < 0) {
      cell.setBackground('#fca5a5'); // red warning
    } else {
      cell.setBackground(null);     // clear — valid number
      cell.setNumberFormat('₹ ##,##,##0');
    }
  }
}

// Phone validation — warning only: red bg, no popup, no block
function validatePhone_(sheet, row, col, value) {
  var cell = sheet.getRange(row, col);
  if (!value || String(value).trim() === '') {
    cell.setBackground(null); return;
  }
  var num = String(value).replace(/\D/g, '');
  if (num.length !== 10) {
    cell.setBackground('#fca5a5'); // Red bg — warning only
  } else {
    cell.setBackground(null); // Valid — clear color
  }
}

function isTextCol_(col, key) {
  return (NUMBER_COLS[key] || []).indexOf(col) === -1;
}

function getDmNo_(sheet, row) {
  var v = sheet.getRange(row, TCO.DM.DM_NO).getValue();
  return (v && String(v).startsWith('TCO')) ? String(v).trim() : null;
}

function findRowByDmNo_(sheet, dmNo, col) {
  if (!sheet || !dmNo) return null;
  var data = sheet.getDataRange().getValues();
  var str  = String(dmNo).trim();
  for (var i = TCO.DATA_ROW - 1; i < data.length; i++) {
    var cellVal = String(data[i][col-1]).trim();
    if (cellVal === str || cellVal.replace(/\s/g,'') === str.replace(/\s/g,'')) return i + 1;
  }
  return null;
}

function generateRtoCode_(regNo) {
  if (!regNo) return '';
  var m = String(regNo).trim().toUpperCase().match(/^([A-Z]{2})(\d{1,2})/);
  if (!m) return '';
  return m[1] + String(parseInt(m[2],10)).padStart(2,'0');
}

function updatePendingDays_(sheet, row) {
  var dmDate = sheet.getRange(row, TCO.RTO.DM_DATE).getValue();
  var rcDate = sheet.getRange(row, TCO.RTO.RC_DATE).getValue();
  if (!dmDate) { sheet.getRange(row, TCO.RTO.PENDING_DAYS).setValue(''); return; }
  var end  = rcDate ? new Date(rcDate) : new Date();
  var days = Math.max(0, Math.round((end - new Date(dmDate)) / 86400000));
  // FIX Point 14: RED + size 12
  styleCell_(
    sheet.getRange(row, TCO.RTO.PENDING_DAYS).setValue(days),
    TCO.STYLE.PENDING
  );
}


// ── MASTER DATA ───────────────────────────────────────────────
function updateMasterData_(ss) {
  try {
    var ms = ss.getSheetByName(TCO.SHEET.MASTER);
    var dm = ss.getSheetByName(TCO.SHEET.DM);
    var ac = ss.getSheetByName(TCO.SHEET.ACC);
    var rt = ss.getSheetByName(TCO.SHEET.RTO);
    if (!ms || !dm) return;

    var dmD = dm.getDataRange().getValues();
    var acD = ac ? ac.getDataRange().getValues() : [];
    var rtD = rt ? rt.getDataRange().getValues() : [];
    var acM = {}, rtM = {};

    for (var i = TCO.DATA_ROW-1; i < acD.length; i++) {
      var k = String(acD[i][TCO.ACC.DM_NO-1]).trim();
      if (k) acM[k] = acD[i];
    }
    for (var j = TCO.DATA_ROW-1; j < rtD.length; j++) {
      var k2 = String(rtD[j][TCO.RTO.DM_NO-1]).trim();
      if (k2) rtM[k2] = rtD[j];
    }

    var lr = ms.getLastRow();
    if (lr >= TCO.DATA_ROW)
      ms.getRange(TCO.DATA_ROW,1,lr-TCO.DATA_ROW+1,ms.getLastColumn()||99).clearContent();

    var rows = [];
    for (var k = TCO.DATA_ROW-1; k < dmD.length; k++) {
      var dmNo = String(dmD[k][0]).trim();
      if (!dmNo.startsWith('TCO')) continue;
      var acc = acM[dmNo] || [], rto = rtM[dmNo] || [], r = [];
      for (var d=0;d<57;d++) r.push(dmD[k][d]||'');
      for (var a=14;a<45;a++) r.push(acc[a]||'');
      for (var x=8;x<30;x++) r.push(rto[x]||'');
      rows.push(r);
    }
    if (rows.length > 0)
      ms.getRange(TCO.DATA_ROW,1,rows.length,rows[0].length).setValues(rows);
  } catch(e) { Logger.log('masterData err: '+e.message); }
}


// ── ONOPEN MENU ───────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ TCO Admin')
    .addItem('📦 Manual Backup',        'manualBackup')
    .addItem('🔒 Apply Protections', 'TCO_applyAllProtections')
    .addItem('🗂️ Borders & Formatting', 'applyBordersAll')
    .addItem('📋 Update Master Data',   'updateMasterDataManual')
    .addItem('🔁 Restore from Backup',  'restoreFromBackup')
    .addSeparator()
    .addSubMenu(ui.createMenu('🗑️ Data Reset')
      .addItem('Reset ALL Sheets',   'resetAllData')
      .addItem('Reset DM Only',      'resetDmOnly')
      .addItem('Reset Account Only', 'resetAccOnly')
      .addItem('Reset RTO Only',     'resetRtoOnly')
      .addItem('Reset Master Data',  'resetMasterOnly'))
    .addToUi();
}

function updateMasterDataManual() {
  updateMasterData_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ Master Data updated!');
}

function manualBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [[TCO.SHEET.DM,TCO.SHEET.BK_DM],[TCO.SHEET.ACC,TCO.SHEET.BK_ACC],[TCO.SHEET.RTO,TCO.SHEET.BK_RTO]]
  .forEach(function(p) {
    var src=ss.getSheetByName(p[0]), bk=ss.getSheetByName(p[1]);
    if (!src||!bk) return;
    bk.clearContents();
    var d=src.getDataRange().getValues();
    bk.getRange(1,1,d.length,d[0].length).setValues(d);
  });
  SpreadsheetApp.getUi().alert('✅ Backup done!\n'+new Date().toLocaleString('en-IN'));
}

function restoreFromBackup() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ Overwrite DM with backup?',ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(TCO.SHEET.BK_DM);
  var t   = ss.getSheetByName(TCO.SHEET.DM);
  if (!src||!t) { ui.alert('❌ Sheet not found'); return; }
  var bd = src.getDataRange().getValues(), ok = false;
  for (var i=TCO.DATA_ROW-1;i<bd.length;i++) {
    if (String(bd[i][0]).startsWith('TCO')&&String(bd[i][5]).trim()!=='') { ok=true; break; }
  }
  if (!ok) { ui.alert('❌ Invalid backup data'); return; }
  t.clearContents();
  t.getRange(1,1,bd.length,bd[0].length).setValues(bd);
  ui.alert('✅ DM restored!\n'+new Date().toLocaleString('en-IN'));
}

function applyBordersAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [TCO.SHEET.DM,TCO.SHEET.ACC,TCO.SHEET.RTO].forEach(function(n) {
    var s=ss.getSheetByName(n); if (!s) return;
    var lr=Math.max(s.getLastRow(),TCO.DATA_ROW+5), lc=s.getLastColumn();
    if (lc<1) return;
    s.getRange(TCO.DATA_ROW,1,lr-TCO.DATA_ROW+1,lc)
     .setBorder(true,true,true,true,true,true,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);
  });
  SpreadsheetApp.getUi().alert('✅ Borders applied!');
}


function applyBackgroundToAllRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [
    {name: TCO.SHEET.DM,  key: 'DM'},
    {name: TCO.SHEET.ACC, key: 'ACC'},
    {name: TCO.SHEET.RTO, key: 'RTO'}
  ];
  sheets.forEach(function(sh) {
    var sheet = ss.getSheetByName(sh.name);
    if (!sheet) return;
    var lastRow = sheet.getMaxRows();
    if (lastRow < TCO.DATA_ROW) return;
    for (var row = TCO.DATA_ROW; row <= lastRow; row++) {
      (AUTO_GEN_COLS[sh.key] || []).forEach(function(c) {
        sheet.getRange(row, c).setBackground('#F8FAFC');
      });
      (AUTO_PULL_COLS[sh.key] || []).forEach(function(c) {
        sheet.getRange(row, c).setBackground('#EEF2FF');
      });
    }
  });
  SpreadsheetApp.getUi().alert('✅ Background colors applied!');
}





