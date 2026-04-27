/**
 * ============================================================
 *  TCO_Operations.gs — v4.0 | FY 2026-27
 *  onEdit Trigger | Auto Sync | Lookups | Master Update
 * ============================================================
 *  Column Numbers: PDF Mapping v3 (EXACT)
 *
 *  RULES:
 *  1. DM NO → Admin fills manually (script only reads)
 *  2. Sync Trigger → DM NO + DATE + BUYER NAME all present
 *  3. Delete Rule → DATE + BUYER both empty → clear ACC+RTO
 *  4. Same Row Rule → DM Row 6 ↔ ACC Row 6 ↔ RTO Row 6
 *  5. General Sync → any manual field change → PDF map → sync
 *  6. RTO PROFIT → auto calc (RTO CHARGES - RTO PAID)
 *  7. MASTER → always hidden, always updated on any sync
 * ============================================================
 */

// ── GLOBAL CONFIG ─────────────────────────────────────────────
var TCO = {

  DATA_ROW : 4,

  SHEET : {
    DM     : 'DM_DISBURSEMENT MEMO',
    ACC    : 'ACCOUNT_PAYMENT_TRACKER',
    RTO    : 'RTO_TRACKER',
    MASTER : 'MASTER_DATA',
    DEALER : 'DEALER_MASTER',
    EMP    : 'TCO_EMPLOYEE_MASTER'
  },

  // ── DM Column Numbers — PDF v3 (55 cols) ──
  DM : {
    DM_NO       :1,  DM_DATE      :2,  MONTH       :3,  DISB_RECV_DATE:4,
    BUYER       :5,  PHONE1       :6,  PHONE2      :7,  EMAIL         :8,
    REG_NO      :9,  MODEL        :10, MFG_YEAR    :11, BANK          :12,
    PRODUCT     :13, LOAN_APPLIED :14, LOAN_APP1   :15, INS1          :16,
    SURAKSHA1   :17, SURAKSHA_TEN1:18, TOTAL_LOAN1 :19, FILE_CHG1     :20,
    ROI1        :21, EMI1         :22, LOAN_TEN1   :23, DISB_AMT1     :24,
    LOAN_APP2   :25, INS2         :26, SURAKSHA2   :27, SURAKSHA_TEN2 :28,
    TOTAL_LOAN2 :29, FILE_CHG2    :30, ROI2        :31, EMI2          :32,
    LOAN_TEN2   :33, DISB_AMT2    :34, RTO_CHARGES :35, PAY_TO_DLR    :36,
    ECHALLAN    :37, DEALER_PAY_ST:38, EMP_ID      :39, EXEC_NAME     :40,
    BRANCH      :41, TEAM_LEADER  :42, DEALER_NAME :43, AUTH_PERSON   :44,
    CONTACT_DLR :45, LOCATION     :46, VEHICLE_OWNER:47,PAYOUT_TCO    :48,
    PAYOUT_CD   :49, PAYOUT_SCORE :50, LOAN_SCORE  :51, RTO_SCORE     :52,
    EXE_INC     :53, NET_SCORE    :54, REMARKS     :55
  },

  // ── ACC Column Numbers — PDF v3 (43 cols) ──
  ACC : {
    DM_NO         :1,  BUYER         :2,  PHONE1        :3,  PHONE2        :4,
    REG_NO        :5,  MODEL         :6,  MFG_YEAR      :7,  BANK          :8,
    PRODUCT       :9,  DEALER        :10, EXEC          :11, DISB_RECV_DATE:12,
    UTR_NO        :13, DISB_RECV     :14, DISB1         :15, DISB2         :16,
    P1_NAME       :17, P1_DATE       :18, P1_AMT        :19, P1_ACC        :20,
    P2_NAME       :21, P2_DATE       :22, P2_AMT        :23, P2_ACC        :24,
    P3_NAME       :25, P3_DATE       :26, P3_AMT        :27, P3_ACC        :28,
    DEALER_PAY_ST :29, PAYOUT_DEALER :30, HOLD_BANK     :31, HOLD_TCO      :32,
    EXE_INC       :33, LOAN_SCORE    :34, PAYOUT_BANK   :35, NET_SCORE     :36,
    RTO_CHARGES   :37, RTO_VENDOR    :38, RTO_PAID      :39, RTO_PAY_DATE  :40,
    RC_STATUS     :41, RTO_PROFIT    :42, REMARKS       :43
  },

  // ── RTO Column Numbers — PDF v3 (30 cols) ──
  RTO : {
    DM_DATE       :1,  DM_NO         :2,  DEALER_PAY_ST :3,  PRODUCT       :4,
    BRANCH        :5,  EXEC          :6,  DEALER        :7,  REG_NO        :8,
    CASE_REC_DATE :9,  RTO_CODE      :10, MODEL         :11, MFG_YEAR      :12,
    CHASSIS       :13, ENGINE        :14, OWNER_SELLER  :15, SELLER_PHONE  :16,
    BUYER         :17, BUYER_PHONE   :18, BANK          :19, CASE_TYPE     :20,
    SCAN_LINK     :21, VENDOR        :22, RECEIPT_NO    :23, RC_STATUS     :24,
    RC_DATE       :25, PENDING_DAYS  :26, ORIG_RC       :27, RC_PROOF      :28,
    SYS_REMARKS   :29, REMARKS       :30
  },

  // ── Cell Styles ──
  STYLE : {
    AUTO_GEN  : { color:'#166534', style:'italic', size:10 },
    AUTO_PULL : { color:'#1A4D8F', style:'italic', size:10 },
    MANUAL    : { color:'#000000', style:'normal', size:10 },
    NOT_FOUND : { color:'#ef4444', style:'italic', size:10 },
    PENDING   : { color:'#ef4444', style:'normal', size:12 },
    KEY_FIELD : { color:'#134e4a', style:'bold',   size:12 },
    DM_NO_ST  : { color:'#5D4037', style:'italic', size:12 },
    REG_NO_ST : { color:'#c2410c', style:'normal', size:12 }
  }
};


// ── SYNC MAP — PDF v3 Mapping ─────────────────────────────────
// Key = source col (manual) → Value = destinations [{s: sheet_key, c: col_no}]
var SYNC_MAP = {

  DM : {
    2  : [{s:'RTO', c:1 }],                         // DM DATE      → RTO DM_DATE
    5  : [{s:'ACC', c:2 }, {s:'RTO', c:17}],        // BUYER NAME   → ACC, RTO
    6  : [{s:'ACC', c:3 }, {s:'RTO', c:18}],        // PHONE 1      → ACC, RTO
    7  : [{s:'ACC', c:4 }],                          // PHONE 2      → ACC
    9  : [{s:'ACC', c:5 }, {s:'RTO', c:8 }],        // REG NO       → ACC, RTO
    10 : [{s:'ACC', c:6 }, {s:'RTO', c:11}],        // MODEL        → ACC, RTO
    11 : [{s:'ACC', c:7 }, {s:'RTO', c:12}],        // MFG YEAR     → ACC, RTO
    12 : [{s:'ACC', c:8 }, {s:'RTO', c:19}],        // BANK         → ACC, RTO
    13 : [{s:'ACC', c:9 }, {s:'RTO', c:4 }],        // PRODUCT      → ACC, RTO
    24 : [{s:'ACC', c:15}],                          // DISB AMT 1   → ACC
    34 : [{s:'ACC', c:16}],                          // DISB AMT 2   → ACC
    35 : [{s:'ACC', c:37}],                          // RTO CHARGES  → ACC
    47 : [{s:'RTO', c:15}]                           // VEHICLE OWNER → RTO
  },

  ACC : {
    12 : [{s:'DM',  c:4 }],                          // DISB RECV DATE → DM
    29 : [{s:'DM',  c:38}, {s:'RTO', c:3}]          // DLR PAY STATUS → DM + RTO
  },

  RTO : {
    22 : [{s:'ACC', c:38}],                          // RTO VENDOR     → ACC
    24 : [{s:'ACC', c:41}]                           // RC STATUS      → ACC
  }
};


// ── COLUMN TYPE MAPS ─────────────────────────────────────────
var CURRENCY_COLS = {
  DM  : [14,15,16,17,19,20,22,24,25,26,29,30,32,34,35,36,37,48,49,53,54],
  ACC : [15,16,19,23,27,30,31,32,33,35,39,42],
  RTO : []
};

var PHONE_COLS = {
  DM  : [6, 7, 45],
  ACC : [],
  RTO : [16]
};

var WORD_COLS = {     // Yellow warning if number found
  DM  : [5, 47],
  ACC : [17, 21, 25],
  RTO : []
};

var SKIP_CAPS = {    // Do not apply CAPS (links/email)
  DM  : [8],         // EMAIL
  ACC : [],
  RTO : [21, 28]     // Links
};


// ─────────────────────────────────────────────────────────────
// ── MAIN TRIGGER ─────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function onEditTrigger(e) {
  if (!e || !e.range) return;
  try {
    var sheet = e.range.getSheet();
    var sName = sheet.getName();
    var row   = e.range.getRow();
    var col   = e.range.getColumn();
    var value = e.range.getValue();

    // Only rows 4+ and only 3 working sheets
    if (row < TCO.DATA_ROW) return;
    if (sName !== TCO.SHEET.DM  &&
        sName !== TCO.SHEET.ACC &&
        sName !== TCO.SHEET.RTO) return;

    var ss  = e.source;
    var key = sName === TCO.SHEET.DM  ? 'DM'  :
              sName === TCO.SHEET.ACC ? 'ACC' : 'RTO';

    // Step 1 — Validation + Formatting
    applyValidation_(sheet, key, row, col, value);

    // Step 2 — Sheet handler
    if      (sName === TCO.SHEET.DM)  handleDmEdit_(ss, sheet, row, col, value);
    else if (sName === TCO.SHEET.ACC) handleAccEdit_(ss, sheet, row, col, value);
    else if (sName === TCO.SHEET.RTO) handleRtoEdit_(ss, sheet, row, col, value);

  } catch(err) {
    Logger.log('onEdit ERROR: ' + err.message + ' | R:' + e.range.getRow() + ' C:' + e.range.getColumn());
  }
}


// ─────────────────────────────────────────────────────────────
// ── DM HANDLER ───────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function handleDmEdit_(ss, sheet, row, col, value) {
  var dm = TCO.DM;

  // ── DATE or BUYER NAME entered/deleted ──────────────────
  if (col === dm.DM_DATE || col === dm.BUYER) {

    var dmNo  = sheet.getRange(row, dm.DM_NO).getValue();
    var date  = sheet.getRange(row, dm.DM_DATE).getValue();
    var buyer = String(sheet.getRange(row, dm.BUYER).getValue()).trim();

    // Admin has not filled DM NO — skip
    if (!dmNo) return;

    // Both deleted → find by DM NO + delete row in ACC/RTO/MASTER
    if (!date && !buyer) {
      clearSyncedRow_(ss, sheet, row);
      return;
    }

    // Generate MONTH when date is entered
    if (col === dm.DM_DATE && date) {
      generateMonth_(sheet, row, date);
    }

    // Both present → initial sync (only if not already synced)
    if (date && buyer) {
      var accSheet    = ss.getSheetByName(TCO.SHEET.ACC);
      var alreadyDone = findRowByDmNo_(accSheet, String(dmNo).trim(), TCO.ACC.DM_NO);

      if (!alreadyDone) {
        doInitialSync_(ss, sheet, row);        // First time — sorted insert
      } else {
        syncField_(ss, sheet, row, col, value, 'DM');  // Already synced — field-level
      }
      return;
    }

    // Only one present — just sync that field
    syncField_(ss, sheet, row, col, value, 'DM');
    return;
  }

  // ── EMP ID → lookup EXEC, BRANCH, TL ──────────────────
  if (col === dm.EMP_ID) {
    var empId = String(value).trim();
    if (!empId) {
      clearEmpFields_(ss, sheet, row);
    } else {
      fillExecDetails_(ss, sheet, row, empId);
    }
    return;
  }

  // ── CONTACT DLR → lookup DEALER NAME, AUTH, LOCATION ───
  if (col === dm.CONTACT_DLR) {
    var contact = String(sheet.getRange(row, dm.CONTACT_DLR).getValue()).trim();
    if (!contact) {
      clearDlrFields_(ss, sheet, row);
    } else {
      fillDealerDetails_(ss, sheet, row, contact);
    }
    return;
  }

  // ── ROI % format ────────────────────────────────────────
  if (col === dm.ROI1 || col === dm.ROI2) {
    sheet.getRange(row, col).setNumberFormat('0.00" %"');
  }

  // ── Currency format ─────────────────────────────────────
  if ((CURRENCY_COLS.DM || []).indexOf(col) !== -1 && value !== '') {
    sheet.getRange(row, col).setNumberFormat('₹ ##,##,##0');
  }

  // ── General field sync (PDF mapping) ────────────────────
  syncField_(ss, sheet, row, col, value, 'DM');

  // ── Auto-Generate RTO CODE when REG NO updates from DM ──
  if (col === dm.REG_NO && value) {
    var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
    var dmNo = sheet.getRange(row, dm.DM_NO).getValue();
    var rtoRow = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);
    if (rtoSheet && rtoRow) {
      var code = generateRtoCode_(value);
      if (code) {
        styleCell_(rtoSheet.getRange(rtoRow, TCO.RTO.RTO_CODE).setValue(code), TCO.STYLE.AUTO_GEN);
      }
    }
  }
}


// ─────────────────────────────────────────────────────────────
// ── ACC HANDLER ──────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function handleAccEdit_(ss, sheet, row, col, value) {
  var acc = TCO.ACC;

  // ── RTO PAID AMOUNT → auto calculate RTO PROFIT ─────────
  if (col === acc.RTO_PAID) {
    var charges = parseFloat(sheet.getRange(row, acc.RTO_CHARGES).getValue()) || 0;
    var paid    = parseFloat(value) || 0;
    var profit  = charges - paid;
    sheet.getRange(row, acc.RTO_PROFIT)
      .setValue(profit)
      .setNumberFormat('₹ ##,##,##0')
      .setFontColor(profit >= 0 ? '#10b981' : '#ef4444')
      .setFontWeight('bold')
      .setFontStyle('italic')
      .setFontSize(10);
    return;
  }

  // ── Currency format ─────────────────────────────────────
  if ((CURRENCY_COLS.ACC || []).indexOf(col) !== -1 && value !== '') {
    sheet.getRange(row, col).setNumberFormat('₹ ##,##,##0');
  }

  // ── Update RTO PENDING DAYS when DISB RECV DATE is entered ──
  if (col === acc.DISB_RECV_DATE) {
    var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
    var dmNoForRto = sheet.getRange(row, acc.DM_NO).getValue();
    var rtoRowForUpdate = findRowByDmNo_(rtoSheet, dmNoForRto, TCO.RTO.DM_NO);
    if (rtoSheet && rtoRowForUpdate) {
      updatePendingDays_(rtoSheet, rtoRowForUpdate);
    }
  }

  // ── General field sync (PDF mapping) ────────────────────
  syncField_(ss, sheet, row, col, value, 'ACC');
}


// ─────────────────────────────────────────────────────────────
// ── OPTIMIZED PAYMENT TRACKER COLOR FORMATTING ───────────────
// ─────────────────────────────────────────────────────────────
  function applyPaymentColors_(sheet, row) {
  try {
    // डेटा रीड करें
    var valL = sheet.getRange(row, 12).getValue(); // Col L (DISB RECV DATE)
    var valQ = sheet.getRange(row, 17).getValue(); // Col Q (PAYER 1 NAME)
    
    var rangeAll = sheet.getRange(row, 17, 1, 12); // Q to AB range

    // 1. अगर Q (Payer Name) खाली है, तो कलर हटा दें
    if (!valQ || valQ.toString().trim() === "") {
      rangeAll.setBackground('#ffffff');
      return;
    }

    // 2. RULE 1: अगर L (Date) खाली है और Q में डेटा है -> पूरा Blue (#cfe2f3)
    if (!valL || valL.toString().trim() === "") {
      rangeAll.setBackground('#cfe2f3'); 
      return;
    }

    // 3. RULE 2: अगर L में डेट है और Q में डेटा है
    // Q to T -> Green (#d9ead3)
    sheet.getRange(row, 17, 1, 4).setBackground('#d9ead3');

    // U to X Logic -> FULL होने पर Green, नहीं तो Yellow (#fff2cc)
    var valU = sheet.getRange(row, 21).getValue();
    var strU = valU ? valU.toString().toUpperCase().trim() : "";
    var colorU = (strU === 'FULL') ? '#d9ead3' : '#fff2cc';
    sheet.getRange(row, 21, 1, 4).setBackground(colorU);

    // Y to AB Logic -> FULL होने पर Green, नहीं तो Yellow (#fff2cc)
    var valY = sheet.getRange(row, 25).getValue();
    var strY = valY ? valY.toString().toUpperCase().trim() : "";
    var colorY = (strY === 'FULL') ? '#d9ead3' : '#fff2cc';
    sheet.getRange(row, 25, 1, 4).setBackground(colorY);

  } catch (e) {
    Logger.log("Color Error: " + e.message);
  }
}


// ─────────────────────────────────────────────────────────────
// ── RTO HANDLER ──────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function handleRtoEdit_(ss, sheet, row, col, value) {
  var rto = TCO.RTO;

  // ── REG NO → generate RTO CODE ──────────────────────────
  if (col === rto.REG_NO && value) {
    var code = generateRtoCode_(value);
    if (code) {
      styleCell_(
        sheet.getRange(row, rto.RTO_CODE).setValue(code),
        TCO.STYLE.AUTO_GEN
      );
    }
    sheet.getRange(row, col)
      .setFontColor(TCO.STYLE.REG_NO_ST.color)
      .setFontSize(TCO.STYLE.REG_NO_ST.size);
  }

  // ── RC STATUS or RC DATE → update PENDING DAYS ──────────
  if (col === rto.RC_STATUS || col === rto.RC_DATE) {
    updatePendingDays_(sheet, row);
  }

  // ── General field sync (PDF mapping) ────────────────────
  syncField_(ss, sheet, row, col, value, 'RTO');
}


// ─────────────────────────────────────────────────────────────
// ── INITIAL SYNC — Sorted Insert in ACC + RTO ────────────────
// ─────────────────────────────────────────────────────────────

function doInitialSync_(ss, dmSheet, dmRow) {
  return withScriptLock_(function() {
    // --- paste the full original body of doInitialSync_ here ---
    // var d = dmSheet.getRange(dmRow, 1, 1, 55).getValues()[0];
    // ... rest of logic (insertRowBefore, setValues, applyRowStyle_, autoSortAllSheets, आदि) ...
  }, 30000);
}


  function doInitialSync_(ss, dmSheet, dmRow) {
  return withScriptLock_(function() {
    var d   = dmSheet.getRange(dmRow, 1, 1, 55).getValues()[0];
    var acc = ss.getSheetByName(TCO.SHEET.ACC);
    var rto = ss.getSheetByName(TCO.SHEET.RTO);
    if (!acc || !rto) return;

    var dm      = TCO.DM;
    var A       = TCO.ACC;
    var R       = TCO.RTO;
    var newDmNo = String(d[dm.DM_NO - 1]).trim();

    // ── Find sorted insert position in ACC ──
    var accInsert = findSortedInsertRow_(acc, newDmNo, A.DM_NO);
    acc.insertRowBefore(accInsert);
    var accData = new Array(43).fill('');
    accData[A.DM_NO-1]         = d[dm.DM_NO-1];
    accData[A.BUYER-1]         = toUpper_(d[dm.BUYER-1]);
    accData[A.PHONE1-1]        = d[dm.PHONE1-1];
    accData[A.PHONE2-1]        = d[dm.PHONE2-1];
    accData[A.REG_NO-1]        = toUpper_(d[dm.REG_NO-1]);
    accData[A.MODEL-1]         = toUpper_(d[dm.MODEL-1]);
    accData[A.MFG_YEAR-1]      = d[dm.MFG_YEAR-1];
    accData[A.BANK-1]          = toUpper_(d[dm.BANK-1]);
    accData[A.PRODUCT-1]       = toUpper_(d[dm.PRODUCT-1]);
    accData[A.DEALER-1]        = toUpper_(d[dm.DEALER_NAME-1]);
    accData[A.EXEC-1]          = toUpper_(d[dm.EXEC_NAME-1]);
    // X (24) और AH (34) का डेटा Account में जाएगा
    accData[A.DISB1 - 1]       = d[dm.DISB_AMT1 - 1] || '';
    accData[A.DISB2 - 1]       = d[dm.DISB_AMT2 - 1] || '';
    accData[A.DEALER_PAY_ST-1] = 'PENDING';
    accData[A.RTO_CHARGES-1]   = d[dm.RTO_CHARGES-1] || '';
    acc.getRange(accInsert, 1, 1, 43).setValues([accData]);
    applyRowStyle_(acc, accInsert, 'ACC');

    // ── Find sorted insert position in RTO ──
    var rtoInsert = findSortedInsertRow_(rto, newDmNo, R.DM_NO);
    rto.insertRowBefore(rtoInsert);
    var rtoData   = new Array(30).fill('');
    var rtoCode   = generateRtoCode_(d[dm.REG_NO-1]);
    rtoData[R.DM_DATE-1]       = d[dm.DM_DATE-1];
    rtoData[R.DM_NO-1]         = d[dm.DM_NO-1];
    rtoData[R.DEALER_PAY_ST-1] = 'PENDING';
    rtoData[R.PRODUCT-1]       = toUpper_(d[dm.PRODUCT-1]);
    rtoData[R.BRANCH-1]        = toUpper_(d[dm.BRANCH-1]);
    rtoData[R.EXEC-1]          = toUpper_(d[dm.EXEC_NAME-1]);
    rtoData[R.DEALER-1]        = toUpper_(d[dm.DEALER_NAME-1]);
    rtoData[R.REG_NO-1]        = toUpper_(d[dm.REG_NO-1]);
    rtoData[R.RTO_CODE-1]      = rtoCode;
    rtoData[R.MODEL-1]         = toUpper_(d[dm.MODEL-1]);
    rtoData[R.MFG_YEAR-1]      = d[dm.MFG_YEAR-1];
    rtoData[R.OWNER_SELLER-1]  = toUpper_(d[dm.VEHICLE_OWNER-1]);
    rtoData[R.BUYER-1]         = toUpper_(d[dm.BUYER-1]);
    rtoData[R.BUYER_PHONE-1]   = d[dm.PHONE1-1];
    rtoData[R.BANK-1]          = toUpper_(d[dm.BANK-1]);
    rtoData[R.RC_STATUS-1]     = 'PENDING';
    rto.getRange(rtoInsert, 1, 1, 30).setValues([rtoData]);
    applyRowStyle_(rto, rtoInsert, 'RTO');
    autoSortAllSheets();
  }, 30000);
}

// पहले (original)
// function clearSyncedRow_(ss, dmSheet, dmRow) { ... deleteRow(...) ... }

// बाद में (use wrapper)
function clearSyncedRow_(ss, dmSheet, dmRow) {
  return withScriptLock_(function() {
    // original clearSyncedRow_ body यहाँ लगाएँ (बिना अतिरिक्त lock)
    var dmNo = String(dmSheet.getRange(dmRow, TCO.DM.DM_NO).getValue()).trim();
    if (!dmNo) return;

    var acc = ss.getSheetByName(TCO.SHEET.ACC);
    var rto = ss.getSheetByName(TCO.SHEET.RTO);
    var mst = ss.getSheetByName(TCO.SHEET.MASTER);

    var rtoRow = rto ? findRowByDmNo_(rto, dmNo, TCO.RTO.DM_NO) : null;
    var accRow = acc ? findRowByDmNo_(acc, dmNo, TCO.ACC.DM_NO) : null;
    var mstRow = mst ? findRowByDmNo_(mst, dmNo, 1)             : null;

    if (rtoRow) rto.deleteRow(rtoRow);
    if (accRow) acc.deleteRow(accRow);
    if (mstRow) mst.deleteRow(mstRow);
  });
}





// यह फंक्शन ACCOUNT और RTO शीट को DM NO के हिसाब से सॉर्ट करेगा
function autoSortAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Account Sheet Sort (Column 1 - DM NO)
  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  if (accSheet) {
    var lrAcc = accSheet.getLastRow();
    if (lrAcc >= TCO.DATA_ROW) {
      accSheet.getRange(TCO.DATA_ROW, 1, lrAcc - TCO.DATA_ROW + 1, accSheet.getLastColumn())
        .sort({column: TCO.ACC.DM_NO, ascending: true});
    }
  }

  // 2. RTO Sheet Sort (Column 2 - DM NO)
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  if (rtoSheet) {
    var lrRto = rtoSheet.getLastRow();
    if (lrRto >= TCO.DATA_ROW) {
      rtoSheet.getRange(TCO.DATA_ROW, 1, lrRto - TCO.DATA_ROW + 1, rtoSheet.getLastColumn())
        .sort({column: TCO.RTO.DM_NO, ascending: true});
    }
  }
}


// ─────────────────────────────────────────────────────────────
// ── CLEAR SYNCED ROW — DATE + BUYER both deleted ─────────────
// ─────────────────────────────────────────────────────────────
function clearSyncedRow_(ss, dmSheet, dmRow) {
  var dmNo = String(dmSheet.getRange(dmRow, TCO.DM.DM_NO).getValue()).trim();
  if (!dmNo) return;

  var acc = ss.getSheetByName(TCO.SHEET.ACC);
  var rto = ss.getSheetByName(TCO.SHEET.RTO);
  var mst = ss.getSheetByName(TCO.SHEET.MASTER);

  // Delete rows by DM NO (delete RTO first, then ACC, then MASTER — no row shift conflict)
  var rtoRow = rto ? findRowByDmNo_(rto, dmNo, TCO.RTO.DM_NO) : null;
  var accRow = acc ? findRowByDmNo_(acc, dmNo, TCO.ACC.DM_NO) : null;
  var mstRow = mst ? findRowByDmNo_(mst, dmNo, 1)             : null;

  if (rtoRow) rto.deleteRow(rtoRow);
  if (accRow) acc.deleteRow(accRow);
  if (mstRow) mst.deleteRow(mstRow);
}


// ─────────────────────────────────────────────────────────────
// ── GENERIC FIELD SYNC — DM NO based ─────────────────────────
// ─────────────────────────────────────────────────────────────

function syncField_(ss, sourceSheet, row, col, value, key) {
  var destinations = (SYNC_MAP[key] || {})[col];
  if (!destinations || destinations.length === 0) return;

  // Get DM NO from source row
  var dmNoCol = key === 'DM'  ? TCO.DM.DM_NO  :
                key === 'ACC' ? TCO.ACC.DM_NO  : TCO.RTO.DM_NO;
  var dmNo = String(sourceSheet.getRange(row, dmNoCol).getValue()).trim();
  if (!dmNo) return;

  var sheetNames = { DM: TCO.SHEET.DM, ACC: TCO.SHEET.ACC, RTO: TCO.SHEET.RTO };
  var dmNoCols   = { DM: TCO.DM.DM_NO, ACC: TCO.ACC.DM_NO, RTO: TCO.RTO.DM_NO };
  var val        = (typeof value === 'string') ? toUpper_(value) : value;

  destinations.forEach(function(dest) {
    var sh      = ss.getSheetByName(sheetNames[dest.s]);
    if (!sh) return;
    var destRow = findRowByDmNo_(sh, dmNo, dmNoCols[dest.s]);
    if (!destRow) return;
    styleCell_(sh.getRange(destRow, dest.c).setValue(val), TCO.STYLE.AUTO_PULL);
  });

  // Earlier: updateMasterData_(ss);
  // Now: mark master as dirty so a time-based trigger does the heavy update in batch
  try {
    markMasterDirty_();
  } catch (e) {
    Logger.log('syncField_ markMasterDirty_ err: ' + (e && e.message ? e.message : e));
  }
}

// ── Find row by DM NO ─────────────────────────────────────────
function findRowByDmNo_(sheet, dmNo, dmNoCol) {
  if (!sheet || !dmNo) return null;
  var data  = sheet.getDataRange().getValues();
  var clean = String(dmNo).trim();
  for (var i = TCO.DATA_ROW - 1; i < data.length; i++) {
    if (String(data[i][dmNoCol - 1]).trim() === clean) return i + 1;
  }
  return null;
}

// ── Find sorted insert position by DM NO ──────────────────────
function findSortedInsertRow_(sheet, newDmNo, dmNoCol) {
  var data    = sheet.getDataRange().getValues();
  var lastRow = TCO.DATA_ROW;

  for (var i = TCO.DATA_ROW - 1; i < data.length; i++) {
    var existing = String(data[i][dmNoCol - 1]).trim();
    if (!existing) continue;
    if (newDmNo < existing) return i + 1;  // Insert before this row
    lastRow = i + 2;                        // Move marker after this row
  }
  return lastRow;  // Append at end
}


// ─────────────────────────────────────────────────────────────
// ── LOOKUP: EMP ID → EXEC NAME, BRANCH, TEAM LEADER ──────────
// ─────────────────────────────────────────────────────────────
function fillExecDetails_(ss, sheet, row, empId) {
  var empSheet = ss.getSheetByName(TCO.SHEET.EMP);
  if (!empSheet) return;

  var data  = empSheet.getDataRange().getValues();
  var found = false;

  // DEALER_MASTER data starts from Row 2 (index 2 = row 3, after version + header)
  for (var i = 2; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() !== empId.toUpperCase()) continue;

    var name   = toUpper_(data[i][1]);  // EMP_NAME   — col 2
    var branch = toUpper_(data[i][7]);  // BRANCH_CODE — col 8
    var tl     = toUpper_(data[i][8]);  // TEAM_LEADER — col 9

    // Fill DM AUTO fields
    styleCell_(sheet.getRange(row, TCO.DM.EXEC_NAME).setValue(name),    TCO.STYLE.AUTO_PULL);
    styleCell_(sheet.getRange(row, TCO.DM.BRANCH).setValue(branch),     TCO.STYLE.AUTO_PULL);
    styleCell_(sheet.getRange(row, TCO.DM.TEAM_LEADER).setValue(tl),    TCO.STYLE.AUTO_PULL);

    // Sync EXEC NAME + BRANCH to ACC + RTO
    var dmNo2    = String(sheet.getRange(row, TCO.DM.DM_NO).getValue()).trim();
    var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
    var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
    var accRow2  = findRowByDmNo_(accSheet, dmNo2, TCO.ACC.DM_NO);
    var rtoRow2  = findRowByDmNo_(rtoSheet, dmNo2, TCO.RTO.DM_NO);
    if (accSheet && accRow2) styleCell_(accSheet.getRange(accRow2, TCO.ACC.EXEC).setValue(name),    TCO.STYLE.AUTO_PULL);
    if (rtoSheet && rtoRow2) {
      styleCell_(rtoSheet.getRange(rtoRow2, TCO.RTO.EXEC).setValue(name),     TCO.STYLE.AUTO_PULL);
      styleCell_(rtoSheet.getRange(rtoRow2, TCO.RTO.BRANCH).setValue(branch), TCO.STYLE.AUTO_PULL);
    }
    found = true;
    break;
  }

  if (!found) {
    [TCO.DM.EXEC_NAME, TCO.DM.BRANCH, TCO.DM.TEAM_LEADER].forEach(function(c) {
      styleCell_(sheet.getRange(row, c).setValue('EMP NOT FOUND'), TCO.STYLE.NOT_FOUND);
    });
  }
}

function clearEmpFields_(ss, sheet, row) {
  // DM के उस row के EXEC_NAME, BRANCH, TEAM_LEADER क्लियर करें
  [TCO.DM.EXEC_NAME, TCO.DM.BRANCH, TCO.DM.TEAM_LEADER].forEach(function(c) {
    sheet.getRange(row, c).setValue('').setFontColor('#000000').setFontStyle('normal');
  });

  // सही ACC / RTO row खोजने के लिए DM_NO लें
  var dmNo = String(sheet.getRange(row, TCO.DM.DM_NO).getValue()).trim();
  if (!dmNo) return;

  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);

  var accRow = accSheet ? findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO) : null;
  var rtoRow = rtoSheet ? findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO) : null;

  if (accSheet && accRow) accSheet.getRange(accRow, TCO.ACC.EXEC).setValue('').setFontStyle('normal');
  if (rtoSheet && rtoRow) {
    rtoSheet.getRange(rtoRow, TCO.RTO.EXEC).setValue('').setFontStyle('normal');
    rtoSheet.getRange(rtoRow, TCO.RTO.BRANCH).setValue('').setFontStyle('normal');
  }
}


// ─────────────────────────────────────────────────────────────
// ── LOOKUP: CONTACT DLR → DEALER NAME, AUTH, LOCATION ────────
// ─────────────────────────────────────────────────────────────
function fillDealerDetails_(ss, sheet, row, contactNo) {
  var dlrSheet = ss.getSheetByName(TCO.SHEET.DEALER);
  if (!dlrSheet) return;

  var data  = dlrSheet.getDataRange().getValues();
  var found = false;

  // Data starts from Row 2 (index 2 = row 3, after version + header)
  for (var i = 2; i < data.length; i++) {
    if (String(data[i][5]).trim() !== String(contactNo).trim()) continue;  // CONTACT NO — col 6

    var dlr  = toUpper_(data[i][1]);  // DEALER NAME    — col 2
    var auth = toUpper_(data[i][4]);  // CONTACT PERSON — col 5
    var loc  = toUpper_(data[i][2]);  // LOCATION       — col 3

    // Fill DM AUTO fields
    styleCell_(sheet.getRange(row, TCO.DM.DEALER_NAME).setValue(dlr),   TCO.STYLE.AUTO_PULL);
    styleCell_(sheet.getRange(row, TCO.DM.AUTH_PERSON).setValue(auth),  TCO.STYLE.AUTO_PULL);
    styleCell_(sheet.getRange(row, TCO.DM.LOCATION).setValue(loc),      TCO.STYLE.AUTO_PULL);

    // Sync DEALER NAME to ACC + RTO (DM NO match karke sahi row par bhejna)
    var dmNo = String(sheet.getRange(row, TCO.DM.DM_NO).getValue()).trim();
    var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
    var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
    
    var accRow = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
    var rtoRow = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);

    if (accSheet && accRow) styleCell_(accSheet.getRange(accRow, TCO.ACC.DEALER).setValue(dlr), TCO.STYLE.AUTO_PULL);
    if (rtoSheet && rtoRow) styleCell_(rtoSheet.getRange(rtoRow, TCO.RTO.DEALER).setValue(dlr), TCO.STYLE.AUTO_PULL);

    found = true;
    break;
  }

  if (!found) {
    [TCO.DM.DEALER_NAME, TCO.DM.AUTH_PERSON, TCO.DM.LOCATION].forEach(function(c) {
      styleCell_(sheet.getRange(row, c).setValue('DLR NOT FOUND'), TCO.STYLE.NOT_FOUND);
    });
  }
}

function clearDlrFields_(ss, sheet, row) {
  [TCO.DM.DEALER_NAME, TCO.DM.AUTH_PERSON, TCO.DM.LOCATION].forEach(function(c) {
    sheet.getRange(row, c).setValue('').setFontColor('#000000').setFontStyle('normal');
  });
  
  // DM NO match karke ACC/RTO se data clear karna
  var dmNo = String(sheet.getRange(row, TCO.DM.DM_NO).getValue()).trim();
  var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  
  var accRow = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
  var rtoRow = findRowByDmNo_(rtoSheet, dmNo, TCO.RTO.DM_NO);

  if (accSheet && accRow) accSheet.getRange(accRow, TCO.ACC.DEALER).setValue('').setFontStyle('normal');
  if (rtoSheet && rtoRow) rtoSheet.getRange(rtoRow, TCO.RTO.DEALER).setValue('').setFontStyle('normal');
}


// ─────────────────────────────────────────────────────────────
// ── MASTER DATA UPDATE — DM NO based ─────────────────────────
// ─────────────────────────────────────────────────────────────
function markMasterDirty_() {
  try {
    PropertiesService.getScriptProperties().setProperty('MASTER_NEEDS_UPDATE', new Date().toISOString());
  } catch (e) {
    Logger.log('markMasterDirty_ ERROR: ' + (e && e.message ? e.message : e));
  }
}

    // ── MASTER Field Map — 97 cols (PDF v3 alphabetical order) ──
    // Each entry: [sheet_key, col_number (1-based)]
    // sheet_key: 'D'=DM, 'A'=ACC, 'R'=RTO
    var FMAP = [
      ['D',1 ], // 1  DM NO
      ['D',2 ], // 2  DM DATE
      ['D',44], // 3  AUTH PERSON (DLR)
      ['D',12], // 4  BANK
      ['D',41], // 5  BRANCH
      ['D',5 ], // 6  BUYER NAME
      ['R',18], // 7  BUYER PHONE
      ['R',9 ], // 8  CASE REC DATE
      ['R',20], // 9  CASE TYPE
      ['R',13], // 10 CHASSIS NO
      ['D',45], // 11 CONTACT NO (DLR)
      ['D',43], // 12 DEALERSHIP NAME
      ['A',29], // 13 DEALERSHIP PAYMENT STATUS
      ['A',15], // 14 DISB AMT 1
      ['A',16], // 15 DISB AMT 2
      ['A',14], // 16 DISB AMT RECEIVED
      ['A',12], // 17 DISB AMT RECEIVING DATE
      ['D',37], // 18 e-Challan AMT (If Any)
      ['D',8 ], // 19 EMAIL
      ['D',22], // 20 EMI 1
      ['D',32], // 21 EMI 2
      ['D',39], // 22 EMP ID
      ['R',14], // 23 ENGINE NO
      ['D',53], // 24 EXE INCENTIVE
      ['A',33], // 25 EXECUTIVE INCENTIVE
      ['D',40], // 26 EXECUTIVE NAME
      ['D',20], // 27 FILE CHG
      ['D',30], // 28 FILE CHG 2
      ['A',31], // 29 HOLD AMT FROM BANK
      ['A',32], // 30 HOLD AMT FROM TCO
      ['D',16], // 31 INS 1
      ['D',26], // 32 INS 2
      ['D',15], // 33 LOAN APP 1
      ['D',25], // 34 LOAN APP 2
      ['D',14], // 35 LOAN APPLIED
      ['A',34], // 36 LOAN SCORE
      ['D',23], // 37 LOAN TENURE 1
      ['D',33], // 38 LOAN TENURE 2
      ['D',46], // 39 LOCATION
      ['D',11], // 40 MFG YEAR
      ['D',10], // 41 MODEL
      ['D',3 ], // 42 MONTH
      ['A',36], // 43 NET SCORE
      ['R',27], // 44 ORIGINAL RC REC STATUS
      ['R',15], // 45 OWNER NAME (SELLER)
      ['A',20], // 46 PAYER 1 ACC DETAILS
      ['A',19], // 47 PAYER 1 AMOUNT
      ['A',18], // 48 PAYER 1 DATE
      ['A',17], // 49 PAYER 1 NAME
      ['A',24], // 50 PAYER 2 ACC DETAILS
      ['A',23], // 51 PAYER 2 AMOUNT
      ['A',22], // 52 PAYER 2 DATE
      ['A',21], // 53 PAYER 2 NAME
      ['A',28], // 54 PAYER 3 ACC DETAILS
      ['A',27], // 55 PAYER 3 AMOUNT
      ['A',26], // 56 PAYER 3 DATE
      ['A',25], // 57 PAYER 3 NAME
      ['D',36], // 58 Payment to DLR (as per DM)
      ['D',49], // 59 PAYOUT CD
      ['A',35], // 60 PAYOUT FROM BANK
      ['D',50], // 61 PAYOUT SCORE
      ['D',48], // 62 PAYOUT TCO
      ['A',30], // 63 PAYOUT TO DEALER
      ['R',26], // 64 PENDING DAYS
      ['D',6 ], // 65 PHONE 1
      ['D',7 ], // 66 PHONE 2
      ['D',13], // 67 PRODUCT
      ['R',25], // 68 RC TRANSFER DATE
      ['R',24], // 69 RC TRANSFER STATUS
      ['D',9 ], // 70 REG NO
      ['D',55], // 71 REMARKS
      ['A',43], // 72 REMARKS (IF ANY)
      ['R',30], // 73 REMARKS / ISSUE
      ['D',21], // 74 ROI 1
      ['D',31], // 75 ROI_2
      ['D',35], // 76 RTO CHARGES
      ['R',10], // 77 RTO CODE
      ['A',39], // 78 RTO PAID AMOUNT
      ['A',40], // 79 RTO PAYMENT DATE
      ['A',42], // 80 RTO PROFIT
      ['R',23], // 81 RTO RECEIPT NO
      ['R',21], // 82 RTO SCAN FILE (LINK)
      ['D',52], // 83 RTO SCORE
      ['R',22], // 84 RTO VENDOR
      ['A',38], // 85 RTO VENDOR NAME
      ['R',16], // 86 SELLER PHONE
      ['D',17], // 87 SURAKSHA 1
      ['D',27], // 88 SURAKSHA 2
      ['D',18], // 89 SURAKSHA TENURE 1
      ['D',28], // 90 SURAKSHA TENURE 2
      ['R',29], // 91 SYSTEM REMARKS
      ['D',42], // 92 TEAM LEADER
      ['D',19], // 93 TOTAL LOAN 1
      ['D',29], // 94 TOTAL LOAN 2
      ['R',28], // 95 TRANSFERRED RC COPY_PROOF (LINK)
      ['A',13], // 96 UTR NO
      ['D',47]  // 97 VEHICLE OWNER NAME
    ];

    // ── Build rows ──
    var rows = [];
    for (var i = TCO.DATA_ROW - 1; i < dmData.length; i++) {
      var dmNo = String(dmData[i][0]).trim();
      if (!dmNo) continue;

      var dRow = dmData[i];
      var aRow = accMap[dmNo] || [];
      var rRow = rtoMap[dmNo] || [];
      var row  = [];

      for (var f = 0; f < FMAP.length; f++) {
        var src = FMAP[f][0];
        var col = FMAP[f][1] - 1;  // 0-based index
        var val = '';
        if      (src === 'D') val = dRow[col] !== undefined ? dRow[col] : '';
        else if (src === 'A') val = aRow[col] !== undefined ? aRow[col] : '';
        else if (src === 'R') val = rRow[col] !== undefined ? rRow[col] : '';
        row.push(val);
      }
      rows.push(row);
    }

    // ── Clear + rewrite MASTER ──
    var lr = ms.getLastRow();
    if (lr >= TCO.DATA_ROW) {
      ms.getRange(TCO.DATA_ROW, 1, lr - TCO.DATA_ROW + 1, 97).clearContent();
    }
    if (rows.length > 0) {
      ms.getRange(TCO.DATA_ROW, 1, rows.length, 97).setValues(rows);
    }
  } catch(err) {
    Logger.log('updateMasterData_ ERROR: ' + err.message);
  }
}


// ─────────────────────────────────────────────────────────────
// ── STYLING ──────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
function styleCell_(cell, style) {
  return cell
    .setFontColor(style.color)
    .setFontStyle(style.style)
    .setFontSize(style.size || 10);
}

function applyRowStyle_(sheet, row, key) {
  var maxCols = key === 'DM' ? 55 : key === 'ACC' ? 43 : 30;
  sheet.getRange(row, 1, 1, maxCols)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor('#000000')
    .setFontStyle('normal')
    .setFontSize(10)
    .setFontFamily('Arial')
    .setBorder(
      true, true, true, true, true, true,
      '#CBD5E1',
      SpreadsheetApp.BorderStyle.SOLID
    );
}

function applyValidation_(sheet, key, row, col, value) {
  var cell   = sheet.getRange(row, col);
  var strVal = (value !== null && value !== undefined) ? String(value).trim() : '';

  // Phone — red bg if not 10 digits
  if ((PHONE_COLS[key] || []).indexOf(col) !== -1) {
    if (!strVal) { cell.setBackground(null); return; }
    cell.setBackground(strVal.replace(/\D/g,'').length !== 10 ? '#fca5a5' : null);
    return;
  }

  // Currency — red if text in currency col
  if ((CURRENCY_COLS[key] || []).indexOf(col) !== -1) {
    if (!strVal) { cell.setBackground(null); return; }
    if (typeof value === 'string' && !/^\d+(\.\d+)?$/.test(strVal)) {
      cell.setBackground('#fca5a5');
    } else {
      cell.setBackground(null);
      if (value !== '') cell.setNumberFormat('₹ ##,##,##0');
    }
    return;
  }

  // Email — lowercase + yellow if no @
  if (key === 'DM' && col === TCO.DM.EMAIL && strVal) {
    cell.setValue(strVal.toLowerCase()).setFontLine('none');
    cell.setBackground(strVal.indexOf('@') === -1 ? '#fef08a' : null);
    return;
  }

  // Word cols — CAPS + yellow if has number
  if ((WORD_COLS[key] || []).indexOf(col) !== -1 && strVal) {
    cell.setValue(strVal.toUpperCase());
    cell.setBackground(/\d/.test(strVal) ? '#fef08a' : null);
    return;
  }

  // Skip CAPS cols (links)
  if ((SKIP_CAPS[key] || []).indexOf(col) !== -1) return;

  // General — CAPS all text
  if (typeof value === 'string' && strVal &&
      strVal.indexOf('http') !== 0 &&
      strVal.indexOf('www')  !== 0) {
    cell.setValue(strVal.toUpperCase());
  }
}


// ─────────────────────────────────────────────────────────────
// ── UTILITY FUNCTIONS ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────

function withScriptLock_(fn, waitMs) {
  waitMs = waitMs || 30000; // default 30s
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(waitMs);
    return fn();
  } catch (e) {
    Logger.log('withScriptLock_ error: ' + (e && e.message ? e.message : e));
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (er) {}
  }
}

// MONTH auto-gen from DM DATE
function generateMonth_(sheet, row, dmDate) {
  try {
    var d  = new Date(dmDate);
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    styleCell_(
      sheet.getRange(row, TCO.DM.MONTH).setValue(mm + '-' + d.getFullYear()),
      TCO.STYLE.AUTO_GEN
    );
  } catch(e) { Logger.log('generateMonth_ err: ' + e.message); }
}

// RTO CODE from REG NO (e.g. DL 01 AB 1234 → DL01)
function generateRtoCode_(regNo) {
  if (!regNo) return '';
  // सारे स्पेस और स्पेशल कैरेक्टर पहले हटा दें
  var cleanReg = String(regNo).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  var m = cleanReg.match(/^([A-Z]{2})(\d{1,2})/);
  if (!m) return '';
  return m[1] + String(parseInt(m[2], 10)).padStart(2, '0');
}

// Safe UPPERCASE
function toUpper_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') return val;
  if (val instanceof Date) return val;
  return String(val).toUpperCase().trim();
}


// ─────────────────────────────────────────────────────────────
// ── PENDING DAYS CALCULATION (AUTO) ──────────────────────────
// ─────────────────────────────────────────────────────────────

function updatePendingDays_(rtoSheet, row) {
  try {
    // rtoSheet से उसी स्प्रेडशीट का संदर्भ लें (trigger / openById दोनों के लिए safe)
    var ss = (rtoSheet && typeof rtoSheet.getParent === 'function') ?
             rtoSheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || !rtoSheet) return;

    var accSheet = ss.getSheetByName(TCO.SHEET.ACC);
    if (!accSheet) return;

    var dmNo = rtoSheet.getRange(row, TCO.RTO.DM_NO).getValue();
    var accRow = findRowByDmNo_(accSheet, dmNo, TCO.ACC.DM_NO);
    if (!accRow) return;

    var startDate = accSheet.getRange(accRow, 12).getValue(); // Col L (Account Sheet)
    var endDate = rtoSheet.getRange(row, 25).getValue();      // Col Y (RTO Sheet)
    var cellZ = rtoSheet.getRange(row, 26);                   // Col Z (RTO Sheet)

    if (!startDate || startDate === "") {
      cellZ.setValue("WAITING FOR DISB").setFontColor("#ef4444").setFontWeight("bold");
      return;
    }

    var start = new Date(startDate);
    var end = (endDate && endDate !== "") ? new Date(endDate) : new Date();

    // Normalize times to midnight to count full days only
    start.setHours(0,0,0,0);
    end.setHours(0,0,0,0);

    var diffInTime = end.getTime() - start.getTime();
    var diffInDays = Math.floor(diffInTime / (1000 * 3600 * 24));

    cellZ.setValue(diffInDays).setFontWeight("bold");

    // Coloring: pending = red, done = green
    if (!endDate) {
      cellZ.setFontColor("#ef4444");
    } else {
      cellZ.setFontColor("#10b981");
    }
  } catch (e) {
    Logger.log('updatePendingDays_ ERROR: ' + (e.message || e));
  }
}

// ─────────────────────────────────────────────────────────────
// ── DAILY RTO REFRESH (12 AM TRIGGER) ────────────────────────
// ─────────────────────────────────────────────────────────────
function dailyRtoRefresh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rtoSheet = ss.getSheetByName(TCO.SHEET.RTO);
  if (!rtoSheet) return;
  
  var lastRow = rtoSheet.getLastRow();

  for (var i = TCO.DATA_ROW; i <= lastRow; i++) {
    updatePendingDays_(rtoSheet, i);
  }
}

// Runs daily — updates all pending RTO rows
function dailyPendingDaysUpdate() {
  var ss  = SpreadsheetApp.openById('16j_0Szvo3WAQrEMgTbJx8oqtkiWbBd3IvaoZ2hdftgI');
  var rto = ss.getSheetByName(TCO.SHEET.RTO);
  if (!rto) return;

  var lr = rto.getLastRow();
  for (var row = TCO.DATA_ROW; row <= lr; row++) {
    var dmNo = rto.getRange(row, TCO.RTO.DM_NO).getValue();
    if (!dmNo) continue;
    var rcSt = String(rto.getRange(row, TCO.RTO.RC_STATUS).getValue()).toUpperCase().trim();
    if (rcSt === 'DONE' || rcSt === 'COMPLETED') continue;
    updatePendingDays_(rto, row);
  }
}

function setupInstallableTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditTrigger') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet('16j_0Szvo3WAQrEMgTbJx8oqtkiWbBd3IvaoZ2hdftgI')
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('✅ Done!');
}

function setupMasterUpdateTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoUpdateMaster') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('autoUpdateMaster')
    .timeBased()
    .everyMinutes(5)
    .create();
  SpreadsheetApp.getUi().alert('✅ Master auto-update set — every 5 mins!');
}

function autoUpdateMaster() {
  var ss = SpreadsheetApp.openById('16j_0Szvo3WAQrEMgTbJx8oqtkiWbBd3IvaoZ2hdftgI');
  updateMasterData_(ss);
}


// ─────────────────────────────────────────────────────────────
// ── SUPER FAST BULK COLOR FIX (1 Second Execution) ───────────
// ─────────────────────────────────────────────────────────────
function forceFixAllColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TCO.SHEET.ACC);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < TCO.DATA_ROW) return;

  // एक साथ सारा डेटा रीड करना (ताकि शीट हैंग न हो)
  var valuesL = sheet.getRange(TCO.DATA_ROW, 12, lastRow - TCO.DATA_ROW + 1, 1).getValues();
  var rangeQ_AB = sheet.getRange(TCO.DATA_ROW, 17, lastRow - TCO.DATA_ROW + 1, 12);
  var valuesQ_AB = rangeQ_AB.getValues();
  
  var backgrounds = [];

  for (var i = 0; i < valuesQ_AB.length; i++) {
    var valL = valuesL[i][0];
    var valQ = valuesQ_AB[i][0];  // Q (Payer 1 Name)
    var valU = valuesQ_AB[i][4];  // U (Payer 1 ACC)
    var valY = valuesQ_AB[i][8];  // Y (Payer 2 ACC)

    var strL = valL ? String(valL).trim() : "";
    var strQ = valQ ? String(valQ).trim() : "";
    var strU = valU ? String(valU).toUpperCase().trim() : "";
    var strY = valY ? String(valY).toUpperCase().trim() : "";

    var cQ_T = '#ffffff', cU_X = '#ffffff', cY_AB = '#ffffff'; // Default White

    if (strQ !== "") {
      if (strL === "") {
        // Rule 1: L Blank है -> पूरा Blue
        cQ_T = '#cfe2f3'; cU_X = '#cfe2f3'; cY_AB = '#cfe2f3';
      } else {
        // Rule 2: L में Date है -> Green & Yellow (साथ में FULL कंडीशन)
        cQ_T = '#d9ead3'; 
        cU_X = (strU === 'FULL') ? '#d9ead3' : '#fff2cc'; 
        cY_AB = (strY === 'FULL') ? '#d9ead3' : '#fff2cc'; 
      }
    }

    backgrounds.push([
      cQ_T, cQ_T, cQ_T, cQ_T,     // Q, R, S, T
      cU_X, cU_X, cU_X, cU_X,     // U, V, W, X
      cY_AB, cY_AB, cY_AB, cY_AB  // Y, Z, AA, AB
    ]);
  }

  // पूरी शीट के कलर्स को एक ही सेकंड में ज़बरदस्ती (Force) अप्लाई करें
  rangeQ_AB.setBackgrounds(backgrounds);
  SpreadsheetApp.getUi().alert('✅ मैजिक! 1 सेकंड में पूरी शीट के कलर्स फिक्स कर दिए गए हैं।');
}

function markMasterDirty_() {
  try {
    PropertiesService.getScriptProperties().setProperty('MASTER_NEEDS_UPDATE', new Date().toISOString());
  } catch (e) {
    Logger.log('markMasterDirty_ err: ' + e.message);
  }
}























