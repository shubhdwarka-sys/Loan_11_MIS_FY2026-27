/**
 * ============================================================
 *  Loan_11_MIS_FY2026-27 — SHEET SETUP SCRIPT
 *  Version   : 3.0
 *  Mapping   : Mapping_v3.pdf + Final_Book.xlsx
 *  Author    : TCO Operations
 *  Date      : Auto-stamped on run
 * ============================================================
 *  SHEETS FROM PDF  : DM (55) | ACCOUNT (43) | RTO (30) | MASTER (97)
 *  SHEETS FROM EXCEL: DEALER_MASTER (10) | TCO_EMPLOYEE_MASTER (27)
 * ============================================================
 *  COLOR LEGEND:
 *   ⬜ MANUAL   = WHITE   → User Entry (Editable)
 *   🟩 DROPDOWN = GREEN   → Select from List (Editable)
 *   🟦 AUTO     = BLUE    → Formula Auto-Pull (LOCKED)
 *   🟨 GEN      = YELLOW  → Script Generated (LOCKED)
 * ============================================================
 */

// ─────────────────────────────────────────────
//  MAIN ENTRY POINT
// ─────────────────────────────────────────────
function setupTCOSystem() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy  HH:mm");

  // ── Sheet Names ──
  var SN = {
    DM     : "DM_DISBURSEMENT MEMO",
    ACC    : "ACCOUNT_PAYMENT_TRACKER",
    RTO    : "RTO_TRACKER",
    MASTER : "MASTER_DATA",
    DEALER : "DEALER_MASTER",
    EMP    : "TCO_EMPLOYEE_MASTER"
  };

  // ── Color Palette ──
  var COLOR = {
    MANUAL   : { col:"#FFFFFF", label:"#333333", data:"#FFFFFF"  },
    AUTO     : { col:"#C9DAF8", label:"#1A4CC8", data:"#EBF2FF"  },
    GEN      : { col:"#FFF2CC", label:"#B45309", data:"#FFFDE7"  },
    DROPDOWN : { col:"#D9EAD3", label:"#276221", data:"#F0FAF0"  },
    HDR_BG   : "#0D1B2A",   // Dark Navy  — Header Row BG
    HDR_FG   : "#FFFFFF",   // White      — Header Row Text
    VER_BG   : "#1F3864",   // Deep Blue  — Version Row BG
    VER_FG   : "#E8F0FE",   // Light Blue — Version Row Text
    ALT1     : "#FFFFFF",   // Row banding color 1
    ALT2     : "#F8FAFE"    // Row banding color 2
  };

  // ── Tab Colors ──
  var TAB = {
    DM     : "#4472C4",   // Blue
    ACC    : "#107C41",   // Green
    RTO    : "#ED7D31",   // Orange
    MASTER : "#7030A0",   // Purple
    DEALER : "#767676",   // Grey
    EMP    : "#767676"    // Grey
  };

  // ══════════════════════════════════════════
  //  SHEET DEFINITIONS (PDF Mapping v3)
  // ══════════════════════════════════════════

  // ── DM_DISBURSEMENT MEMO — 55 Columns ──
  var DM_COLS = [
    ["DM NO",                        "MANUAL"  ],
    ["DM DATE",                      "MANUAL"  ],
    ["MONTH",                        "GEN"     ],
    ["DISB AMT RECEIVING DATE",      "AUTO"    ],
    ["BUYER NAME",                   "MANUAL"  ],
    ["PHONE 1",                      "MANUAL"  ],
    ["PHONE 2",                      "MANUAL"  ],
    ["EMAIL",                        "MANUAL"  ],
    ["REG NO",                       "MANUAL"  ],
    ["MODEL",                        "MANUAL"  ],
    ["MFG YEAR",                     "MANUAL"  ],
    ["BANK",                         "DROPDOWN"],
    ["PRODUCT",                      "DROPDOWN"],
    ["LOAN APPLIED",                 "MANUAL"  ],
    ["LOAN APP 1",                   "MANUAL"  ],
    ["INS 1",                        "MANUAL"  ],
    ["SURAKSHA 1",                   "MANUAL"  ],
    ["SURAKSHA TENURE 1",            "MANUAL"  ],
    ["TOTAL LOAN 1",                 "MANUAL"  ],
    ["FILE CHG",                     "MANUAL"  ],
    ["ROI 1",                        "MANUAL"  ],
    ["EMI 1",                        "MANUAL"  ],
    ["LOAN TENURE 1",                "MANUAL"  ],
    ["DISB AMT 1",                   "MANUAL"  ],
    ["LOAN APP 2",                   "MANUAL"  ],
    ["INS 2",                        "MANUAL"  ],
    ["SURAKSHA 2",                   "MANUAL"  ],
    ["SURAKSHA TENURE 2",            "MANUAL"  ],
    ["TOTAL LOAN 2",                 "MANUAL"  ],
    ["FILE CHG 2",                   "MANUAL"  ],
    ["ROI_2",                        "MANUAL"  ],
    ["EMI 2",                        "MANUAL"  ],
    ["LOAN TENURE 2",                "MANUAL"  ],
    ["DISB AMT 2",                   "MANUAL"  ],
    ["RTO CHARGES",                  "MANUAL"  ],
    ["Payment to DLR (as per DM)",   "MANUAL"  ],
    ["e-Challan AMT (If Any)",       "MANUAL"  ],
    ["DEALERSHIP PAYMENT STATUS",    "AUTO"    ],
    ["EMP ID",                       "MANUAL"  ],
    ["EXECUTIVE NAME",               "AUTO"    ],
    ["BRANCH",                       "AUTO"    ],
    ["TEAM LEADER",                  "AUTO"    ],
    ["DEALERSHIP NAME",              "AUTO"    ],
    ["AUTH PERSON (DLR)",            "AUTO"    ],
    ["CONTACT NO (DLR)",             "MANUAL"  ],
    ["LOCATION",                     "AUTO"    ],
    ["VEHICLE OWNER NAME",           "MANUAL"  ],
    ["PAYOUT TCO",                   "MANUAL"  ],
    ["PAYOUT CD",                    "MANUAL"  ],
    ["PAYOUT SCORE",                 "MANUAL"  ],
    ["LOAN SCORE",                   "MANUAL"  ],
    ["RTO SCORE",                    "MANUAL"  ],
    ["EXE INCENTIVE",                "MANUAL"  ],
    ["NET SCORE",                    "MANUAL"  ],
    ["REMARKS",                      "MANUAL"  ]
  ];

  // ── ACCOUNT_PAYMENT_TRACKER — 43 Columns ──
  var ACC_COLS = [
    ["DM NO",                        "AUTO"    ],
    ["BUYER NAME",                   "AUTO"    ],
    ["PHONE 1",                      "AUTO"    ],
    ["PHONE 2",                      "AUTO"    ],
    ["REG NO",                       "AUTO"    ],
    ["MODEL",                        "AUTO"    ],
    ["MFG YEAR",                     "AUTO"    ],
    ["BANK",                         "AUTO"    ],
    ["PRODUCT",                      "AUTO"    ],
    ["DEALERSHIP NAME",              "AUTO"    ],
    ["EXECUTIVE NAME",               "AUTO"    ],
    ["DISB AMT RECEIVING DATE",      "MANUAL"  ],
    ["UTR NO",                       "MANUAL"  ],
    ["DISB AMT RECEIVED",            "MANUAL"  ],
    ["DISB AMT 1",                   "AUTO"    ],
    ["DISB AMT 2",                   "AUTO"    ],
    ["PAYER 1 NAME",                 "MANUAL"  ],
    ["PAYER 1 DATE",                 "MANUAL"  ],
    ["PAYER 1 AMOUNT",               "MANUAL"  ],
    ["PAYER 1 ACC DETAILS",          "MANUAL"  ],
    ["PAYER 2 NAME",                 "MANUAL"  ],
    ["PAYER 2 DATE",                 "MANUAL"  ],
    ["PAYER 2 AMOUNT",               "MANUAL"  ],
    ["PAYER 2 ACC DETAILS",          "MANUAL"  ],
    ["PAYER 3 NAME",                 "MANUAL"  ],
    ["PAYER 3 DATE",                 "MANUAL"  ],
    ["PAYER 3 AMOUNT",               "MANUAL"  ],
    ["PAYER 3 ACC DETAILS",          "MANUAL"  ],
    ["DEALERSHIP PAYMENT STATUS",    "DROPDOWN"],
    ["PAYOUT TO DEALER",             "MANUAL"  ],
    ["HOLD AMT FROM BANK",           "MANUAL"  ],
    ["HOLD AMT FROM TCO",            "MANUAL"  ],
    ["EXECUTIVE INCENTIVE",          "MANUAL"  ],
    ["LOAN SCORE",                   "MANUAL"  ],
    ["PAYOUT FROM BANK",             "MANUAL"  ],
    ["NET SCORE",                    "MANUAL"  ],
    ["RTO CHARGES",                  "AUTO"    ],
    ["RTO VENDOR NAME",              "AUTO"    ],
    ["RTO PAID AMOUNT",              "MANUAL"  ],
    ["RTO PAYMENT DATE",             "MANUAL"  ],
    ["RC TRANSFER STATUS",           "AUTO"    ],
    ["RTO PROFIT",                   "GEN"     ],
    ["REMARKS (IF ANY)",             "MANUAL"  ]
  ];

  // ── RTO_TRACKER — 30 Columns ──
  var RTO_COLS = [
    ["DM DATE",                          "AUTO"    ],
    ["DM NO",                            "AUTO"    ],
    ["DEALERSHIP PAYMENT STATUS",        "AUTO"    ],
    ["PRODUCT",                          "AUTO"    ],
    ["BRANCH",                           "AUTO"    ],
    ["EXECUTIVE NAME",                   "AUTO"    ],
    ["DEALERSHIP NAME",                  "AUTO"    ],
    ["REG NO",                           "AUTO"    ],
    ["CASE REC DATE",                    "MANUAL"  ],
    ["RTO CODE",                         "GEN"     ],
    ["MODEL",                            "AUTO"    ],
    ["MFG YEAR",                         "AUTO"    ],
    ["CHASSIS NO",                       "MANUAL"  ],
    ["ENGINE NO",                        "MANUAL"  ],
    ["OWNER NAME (SELLER)",              "AUTO"    ],
    ["SELLER PHONE",                     "MANUAL"  ],
    ["BUYER NAME",                       "AUTO"    ],
    ["BUYER PHONE",                      "AUTO"    ],
    ["BANK",                             "AUTO"    ],
    ["CASE TYPE",                        "MANUAL"  ],
    ["RTO SCAN FILE (LINK)",             "MANUAL"  ],
    ["RTO VENDOR",                       "DROPDOWN"],
    ["RTO RECEIPT NO",                   "MANUAL"  ],
    ["RC TRANSFER STATUS",               "DROPDOWN"],
    ["RC TRANSFER DATE",                 "MANUAL"  ],
    ["PENDING DAYS",                     "GEN"     ],
    ["ORIGINAL RC REC STATUS",           "DROPDOWN"],
    ["TRANSFERRED RC COPY_PROOF (LINK)", "MANUAL"  ],
    ["SYSTEM REMARKS",                   "MANUAL"  ],
    ["REMARKS / ISSUE",                  "MANUAL"  ]
  ];

  // ── MASTER_DATA — 97 Columns (PDF v3) — all GEN ──
  var MASTER_HEADERS = [
    "DM NO","DM DATE","AUTH PERSON (DLR)","BANK","BRANCH","BUYER NAME","BUYER PHONE",
    "CASE REC DATE","CASE TYPE","CHASSIS NO","CONTACT NO (DLR)","DEALERSHIP NAME",
    "DEALERSHIP PAYMENT STATUS","DISB AMT 1","DISB AMT 2","DISB AMT RECEIVED",
    "DISB AMT RECEIVING DATE","e-Challan AMT (If Any)","EMAIL","EMI 1","EMI 2","EMP ID",
    "ENGINE NO","EXE INCENTIVE","EXECUTIVE INCENTIVE","EXECUTIVE NAME","FILE CHG","FILE CHG 2",
    "HOLD AMT FROM BANK","HOLD AMT FROM TCO","INS 1","INS 2","LOAN APP 1","LOAN APP 2",
    "LOAN APPLIED","LOAN SCORE","LOAN TENURE 1","LOAN TENURE 2","LOCATION","MFG YEAR",
    "MODEL","MONTH","NET SCORE","ORIGINAL RC REC STATUS","OWNER NAME (SELLER)",
    "PAYER 1 ACC DETAILS","PAYER 1 AMOUNT","PAYER 1 DATE","PAYER 1 NAME",
    "PAYER 2 ACC DETAILS","PAYER 2 AMOUNT","PAYER 2 DATE","PAYER 2 NAME",
    "PAYER 3 ACC DETAILS","PAYER 3 AMOUNT","PAYER 3 DATE","PAYER 3 NAME",
    "Payment to DLR (as per DM)","PAYOUT CD","PAYOUT FROM BANK","PAYOUT SCORE",
    "PAYOUT TCO","PAYOUT TO DEALER","PENDING DAYS","PHONE 1","PHONE 2","PRODUCT",
    "RC TRANSFER DATE","RC TRANSFER STATUS","REG NO","REMARKS","REMARKS (IF ANY)",
    "REMARKS / ISSUE","ROI 1","ROI_2","RTO CHARGES","RTO CODE","RTO PAID AMOUNT",
    "RTO PAYMENT DATE","RTO PROFIT","RTO RECEIPT NO","RTO SCAN FILE (LINK)","RTO SCORE",
    "RTO VENDOR","RTO VENDOR NAME","SELLER PHONE","SURAKSHA 1","SURAKSHA 2",
    "SURAKSHA TENURE 1","SURAKSHA TENURE 2","SYSTEM REMARKS","TEAM LEADER",
    "TOTAL LOAN 1","TOTAL LOAN 2","TRANSFERRED RC COPY_PROOF (LINK)","UTR NO",
    "VEHICLE OWNER NAME"
  ];

  // ── DEALER_MASTER — 10 Columns (from Excel) ──
  var DEALER_HEADERS = [
    "DEALER CODE","DEALER NAME","LOCATION","CITY","CONTACT PERSON",
    "CONTACT NO","DLR EMAIL","EMP ID","EXECUTIVE NAME","BRANCHES"
  ];

  // ── TCO_EMPLOYEE_MASTER — 27 Columns (from Excel) ──
  var EMP_HEADERS = [
    "EMP_ID","EMP_NAME","FATHER_NAME","DOB","GENDER","DESIGNATION","DEPARTMENT",
    "BRANCH_CODE","TEAM_LEADER","TRAINING_DATE","DOJ","PROBATION_END_DATE",
    "ON_ROLL_STATUS","STATUS","MOBILE_PERSONAL","MOBILE_OFFICIAL","EMAIL_PERSONAL",
    "EMAIL_OFFICIAL","EMERGENCY_CONTACT_RELATION","EMERGENCY_CONTACT_NAME",
    "EMERGENCY_MOBILE","SALARY_BASIC","PAN_NO","AADHAR_NO","POLICE_CLEARANCE_CERT",
    "RESUME_FILE_LINK","PASSPORT_PHOTO_LINK"
  ];

  // ══════════════════════════════════════════
  //  STEP 1 — RESET ALL SHEETS
  // ══════════════════════════════════════════
  resetSheet_(ss, SN.DM);
  resetSheet_(ss, SN.ACC);
  resetSheet_(ss, SN.RTO);
  resetSheet_(ss, SN.MASTER);
  resetSheet_(ss, SN.DEALER);
  resetSheet_(ss, SN.EMP);

  // ══════════════════════════════════════════
  //  STEP 2 — BUILD WORKING SHEETS (User Editable)
  // ══════════════════════════════════════════
  buildWorkSheet_(ss, SN.DM,  DM_COLS,  TAB.DM,  "DM",  COLOR, now);
  buildWorkSheet_(ss, SN.ACC, ACC_COLS, TAB.ACC, "ACC", COLOR, now);
  buildWorkSheet_(ss, SN.RTO, RTO_COLS, TAB.RTO, "RTO", COLOR, now);

  // ══════════════════════════════════════════
  //  STEP 3 — BUILD MASTER DATA (Auto Generated)
  // ══════════════════════════════════════════
  buildMasterSheet_(ss, SN.MASTER, MASTER_HEADERS, TAB.MASTER, COLOR, now);

  // ══════════════════════════════════════════
  //  STEP 4 — BUILD REFERENCE TABLES
  // ══════════════════════════════════════════
  buildRefSheet_(ss, SN.DEALER, DEALER_HEADERS, TAB.DEALER, COLOR, now);
  buildRefSheet_(ss, SN.EMP,    EMP_HEADERS,    TAB.EMP,    COLOR, now);

  // ══════════════════════════════════════════
  //  STEP 5 — LOCK & HIDE MASTER/REF SHEETS
  // ══════════════════════════════════════════
  lockAndHideSheet_(ss, SN.MASTER);
  lockAndHideSheet_(ss, SN.DEALER);
  lockAndHideSheet_(ss, SN.EMP);

  // ══════════════════════════════════════════
  //  STEP 6 — SET SHEET ORDER
  // ══════════════════════════════════════════
  var order = [SN.DM, SN.ACC, SN.RTO, SN.MASTER, SN.DEALER, SN.EMP];
  for (var i = 0; i < order.length; i++) {
    var sh = ss.getSheetByName(order[i]);
    if (sh) ss.setActiveSheet(sh), ss.moveActiveSheet(i + 1);
  }
  ss.setActiveSheet(ss.getSheetByName(SN.DM));

  // ══════════════════════════════════════════
  //  DONE — ALERT
  // ══════════════════════════════════════════
  SpreadsheetApp.getUi().alert(
    "✅  Loan_11_MIS_FY2026-27 — v3.0\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "  DM Sheet          → 55 Columns  ✔\n" +
    "  Account Sheet     → 43 Columns  ✔\n" +
    "  RTO Sheet         → 30 Columns  ✔\n" +
    "  Master Data       → 97 Columns  🔒\n" +
    "  Dealer Master     → 10 Columns  🔒\n" +
    "  Employee Master   → 27 Columns  🔒\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "  Setup: " + now + "\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "  System Ready!"
  );
}


// ─────────────────────────────────────────────
//  RESET SHEET — Clear + Remove Protections
// ─────────────────────────────────────────────
function resetSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) { ss.insertSheet(name); sh = ss.getSheetByName(name); }

  // Unhide if hidden (needed to edit)
  try { sh.showSheet(); } catch(e) {}

  // Remove all protections
  var rp = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < rp.length; i++) rp[i].remove();
  var sp = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var j = 0; j < sp.length; j++) sp[j].remove();

  // Remove bandings
  var bd = sh.getBandings();
  for (var k = 0; k < bd.length; k++) bd[k].remove();

  sh.clear();
  sh.clearConditionalFormatRules();
}


// ─────────────────────────────────────────────
//  BUILD WORKING SHEET (DM / ACC / RTO)
// ─────────────────────────────────────────────
function buildWorkSheet_(ss, name, cols, tabColor, code, C, now) {
  var sheet   = ss.getSheetByName(name);
  var numCols = cols.length;
  var DATA_ROWS = 1000;  // rows 4 → 1003

  // ── Tab Color ──
  sheet.setTabColor(tabColor);

  // ── ROW 1: VERSION BAR ──
  sheet.getRange(1, 1, 1, numCols).merge();
  sheet.getRange(1, 1)
    .setValue("📋  Loan_11_MIS_FY2026-27   |   " + code + "   |   v3.0   |   Setup: " + now)
    .setBackground(C.VER_BG)
    .setFontColor(C.VER_FG)
    .setFontSize(10)
    .setFontFamily("Arial")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 28);

  // ── ROW 2: COLUMN HEADERS ──
  var headers = cols.map(function(c){ return c[0]; });
  var hrng    = sheet.getRange(2, 1, 1, numCols);
  hrng.setValues([headers])
      .setBackground(C.HDR_BG)
      .setFontColor(C.HDR_FG)
      .setFontSize(9)
      .setFontFamily("Arial")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(false);
  sheet.setRowHeight(2, 38);

  // ── ROW 3: TYPE LABELS ──
  var labels = cols.map(function(c){
    var t = c[1];
    if (t === "AUTO")     return "↺  AUTO PULL";
    if (t === "GEN")      return "⚙  AUTO GEN";
    if (t === "DROPDOWN") return "▾  DROPDOWN";
    return "✏  MANUAL";
  });
  sheet.getRange(3, 1, 1, numCols).setValues([labels]);
  sheet.setRowHeight(3, 22);

  // ── PER-COLUMN STYLING ──
  for (var i = 0; i < numCols; i++) {
    var col  = i + 1;
    var type = cols[i][1];
    var name_ = cols[i][0];
    var palette = C[type];

    // Type label row (row 3)
    sheet.getRange(3, col)
      .setBackground(palette.col)
      .setFontColor(palette.label)
      .setFontSize(8)
      .setFontFamily("Arial")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Header bottom border (type-color indicator)
    sheet.getRange(2, col).setBorder(
      null, null, true, null, null, null,
      type === "MANUAL" ? "#888888" : palette.col,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

    // Data area background (rows 4 onward)
    if (type !== "MANUAL") {
      sheet.getRange(4, col, DATA_ROWS, 1).setBackground(palette.data);
    }

    // Column width
    var w = 120;
    if (name_.indexOf("REMARKS") > -1 || name_.indexOf("LINK") > -1) w = 210;
    else if (name_.indexOf("NAME") > -1 || name_.indexOf("DETAILS") > -1) w = 145;
    else if (name_.indexOf("AMT") > -1  || name_.indexOf("AMOUNT") > -1)  w = 130;
    else if (name_.indexOf("DATE") > -1) w = 125;
    sheet.setColumnWidth(col, w);
  }

  // ── ALTERNATING ROW COLORS (data area — MANUAL cols only) ──
  var cfRules = [];
  var evenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(MOD(ROW(),2)=0,TRUE)")
    .setBackground(C.ALT2)
    .setRanges([sheet.getRange(4, 1, DATA_ROWS, numCols)])
    .build();
  cfRules.push(evenRule);
  sheet.setConditionalFormatRules(cfRules);

  // ── FREEZE TOP 3 ROWS ──
  sheet.setFrozenRows(3);

  // ── PROTECT AUTO / GEN COLUMNS ──
  for (var j = 0; j < numCols; j++) {
    var t = cols[j][1];
    if (t === "AUTO" || t === "GEN") {
      var lock = sheet.getRange(4, j + 1, DATA_ROWS, 1).protect();
      lock.setDescription("🔒 " + cols[j][0] + " — Auto Column");
      lock.setWarningOnly(false);
      lock.removeEditors(lock.getEditors());
      if (lock.canDomainEdit()) lock.setDomainEdit(false);
    }
  }

  // ── HEADER ROW PROTECTION (rows 1-3) ──
  var hdrLock = sheet.getRange(1, 1, 3, numCols).protect();
  hdrLock.setDescription("🔒 Header Rows — System");
  hdrLock.setWarningOnly(false);
  hdrLock.removeEditors(hdrLock.getEditors());
  if (hdrLock.canDomainEdit()) hdrLock.setDomainEdit(false);
}


// ─────────────────────────────────────────────
//  BUILD MASTER DATA SHEET (97 cols — hidden/locked)
// ─────────────────────────────────────────────
function buildMasterSheet_(ss, name, headers, tabColor, C, now) {
  var sheet   = ss.getSheetByName(name);
  var numCols = headers.length;

  sheet.setTabColor(tabColor);

  // Version row
  sheet.getRange(1, 1, 1, numCols).merge();
  sheet.getRange(1, 1)
    .setValue("📋  MASTER_DATA   |   97 Columns   |   Auto Generated   |   v3.0   |   " + now)
    .setBackground(C.VER_BG)
    .setFontColor(C.VER_FG)
    .setFontSize(10)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 28);

  // Header row
  sheet.getRange(2, 1, 1, numCols)
    .setValues([headers])
    .setBackground(C.HDR_BG)
    .setFontColor(C.HDR_FG)
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 35);

  // Type row — all GEN
  var types = headers.map(function(){ return "⚙  AUTO GEN"; });
  sheet.getRange(3, 1, 1, numCols)
    .setValues([types])
    .setBackground(C.GEN.col)
    .setFontColor(C.GEN.label)
    .setFontSize(8)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 22);

  // Data area tint — yellow
  sheet.getRange(4, 1, 997, numCols).setBackground(C.GEN.data);

  // Column widths
  for (var i = 1; i <= numCols; i++) sheet.setColumnWidth(i, 130);

  sheet.setFrozenRows(3);
}


// ─────────────────────────────────────────────
//  BUILD REFERENCE TABLE (DEALER / EMP)
// ─────────────────────────────────────────────
function buildRefSheet_(ss, name, headers, tabColor, C, now) {
  var sheet   = ss.getSheetByName(name);
  var numCols = headers.length;

  sheet.setTabColor(tabColor);

  // Version row
  sheet.getRange(1, 1, 1, numCols).merge();
  sheet.getRange(1, 1)
    .setValue("📋  " + name + "   |   " + numCols + " Columns   |   Reference Table   |   v3.0   |   " + now)
    .setBackground(C.VER_BG)
    .setFontColor(C.VER_FG)
    .setFontSize(10)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 28);

  // Header row
  sheet.getRange(2, 1, 1, numCols)
    .setValues([headers])
    .setBackground(C.HDR_BG)
    .setFontColor(C.HDR_FG)
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 35);
  sheet.setFrozenRows(2);

  for (var i = 1; i <= numCols; i++) sheet.setColumnWidth(i, 150);
}


// ─────────────────────────────────────────────
//  LOCK AND HIDE SHEET
// ─────────────────────────────────────────────
function lockAndHideSheet_(ss, name) {
  var sh   = ss.getSheetByName(name);
  var prot = sh.protect();
  prot.setDescription("🔒 SYSTEM SHEET — Owner Only — " + name);
  prot.setWarningOnly(false);
  prot.removeEditors(prot.getEditors());
if (prot.canDomainEdit()) prot.setDomainEdit(false);
prot.addEditor('admin.loan11@gmail.com');
sh.hideSheet();
}


// ══════════════════════════════════════════════════════
//  BONUS UTILITIES — Run these separately as needed
// ══════════════════════════════════════════════════════

/**
 * REFRESH VERSION TIMESTAMP
 * Run this whenever you re-deploy or update the sheet.
 * Updates the version bar date on all 3 working sheets.
 */
function refreshVersionStamp() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var now = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy  HH:mm");
  var sheets = {
    "DM_DISBURSEMENT MEMO"     : "DM",
    "ACCOUNT_PAYMENT_TRACKER"  : "ACC",
    "RTO_TRACKER"              : "RTO"
  };
  for (var sname in sheets) {
    var sh = ss.getSheetByName(sname);
    if (!sh) continue;
    // Unprotect header row temporarily
    var protections = sh.getRange(1,1,1,1).getProtections(SpreadsheetApp.ProtectionType.RANGE);
    // Direct edit via script always works — protections apply to UI users only
    sh.getRange(1, 1)
      .setValue("📋  Loan_11_MIS_FY2026-27   |   " + sheets[sname] + "   |   v3.0   |   Updated: " + now);
  }
  SpreadsheetApp.getUi().alert("✅ Version stamps refreshed!\n" + now);
}


/**
 * SHOW ALL HIDDEN SHEETS (Admin Use)
 * Temporarily reveals MASTER, DEALER, EMP for admin viewing.
 */
function adminShowAllSheets() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var hidden  = ["MASTER_DATA","DEALER_MASTER","TCO_EMPLOYEE_MASTER"];
  var ui      = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "⚠️  ADMIN ACTION",
    "Reveal all hidden sheets?\nThis is for Admin review only.",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  for (var i = 0; i < hidden.length; i++) {
    var sh = ss.getSheetByName(hidden[i]);
    if (sh) sh.showSheet();
  }
  ui.alert("✅ Hidden sheets are now visible.\nRun adminHideAllSheets() when done.");
}

/**
 * HIDE MASTER SHEETS AGAIN (Admin Use)
 */
function adminHideAllSheets() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var hidden = ["MASTER_DATA","DEALER_MASTER","TCO_EMPLOYEE_MASTER"];
  for (var i = 0; i < hidden.length; i++) {
    var sh = ss.getSheetByName(hidden[i]);
    if (sh) sh.hideSheet();
  }
  SpreadsheetApp.getUi().alert("✅ Master sheets hidden again.");
}

