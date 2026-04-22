// ============================================================
// TCO_SheetSetup.gs — v3.0 | FY 2026-27
// All 11 Sheets | Row colors | Validations | Number formats
// Run: setupTCOSystem()
// ============================================================

var SETUP = {
  SHEETS : {
    VISIBLE : ['DM_DISBURSEMENT MEMO','ACCOUNT_PAYMENT_TRACKER','RTO_TRACKER'],
    HIDDEN  : ['MASTER_DATA','DEALER_MASTER','TCO_EMPLOYEE_MASTER',
               'AUTH_AUDIT_LOG','_BACKUP_DM','_BACKUP_ACCOUNT','_BACKUP_RTO']
  },
  // Row 2: Light bg, dark font (same all sheets)
  ROW2 : { BG: '#EFF6FF', FONT: '#1e293b' },

  // Row 1 section colors — DM Sheet
  DM_COLORS : {
    KEY      : '#374151',  // Charcoal
    CUSTOMER : '#065f46',  // Dark Green
    BANK1    : '#1e3a5f',  // Navy
    BANK2    : '#1e40af',  // Royal Blue
    CHARGES  : '#92400e',  // Dark Amber
    EXEC     : '#4c1d95',  // Purple
    DEALER   : '#064e3b',  // Forest Green
    PAYOUT   : '#78350f',  // Brown
    REMARKS  : '#6b7280'   // Gray
  },
  // Row 1 section colors — ACC Sheet
  ACC_COLORS : {
    FROM_DM  : '#1e3a5f',  // Navy
    PAYMENT  : '#065f46',  // Green
    PAYER1   : '#6d28d9',  // Violet
    PAYER2   : '#5b21b6',  // Deep Violet
    PAYER3   : '#4c1d95',  // Purple
    DEALER   : '#92400e',  // Amber
    HOLD     : '#991b1b',  // Dark Red
    RTO      : '#0369a1',  // Steel Blue
    REMARKS  : '#374151'   // Charcoal
  },
  // Row 1 section colors — RTO Sheet
  RTO_COLORS : {
    KEY      : '#374151',  // Charcoal
    STATUS   : '#991b1b',  // Dark Red
    FROM_DM  : '#1e3a5f',  // Navy
    PROCESS  : '#065f46',  // Green
    VEHICLE  : '#92400e',  // Amber
    OWNER    : '#4c1d95',  // Purple
    FILING   : '#0369a1',  // Steel Blue
    PENDING  : '#7c2d12',  // Dark Orange
    RC       : '#374151'   // Charcoal
  },
  // Cell type colors
  TYPE : {
    AUTO_GEN  : '#166534',  // Dark Green
    AUTO_PULL : '#1A4D8F',  // Deep Blue
    MANUAL    : '#000000',  // Black
    REG_NO    : '#c2410c',  // Dark Orange
    DM_NO     : '#5D4037',  // Dark Brown
  },
  BORDER : '#CBD5E1'
};


// ── MAIN ──────────────────────────────────────────────────────
function setupTCOSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.alert('🚀 TCO Setup Starting...\nPlease wait 60-90 seconds.');

  createAllSheets_(ss);
  setDmHeaders_(ss);
  setAccHeaders_(ss);
  setRtoHeaders_(ss);
  setDealerMasterHeaders_(ss);
  setEmpMasterHeaders_(ss);
  setUserMasterHeaders_(ss);
  setMasterDataHeaders_(ss);
  setAuditLogHeaders_(ss);
  setBackupSheetLabels_(ss);
  applyAllFormatting_(ss);
  setValidations_(ss);
  manageSheetVisibility_(ss);
  setColumnWidths_(ss);

  ui.alert(
    '✅ Setup Complete! All 11 Sheets Ready.\n\n' +
    'Next: ⚙️ TCO Admin → Apply Sheet Protections'
  );
}

function createAllSheets_(ss) {
  var all = SETUP.SHEETS.VISIBLE.concat(SETUP.SHEETS.HIDDEN);
  var ex  = ss.getSheets().map(function(s){ return s.getName(); });
  all.forEach(function(n){ if (ex.indexOf(n)===-1) ss.insertSheet(n); });
}


// ── DM SHEET ──────────────────────────────────────────────────
function setDmHeaders_(ss) {
  var s = ss.getSheetByName('DM_DISBURSEMENT MEMO');
  if (!s) return;
  s.getRange(1,1,3,57).clear();

  // Row 1 — each section unique color
  var c = SETUP.DM_COLORS;
  var dmSec = [
    {st:1, en:2,  label:'🔑  KEY',                  color:c.KEY},
    {st:3, en:10, label:'🪪  CUSTOMER DETAILS',      color:c.CUSTOMER},
    {st:11,en:26, label:'🏦  BANK 1 — LOAN DETAILS', color:c.BANK1},
    {st:27,en:37, label:'🏦  BANK 2 — LOAN DETAILS', color:c.BANK2},
    {st:38,en:39, label:'💲  CHARGES & STATUS',       color:c.CHARGES},
    {st:40,en:43, label:'👤  EXECUTIVE',              color:c.EXEC},
    {st:44,en:47, label:'🏪  DEALERSHIP',             color:c.DEALER},
    {st:48,en:56, label:'💰  PAYOUT & SCORE',         color:c.PAYOUT},
    {st:57,en:57, label:'📝  REMARKS',                color:c.REMARKS}
  ];
  setSectionRow_(s, dmSec);

  // Row 2 — light bg, dark font
  var h2 = [
    'DM NO','DM DATE','MONTH','DM RECEIVED DATE\n(FROM ACCOUNTS)','NO. OF DAYS',
    'CUSTOMER NAME\n(BUYER)','PHONE NO','2ND PHONE NO','EMAIL ID','REGISTRATION NO',
    'VEHICLE OWNER\n(SELLER)','VEHICLE MODEL','MANUFACTURING\nYEAR','FINANCED BANK NAME','PRODUCT',
    'LOAN AMOUNT\nAPPLIED','LOAN AMOUNT\nAPPROVED 1','LOAN TENURE 1','INSURANCE\nAMOUNT 1',
    'LOAN SURAKHSA 1','TOTAL LOAN\nAMOUNT 1','SURAKHSA\nTENURE 1','FILE CHARGE 1',
    'ROI (%) 1','EMI 1','DISBURSEMENT\nAMOUNT 1',
    '2ND FINANCED\nBANK NAME','LOAN AMOUNT\nAPPROVED 2','LOAN TENURE 2','INSURANCE\nAMOUNT 2',
    'LOAN SURAKHSA 2','TOTAL LOAN\nAMOUNT 2','SURAKHSA\nTENURE 2','FILE CHARGE 2',
    'ROI (%) 2','EMI 2','DISBURSEMENT\nAMOUNT 2',
    'e-Challan AMT\n(If Any)','DEALERSHIP\nPAYMENT STATUS','EMP ID',
    'EXECUTIVE NAME','BRANCH','TEAM LEADER','DEALERSHIP NAME',
    'AUTH PERSON (DLR)','CONTACT NO (DLR)','LOCATION','SELLER NAME (AUTO)',
    'PAYOUT TCO','PAYOUT CD','PAYOUT SCORE','LOAN SCORE',
    'RTO CHARGES','RTO SCORE','EXE INCENTIVE','NET SCORE','REMARKS'
  ];
  setHeaderRow2_(s, h2);

  // Row 3 — type indicators with colors
  var t3 = [
    {v:'⚡ AUTO GEN',  c:SETUP.TYPE.AUTO_GEN},
    {v:'✏ MANUAL',    c:SETUP.TYPE.MANUAL},
    {v:'⚡ AUTO GEN',  c:SETUP.TYPE.AUTO_GEN},
    {v:'↺ AUTO PULL', c:SETUP.TYPE.AUTO_PULL},
    {v:'⚡ AUTO GEN',  c:SETUP.TYPE.AUTO_GEN},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.REG_NO},   // REG NO — orange hint
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'▾ DROPDOWN',c:SETUP.TYPE.MANUAL},
    {v:'▾ DROPDOWN',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'% PERCENT',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'% PERCENT',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✎ FREE TEXT',c:SETUP.TYPE.MANUAL}
  ];
  setTypeRow3_(s, t3);

  finalizeRows_(s, true);
}


// ── ACC SHEET ─────────────────────────────────────────────────
function setAccHeaders_(ss) {
  var s = ss.getSheetByName('ACCOUNT_PAYMENT_TRACKER');
  if (!s) return;
  s.getRange(1,1,3,45).clear();

  var c = SETUP.ACC_COLORS;
  var accSec = [
    {st:1, en:14, label:'🔑  AUTO FROM DM ▶',     color:c.FROM_DM},
    {st:15,en:17, label:'💳  PAYMENT RECEIVED',    color:c.PAYMENT},
    {st:18,en:21, label:'💸  PAYER 1',             color:c.PAYER1},
    {st:22,en:25, label:'💸  PAYER 2',             color:c.PAYER2},
    {st:26,en:29, label:'💸  PAYER 3',             color:c.PAYER3},
    {st:30,en:32, label:'🏪  DEALER PAYMENT',      color:c.DEALER},
    {st:33,en:38, label:'💲  HOLD & INCENTIVE',    color:c.HOLD},
    {st:39,en:44, label:'🔧  RTO SECTION',         color:c.RTO},
    {st:45,en:45, label:'📝  REMARKS',             color:c.REMARKS}
  ];
  setSectionRow_(s, accSec);

  var h2 = [
    'DM DATE','DM NO','BUYER NAME','PHONE NO','2ND PHONE NO','REGISTRATION NO',
    'VEHICLE MODEL','MANUFACTURING\nYEAR','FINANCED BANK NAME','PRODUCT',
    'DEALERSHIP NAME','EXECUTIVE NAME','DISB AMOUNT 1\n(BANK)','DISB AMOUNT 2\n(CM)',
    'DISB AMT RECEIVED','DISB AMT\nRECEIVING DATE','UTR NO',
    'PAYER 1 NAME','PAYER 1 DATE','PAYER 1 AMOUNT','PAYER 1 ACC DETAILS',
    'PAYER 2 NAME','PAYER 2 DATE','PAYER 2 AMOUNT','PAYER 2 ACC DETAILS',
    'PAYER 3 NAME','PAYER 3 DATE','PAYER 3 AMOUNT','PAYER 3 ACC DETAILS',
    'DEALERSHIP\nPAYMENT STATUS','DEALERSHIP\nPAYMENT DATE','PAYOUT TO DEALER',
    'HOLD AMT FROM BANK','HOLD AMT FROM TCO','EXECUTIVE INCENTIVE',
    'LOAN SCORE','PAYOUT FROM BANK','NET SCORE',
    'RTO CHARGES','RTO VENDOR NAME','RTO PAID AMOUNT','RTO PAYMENT DATE',
    'RC TRANSFER STATUS','RTO PROFIT\n(CHARGES-PAID)','REMARKS (IF ANY)'
  ];
  setHeaderRow2_(s, h2);

  var t3 = [
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.REG_NO},  // REG NO orange hint
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'▾ DROPDOWN',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'⚡ AUTO GEN',c:SETUP.TYPE.AUTO_GEN},{v:'✎ FREE TEXT',c:SETUP.TYPE.MANUAL}
  ];
  setTypeRow3_(s, t3);
  finalizeRows_(s, true);
}


// ── RTO SHEET ─────────────────────────────────────────────────
function setRtoHeaders_(ss) {
  var s = ss.getSheetByName('RTO_TRACKER');
  if (!s) return;
  s.getRange(1,1,3,30).clear();

  var c = SETUP.RTO_COLORS;
  var rtoSec = [
    {st:1, en:2,  label:'🔑  LINK KEY',         color:c.KEY},
    {st:3, en:3,  label:'📋  STATUS',            color:c.STATUS},
    {st:4, en:8,  label:'👤  AUTO FROM DM ▶',   color:c.FROM_DM},
    {st:9, en:10, label:'🏛  RTO PROCESS',       color:c.PROCESS},
    {st:11,en:14, label:'🚗  VEHICLE',           color:c.VEHICLE},
    {st:15,en:19, label:'📋  OWNER & BANK',      color:c.OWNER},
    {st:20,en:25, label:'🏛  RTO FILING',        color:c.FILING},
    {st:26,en:26, label:'⏱  PENDING DAYS',      color:c.PENDING},
    {st:27,en:30, label:'📦  RC & REMARKS',      color:c.RC}
  ];
  setSectionRow_(s, rtoSec);

  var h2 = [
    'DM DATE','DM NO','DEALERSHIP PAYMENT\nSTATUS',
    'PRODUCT','BRANCH','EXECUTIVE NAME','DEALERSHIP NAME','REGISTRATION NO',
    'RTO CASE_FILE\nREC DATE','RTO CODE',
    'VEHICLE MODEL','MANUFACTURING\nYEAR','CHASSIS NO','ENGINE NO',
    'SELLER NAME','SELLER PHONE','BUYER NAME','BUYER PHONE',
    'FINANCED BANK NAME','CASE TYPE',
    'RTO SCAN FILE\n(LINK)','RTO VENDOR','RTO RECEIPT NO',
    'RC TRANSFER\nSTATUS','RC TRANSFER DATE','PENDING DAYS',
    'ORIGINAL RC\nREC STATUS','TRANSFERRED RC\nCOPY_PROOF (LINK)',
    'SYSTEM REMARKS','REMARKS / ISSUE'
  ];
  setHeaderRow2_(s, h2);

  var t3 = [
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.REG_NO},  // REG NO orange hint
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'⚡ AUTO GEN',c:SETUP.TYPE.AUTO_GEN},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},{v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'↺ AUTO PULL',c:SETUP.TYPE.AUTO_PULL},
    {v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'▾ DROPDOWN',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'▾ DROPDOWN',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'⚡ AUTO GEN',c:SETUP.TYPE.AUTO_GEN},
    {v:'▾ DROPDOWN',c:SETUP.TYPE.MANUAL},{v:'✏ MANUAL',c:SETUP.TYPE.MANUAL},
    {v:'✎ FREE TEXT',c:SETUP.TYPE.MANUAL},{v:'✎ FREE TEXT',c:SETUP.TYPE.MANUAL}
  ];
  setTypeRow3_(s, t3);
  finalizeRows_(s, true);
}


// ── SUPPORT SHEETS ────────────────────────────────────────────
function setDealerMasterHeaders_(ss) {
  var s = ss.getSheetByName('DEALER_MASTER'); if (!s) return;
  s.getRange(1,1,3,10).clear();
  s.getRange(1,1,1,10).merge().setValue('📊  DEALER_MASTER — TCO Authorized Dealers')
   .setBackground('#065f46').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(14)
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  setHeaderRow2_(s,['DEALER CODE','DEALER NAME','LOCATION','CITY','CONTACT PERSON',
    'CONTACT NO','DLR EMAIL','EMP ID','EXECUTIVE NAME','BRANCHES']);
  s.getRange(3,1,1,10).setValues([['DLR001','Dealer Name','Area','City',
    'Owner/Manager','10-digit','email@domain.com','TCO_ID','Exec Name','Branch']])
   .setBackground('#D1FAE5').setFontColor('#065f46').setFontSize(9).setFontStyle('italic')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(1,40);s.setRowHeight(2,40);s.setRowHeight(3,22);s.setFrozenRows(2);
  s.getRange(4,1,200,10).setFontWeight('bold').setHorizontalAlignment('center')
   .setVerticalAlignment('middle').setFontSize(10)
   .setBorder(true,true,true,true,true,true,SETUP.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  [90,180,160,100,160,130,180,90,150,150].forEach(function(w,i){s.setColumnWidth(i+1,w);});
}

function setEmpMasterHeaders_(ss) {
  var s = ss.getSheetByName('TCO_EMPLOYEE_MASTER'); if (!s) return;
  s.getRange(1,1,3,27).clear();
  var sec27 = [
    {st:1,en:5,label:'🪪  IDENTITY',color:'#1e3a5f'},
    {st:6,en:9,label:'💼  JOB DETAILS',color:'#065f46'},
    {st:10,en:14,label:'📅  DATES & STATUS',color:'#92400e'},
    {st:15,en:17,label:'📞  CONTACT',color:'#4c1d95'},
    {st:18,en:21,label:'📧  EMAIL & EMERGENCY',color:'#0369a1'},
    {st:22,en:24,label:'💰  FINANCE',color:'#78350f'},
    {st:25,en:27,label:'📄  DOCUMENTS',color:'#374151'}
  ];
  setSectionRow_(s, sec27);
  setHeaderRow2_(s,['EMP_ID','EMP_NAME','FATHER_NAME','DOB','GENDER',
    'DESIGNATION','DEPARTMENT','BRANCH_CODE','TEAM_LEADER',
    'TRAINING_DATE','DOJ','PROBATION_END_DATE','ON_ROLL_STATUS','STATUS',
    'MOBILE_PERSONAL','MOBILE_OFFICIAL','EMAIL_PERSONAL','EMAIL_OFFICIAL',
    'EMERGENCY_CONTACT_RELATION','EMERGENCY_CONTACT_NAME','EMERGENCY_MOBILE',
    'SALARY_BASIC','PAN_NO','AADHAR_NO',
    'POLICE_CLEARANCE_CERT','RESUME_FILE_LINK','PASSPORT_PHOTO_LINK']);
  s.getRange(3,1,1,27).setValues([['TCO001','FULL NAME','FATHER','DD/MM/YYYY','M/F',
    'Designation','Dept','Branch','TL','DD/MM/YYYY','DD/MM/YYYY','DD/MM/YYYY',
    'On-roll','ACTIVE','10-digit','10-digit','personal@','official@',
    'Father/Mother','Name','10-digit','₹Amount','PAN','Aadhar',
    'YES/NO','Link','Link']])
   .setBackground('#DBEAFE').setFontColor('#1e3a5f').setFontSize(9).setFontStyle('italic')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(1,36);s.setRowHeight(2,45);s.setRowHeight(3,22);
  s.setFrozenRows(2);
  s.getRange(4,1,200,27).setFontWeight('bold').setHorizontalAlignment('center')
   .setVerticalAlignment('middle').setFontSize(10)
   .setBorder(true,true,true,true,true,true,SETUP.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  dropdownRule_(s,4,14,500,['ACTIVE','PROBATION','CONFIRMED','NOTICE PERIOD','RESIGNED','TERMINATED']);
  dropdownRule_(s,4,5,500,['MALE','FEMALE','OTHER']);
  var w={1:90,2:160,3:150,4:100,5:80,6:140,7:140,8:140,9:150,13:110,14:110,17:180,18:180,22:100,23:130,24:150,26:130,27:130};
  Object.keys(w).forEach(function(c){s.setColumnWidth(Number(c),w[c]);});
}

function setUserMasterHeaders_(ss) {
  var s = ss.getSheetByName('USER_MASTER'); if (!s) return;
  s.getRange(1,1,3,21).clear();
  var secU = [
    {st:1,en:7,label:'👤  USER INFO',color:'#1e3a5f'},
    {st:8,en:9,label:'📸  PROFILE PHOTO',color:'#374151'},
    {st:10,en:13,label:'📋  DM PERMISSIONS',color:'#065f46'},
    {st:14,en:17,label:'💰  ACCOUNT PERMISSIONS',color:'#92400e'},
    {st:18,en:21,label:'🚗  RTO PERMISSIONS',color:'#0369a1'}
  ];
  setSectionRow_(s, secU);
  setHeaderRow2_(s,['EMP ID','FULL NAME','PASSWORD','ROLE','BRANCH','EMAIL','IS ACTIVE',
    'PROFILE PHOTO LINK','PHOTO THUMB LINK',
    'DM — CAN VIEW','DM — EDIT COLS','DM — BRANCH FILTER','DM — OWN CASES ONLY',
    'ACC — CAN VIEW','ACC — EDIT COLS','ACC — BRANCH FILTER','ACC — OWN CASES ONLY',
    'RTO — CAN VIEW','RTO — EDIT COLS','RTO — BRANCH FILTER','RTO — OWN CASES ONLY']);
  s.getRange(3,1,1,21).setValues([['MGT01','Full Name','password',
    'MANAGEMENT/SALES/ACCOUNTS/RTO/ADMIN','Branch or ALL','email@gmail.com','YES/NO',
    'Drive link','Thumb link','YES/NO','Col nos. or ALL','Branch or ALL','YES/NO',
    'YES/NO','Col nos. or ALL','Branch or ALL','YES/NO',
    'YES/NO','Col nos. or ALL','Branch or ALL','YES/NO']])
   .setBackground('#FEF3C7').setFontColor('#92400e').setFontSize(9).setFontStyle('italic')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(1,36);s.setRowHeight(2,45);s.setRowHeight(3,22);s.setFrozenRows(2);
  s.getRange(4,1,1,21).setValues([['MGT01','ADMIN','admin@123','MANAGEMENT','ALL',
    'admin.loan11@gmail.com','YES','','',
    'YES','ALL','ALL','NO','YES','ALL','ALL','NO','YES','ALL','ALL','NO']])
   .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
   .setFontSize(10).setBorder(true,true,true,true,true,true,SETUP.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  dropdownRule_(s,4,4,200,['SALES','ACCOUNTS','RTO','MANAGEMENT','ADMIN']);
  dropdownRule_(s,4,7,200,['YES','NO']);
  s.getRange(5,1,196,21).setFontWeight('bold').setHorizontalAlignment('center')
   .setVerticalAlignment('middle').setFontSize(10)
   .setBorder(true,true,true,true,true,true,SETUP.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  [90,150,110,130,110,180,80].forEach(function(w,i){s.setColumnWidth(i+1,w);});
  s.setColumnWidths(8,2,130);s.setColumnWidths(10,12,100);
}

function setMasterDataHeaders_(ss) {
  var s = ss.getSheetByName('MASTER_DATA'); if (!s) return;
  s.getRange(1,1,2,99).clear();
  s.getRange(1,1,1,57).merge().setValue('📋  DM DATA — 57 COLS')
   .setBackground('#1F3864').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13)
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.getRange(1,58,1,31).merge().setValue('💰  ACCOUNT DATA — 31 COLS')
   .setBackground('#166534').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13)
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.getRange(1,89,1,11).merge().setValue('🚗  RTO DATA — 11 COLS')
   .setBackground('#7c2d12').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13)
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  var allH=['DM NO','DM DATE','MONTH','DM RECEIVED DATE','NO. OF DAYS',
    'CUSTOMER NAME (BUYER)','PHONE NO','2ND PHONE NO','EMAIL ID','REGISTRATION NO',
    'VEHICLE OWNER (SELLER)','VEHICLE MODEL','MANUFACTURING YEAR','FINANCED BANK NAME','PRODUCT',
    'LOAN AMOUNT APPLIED','LOAN AMOUNT APPROVED 1','LOAN TENURE 1','INSURANCE AMOUNT 1',
    'LOAN SURAKHSA 1','TOTAL LOAN AMOUNT 1','SURAKHSA TENURE 1','FILE CHARGE 1',
    'ROI (%) 1','EMI 1','DISBURSEMENT AMOUNT 1',
    '2ND FINANCED BANK NAME','LOAN AMOUNT APPROVED 2','LOAN TENURE 2','INSURANCE AMOUNT 2',
    'LOAN SURAKHSA 2','TOTAL LOAN AMOUNT 2','SURAKHSA TENURE 2','FILE CHARGE 2',
    'ROI (%) 2','EMI 2','DISBURSEMENT AMOUNT 2',
    'e-Challan AMT (If Any)','DEALER PAYMENT STATUS','EMP ID',
    'EXECUTIVE NAME','BRANCH','TEAM LEADER','DEALERSHIP NAME',
    'AUTH PERSON (DLR)','CONTACT NO (DLR)','LOCATION','SELLER NAME (AUTO)',
    'PAYOUT TCO','PAYOUT CD','PAYOUT SCORE','LOAN SCORE',
    'RTO CHARGES','RTO SCORE','EXE INCENTIVE','NET SCORE','DM REMARKS',
    'DISB AMT RECEIVED','DISB AMT RECEIVING DATE','UTR NO',
    'PAYER 1 NAME','PAYER 1 DATE','PAYER 1 AMOUNT','PAYER 1 ACC DETAILS',
    'PAYER 2 NAME','PAYER 2 DATE','PAYER 2 AMOUNT','PAYER 2 ACC DETAILS',
    'PAYER 3 NAME','PAYER 3 DATE','PAYER 3 AMOUNT','PAYER 3 ACC DETAILS',
    'DEALERSHIP PAYMENT DATE','PAYOUT TO DEALER',
    'HOLD AMT FROM BANK','HOLD AMT FROM TCO','EXECUTIVE INCENTIVE',
    'LOAN SCORE (ACC)','PAYOUT FROM BANK','NET SCORE (ACC)',
    'RTO VENDOR NAME','RTO PAID AMOUNT','RTO PAYMENT DATE',
    'RC TRANSFER STATUS','RTO PROFIT','ACC REMARKS',
    'DEALERSHIP PAYMENT STATUS (ACC)','DISB AMT 1 (BANK)',
    'RTO CODE','RTO CASE_FILE REC DATE','CHASSIS NO','ENGINE NO',
    'SELLER PHONE','CASE TYPE','RTO SCAN FILE (LINK)',
    'RTO RECEIPT NO','RC TRANSFER DATE','ORIGINAL RC REC STATUS','RTO REMARKS / ISSUE'];
  setHeaderRow2_(s, allH);
  s.setRowHeight(1,36);s.setRowHeight(2,40);s.setFrozenRows(2);
}

function setAuditLogHeaders_(ss) {
  var s = ss.getSheetByName('AUTH_AUDIT_LOG'); if (!s) return;
  s.getRange(1,1,2,6).clear();
  s.getRange(1,1,1,6).merge().setValue('📊  AUTH_AUDIT_LOG — Login History')
   .setBackground('#374151').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13)
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  setHeaderRow2_(s,['LOGIN TIMESTAMP','EMP ID','NAME','ROLE','BRANCH','SOURCE']);
  s.setRowHeight(1,36);s.setRowHeight(2,36);s.setFrozenRows(2);
  [160,90,150,110,120,120].forEach(function(w,i){s.setColumnWidth(i+1,w);});
}

function setBackupSheetLabels_(ss) {
  ['_BACKUP_DM','_BACKUP_ACCOUNT','_BACKUP_RTO'].forEach(function(n) {
    var sh=ss.getSheetByName(n); if (!sh) return;
    sh.getRange(1,1).setValue('AUTO BACKUP — '+n)
      .setFontWeight('bold').setFontSize(12)
      .setBackground('#374151').setFontColor('#FFFFFF');
    sh.setRowHeight(1,32);
  });
}


// ── VALIDATIONS ───────────────────────────────────────────────
function setValidations_(ss) {
  var dm  = ss.getSheetByName('DM_DISBURSEMENT MEMO');
  var acc = ss.getSheetByName('ACCOUNT_PAYMENT_TRACKER');
  var rto = ss.getSheetByName('RTO_TRACKER');

  // ── 0. CLEAR ALL EXISTING VALIDATIONS FIRST ───────────────
  // Removes old/conflicting validations before applying new ones
  [dm, acc, rto].forEach(function(sh) {
    if (!sh) return;
    var lastRow = Math.max(sh.getLastRow(), 4);
    var lastCol = sh.getLastColumn() || 57;
    sh.getRange(4, 1, lastRow, lastCol).clearDataValidations();
  });

  // ── 1. DROPDOWNS — Strict block ───────────────────────────
  if (dm) {
    dropdownRule_(dm,4,14,500,['SBI','HDFC','KOTAK','ICICI','AXIS','PNB','BOB',
      'UNION BANK','CANARA','FEDERAL','YES BANK','IDFC','OTHER']);
    dropdownRule_(dm,4,15,500,['USED','NEW','TWO WHEELER']);
  }
  if (acc) dropdownRule_(acc,4,30,500,['PENDING','PAID','PARTIAL']);
  if (rto) {
    dropdownRule_(rto,4,22,500,['VENDOR 1','VENDOR 2','VENDOR 3','VENDOR 4','OTHER']);
    dropdownRule_(rto,4,24,500,['PENDING','IN PROCESS','COMPLETED']);
    dropdownRule_(rto,4,27,500,['RECEIVED','PENDING','NOT APPLICABLE']);
  }

  // ── 2. DATE — Strict block ─────────────────────────────────
  if (dm)  dateRule_(dm,  4, 2,  500);
  if (acc) {
    dateRule_(acc,4,16,500); dateRule_(acc,4,19,500);
    dateRule_(acc,4,23,500); dateRule_(acc,4,27,500);
    dateRule_(acc,4,31,500); dateRule_(acc,4,42,500);
  }
  if (rto) { dateRule_(rto,4,9,500); dateRule_(rto,4,25,500); }
  // RTO Seller Phone — 10 digit
          
  // ── 3. TENURE — Strict block (2 digit max) ────────────────
  if (dm) {
    numLenRule_(dm,4,18,500,2);  // R: LOAN TENURE 1
    numLenRule_(dm,4,22,500,2);  // V: SURAKHSA TENURE 1
    numLenRule_(dm,4,29,500,2);  // AC: LOAN TENURE 2
    numLenRule_(dm,4,33,500,2);  // AG: SURAKHSA TENURE 2
  }

  // ── 4. MFG YEAR — Strict block (4 digit) ──────────────────
  if (dm) numLenRule_(dm,4,13,500,4);

  // ── 5. NUMBER FORMAT — Currency columns ───────────────────
  // Format only — no validation popup (warning via onEdit script)
  var dmCurr  = [16,17,19,20,21,23,25,26,28,30,31,32,34,36,37,38,49,50,53,55];  // Removed score cols: 51,52,54,56
  var accCurr = [15,20,24,28,32,33,34,35,37,41];
  if (dm)  applyFormat_(dm, 4, dmCurr,  500, '₹ ##,##,##0');
  if (acc) applyFormat_(acc,4, accCurr, 500, '₹ ##,##,##0');

  // ── 6. ROI % FORMAT ───────────────────────────────────────
  if (dm) {
    applyFormat_(dm,4,[24],500,'0.00" %"');
    applyFormat_(dm,4,[35],500,'0.00" %"');
  }

  // ── 6b. SCORE COLS — plain number (no currency, no comma) ──
  if (dm) applyFormat_(dm,4,[51,52,54,56],500,'0');  // PAYOUT SCORE, LOAN SCORE, RTO SCORE, NET SCORE

  // ── 7. TENURE number format ───────────────────────────────
  if (dm) applyFormat_(dm,4,[18,22,29,33],500,'0');

  // ── 8. DATE FORMAT: 6-Apr-2026 ────────────────────────────
  if (dm)  applyFormat_(dm, 4,[2], 500,'d-mmm-yyyy');
  if (acc) applyFormat_(acc,4,[1,16,19,23,27,31,42],500,'d-mmm-yyyy');
  if (rto) applyFormat_(rto,4,[1,9,25],500,'d-mmm-yyyy');

  // ── 12. SPECIAL SIZE 12 COLS ────────────────────────────────
  // DM sheet special cols
  if (dm) {
    [[1],[10],[16],[26],[37],[39],[55],[56]].forEach(function(arr) {
      dm.getRange(4,arr[0],500,1).setFontSize(12);
    });
  }
  // ACC sheet special cols
  if (acc) {
    [[2],[6],[30]].forEach(function(arr) {
      acc.getRange(4,arr[0],500,1).setFontSize(12);
    });
  }
  // RTO sheet special cols
  if (rto) {
    [[2],[8],[24],[26]].forEach(function(arr) {
      rto.getRange(4,arr[0],500,1).setFontSize(12);
    });
  }

  // NOTE: Word, Word+No, Email, Phone validations are handled
  // via onEdit script (warning color only — no popup, no block)
}


// ── VALIDATION HELPERS ────────────────────────────────────────

// Helper: column number → letter (A, B, ... Z, AA, AB...)
function colLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter  = String.fromCharCode(65 + rem) + letter;
    col     = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Dropdown
function dropdownRule_(sh,sr,col,rows,vals) {
  sh.getRange(sr,col,rows,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(vals,true)
      .setAllowInvalid(false).build());
}

// Date only — strict block + format 6-Apr-2026
function dateRule_(sh,sr,col,rows) {
  sh.getRange(sr,col,rows,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDate()
      .setHelpText('Date only — format: 6-Apr-2026')
      .setAllowInvalid(false).build());
}

// Phone — 10 digit only
function phoneRule_(sh,sr,col,rows) {
  if (!sh) return;
  var cl = colLetter_(col);
  sh.getRange(sr,col,rows,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied(
        '=AND(ISNUMBER('+cl+sr+'),LEN(TEXT('+cl+sr+',"0"))=10)')
      .setHelpText('10-digit mobile number only')
      .setAllowInvalid(false).build());
}

// Number length — 2 digit or 4 digit (strict block)
function numLenRule_(sh,sr,col,rows,digits) {
  if (!sh) return;
  var maxVal = Math.pow(10,digits) - 1;
  var minVal = digits === 4 ? 1900 : 0;  // 0 allowed = no data
  sh.getRange(sr,col,rows,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(minVal,maxVal)
      .setHelpText('Max '+digits+' digits allowed')
      .setAllowInvalid(false).build());
}

// Word-only validation removed — handled by onEdit warning colors

// Word+Num validation removed — handled by onEdit warning colors

// Currency — number only + ₹ format with space (₹ 8,00,000)
function currencyRule_(sh,sr,cols,rows) {
  if (!sh) return;
  cols.forEach(function(col) {
    var cl = colLetter_(col);
    // Validation: numbers only (greater than or equal to 0)
    sh.getRange(sr,col,rows,1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setHelpText('Numbers only — no text or special characters')
        .setAllowInvalid(false).build());
    // Format: ₹ 8,00,000 (Indian format with space)
    sh.getRange(sr,col,rows,1).setNumberFormat('₹ ##,##,##0');
  });
}

// Percentage — number only + format: 12.54 % (space before %)
function percentRule_(sh,sr,col,rows) {
  if (!sh) return;
  sh.getRange(sr,col,rows,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0,100)
      .setHelpText('Enter percentage: e.g. 12.54')
      .setAllowInvalid(false).build());
  // Format: 12.54 % (gap between number and %)
  sh.getRange(sr,col,rows,1).setNumberFormat('0.00" %"');
}

// Apply number format to multiple columns
function applyFormat_(sh,sr,cols,rows,fmt) {
  cols.forEach(function(c) {
    var maxRows = sh.getMaxRows() - sr + 1;
    sh.getRange(sr,c,maxRows,1).setNumberFormat(fmt);
  });
}


// ── FORMATTING ────────────────────────────────────────────────
function applyAllFormatting_(ss) {
  var map = {'DM_DISBURSEMENT MEMO':57,'ACCOUNT_PAYMENT_TRACKER':45,'RTO_TRACKER':30};
  Object.keys(map).forEach(function(sn) {
    var sh = ss.getSheetByName(sn); if (!sh) return;
    sh.getRange(4,2,200,map[sn]-1)
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setFontColor('#000000').setFontSize(10).setFontFamily('Arial')
      .setBorder(true,true,true,true,true,true,SETUP.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    for (var r=4;r<=60;r++) sh.setRowHeight(r,30);
  });
}

function manageSheetVisibility_(ss) {
  SETUP.SHEETS.VISIBLE.forEach(function(n){var s=ss.getSheetByName(n);if(s)s.showSheet();});
  SETUP.SHEETS.HIDDEN.forEach(function(n){var s=ss.getSheetByName(n);if(s)s.hideSheet();});
}

function setColumnWidths_(ss) {
  var dm=ss.getSheetByName('DM_DISBURSEMENT MEMO');
  if (dm) {
    dm.setColumnWidth(1,130);dm.setColumnWidth(2,100);dm.setColumnWidth(3,80);
    dm.setColumnWidth(6,160);dm.setColumnWidth(10,120);dm.setColumnWidth(14,110);
    dm.setColumnWidths(16,11,100);dm.setColumnWidth(40,90);
    dm.setColumnWidth(41,150);dm.setColumnWidth(42,120);dm.setColumnWidth(44,160);
  }
  var acc=ss.getSheetByName('ACCOUNT_PAYMENT_TRACKER');
  if (acc) {
    acc.setColumnWidth(2,130);acc.setColumnWidth(3,160);
    acc.setColumnWidth(6,120);acc.setColumnWidth(11,150);acc.setColumnWidth(30,120);
  }
  var rto=ss.getSheetByName('RTO_TRACKER');
  if (rto) {
    rto.setColumnWidth(2,130);rto.setColumnWidth(8,120);rto.setColumnWidth(10,80);
    rto.setColumnWidth(22,130);rto.setColumnWidth(24,120);rto.setColumnWidth(26,90);
  }
}


// ── SHARED HEADER HELPERS ─────────────────────────────────────
function setSectionRow_(sheet, sections) {
  sections.forEach(function(sec) {
    var r = sheet.getRange(1, sec.st, 1, sec.en - sec.st + 1);
    if (sec.st !== sec.en) r.merge();
    r.setValue(sec.label)
     .setBackground(sec.color).setFontColor('#FFFFFF')
     .setFontWeight('bold').setFontSize(13)
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
}

function setHeaderRow2_(sheet, headers) {
  sheet.getRange(2,1,1,headers.length).setValues([headers])
    .setBackground(SETUP.ROW2.BG)       // Light background
    .setFontColor(SETUP.ROW2.FONT)      // Dark font
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function setTypeRow3_(sheet, typeArr) {
  // Set values
  var vals = typeArr.map(function(t){ return t.v; });
  sheet.getRange(3,1,1,vals.length).setValues([vals])
    .setBackground('#F8FAFC')
    .setFontSize(9).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  // Set individual font colors per type
  typeArr.forEach(function(t,i) {
    sheet.getRange(3, i+1).setFontColor(t.c);
  });
}

function finalizeRows_(sheet, withType) {
  sheet.setRowHeight(1, 36);
  sheet.setRowHeight(2, 50);
  if (withType) sheet.setRowHeight(3, 22);
  sheet.setFrozenRows(withType ? 3 : 2);
}
