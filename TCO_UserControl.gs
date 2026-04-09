// ============================================================
// TCO_UserControl.gs — v4.0 | FY 2026-27
// Sheet Protection System
// 
// HIDDEN SHEETS (7) : Lock + Hide — Admin only
// VISIBLE SHEETS (3): DM, Account, RTO — Open for editors
// AUTO COLS          : Lock — Admin only
// MANUAL COLS        : Open — All editors
// ============================================================

var ADMIN_EMAIL = 'admin.loan11@gmail.com';
var SHEET_ID    = '1yGs3OfRX9UdRUEMYzKHr1yt0rUaDQ5rBjoxg6aeNBgM';

var VISIBLE_SHEETS = [
  'DM_DISBURSEMENT MEMO',
  'ACCOUNT_PAYMENT_TRACKER',
  'RTO_TRACKER'
];

var HIDDEN_SHEETS = [
  'MASTER_DATA',
  'DEALER_MASTER',
  'TCO_EMPLOYEE_MASTER',
  'AUTH_AUDIT_LOG',
  '_BACKUP_DM',
  '_BACKUP_ACCOUNT',
  '_BACKUP_RTO'
];

// AUTO cols to LOCK per sheet (1-based col numbers)
var AUTO_COLS = {
  'DM_DISBURSEMENT MEMO'    : [1,3,4,5,39,41,42,43,44,45,47,48],
  'ACCOUNT_PAYMENT_TRACKER' : [1,2,3,4,5,6,7,8,9,10,11,12,13,14,39,40,43,44],
  'RTO_TRACKER'             : [1,2,3,4,5,6,7,8,10,11,12,15,17,18,19,26]
};

// ── STEP 1: Lock 7 hidden sheets ─────────────────────────────
function TCO_lockHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Remove old protections from hidden sheets
  HIDDEN_SHEETS.forEach(function(n) {
    var sh = ss.getSheetByName(n);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
  });

  // Hide + lock 7 hidden sheets
  HIDDEN_SHEETS.forEach(function(n) {
    var sh = ss.getSheetByName(n);
    if (!sh) return;
    sh.hideSheet();
    var p = sh.protect();
    p.setDescription('TCO HIDDEN — ' + n);
    p.setWarningOnly(false);
    var editors = p.getEditors();
    if (editors.length) p.removeEditors(editors);
    if (p.canDomainEdit()) p.setDomainEdit(false);
    p.addEditor(ADMIN_EMAIL);
  });

  // Show 3 visible sheets
  VISIBLE_SHEETS.forEach(function(n) {
    var sh = ss.getSheetByName(n);
    if (sh) sh.showSheet();
  });

  ui.alert(
    '✅ Hidden Sheets Locked!\n\n' +
    '7 Hidden : Locked + Hidden (Admin only)\n' +
    '3 Visible: DM, Account, RTO — Open\n\n' +
    'Date: ' + new Date().toLocaleString('en-IN')
  );
}

// ── STEP 2: Lock auto cols on 3 visible sheets ────────────────
function TCO_lockAutoCols() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Remove old range protections on visible sheets
  VISIBLE_SHEETS.forEach(function(n) {
    var sh = ss.getSheetByName(n);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
  });

  var totalLocked = 0;

  VISIBLE_SHEETS.forEach(function(sheetName) {
    var sheet   = ss.getSheetByName(sheetName);
    if (!sheet) return;
    var autoCols = AUTO_COLS[sheetName] || [];
    var maxRow   = sheet.getMaxRows();

    autoCols.forEach(function(col) {
      var rp = sheet.getRange(1, col, maxRow, 1).protect();
      rp.setDescription('AUTO | ' + sheetName + ' | COL' + col);
      rp.setWarningOnly(false);
      var editors = rp.getEditors();
      if (editors.length) rp.removeEditors(editors);
      if (rp.canDomainEdit()) rp.setDomainEdit(false);
      rp.addEditor(ADMIN_EMAIL);
      totalLocked++;
    });
  });

  ui.alert(
    '✅ Auto Cols Locked!\n\n' +
    'DM Sheet      : ' + AUTO_COLS['DM_DISBURSEMENT MEMO'].length + ' cols locked\n' +
    'Account Sheet : ' + AUTO_COLS['ACCOUNT_PAYMENT_TRACKER'].length + ' cols locked\n' +
    'RTO Sheet     : ' + AUTO_COLS['RTO_TRACKER'].length + ' cols locked\n\n' +
    'Total : ' + totalLocked + ' col ranges\n' +
    'Manual cols : Open for all editors\n\n' +
    'Date: ' + new Date().toLocaleString('en-IN')
  );
}

// ── RUN BOTH: Lock hidden + Lock auto cols ────────────────────
function TCO_applyAllProtections() {
  TCO_lockHiddenSheets();
  TCO_lockAutoCols();
}

// ── REMOVE ALL PROTECTIONS (emergency) ───────────────────────
function TCO_removeProtections() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.alert('⚠️ Remove ALL protections?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
  ui.alert('✅ All protections removed.');
}

// ── WEB APP ENTRY POINT ───────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TCO Operations System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── DASHBOARD FUNCTIONS ───────────────────────────────────────
function getSalesDashboard(empId) {
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID),dm=ss.getSheetByName('DM_DISBURSEMENT MEMO'),data=dm.getDataRange().getValues();
    var stats={total:0,disbursed:0,approved:0,pending:0,totalLoan:0,totalInc:0},bankPend=[],transPend=[],monthly={},recent=[];
    for (var i=3;i<data.length;i++) {
      var r=data[i];
      if (!r[0]||!String(r[0]).startsWith('TCO')) continue;
      if (String(r[39]).trim()!==String(empId).trim()) continue;
      stats.total++;
      var loan=parseFloat(r[15])||0,disb=parseFloat(r[25])||0,inc=parseFloat(r[54])||0;
      stats.totalLoan+=loan;stats.totalInc+=inc;
      if (disb>0) stats.disbursed++; else stats.pending++;
      var mon=String(r[2]).trim();
      if (!monthly[mon]) monthly[mon]={cases:0};
      monthly[mon].cases++;
      if (r[13]&&!disb) bankPend.push({dmNo:r[0],buyer:r[5],bank:r[13],loanAmt:loan,loginDate:formatDate_(r[1]),status:'PENDING',approvalDate:''});
      if (disb>0) transPend.push({dmNo:r[0],buyer:r[5],regNo:r[9],disbDate:formatDate_(r[1]),rtoStatus:String(r[38]||'PENDING'),rcStatus:String(r[38]||'PENDING'),alert:calcAlert_(r[1])});
      recent.push({dmNo:r[0],buyer:r[5],loginDate:formatDate_(r[1]),vehicle:r[11],loanAmt:loan,status:disb>0?'DISBURSED':'PENDING',bank:r[13]});
    }
    return {stats:stats,bankPending:bankPend.slice(0,50),transferPending:transPend.slice(0,50),monthly:monthly,recent:recent.slice(0,100)};
  } catch(e) { return {error:e.message}; }
}

function getAccountsDashboard() {
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID),acc=ss.getSheetByName('ACCOUNT_PAYMENT_TRACKER'),data=acc.getDataRange().getValues();
    var stats={totalCases:0,disbursedCases:0,totalDisb:0,totalPayout:0,totalHold:0,dealerPending:0},payPend=[],dealerPay=[],holdCases=[],branchDisb={};
    for (var i=3;i<data.length;i++) {
      var r=data[i];
      if (!r[1]||!String(r[1]).startsWith('TCO')) continue;
      stats.totalCases++;
      var d1=parseFloat(r[12])||0,d2=parseFloat(r[13])||0,recv=parseFloat(r[14])||0,hold=parseFloat(r[32])||0,holdTco=parseFloat(r[33])||0,payout=parseFloat(r[36])||0;
      if (recv>0) stats.disbursedCases++;
      stats.totalDisb+=d1+d2;stats.totalPayout+=payout;stats.totalHold+=hold+holdTco;
      if (String(r[29]||'PENDING').trim()==='PENDING') { stats.dealerPending++; dealerPay.push({dmNo:r[1],buyer:r[2],dealer:r[10],branch:r[11],disbAmt:d1+d2,payoutDealer:parseFloat(r[31])||0,disbDate:formatDate_(r[0])}); }
      var p1s=recv>0?'RECEIVED':'PENDING';
      payPend.push({dmNo:r[1],buyer:r[2],branch:r[11]||'—',p1Status:p1s,p1Amt:d1,p2Status:d2>0?p1s:'—',p2Amt:d2,holdAmt:hold+holdTco,challan:0});
      if (hold+holdTco>0) holdCases.push({dmNo:r[1],buyer:r[2],branch:r[11]||'—',holdAmt:hold+holdTco,challan:0,status:p1s});
      var bn=String(r[11]||'UNKNOWN'); if (!branchDisb[bn]) branchDisb[bn]=0; branchDisb[bn]+=d1+d2;
    }
    return {stats:stats,payPending:payPend,dealerPay:dealerPay,holdCases:holdCases,branchDisb:branchDisb};
  } catch(e) { return {error:e.message}; }
}

function getRtoDashboard() {
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID),rto=ss.getSheetByName('RTO_TRACKER'),data=rto.getDataRange().getValues();
    var stats={totalCases:0,rcPendingCount:0,critical:0,warning:0,completed:0},rcPending=[],vendors={},rtoCodeMap={};
    for (var i=3;i<data.length;i++) {
      var r=data[i];
      if (!r[1]||!String(r[1]).startsWith('TCO')) continue;
      stats.totalCases++;
      var rcs=String(r[23]||'PENDING').trim().toUpperCase(),pd=parseInt(r[25])||0,vendor=String(r[21]||'UNASSIGNED').trim(),rtoCode=String(r[9]||'—').trim();
      var alert=pd>=50?'CRITICAL':pd>=30?'WARNING':'OK';
      if (rcs!=='COMPLETED'&&rcs!=='DONE') { stats.rcPendingCount++; if(alert==='CRITICAL')stats.critical++; if(alert==='WARNING')stats.warning++; rcPending.push({dmNo:r[1],buyer:r[16],regNo:r[7],branch:r[4],pendingDays:pd,rtoStatus:'IN PROCESS',rcStatus:rcs,alert:alert,disbDate:formatDate_(r[0])}); } else stats.completed++;
      if (!vendors[vendor]) vendors[vendor]={vendor:vendor,total:0,completed:0,totalDays:0};
      vendors[vendor].total++;vendors[vendor].totalDays+=pd;
      if (rcs==='COMPLETED') vendors[vendor].completed++;
      if (rtoCode&&rtoCode!=='—') { if(!rtoCodeMap[rtoCode])rtoCodeMap[rtoCode]=0; rtoCodeMap[rtoCode]++; }
    }
    rcPending.sort(function(a,b){return b.pendingDays-a.pendingDays;});
    var vl=Object.keys(vendors).map(function(v){var vd=vendors[v];return{vendor:v,total:vd.total,completed:vd.completed,pending:vd.total-vd.completed,avgDays:vd.total>0?Math.round(vd.totalDays/vd.total):0};});
    return {stats:stats,rcPending:rcPending,vendors:vl,rtoCodeMap:rtoCodeMap};
  } catch(e) { return {error:e.message}; }
}

function getManagementDashboard() {
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID),dm=ss.getSheetByName('DM_DISBURSEMENT MEMO'),data=dm.getDataRange().getValues();
    var stats={totalCases:0,disbursed:0,pending:0,totalLoan:0,totalDisb:0,branches:0},branchMap={},monthlyMap={},bankMap={},execMap={};
    for (var i=3;i<data.length;i++) {
      var r=data[i];
      if (!r[0]||!String(r[0]).startsWith('TCO')) continue;
      stats.totalCases++;
      var loan=parseFloat(r[15])||0,disb=parseFloat(r[25])||0,inc=parseFloat(r[54])||0;
      var branch=String(r[41]||'UNKNOWN').trim(),month=String(r[2]||'').trim(),bank=String(r[13]||'UNKNOWN').trim(),exec=String(r[40]||'UNKNOWN').trim();
      stats.totalLoan+=loan;stats.totalDisb+=disb;
      if (disb>0) stats.disbursed++; else stats.pending++;
      if (!branchMap[branch]) { branchMap[branch]={cases:0,disbursed:0,pending:0,totalLoan:0,totalDisb:0,incentive:0};stats.branches++; }
      branchMap[branch].cases++;branchMap[branch].totalLoan+=loan;branchMap[branch].totalDisb+=disb;branchMap[branch].incentive+=inc;
      if (disb>0) branchMap[branch].disbursed++; else branchMap[branch].pending++;
      if (month) { if(!monthlyMap[month])monthlyMap[month]={cases:0,disbAmount:0,loanAmount:0};monthlyMap[month].cases++;monthlyMap[month].disbAmount+=disb;monthlyMap[month].loanAmount+=loan; }
      if (!bankMap[bank]) bankMap[bank]=0;bankMap[bank]++;
      if (!execMap[exec]) execMap[exec]={name:exec,branch:branch,cases:0,disbursed:0,totalLoan:0,incentive:0};
      execMap[exec].cases++;execMap[exec].totalLoan+=loan;execMap[exec].incentive+=inc;
      if (disb>0) execMap[exec].disbursed++;
    }
    return {stats:stats,branchMap:branchMap,monthlyMap:monthlyMap,bankMap:bankMap,leaderboard:Object.values(execMap).sort(function(a,b){return b.disbursed-a.disbursed;})};
  } catch(e) { return {error:e.message}; }
}

function getAdminDashboard() {
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID),us=ss.getSheetByName('USER_MASTER'),data=us.getDataRange().getValues(),users=[];
    for (var i=3;i<data.length;i++) { var r=data[i]; if(!r[0])continue; users.push({empId:r[0],name:r[1],role:r[3],branch:r[4],isActive:String(r[6]).toUpperCase()==='YES'}); }
    return {users:users};
  } catch(e) { return {error:e.message}; }
}

function formatDate_(val) {
  if (!val) return '—';
  try { var d=new Date(val); if(isNaN(d))return String(val); return d.getDate()+'/'+(d.getMonth()+1)+'/'+d.getFullYear(); } catch(e) { return String(val); }
}

function calcAlert_(dmDate) {
  if (!dmDate) return 'OK';
  var days=Math.round((new Date()-new Date(dmDate))/86400000);
  return days>=50?'CRITICAL':days>=30?'WARNING':'OK';
}