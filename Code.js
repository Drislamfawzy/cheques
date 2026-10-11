// المصفوفة المركزية للصلاحيات
var ALLOWED_MANAGERS = [
"egyptianplast@gmail.com",
"adhamamro55@gmail.com",
"islamfawzy86@icloud.com"
];

// الاحتياطي الحديدي (تم تجميد 200 ألف لا تظهر في المتاح)
var CASH_FLOOR_RESERVE = 200000;

function doGet(e) {
ensureSystemTabs();
var htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
htmlOutput.setTitle('نظام إدارة الشيكات والخزينة')
.addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no');
return htmlOutput;
}
function include(filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

var TZ = Session.getScriptTimeZone();
var BANKS = ["NBE", "BM", "CIB", "MTF"];
var CHEQUES_HEADERS = [
"Cheque_ID", "Type", "Party_Name", "Bank_Name", "Reference_Number / Cheque_Number",
"Amount", "Cheque_Date", "Due_Date", "Status", "Received_Date", "Cleared_Date",
"Drive_Original", "Drive_Cleared", "Calendar_Event_ID", "Notes", "Created_At",
"Updated_At", "Instrument_Type", "Is_Settled", "Invoice_Date"
];
var SETTINGS_HEADERS = [
"Bank_Name", "Opening_Balance", "Last_Actual_Balance", "Last_Reconciliation_Date", "Variance_Alert_Threshold_Percent"
];

var LOG_HEADERS = [
"Entry_ID", "Entry_Type", "Bank_Name", "Amount", "Direction", "Reason", "Date", "Created_At"
];

// أعمدة شيت القروض الجديد
var LOANS_HEADERS = [
"Loan_ID", "Loan_Name", "Bank_Name", "Monthly_Amount", "Start_Date", "End_Date", "Total_Installments", "Day_Of_Month", "Active", "Created_At"
];

var CACHE_TTL_SECONDS = 30;

function getCache_(key) {
try {
var cache = CacheService.getScriptCache();
var val = cache.get(key);
return val ? JSON.parse(val) : null;
} catch (e) {
return null;
}
}

function setCache_(key, value) {
try {
CacheService.getScriptCache().put(key, JSON.stringify(value), CACHE_TTL_SECONDS);
} catch (e) {}
}

function clearAppCache_() {
try {
CacheService.getScriptCache().removeAll([
"CACHE_DASHBOARD_STATS", "CACHE_BANK_PANEL", "CACHE_ACTIVE_ASC", "CACHE_ACTIVE_DESC",
"CACHE_CLEARED_ASC", "CACHE_CLEARED_DESC", "CACHE_DECISION_DASHBOARD",
"CACHE_CASHFLOW_7", "CACHE_CASHFLOW_30", "CACHE_CASHFLOW_90", "CACHE_DAILY_LADDER"
]);
} catch (e) {}
}

function getSheet(name) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(name);
if (!sheet) {
sheet = ss.insertSheet(name);
if (name === "CHEQUES") sheet.appendRow(CHEQUES_HEADERS);
else if (name === "SETTINGS") sheet.appendRow(SETTINGS_HEADERS);
else if (name === "LOG") sheet.appendRow(LOG_HEADERS);
else if (name === "LOANS") sheet.appendRow(LOANS_HEADERS);
}
return sheet;
}

function ensureSystemTabs() {
var settingsSheet = getSheet("SETTINGS");
if (settingsSheet.getLastRow() <= 1) {
if (settingsSheet.getLastRow() === 0) settingsSheet.appendRow(SETTINGS_HEADERS);
BANKS.forEach(function(bank) {
settingsSheet.appendRow([bank, 0, 0, "", 10]);
});
} else {
var data = settingsSheet.getDataRange().getValues();
var existingBanks = [];
for (var i = 1; i < data.length; i++) {
existingBanks.push(data[i][0].toString().trim());
}
BANKS.forEach(function(bank) {
if (existingBanks.indexOf(bank) === -1) {
settingsSheet.appendRow([bank, 0, 0, "", 10]);
}
});
}
getSheet("LOG");
ensureLoanTab_(); 
}

function ensureLoanTab_() {
var loanSheet = getSheet("LOANS");
if (loanSheet.getLastRow() <= 1) {
if (loanSheet.getLastRow() === 0) loanSheet.appendRow(LOANS_HEADERS);
var now = new Date();
var seedLoans = [
  ["LOAN-NBE-01", "قسط البنك الأهلي", "NBE", 23450, "", "2027-06-30", "", 20, "Y", now],
  ["LOAN-BM-01", "قسط بنك مصر", "BM", 112000, "", "2027-09-30", "", 4, "Y", now],
  ["LOAN-FAB-01", "قسط بنك FAB", "FAB", 150000, "", "2027-08-31", "", 4, "Y", now],
["LOAN-NBE-MACH-01", "قسط أهلي آلات 2026", "NBE", 94000, "2027-03-10", "", 57, 10, "Y", now]
];
seedLoans.forEach(function(row) {
  loanSheet.appendRow(row);
});
}
}

function formatDateStr(val) {
if (!val) return "";
if (val instanceof Date) {
if (isNaN(val.getTime())) return "";
return Utilities.formatDate(val, TZ, "yyyy-MM-dd");
}
return val.toString().trim();
}

function parseLocalDate(dateStr) {
if (!dateStr) return null;
var parts = dateStr.split("-");
if (parts.length !== 3) {
var d = new Date(dateStr);
return isNaN(d.getTime()) ? null : d;
}
var d2 = new Date(parseInt(parts[0],10), parseInt(parts[1],10)-1, parseInt(parts[2],10));
return isNaN(d2.getTime()) ? null : d2;
}

function isItemCleared(isSettledValue) {
return isSettledValue === "Y";
}

function getOrCreateFolder(name) {
var folders = DriveApp.getFoldersByName(name);
if (folders.hasNext()) return folders.next();
return DriveApp.createFolder(name);
}

function getTargetFolder(instrumentType) {
if (instrumentType === "BANK_TRANSFER" || instrumentType === "CASH") {
try { return DriveApp.getFolderById("1CtoyZvLxaQa3W91GtXfqowTdSrmNsgn6"); } catch (e) { return getOrCreateFolder("03_BANK_TRANSFERS"); }
}
return getOrCreateFolder("01_ORIGINAL_CHEQUES");
}

function findImageInDriveByChequeNumber(chequeNum, instrumentType) {
try {
if (!chequeNum) return "";
var folder = getTargetFolder(instrumentType);
var files = folder.getFiles();
while (files.hasNext()) {
var file = files.next();
var fileName = file.getName();
if (fileName.indexOf(chequeNum) > -1 && fileName.indexOf("-doc") > -1) { return file.getUrl(); }
}
} catch (e) {}
return "";
}

function getImagesMapFromFormResponses() {
var map = {};
try {
var sheet = getSheet("Form Responses 1");
if (!sheet || sheet.getLastRow() <= 1) return map;
var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });
var numIdx = -1, imgIdx = -1;
for (var h = 0; h < headers.length; h++) {
var hn = headers[h];
if (hn.includes("رقم الشيك") || hn.toLowerCase().includes("cheque_number")) numIdx = h;
if (hn.includes("صورة الشيك") || hn.toLowerCase().includes("image") || hn.toLowerCase().includes("file")) imgIdx = h;
}
if (numIdx === -1) numIdx = 1;
if (imgIdx === -1) imgIdx = 2;
for (var i = 1; i < data.length; i++) {
var num = data[i][numIdx] ? data[i][numIdx].toString().trim() : "";
var url = data[i][imgIdx] ? data[i][imgIdx].toString().trim() : "";
if (num && url) map[num] = url;
}
} catch (e) {}
return map;
}

function uploadChequeImage(refNumber, base64Data, mimeType, fileNameHint, instrumentType) {
try {
if (!refNumber || !base64Data) return { success:false, error:"بيانات ناقصة لرفع الصورة." };
var folder = getTargetFolder(instrumentType);
var cleanBase64 = base64Data.indexOf(",") > -1 ? base64Data.split(",")[1] : base64Data;
var bytes = Utilities.base64Decode(cleanBase64);
var prefix = (instrumentType === "BANK_TRANSFER") ? "TRF-" : "CHQ-";
var blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", prefix + refNumber + "-" + (fileNameHint || "doc"));
var file = folder.createFile(blob);
file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
ALLOWED_MANAGERS.forEach(function(email) {
  try { file.addViewer(email); } catch(e) {}
});

var url = file.getUrl();
var lock = LockService.getScriptLock();

try {
  lock.waitLock(10000);
  var sheet = getSheet("CHEQUES");
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return h.toString().trim(); });
    var numIdx = headers.indexOf("Reference_Number / Cheque_Number"); if (numIdx === -1) numIdx = 4;
    var origIdx = headers.indexOf("Drive_Original"); if (origIdx === -1) origIdx = 11;
    var updIdx = headers.indexOf("Updated_At"); if (updIdx === -1) updIdx = 16;
    
    for (var i = 1; i < data.length; i++) {
      var rowNum = data[i][numIdx] ? data[i][numIdx].toString().trim() : "";
      if (rowNum === refNumber.toString().trim()) {
        sheet.getRange(i+1, origIdx+1).setValue(url);
        if (updIdx < data[i].length) sheet.getRange(i+1, updIdx+1).setValue(new Date());
        break;
      }
    }
  }
} finally {
  lock.releaseLock();
}
clearAppCache_();
return { success:true, url: url };
} catch (error) {
return { success:false, error: error.toString() };
}
}

function getEmptyStats() {
return { INCOMING: { count:0, total:0, today:0, week:0, overdue:0 }, OUTGOING: { count:0, total:0, today:0, week:0, overdue:0 } };
}

function computeLiquidityStatus(liquidity, outgoingUpcoming) {
if (outgoingUpcoming <= 0) return "safe";
var ratio = liquidity / outgoingUpcoming;
if (ratio >= 1.2) return "safe";
if (ratio >= 0.9) return "warning";
return "danger";
}

function getDashboardStats() {
try {
CacheService.getScriptCache().remove("CACHE_DASHBOARD_STATS");
var cached = getCache_("CACHE_DASHBOARD_STATS");
if (cached) return cached;

var bankPanels = getBankAccountsPanel();
var totalActualBankBalance = 0;
var totalMTFBalance = 0;
var minConfidence = "high";

for (var b = 0; b < bankPanels.length; b++) {
  if (bankPanels[b].bank === "MTF") {
    totalMTFBalance += bankPanels[b].estimatedActualBalance;
  } else {
    totalActualBankBalance += bankPanels[b].estimatedActualBalance;
  }
  if (bankPanels[b].confidence === "low") minConfidence = "low";
  else if (bankPanels[b].confidence === "medium" && minConfidence !== "low") minConfidence = "medium";
}

var sheet = getSheet("CHEQUES");
var stats = getEmptyStats();
stats.todayTasks = [];

if (!sheet || sheet.getLastRow() <= 1) {
  stats.liquidityStatus = computeLiquidityStatus(totalActualBankBalance + totalMTFBalance, 0);
  stats.netLiquiditySnapshot = (totalActualBankBalance + totalMTFBalance) - CASH_FLOOR_RESERVE;
  stats.overallConfidence = minConfidence;
  stats.cashFloor = CASH_FLOOR_RESERVE;
  setCache_("CACHE_DASHBOARD_STATS", stats);
  return stats;
}

var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });
var typeIdx = headers.indexOf("Type"); if (typeIdx===-1) typeIdx=1;
var amountIdx = headers.indexOf("Amount"); if (amountIdx===-1) amountIdx=5;
var dueDateIdx = headers.indexOf("Due_Date"); if (dueDateIdx===-1) dueDateIdx=7;
var statusIdx = headers.indexOf("Is_Settled"); if (statusIdx===-1) statusIdx=18;
var numIdx = headers.indexOf("Reference_Number / Cheque_Number"); if (numIdx===-1) numIdx=4;
var partyIdx = headers.indexOf("Party_Name"); if (partyIdx===-1) partyIdx=2;

var today = new Date(); today.setHours(0,0,0,0);
var in7 = new Date(today.getTime() + 7*86400000);

for (var i=1; i<data.length; i++) {
  var row = data[i];
  var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
  var status = row[statusIdx] ? row[statusIdx].toString().trim() : "";
  var amount = parseFloat(row[amountIdx]) || 0;
  if ((type !== "INCOMING" && type !== "OUTGOING") || isItemCleared(status)) continue;

  stats[type].count += 1;
  stats[type].total += amount;

  var dueDate = parseLocalDate(formatDateStr(row[dueDateIdx]));
  if (dueDate) {
    dueDate.setHours(0,0,0,0);
    var isOverdue = dueDate.getTime() < today.getTime();
    var isToday = dueDate.getTime() === today.getTime();

    if (isOverdue) stats[type].overdue += amount;
    else if (isToday) stats[type].today += amount;
    else if (dueDate.getTime() <= in7.getTime()) stats[type].week += amount;

    if (isOverdue || isToday) {
      stats.todayTasks.push({
        Cheque_Number: row[numIdx] ? row[numIdx].toString().trim() : "",
        Type: type,
        Party_Name: row[partyIdx] ? row[partyIdx].toString().trim() : "",
        Amount: amount,
        Is_Overdue: isOverdue
      });
    }
  }
}

var incomingUpcoming = stats.INCOMING.today + stats.INCOMING.week;
var outgoingUpcoming = stats.OUTGOING.overdue + stats.OUTGOING.today + stats.OUTGOING.week;
var totalAvailableLiquidity = totalActualBankBalance + totalMTFBalance + incomingUpcoming;

stats.liquidityStatus = computeLiquidityStatus(totalAvailableLiquidity, outgoingUpcoming);
stats.netLiquiditySnapshot = (totalAvailableLiquidity - outgoingUpcoming) - CASH_FLOOR_RESERVE;
stats.overallConfidence = minConfidence;
stats.cashFloor = CASH_FLOOR_RESERVE;

setCache_("CACHE_DASHBOARD_STATS", stats);
return stats;
} catch (error) {
return { error: error.toString() };
}
}

function getActiveChequesList(sortDir) { return getChequesByStatusFilter(false, sortDir || "ASC"); }
function getClearedChequesList(sortDir) { return getChequesByStatusFilter(true, sortDir || "DESC"); }

function getChequesByStatusFilter(isCleared, sortDir) {
try {
var dir0 = (sortDir === "DESC") ? "DESC" : "ASC";
var cacheKey = (isCleared ? "CACHE_CLEARED_" : "CACHE_ACTIVE_") + dir0;
var cached = getCache_(cacheKey);
if (cached) return cached;

var sheet = getSheet("CHEQUES");
if (!sheet || sheet.getLastRow() <= 1) return [];
var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });

var idIdx = headers.indexOf("Cheque_ID"); if(idIdx===-1) idIdx=0;
var typeIdx = headers.indexOf("Type"); if(typeIdx===-1) typeIdx=1;
var partyIdx = headers.indexOf("Party_Name"); if(partyIdx===-1) partyIdx=2;
var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx===-1) bankIdx=3;
var numIdx = headers.indexOf("Reference_Number / Cheque_Number"); if(numIdx===-1) numIdx=4;
var amtIdx = headers.indexOf("Amount"); if(amtIdx===-1) amtIdx=5;
var dueIdx = headers.indexOf("Due_Date"); if(dueIdx===-1) dueIdx=7;
var statIdx = headers.indexOf("Is_Settled"); if(statIdx===-1) statIdx=18;
var clearedDateIdx = headers.indexOf("Cleared_Date"); if(clearedDateIdx===-1) clearedDateIdx=10;
var origIdx = headers.indexOf("Drive_Original"); if(origIdx===-1) origIdx=11;
var clearedImgIdx = headers.indexOf("Drive_Cleared"); if(clearedImgIdx===-1) clearedImgIdx=12;
var instIdx = headers.indexOf("Instrument_Type"); if(instIdx===-1) instIdx=17;

var formImagesMap = getImagesMapFromFormResponses();
var cheques = [];

for (var i=1; i<data.length; i++) {
  var row = data[i];
  var status = row[statIdx] ? row[statIdx].toString().trim() : "";
  var itemCleared = isItemCleared(status);
  if (isCleared !== itemCleared) continue;

  var chequeNum = row[numIdx] ? row[numIdx].toString().trim() : "";
  var dueStr = formatDateStr(row[dueIdx]);
  var clearedStr = formatDateStr(row[clearedDateIdx]);
  var instType = row[instIdx] ? row[instIdx].toString().trim() : "CHEQUE";
  var originalUrl = row[origIdx] || formImagesMap[chequeNum] || "";

  cheques.push({
    Cheque_ID: row[idIdx] || ("CHQ-"+chequeNum),
    Type: row[typeIdx] || "INCOMING",
    Party_Name: row[partyIdx] || "",
    Bank_Name: row[bankIdx] || "",
    Cheque_Number: chequeNum,
    Amount: parseFloat(row[amtIdx]) || 0,
    Due_Date: dueStr,
    Cleared_Date: clearedStr,
    Status: status,
    Drive_Original: originalUrl,
    Drive_Cleared: row[clearedImgIdx] || "",
    Instrument_Type: instType
  });
}

var dir = (sortDir === "DESC") ? -1 : 1;
cheques.sort(function(a, b) {
  var da = a.Due_Date ? parseLocalDate(a.Due_Date) : null;
  var db = b.Due_Date ? parseLocalDate(b.Due_Date) : null;
  if (!da && !db) return 0;
  if (!da) return 1;
  if (!db) return -1;
  return dir * (da.getTime() - db.getTime());
});

setCache_(cacheKey, cheques);
return cheques;
} catch (error) {
return [];
}
}

function refreshMissingImageLinks() {
var lock = LockService.getScriptLock();
try {
lock.waitLock(10000);
var sheet = getSheet("CHEQUES");
if (!sheet || sheet.getLastRow() <= 1) return { success: true, updated: 0 };
var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });
var numIdx = headers.indexOf("Reference_Number / Cheque_Number"); if(numIdx===-1) numIdx=4;
var origIdx = headers.indexOf("Drive_Original"); if(origIdx===-1) origIdx=11;
var instIdx = headers.indexOf("Instrument_Type"); if(instIdx===-1) instIdx=17;

var updated = 0;
for (var i = 1; i < data.length; i++) {
  var chequeNum = data[i][numIdx] ? data[i][numIdx].toString().trim() : "";
  var existing = data[i][origIdx];
  if (!existing && chequeNum) {
    var instType = data[i][instIdx] ? data[i][instIdx].toString().trim() : "CHEQUE";
    var found = findImageInDriveByChequeNumber(chequeNum, instType);
    if (found) {
      sheet.getRange(i + 1, origIdx + 1).setValue(found);
      updated++;
    }
  }
}
clearAppCache_();
return { success: true, updated: updated };
} catch (error) {
return { success: false, error: error.toString() };
} finally {
  lock.releaseLock();
}
}

function saveChequeToSheet(chequeData) {
var lock = LockService.getScriptLock();
try {
lock.waitLock(10000);
var sheet = getSheet("CHEQUES");
var chequeNum = chequeData.Cheque_Number ? chequeData.Cheque_Number.toString().trim() : "";
if (!chequeNum) return { success:false, error:"رقم المعاملة أو الشيك إلزامي." };

var prefix = (chequeData.Instrument_Type === "BANK_TRANSFER") ? "TRF-" : "CHQ-";
var chequeId = prefix + chequeNum;
var now = new Date();
sheet.appendRow([
  chequeId,
  chequeData.Type || "INCOMING",
  chequeData.Party_Name || "",
  chequeData.Bank_Name || "",
  chequeNum,
  parseFloat(chequeData.Amount) || 0,
  chequeData.Cheque_Date || "",
  chequeData.Due_Date || "",
  chequeData.Instrument_Type === "BANK_TRANSFER" ? "تم التنفيذ" : "مستلم",
  Utilities.formatDate(now, TZ, "yyyy-MM-dd"),
  "", "", "", "",
  chequeData.Notes || "",
  now, now,
  chequeData.Instrument_Type || "CHEQUE",
  "N",
  chequeData.Invoice_Date || ""
]);

clearAppCache_();
return { success:true, chequeId: chequeId, chequeNumber: chequeNum };
} catch (error) {
return { success:false, error: "خطأ في حفظ البيانات: " + error.toString() };
} finally {
lock.releaseLock();
}
}
function quickClearChequesBatch(chequeNumbers, clearanceNote) {
var lock = LockService.getScriptLock();
try {
lock.waitLock(15000);
var sheet = getSheet("CHEQUES");
if (!sheet) return { success:false, error:"لم يتم العثور على جدول المعاملات." };

var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });

var numIdx = headers.indexOf("Reference_Number / Cheque_Number"); if(numIdx===-1) numIdx=4;
var statusIdx = headers.indexOf("Status"); if(statusIdx===-1) statusIdx=8;
var settledIdx = headers.indexOf("Is_Settled"); if(settledIdx===-1) settledIdx=18;
var clearedDateIdx = headers.indexOf("Cleared_Date"); if(clearedDateIdx===-1) clearedDateIdx=10;
var notesIdx = headers.indexOf("Notes"); if(notesIdx===-1) notesIdx=14;
var updIdx = headers.indexOf("Updated_At"); if(updIdx===-1) updIdx=16;
var origIdx = headers.indexOf("Drive_Original"); if(origIdx===-1) origIdx=11;
var clearedImgIdx = headers.indexOf("Drive_Cleared"); if(clearedImgIdx===-1) clearedImgIdx=12;
var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx===-1) bankIdx=3; // إضافة التحقق من البنك

var updatedCount = 0;
var clearedFolder = null;

for (var i=1; i<data.length; i++) {
  var rowNum = data[i][numIdx] ? data[i][numIdx].toString().trim() : "";
  if (chequeNumbers.indexOf(rowNum) !== -1) {
    var status = data[i][statusIdx] ? data[i][statusIdx].toString().trim() : "";
    if(isItemCleared(status)) continue;

    // التحقق الصارم من وجود بنك محدد للخصم أو الإيداع
    var bankName = data[i][bankIdx] ? data[i][bankIdx].toString().trim() : "";
    if (!bankName || bankName === "غير محدد" || bankName === "-") {
        return { success:false, error: "المعاملة رقم " + rowNum + " لا تحتوي على بنك محدد. يرجى تعديلها يدوياً وإضافة البنك ليتم الخصم أو الإيداع منه بشكل صحيح." };
    }

    var r = i+1;
    sheet.getRange(r, statusIdx+1).setValue("تم التحصيل");
    sheet.getRange(r, settledIdx+1).setValue("Y");
    sheet.getRange(r, clearedDateIdx+1).setValue(Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"));

    var origUrl = data[i][origIdx];
    if (origUrl) {
      try {
        var m = origUrl.toString().match(/\/d\/(.*?)\/(view|edit)/) || origUrl.toString().match(/id=([a-zA-Z0-9_-]+)/);
        if (m && m[1]) {
          var file = DriveApp.getFileById(m[1]);
          if (!clearedFolder) clearedFolder = getOrCreateFolder("02_CLEARED_CHEQUES");
          file.moveTo(clearedFolder);
          sheet.getRange(r, clearedImgIdx+1).setValue(file.getUrl());
          sheet.getRange(r, origIdx+1).setValue("");
        }
      } catch(e) {}
    }
    if (clearanceNote) {
      var oldNote = data[i][notesIdx];
      sheet.getRange(r, notesIdx+1).setValue(oldNote ? oldNote + " | " + clearanceNote : clearanceNote);
    }
    sheet.getRange(r, updIdx+1).setValue(new Date());
    updatedCount++;
  }
}

if (updatedCount > 0) {
  clearAppCache_();
  return { success:true, updated: updatedCount };
} else {
  return { success:false, error: "لم يتم العثور على أرقام المعاملات أو تم تحصيلها مسبقاً." };
}
} catch (error) {
return { success:false, error: error.toString() };
} finally {
lock.releaseLock();
}
}

function quickClearCheque(chequeNumber, clearanceNote) {
return quickClearChequesBatch([chequeNumber.toString().trim()], clearanceNote);
}
function getBankAccountsPanel() {
try {
var cachedPanel = getCache_("CACHE_BANK_PANEL");
if (cachedPanel) return cachedPanel;

var settingsSheet = getSheet("SETTINGS");
var sData = settingsSheet.getDataRange().getValues();
var sHeaders = sData[0].map(function(h){ return h.toString().trim(); });

var sBankIdx = sHeaders.indexOf("Bank_Name"); if(sBankIdx===-1) sBankIdx=0;
var openBalIdx = sHeaders.indexOf("Opening_Balance"); if(openBalIdx===-1) openBalIdx=1;
var actBalIdx = sHeaders.indexOf("Last_Actual_Balance"); if(actBalIdx===-1) actBalIdx=2;
var recDateIdx = sHeaders.indexOf("Last_Reconciliation_Date"); if(recDateIdx===-1) recDateIdx=3;

var settingsMap = {};
var bankRecTimes = {}; 

for (var i=1; i<sData.length; i++) {
  var bName = sData[i][sBankIdx] ? sData[i][sBankIdx].toString().trim() : "";
  if (bName) {
    var recDateStr = formatDateStr(sData[i][recDateIdx]);
    settingsMap[bName] = {
      openingBalance: parseFloat(sData[i][openBalIdx]) || 0,
      actualBalance: parseFloat(sData[i][actBalIdx]) || 0,
      reconciliationDate: recDateStr
    };
    var rd = parseLocalDate(recDateStr);
    bankRecTimes[bName] = rd ? rd.getTime() : 0;
  }
}

var chequesSheet = getSheet("CHEQUES");
var clearedCheques = { INCOMING: {}, OUTGOING: {} };
var postRecClearedCheques = { INCOMING: {}, OUTGOING: {} }; 

if (chequesSheet && chequesSheet.getLastRow() > 1) {
  var cData = chequesSheet.getDataRange().getValues();
  var cHeaders = cData[0].map(function(h){ return h.toString().trim(); });
  var cTypeIdx = cHeaders.indexOf("Type"); if(cTypeIdx===-1) cTypeIdx=1;
  var cBankIdx = cHeaders.indexOf("Bank_Name"); if(cBankIdx===-1) cBankIdx=3;
  var cAmtIdx = cHeaders.indexOf("Amount"); if(cAmtIdx===-1) cAmtIdx=5;
  var cStatIdx = cHeaders.indexOf("Is_Settled"); if(cStatIdx===-1) cStatIdx=18;
  var clearedDateIdx = cHeaders.indexOf("Cleared_Date"); if(clearedDateIdx===-1) clearedDateIdx=10;

  for (var j=1; j<cData.length; j++) {
    var row = cData[j];
    var type = row[cTypeIdx] ? row[cTypeIdx].toString().trim() : "";
    var bank = row[cBankIdx] ? row[cBankIdx].toString().trim() : "";
    var status = row[cStatIdx] ? row[cStatIdx].toString().trim() : "";
    var amount = parseFloat(row[cAmtIdx]) || 0;

    if (bank && (type === "INCOMING" || type === "OUTGOING") && isItemCleared(status)) {
      clearedCheques[type][bank] = (clearedCheques[type][bank] || 0) + amount;
      
      var cDateObj = parseLocalDate(formatDateStr(row[clearedDateIdx]));
      var cTime = cDateObj ? cDateObj.getTime() : 0;
      var recTime = bankRecTimes[bank] || 0;
      
      // التعديل الجذري هنا: >= بدلاً من > لضمان خصم معاملات نفس اليوم
      if (cTime >= recTime) {
        postRecClearedCheques[type][bank] = (postRecClearedCheques[type][bank] || 0) + amount;
      }
    }
  }
}

var logSheet = getSheet("LOG");
var logSums = { IN: {}, OUT: {} };
var postRecLogSums = { IN: {}, OUT: {} }; 

if (logSheet && logSheet.getLastRow() > 1) {
  var lData = logSheet.getDataRange().getValues();
  var lHeaders = lData[0].map(function(h){ return h.toString().trim(); });
  var lTypeIdx = lHeaders.indexOf("Entry_Type"); if(lTypeIdx===-1) lTypeIdx=1;
  var lBankIdx = lHeaders.indexOf("Bank_Name"); if(lBankIdx===-1) lBankIdx=2;
  var lAmtIdx = lHeaders.indexOf("Amount"); if(lAmtIdx===-1) lAmtIdx=3;
  var lDirIdx = lHeaders.indexOf("Direction"); if(lDirIdx===-1) lDirIdx=4;
  var lDateIdx = lHeaders.indexOf("Date"); if(lDateIdx===-1) lDateIdx=6;

  for (var k=1; k<lData.length; k++) {
    var lRow = lData[k];
    var entryType = lRow[lTypeIdx] ? lRow[lTypeIdx].toString().trim() : "";
    var bank = lRow[lBankIdx] ? lRow[lBankIdx].toString().trim() : "";
    var dir = lRow[lDirIdx] ? lRow[lDirIdx].toString().trim() : "";
    var amt = parseFloat(lRow[lAmtIdx]) || 0;

    if (bank && (entryType === "MANUAL" || entryType === "AUTO_RECURRING") && (dir === "IN" || dir === "OUT")) {
      logSums[dir][bank] = (logSums[dir][bank] || 0) + amt;
      
      var lDateObj = parseLocalDate(formatDateStr(lRow[lDateIdx]));
      var lTime = lDateObj ? lDateObj.getTime() : 0;
      var recTime = bankRecTimes[bank] || 0;
      if (lTime >= recTime) { // التعديل هنا أيضاً
        postRecLogSums[dir][bank] = (postRecLogSums[dir][bank] || 0) + amt;
      }
    }
  }
}

var panelResult = [];
var today = new Date(); today.setHours(0,0,0,0);

BANKS.forEach(function(bank) {
  var st = settingsMap[bank] || { openingBalance: 0, actualBalance: 0, reconciliationDate: "" };
  var incCleared = clearedCheques.INCOMING[bank] || 0;
  var outCleared = clearedCheques.OUTGOING[bank] || 0;
  var logIn = logSums.IN[bank] || 0;
  var logOut = logSums.OUT[bank] || 0;

  var expectedBalance = st.openingBalance + incCleared - outCleared + logIn - logOut;
  var variance = st.actualBalance - expectedBalance;
  var variancePercent = st.actualBalance !== 0 ? Math.abs(variance) / st.actualBalance * 100 : null;

  var incPostRec = postRecClearedCheques.INCOMING[bank] || 0;
  var outPostRec = postRecClearedCheques.OUTGOING[bank] || 0;
  var logInPostRec = postRecLogSums.IN[bank] || 0;
  var logOutPostRec = postRecLogSums.OUT[bank] || 0;
  var estimatedActualBalance = st.actualBalance + incPostRec - outPostRec + logInPostRec - logOutPostRec;

  var confidence = "low";
  if (st.reconciliationDate) {
    var recDateObj = parseLocalDate(st.reconciliationDate);
    if (recDateObj) {
      recDateObj.setHours(0,0,0,0);
      var diffDays = Math.floor((today.getTime() - recDateObj.getTime()) / 86400000);
      if (diffDays <= 3) confidence = "high";
      else if (diffDays <= 7) confidence = "medium";
      else confidence = "low";
    }
  }

  panelResult.push({
    bank: bank,
    expectedBalance: expectedBalance,
    actualBalance: st.actualBalance,
    estimatedActualBalance: estimatedActualBalance,
    variance: variance,
    variancePercent: variancePercent,
    confidence: confidence,
    lastReconciliationDate: st.reconciliationDate
  });
});

setCache_("CACHE_BANK_PANEL", panelResult);
return panelResult;
} catch (error) {
return [];
}
}

function getHomeScreenData() {
try {
var stats = getDashboardStats();
if (!stats) stats = getEmptyStats();
if (stats.error) {
stats.bankPanels = [];
return stats;
}
stats.bankPanels = getBankAccountsPanel();
return stats;
} catch (error) {
return { error: error.toString(), bankPanels: [] };
}
}

function saveQuickTransaction(bankName, amount, direction, reason, toBankName) {
var lock = LockService.getScriptLock();
try {
lock.waitLock(10000);
ensureSystemTabs();
var amt = parseFloat(amount);
if (isNaN(amt) || amt <= 0) return { success: false, error: "المبلغ غير صحيح." };

var logSheet = getSheet("LOG");
var now = new Date();
var dateStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");

if (direction === "TRANSFER") {
  if (!bankName || !toBankName || bankName === toBankName) return { success: false, error: "اختر حسابين مختلفين للتحويل." };
  if (BANKS.indexOf(bankName) === -1 || BANKS.indexOf(toBankName) === -1) return { success: false, error: "البنك غير صحيح." };
  
  logSheet.appendRow(["LOG-" + now.getTime() + "A", "MANUAL", bankName, amt, "OUT", "تحويل إلى " + toBankName + (reason ? " - "+reason : ""), dateStr, now]);
  logSheet.appendRow(["LOG-" + now.getTime() + "B", "MANUAL", toBankName, amt, "IN", "تحويل من " + bankName + (reason ? " - "+reason : ""), dateStr, now]);
  
} else {
  if (BANKS.indexOf(bankName) === -1) return { success: false, error: "البنك غير صحيح." };
  if (direction !== "IN" && direction !== "OUT") return { success: false, error: "الاتجاه غير صحيح." };
  logSheet.appendRow(["LOG-" + now.getTime(), "MANUAL", bankName, amt, direction, reason || "حركة سريعة", dateStr, now]);
}

clearAppCache_();
return { success: true };
} catch (error) {
return { success: false, error: error.toString() };
} finally {
lock.releaseLock();
}
}

function saveReconciliation(bankName, actualBalance) {
var lock = LockService.getScriptLock();
try {
lock.waitLock(10000);
ensureSystemTabs();
if (BANKS.indexOf(bankName) === -1) return { success: false, error: "البنك غير صحيح." };
var actBal = parseFloat(actualBalance);
if (isNaN(actBal)) return { success: false, error: "الرصيد غير صحيح." };

var settingsSheet = getSheet("SETTINGS");
var sData = settingsSheet.getDataRange().getValues();
var todayStr = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
var threshold = 10;
var expectedCalculated = 0;

var panel = getBankAccountsPanel();
for (var p = 0; p < panel.length; p++) {
  if (panel[p].bank === bankName) expectedCalculated = panel[p].expectedBalance;
}
var varianceValue = actBal - expectedCalculated;

for (var i = 1; i < sData.length; i++) {
  if (sData[i][0].toString().trim() === bankName) {
    threshold = parseFloat(sData[i][4]) || 10;
    settingsSheet.getRange(i + 1, 3).setValue(actBal);
    settingsSheet.getRange(i + 1, 4).setValue(todayStr);
    break;
  }
}

var logSheet = getSheet("LOG");
var now = new Date();
logSheet.appendRow(["REC-" + now.getTime(), "RECONCILE", bankName, varianceValue, "IN", "تسوية رصيد فعلي", todayStr, now]);

clearAppCache_();
var newPanel = getBankAccountsPanel();
for (var p2 = 0; p2 < newPanel.length; p2++) {
  if (newPanel[p2].bank === bankName && newPanel[p2].variancePercent !== null && newPanel[p2].variancePercent > threshold) {
    try {
      var recipient = "adhamamro55@gmail.com";
      var subject = "⚠️ تنبيه انحراف رصيد البنك: " + bankName;
      var body = "مرحباً،\n\nتجاوزت نسبة الفارق بين الرصيد المتوقع والفعلـي للبنك " + bankName + " نسبة التنبيه (" + threshold + "%).\nنسبة الفارق: " + newPanel[p2].variancePercent.toFixed(2) + "%\n\nنظام الخزينة.";
      GmailApp.sendEmail(recipient, subject, body);
    } catch(e) {}
  }
}
return { success: true };
} catch (error) {
return { success: false, error: error.toString() };
} finally {
lock.releaseLock();
}
}

function getVarianceTrend(bankName) {
try {
ensureSystemTabs();
var logSheet = getSheet("LOG");
if (!logSheet || logSheet.getLastRow() <= 1) return [];
var data = logSheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });

var typeIdx = headers.indexOf("Entry_Type"); if(typeIdx===-1) typeIdx=1;
var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx===-1) bankIdx=2;
var amtIdx = headers.indexOf("Amount"); if(amtIdx===-1) amtIdx=3;
var dateIdx = headers.indexOf("Date"); if(dateIdx===-1) dateIdx=6;

var trend = [];
for (var i = 1; i < data.length; i++) {
  var row = data[i];
  if (row[typeIdx] && row[typeIdx].toString().trim() === "RECONCILE" && row[bankIdx] && row[bankIdx].toString().trim() === bankName) {
    trend.push({ date: formatDateStr(row[dateIdx]), actual: parseFloat(row[amtIdx]) || 0 });
  }
}
return trend;
} catch (error) {
return [];
}
}

function checkReconciliationReminders() {
try {
ensureSystemTabs();
var panel = getBankAccountsPanel();
var overdueBanks = [];
var today = new Date(); today.setHours(0,0,0,0);

panel.forEach(function(item) {
  if (!item.lastReconciliationDate) {
    overdueBanks.push(item.bank + " (لم يتم التسوية نهائياً)");
  } else {
    var d = parseLocalDate(item.lastReconciliationDate);
    if (d) {
      d.setHours(0,0,0,0);
      var diffDays = Math.floor((today.getTime() - d.getTime()) / 86400000);
      if (diffDays > 7) overdueBanks.push(item.bank + " (منذ " + diffDays + " يوم)");
    }
  }
});

if (overdueBanks.length > 0) {
  var recipient = "adhamamro55@gmail.com";
  var subject = "⏰ تذكير دوري: بنوك تحتاج إلى تسوية";
  var body = "مرحباً،\n\nالبنوك التالية تجاوزت 7 أيام بدون تسوية:\n- " + overdueBanks.join("\n- ") + "\n\nنظام الخزينة.";
  GmailApp.sendEmail(recipient, subject, body);
}
} catch(e) {}
}

function saveRecurringTemplate(bankName, amount, direction, reason, dayOfMonth) {
try {
ensureSystemTabs();
if (BANKS.indexOf(bankName) === -1) return { success: false, error: "البنك غير صحيح." };
var logSheet = getSheet("LOG");
var now = new Date();
logSheet.appendRow(["TMP-" + now.getTime(), "TEMPLATE", bankName, parseFloat(amount) || 0, direction, reason || "قالب", dayOfMonth.toString(), now]);
return { success: true };
} catch (error) {
return { success: false, error: error.toString() };
}
}

function runMonthlyRecurringEntries() {
var lock = LockService.getScriptLock();
try {
lock.waitLock(15000);
ensureSystemTabs();
var logSheet = getSheet("LOG");
if (!logSheet || logSheet.getLastRow() <= 1) return;
var data = logSheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });

var typeIdx = headers.indexOf("Entry_Type"); if(typeIdx===-1) typeIdx=1;
var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx===-1) bankIdx=2;
var amtIdx = headers.indexOf("Amount"); if(amtIdx===-1) amtIdx=3;
var dirIdx = headers.indexOf("Direction"); if(dirIdx===-1) dirIdx=4;
var reasonIdx = headers.indexOf("Reason"); if(reasonIdx===-1) reasonIdx=5;
var dateColIdx = headers.indexOf("Date"); if(dateColIdx===-1) dateColIdx=6;

var today = new Date();
var currentDay = today.getDate();
var currentYearMonth = Utilities.formatDate(today, TZ, "yyyy-MM");
var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

var alreadyCopied = {};
for (var i = 1; i < data.length; i++) {
  var t = data[i][typeIdx] ? data[i][typeIdx].toString().trim() : "";
  if (t === "AUTO_RECURRING") {
    var created = data[i][7];
    if (created instanceof Date) {
      var ym = Utilities.formatDate(created, TZ, "yyyy-MM");
      var rsn = data[i][reasonIdx] ? data[i][reasonIdx].toString().trim() : "";
      if (ym === currentYearMonth) alreadyCopied[rsn] = true;
    }
  }
}

for (var j = 1; j < data.length; j++) {
  var row = data[j];
  var t = row[typeIdx] ? row[typeIdx].toString().trim() : "";
  if (t === "TEMPLATE") {
    var targetDay = parseInt(row[dateColIdx], 10);
    var effectiveTargetDay = targetDay > lastDayOfMonth ? lastDayOfMonth : targetDay;
    var reason = row[reasonIdx] ? row[reasonIdx].toString().trim() : "قالب";
    if (effectiveTargetDay === currentDay && !alreadyCopied[reason]) {
      logSheet.appendRow(["AUT-" + new Date().getTime() + "-" + j, "AUTO_RECURRING", row[bankIdx], row[amtIdx], row[dirIdx], reason, Utilities.formatDate(today, TZ, "yyyy-MM-dd"), new Date()]);
    }
  }
}
clearAppCache_();
} catch(e) {} finally { lock.releaseLock(); }
}

function runMonthlyLoanInstallments() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    ensureSystemTabs();
    var loanSheet = getSheet("LOANS");
    if (!loanSheet || loanSheet.getLastRow() <= 1) return;
    
    var lData = loanSheet.getDataRange().getValues();
    var lHeaders = lData[0].map(function(h){ return h.toString().trim(); });
    
    var idIdx = lHeaders.indexOf("Loan_ID"); if (idIdx === -1) idIdx = 0;
    var nameIdx = lHeaders.indexOf("Loan_Name"); if (nameIdx === -1) nameIdx = 1;
    var bankIdx = lHeaders.indexOf("Bank_Name"); if (bankIdx === -1) bankIdx = 2;
    var amtIdx = lHeaders.indexOf("Monthly_Amount"); if (amtIdx === -1) amtIdx = 3;
    var startIdx = lHeaders.indexOf("Start_Date"); if (startIdx === -1) startIdx = 4;
    var endIdx = lHeaders.indexOf("End_Date"); if (endIdx === -1) endIdx = 5;
    var totalInsIdx = lHeaders.indexOf("Total_Installments"); if (totalInsIdx === -1) totalInsIdx = 6;
    var dayIdx = lHeaders.indexOf("Day_Of_Month"); if (dayIdx === -1) dayIdx = 7;
    var activeIdx = lHeaders.indexOf("Active"); if (activeIdx === -1) activeIdx = 8;

    var today = new Date();
    today.setHours(0,0,0,0);
    var currentDay = today.getDate();
    var currentYearMonth = Utilities.formatDate(today, TZ, "yyyy-MM");
    var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

    var chequesSheet = getSheet("CHEQUES");
    var cData = chequesSheet.getDataRange().getValues();
    var cHeaders = cData[0].map(function(h){ return h.toString().trim(); });
    var cNumIdx = cHeaders.indexOf("Reference_Number / Cheque_Number"); if (cNumIdx === -1) cNumIdx = 4;
    var cDueIdx = cHeaders.indexOf("Due_Date"); if (cDueIdx === -1) cDueIdx = 7;

  var existingInstallments = {};
    for (var c = 1; c < cData.length; c++) {
      var refNum = cData[c][cNumIdx] ? cData[c][cNumIdx].toString().trim() : "";
      if (refNum) {
        existingInstallments[refNum] = true;
      }
    }

    for (var i = 1; i < lData.length; i++) {
      var row = lData[i];
      var isActive = row[activeIdx] ? row[activeIdx].toString().trim() : "";
      if (isActive !== "Y") continue;

      var loanId = row[idIdx] ? row[idIdx].toString().trim() : "";
      var loanName = row[nameIdx] ? row[nameIdx].toString().trim() : "";
      var bankName = row[bankIdx] ? row[bankIdx].toString().trim() : "NBE";
      var monthlyAmt = parseFloat(row[amtIdx]) || 0;
      var targetDay = parseInt(row[dayIdx], 10) || 1;
      var effectiveTargetDay = targetDay > lastDayOfMonth ? lastDayOfMonth : targetDay;
      
      var startDate = parseLocalDate(formatDateStr(row[startIdx]));
      var endDate = parseLocalDate(formatDateStr(row[endIdx]));
      
      // تم إلغاء شرط التطابق اليومي لتعمل الدالة من أول يوم في الشهر
      if (startDate && today.getTime() < startDate.getTime()) continue;
      if (endDate && today.getTime() > endDate.getTime()) continue;

      var expectedRefNum = loanId + "-" + currentYearMonth.replace("-", "");
      if (!existingInstallments[expectedRefNum]) {
        var dueCurrentStr = currentYearMonth + "-" + (effectiveTargetDay < 10 ? "0" + effectiveTargetDay : effectiveTargetDay);
        chequesSheet.appendRow([
            "CHQ-" + loanId + "-" + currentYearMonth.replace("-", ""),
            "OUTGOING", bankName, bankName,
            loanId + "-" + currentYearMonth.replace("-", ""),
            monthlyAmt, Utilities.formatDate(today, TZ, "yyyy-MM-dd"), dueCurrentStr,
            "تم التنفيذ", Utilities.formatDate(today, TZ, "yyyy-MM-dd"), "", "", "", "",
            "[قسط بنك] " + loanName, today, today, "BANK_TRANSFER", "N", ""
        ]);
      }
    }
    clearAppCache_();
  } catch(e) {} finally { lock.releaseLock(); }
}

function archiveOldClearedCheques(monthsOld) {
try {
monthsOld = parseInt(monthsOld, 10) || 6;
var sheet = getSheet("CHEQUES");
if (!sheet || sheet.getLastRow() <= 1) return { success: true, archived: 0 };

var data = sheet.getDataRange().getValues();
var headers = data[0];
var headerNames = headers.map(function(h){ return h.toString().trim(); });
var statIdx = headerNames.indexOf("Is_Settled"); if (statIdx === -1) statIdx = 18;
var clearedDateIdx = headerNames.indexOf("Cleared_Date"); if (clearedDateIdx === -1) clearedDateIdx = 10;

var archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CHEQUES_ARCHIVE");
if (!archiveSheet) {
  archiveSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("CHEQUES_ARCHIVE");
  archiveSheet.appendRow(headers);
}

var cutoff = new Date();
cutoff.setMonth(cutoff.getMonth() - monthsOld);
cutoff.setHours(0,0,0,0);

var rowsToArchive = [];
var rowIndexesToDelete = [];

for (var i = 1; i < data.length; i++) {
  var status = data[i][statIdx] ? data[i][statIdx].toString().trim() : "";
  if (!isItemCleared(status)) continue;
  var clearedStr = formatDateStr(data[i][clearedDateIdx]);
  var clearedDateObj = parseLocalDate(clearedStr);
  if (clearedDateObj && clearedDateObj.getTime() < cutoff.getTime()) {
    rowsToArchive.push(data[i]);
    rowIndexesToDelete.push(i + 1);
  }
}

if (rowsToArchive.length > 0) {
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, headers.length).setValues(rowsToArchive);
  rowIndexesToDelete.sort(function(a,b){ return b - a; });
  rowIndexesToDelete.forEach(function(rowNum) { sheet.deleteRow(rowNum); });
  clearAppCache_();
}
return { success: true, archived: rowsToArchive.length };
} catch (error) {
return { success: false, error: error.toString() };
}
}

function setupTriggers() {
  var scripts = ScriptApp.getProjectTriggers();
  scripts.forEach(function(trig) { ScriptApp.deleteTrigger(trig); });
  
  // التعديل هنا: إضافة onWeekDay لتحديد يوم الأسبوع (الخميس مثلاً)
  ScriptApp.newTrigger("checkReconciliationReminders").timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.THURSDAY).create();
  
  ScriptApp.newTrigger("runMonthlyRecurringEntries").timeBased().everyDays(1).create();
  ScriptApp.newTrigger("runMonthlyLoanInstallments").timeBased().everyDays(1).create();
  ScriptApp.newTrigger("dailyMTFTriggerAlert").timeBased().atHour(9).everyDays(1).create();
}

// محرك مدرج السيولة اليومي (Daily Liquidity Ladder Engine)
function getDailyLiquidityLadder(daysCount) {
  try {
    daysCount = parseInt(daysCount, 10) || 30;
    var cacheKey = "CACHE_DAILY_LADDER_" + daysCount;
    var cached = getCache_(cacheKey);
    if (cached) return cached;

    var bankPanels = getBankAccountsPanel();
    var openingBalance = 0;
    bankPanels.forEach(function(p) {
      openingBalance += p.estimatedActualBalance;
    });

    var sheet = getSheet("CHEQUES");
    var dailyMap = {};
    
    var today = new Date(); today.setHours(0,0,0,0);

    // دمج الالتزامات والأقساط الثابتة من LOANS
    var loanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOANS");
    var loansList = [];
    if (loanSheet && loanSheet.getLastRow() > 1) {
      var lData = loanSheet.getDataRange().getValues();
      var lHeaders = lData[0].map(function(h){ return h.toString().trim(); });
      var lAmtIdx = lHeaders.indexOf("Monthly_Amount"); if (lAmtIdx === -1) lAmtIdx = 3;
      var lStartIdx = lHeaders.indexOf("Start_Date"); if (lStartIdx === -1) lStartIdx = 4;
      var lEndIdx = lHeaders.indexOf("End_Date"); if (lEndIdx === -1) lEndIdx = 5;
      var lDayIdx = lHeaders.indexOf("Day_Of_Month"); if (lDayIdx === -1) lDayIdx = 7;
      var lActiveIdx = lHeaders.indexOf("Active"); if (lActiveIdx === -1) lActiveIdx = 8;
      var lNameIdx = lHeaders.indexOf("Loan_Name"); if (lNameIdx === -1) lNameIdx = 1;

      for (var l = 1; l < lData.length; l++) {
        if (lData[l][lActiveIdx] && lData[l][lActiveIdx].toString().trim() === "Y") {
          loansList.push({
            name: lData[l][lNameIdx] || "قسط بنك",
            amount: parseFloat(lData[l][lAmtIdx]) || 0,
            day: parseInt(lData[l][lDayIdx], 10) || 1,
            start: parseLocalDate(formatDateStr(lData[l][lStartIdx])),
            end: parseLocalDate(formatDateStr(lData[l][lEndIdx]))
          });
        }
      }
    }

    // توليد الأيام مستقبلاً وتعبئة الأقساط الثابتة
    var ladderDays = [];
    var runningBalance = openingBalance;

    for (var d = 0; d < daysCount; d++) {
      var currDate = new Date(today.getTime() + d * 86400000);
      var dateStr = Utilities.formatDate(currDate, TZ, "yyyy-MM-dd");
      var dayName = Utilities.formatDate(currDate, TZ, "EEEE"); // اسم اليوم
      
      var dayOutflows = 0;
      var detailsOut = [];

      // التحقق من الأقساط في هذا اليوم من الشهر
      var cDay = currDate.getDate();
      loansList.forEach(function(ln) {
        var lastDay = new Date(currDate.getFullYear(), currDate.getMonth() + 1, 0).getDate();
        var effDay = ln.day > lastDay ? lastDay : ln.day;
        if (cDay === effDay) {
          if ((!ln.start || currDate.getTime() >= ln.start.getTime()) && (!ln.end || currDate.getTime() <= ln.end.getTime())) {
            dayOutflows += ln.amount;
            detailsOut.push({ party: ln.name, amount: ln.amount });
          }
        }
      });

      dailyMap[dateStr] = {
        date: dateStr,
        dayName: dayName,
        inflows: 0,
        outflows: dayOutflows,
        detailsIn: [],
        detailsOut: detailsOut
      };
    }

    // جلب الشيكات والتحويلات المعلقة من شيت CHEQUES
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h){ return h.toString().trim(); });
      var typeIdx = headers.indexOf("Type"); if (typeIdx === -1) typeIdx = 1;
      var amtIdx = headers.indexOf("Amount"); if (amtIdx === -1) amtIdx = 5;
      var dueIdx = headers.indexOf("Due_Date"); if (dueIdx === -1) dueIdx = 7;
      var statIdx = headers.indexOf("Is_Settled"); if (statIdx === -1) statIdx = 18;
      var partyIdx = headers.indexOf("Party_Name"); if (partyIdx === -1) partyIdx = 2;

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
        var status = row[statIdx] ? row[statIdx].toString().trim() : "";
        if (isItemCleared(status)) continue; // تخطي المسدد

        var amount = parseFloat(row[amtIdx]) || 0;
        var dueStr = formatDateStr(row[dueIdx]);
        if (!dueStr || !dailyMap[dueStr]) continue;

        var party = row[partyIdx] ? row[partyIdx].toString().trim() : "غير محدد";

        if (type === "INCOMING") {
          dailyMap[dueStr].inflows += amount;
          dailyMap[dueStr].detailsIn.push({ party: party, amount: amount });
        } else if (type === "OUTGOING") {
          dailyMap[dueStr].outflows += amount;
          dailyMap[dueStr].detailsOut.push({ party: party, amount: amount });
        }
      }
    }

    // تجميع الحسابات اليومية التصاعدية (Running Balance)
    var keys = Object.keys(dailyMap).sort();
    keys.forEach(function(k) {
      var item = dailyMap[k];
      var netFlow = item.inflows - item.outflows;
      runningBalance += netFlow;

      ladderDays.push({
        date: item.date,
        dayName: item.dayName,
        inflows: item.inflows,
        outflows: item.outflows,
        netFlow: netFlow,
        endingBalance: runningBalance,
        isSafe: runningBalance >= CASH_FLOOR_RESERVE,
        detailsIn: item.detailsIn,
        detailsOut: item.detailsOut
      });
    });

    var result = {
      openingBalance: openingBalance,
      days: ladderDays
    };

    setCache_(cacheKey, result);
    return result;
  } catch (error) {
    return { error: error.toString() };
  }
}

function getCashFlowForecast(periodDays) {
try {
periodDays = parseInt(periodDays, 10) || 30;
var cfCacheKey = "CACHE_CASHFLOW_" + periodDays;
CacheService.getScriptCache().remove(cfCacheKey);

var sheet = getSheet("CHEQUES");
var result = { periodDays: periodDays, bankBalance: 0, incoming: 0, outgoing: 0, overdueIncoming: 0, overdueOutgoing: 0, net: 0, projectedBalance: 0, status: "safe", incomingList: [], outgoingList: [] };

var bankPanels = getBankAccountsPanel();
var totalActualBankBalance = 0;
for (var b = 0; b < bankPanels.length; b++) {
  totalActualBankBalance += bankPanels[b].estimatedActualBalance;
}
result.bankBalance = totalActualBankBalance;

if (!sheet || sheet.getLastRow() <= 1) {
  result.projectedBalance = totalActualBankBalance;
  result.status = computeLiquidityStatus(totalActualBankBalance, 0);
  setCache_(cfCacheKey, result);
  return result;
}

var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });
var typeIdx = headers.indexOf("Type"); if(typeIdx === -1) typeIdx = 1;
var amtIdx = headers.indexOf("Amount"); if(amtIdx === -1) amtIdx = 5;
var dueIdx = headers.indexOf("Due_Date"); if(dueIdx === -1) dueIdx = 7;
var statIdx = headers.indexOf("Is_Settled"); if(statIdx === -1) statIdx = 18;
var partyIdx = headers.indexOf("Party_Name"); if(partyIdx === -1) partyIdx = 2;
var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx === -1) bankIdx = 3;

var today = new Date(); today.setHours(0,0,0,0);
var end = new Date(today.getTime() + periodDays * 86400000);

for (var i = 1; i < data.length; i++) {
  var row = data[i];
  var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
  var status = row[statIdx] ? row[statIdx].toString().trim() : "";
  if ((type !== "INCOMING" && type !== "OUTGOING") || isItemCleared(status)) continue;
  
  var amount = parseFloat(row[amtIdx]) || 0;
  var dueDate = parseLocalDate(formatDateStr(row[dueIdx]));
  if (!dueDate) continue;
  dueDate.setHours(0,0,0,0);

  var party = row[partyIdx] || "غير محدد";
  var bank = row[bankIdx] || "-";
  var dueStr = formatDateStr(row[dueIdx]);
  var imgLink = "";
  try { imgLink = row[11] || row[10] || row[13] || ""; } catch(e) {}

  if (dueDate.getTime() < today.getTime()) {
    if (type === "INCOMING") {
      result.overdueIncoming += amount;
      result.incomingList.push({Party: party, Bank: bank, Amount: amount, Date: dueStr, Img: imgLink});
    } else {
      result.overdueOutgoing += amount;
      result.outgoingList.push({Party: party, Bank: bank, Amount: amount, Date: dueStr, Img: imgLink});
    }
  } else if (dueDate.getTime() <= end.getTime()) {
    if (type === "INCOMING") {
      result.incoming += amount;
      result.incomingList.push({Party: party, Bank: bank, Amount: amount, Date: dueStr, Img: imgLink});
    } else {
      result.outgoing += amount;
      result.outgoingList.push({Party: party, Bank: bank, Amount: amount, Date: dueStr, Img: imgLink});
    }
  }
}

result.incoming += 0; 
result.outgoing += result.overdueOutgoing; 
result.net = result.incoming - result.outgoing;
result.projectedBalance = result.bankBalance + result.net;

var totalAvailableLiquidity = result.bankBalance + result.incoming;
result.status = computeLiquidityStatus(totalAvailableLiquidity, result.outgoing);

result.incomingList.sort(function(a, b){ return new Date(a.Date) - new Date(b.Date); });
result.outgoingList.sort(function(a, b){ return new Date(a.Date) - new Date(b.Date); });

setCache_(cfCacheKey, result);
return result;
} catch (error) { 
return { error: error.toString() }; 
}
}

function simulateShockScenario(delayPercent, horizonDays) {
  try {
    var delay = parseFloat(delayPercent) || 0;
    var horizon = parseInt(horizonDays, 10) || 30;
    
    var sheet = getSheet("CHEQUES");
    var incoming = 0;
    var materialsDue = 0;
    var installmentsDue = 0;

    var today = new Date(); today.setHours(0,0,0,0);
    var end = new Date(today.getTime() + horizon * 86400000);

    var loanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOANS");
    if (loanSheet && loanSheet.getLastRow() > 1) {
      var lData = loanSheet.getDataRange().getValues();
      var lHeaders = lData[0].map(function(h){ return h.toString().trim(); });
      
      var amtIdx = lHeaders.indexOf("Monthly_Amount"); if (amtIdx === -1) amtIdx = 3;
      var startIdx = lHeaders.indexOf("Start_Date"); if (startIdx === -1) startIdx = 4;
      var endIdx = lHeaders.indexOf("End_Date"); if (endIdx === -1) endIdx = 5;
      var dayIdx = lHeaders.indexOf("Day_Of_Month"); if (dayIdx === -1) dayIdx = 7;
      var activeIdx = lHeaders.indexOf("Active"); if (activeIdx === -1) activeIdx = 8;

      for (var l = 1; l < lData.length; l++) {
        var lRow = lData[l];
        if (lRow[activeIdx] && lRow[activeIdx].toString().trim() === "Y") {
          var mAmount = parseFloat(lRow[amtIdx]) || 0;
          var tDay = parseInt(lRow[dayIdx], 10) || 1;
          var sDate = parseLocalDate(formatDateStr(lRow[startIdx]));
          var eDate = parseLocalDate(formatDateStr(lRow[endIdx]));

          var checkDate = new Date(today.getFullYear(), today.getMonth(), tDay);
          if (checkDate.getTime() < today.getTime()) {
            checkDate.setMonth(checkDate.getMonth() + 1);
          }

          while (checkDate.getTime() <= end.getTime()) {
            if ((!sDate || checkDate.getTime() >= sDate.getTime()) && (!eDate || checkDate.getTime() <= eDate.getTime())) {
              installmentsDue += mAmount;
            }
            checkDate.setMonth(checkDate.getMonth() + 1);
          }
        }
      }
    }

    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h){ return h.toString().trim(); });
      var typeIdx = headers.indexOf("Type"); if (typeIdx === -1) typeIdx = 1;
      var amtIdx = headers.indexOf("Amount"); if (amtIdx === -1) amtIdx = 5;
      var dueIdx = headers.indexOf("Due_Date"); if (dueIdx === -1) dueIdx = 7;
      var statIdx = headers.indexOf("Is_Settled"); if (statIdx === -1) statIdx = 18;
      var notesIdx = headers.indexOf("Notes"); if (notesIdx === -1) notesIdx = 14;

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
        var status = row[statIdx] ? row[statIdx].toString().trim() : "";
        if (isItemCleared(status)) continue;

        var amount = parseFloat(row[amtIdx]) || 0;
        var dueDate = parseLocalDate(formatDateStr(row[dueIdx]));
        if (!dueDate) continue;
        dueDate.setHours(0,0,0,0);

        if (dueDate.getTime() <= end.getTime()) {
          if (type === "INCOMING") {
            incoming += amount;
          } else if (type === "OUTGOING") {
            var notes = row[notesIdx] ? row[notesIdx].toString() : "";
            if (notes.indexOf("[قسط بنك]") === -1) {
              materialsDue += amount;
            }
          }
        }
      }
    }

    var incomingAfterShock = incoming * (1 - (delay / 100));

    var bankPanels = getBankAccountsPanel();
    var totalActualBankBalance = 0;
    for (var b = 0; b < bankPanels.length; b++) {
      totalActualBankBalance += bankPanels[b].estimatedActualBalance;
    }

    var totalAvailableUnderStress = totalActualBankBalance + incomingAfterShock;
    var remaining = totalAvailableUnderStress;
    var installmentsCovered = remaining >= installmentsDue;
    remaining = remaining - installmentsDue;
    
    var materialsCovered = remaining >= materialsDue;
    remaining = remaining - materialsDue;

    return {
      success: true,
      incomingAfterShock: incomingAfterShock,
      totalAvailableUnderStress: totalAvailableUnderStress,
      installmentsDue: installmentsDue,
      installmentsCovered: installmentsCovered,
      materialsDue: materialsDue,
      materialsCovered: materialsCovered,
      finalRemaining: remaining,
      horizonDays: horizon,
      delayPercent: delay
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getDecisionDashboardData_() {
try {
var ddCached = getCache_("CACHE_DECISION_DASHBOARD");
if (ddCached) return ddCached;

var sheet = getSheet("CHEQUES");
var out = {
  maturityLadder: [
    {label:"متأخرة ⚠️", incoming:0, outgoing:0},
    {label:"0-7 أيام", incoming:0, outgoing:0},
    {label:"8-30 يوم", incoming:0, outgoing:0},
    {label:"31-90 يوم", incoming:0, outgoing:0},
    {label:"أكثر من 90 يوم", incoming:0, outgoing:0}
  ],
  bankExposure: [],
  concentration: { INCOMING: [], OUTGOING: [] },
  overdueList: [],
  dsoMetrics: {}, 
  supplierConcentrationWarning: null
};
if (!sheet || sheet.getLastRow() <= 1) return out;

var data = sheet.getDataRange().getValues();
var headers = data[0].map(function(h){ return h.toString().trim(); });
var typeIdx = headers.indexOf("Type"); if(typeIdx===-1) typeIdx=1;
var partyIdx = headers.indexOf("Party_Name"); if(partyIdx===-1) partyIdx=2;
var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx===-1) bankIdx=3;
var numIdx = headers.indexOf("Reference_Number / Cheque_Number"); if(numIdx===-1) numIdx=4;
var amtIdx = headers.indexOf("Amount"); if(amtIdx===-1) amtIdx=5;
var dueIdx = headers.indexOf("Due_Date"); if(dueIdx===-1) dueIdx=7;
var statIdx = headers.indexOf("Is_Settled"); if(statIdx===-1) statIdx=18;
var invoiceDateIdx = headers.indexOf("Invoice_Date"); if(invoiceDateIdx===-1) invoiceDateIdx=19;
var clearedDateIdx = headers.indexOf("Cleared_Date"); if(clearedDateIdx===-1) clearedDateIdx=10;

var today = new Date(); today.setHours(0,0,0,0);
var bankMap = {};
var partyMap = { INCOMING:{}, OUTGOING:{} };
var totalByType = { INCOMING:0, OUTGOING:0 };

var customerDsoStats = {}; 
var outgoingSupplierTotals = {};
var totalOutgoingOutstandingAmount = 0;

for (var i=1; i<data.length; i++) {
  var row = data[i];
  var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
  var status = row[statIdx] ? row[statIdx].toString().trim() : "";
  var amount = parseFloat(row[amtIdx]) || 0;
  var party = row[partyIdx] ? row[partyIdx].toString().trim() : "غير محدد";
  var bank = row[bankIdx] ? row[bankIdx].toString().trim() : "غير محدد";
  var dueDate = parseLocalDate(formatDateStr(row[dueIdx]));
  var invoiceDate = parseLocalDate(formatDateStr(row[invoiceDateIdx]));
  var clearedDate = parseLocalDate(formatDateStr(row[clearedDateIdx]));

  if (type === "OUTGOING") {
    if (!isItemCleared(status)) {
        totalOutgoingOutstandingAmount += amount;
        outgoingSupplierTotals[party] = (outgoingSupplierTotals[party] || 0) + amount;
    }
  }

  if (type === "INCOMING" && party !== "غير محدد") {
    customerDsoStats[party] = customerDsoStats[party] || { totalAmount: 0, totalDays: 0, count: 0, overdueCount: 0 };
    customerDsoStats[party].totalAmount += amount;
    customerDsoStats[party].count += 1;
    
    if (invoiceDate) {
      var endDateForDSO = (isItemCleared(status) && clearedDate) ? clearedDate : today;
      var dsoDays = Math.floor((endDateForDSO.getTime() - invoiceDate.getTime()) / 86400000);
      if (dsoDays >= 0) {
        customerDsoStats[party].totalDays = (customerDsoStats[party].totalDays || 0) + dsoDays;
      }
    }
    if (!isItemCleared(status) && dueDate && dueDate.getTime() < today.getTime()) {
      customerDsoStats[party].overdueCount += 1;
    }
  }

  if ((type!=="INCOMING" && type!=="OUTGOING") || isItemCleared(status)) continue;

  totalByType[type] += amount;
  bankMap[bank] = bankMap[bank] || { bank:bank, incoming:0, outgoing:0 };
  bankMap[bank][type.toLowerCase()] += amount;
  partyMap[type][party] = (partyMap[type][party] || 0) + amount;

  if (dueDate) {
    dueDate.setHours(0,0,0,0);
    var diffDays = Math.floor((dueDate.getTime() - today.getTime()) / 86400000);
    var bucket;
    if (diffDays < 0) bucket = 0;
    else if (diffDays <= 7) bucket = 1;
    else if (diffDays <= 30) bucket = 2;
    else if (diffDays <= 90) bucket = 3;
    else bucket = 4;
    out.maturityLadder[bucket][type.toLowerCase()] += amount;

    if (diffDays < 0) {
      out.overdueList.push({
        Cheque_Number: row[numIdx], Type:type, Party_Name:party, Bank_Name:bank,
        Amount:amount, Due_Date: formatDateStr(row[dueIdx]), DaysOverdue: -diffDays
      });
    }
  }
}

var processedDso = {};
Object.keys(customerDsoStats).forEach(function(cust) {
  var st = customerDsoStats[cust];
  var avgDso = st.count > 0 && st.totalDays ? Math.round(st.totalDays / st.count) : 0;
  var isToxic = (st.totalAmount > 100000 && st.overdueCount >= 2) || avgDso > 75;
  
  processedDso[cust] = {
    avgDso: avgDso,
    totalVolume: st.totalAmount,
    overdueCount: st.overdueCount,
    isToxic: isToxic
  };
});
out.dsoMetrics = processedDso;

if (totalOutgoingOutstandingAmount > 0) {
  Object.keys(outgoingSupplierTotals).forEach(function(supp) {
    var ratio = (outgoingSupplierTotals[supp] / totalOutgoingOutstandingAmount) * 100;
    if (ratio >= 20) {
      out.supplierConcentrationWarning = {
        supplier: supp,
        percentage: ratio.toFixed(1),
        totalPaid: outgoingSupplierTotals[supp]
      };
    }
  });
}

out.bankExposure = Object.keys(bankMap).map(function(k){ return bankMap[k]; })
  .sort(function(a,b){ return (b.incoming+b.outgoing) - (a.incoming+a.outgoing); });

["INCOMING","OUTGOING"].forEach(function(t){
  var total = totalByType[t] || 0;
  out.concentration[t] = Object.keys(partyMap[t]).map(function(p){
    return { party:p, amount:partyMap[t][p], percent: total>0 ? (partyMap[t][p]/total*100) : 0 };
  }).sort(function(a,b){ return b.amount - a.amount; }).slice(0,5);
});

out.overdueList.sort(function(a,b){ return b.DaysOverdue - a.DaysOverdue; });
setCache_("CACHE_DECISION_DASHBOARD", out);
return out;
} catch (error) {
return { error: error.toString() };
}
}

function getDecisionDashboard() {
try {
// تم إيقاف فحص الإيميل عشان جوجل بيحجب إيميلات الـ gmail العادية
var data = getDecisionDashboardData_();
data.authorized = true;
return data;
} catch (error) { return { authorized: false, error: error.toString() }; }
}

function sendEmailReportToServer(cheques) {
try {
var recipient = "adhamamro55@gmail.com";
var subject = "📊 تقرير المعاملات والسيولة العاجلة - نظام الشيكات";
var htmlBody = "<div dir='rtl' style='font-family: Cairo, Tahoma, sans-serif; padding: 20px; background-color: #f8f9fa;'>";
htmlBody += "<h2 style='color:#198754;'>تقرير المعاملات المعلقة</h2><p>مرحباً، إليك كشف المعاملات الحالية في النظام:</p>";
htmlBody += "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse:collapse;width:100%;background:#fff;'>";
htmlBody += "<tr style='background:#343a40;color:#fff;'><th>الرقم المرجعي</th><th>النوع</th><th>الطرف</th><th>البنك</th><th>المبلغ</th><th>تاريخ الاستحقاق</th><th>الصورة</th></tr>";

var attachments = [];
if (cheques && cheques.length > 0) {
  cheques.forEach(function(chq) {
    var typeText = chq.Type === 'INCOMING' ? 'وارد' : 'صادر';
    var imgLinkCell = chq.Drive_Original ? "<a href='"+chq.Drive_Original+"' target='_blank'>عرض الصورة</a>" : "لا توجد صورة";
    htmlBody += "<tr><td>"+chq.Cheque_Number+"</td><td>"+typeText+"</td><td>"+(chq.Party_Name||'-')+"</td><td>"+(chq.Bank_Name||'-')+"</td><td>"+Number(chq.Amount).toLocaleString('ar-EG')+" ج.م</td><td>"+(chq.Due_Date||'-')+"</td><td>"+imgLinkCell+"</td></tr>";
    if (chq.Drive_Original) {
      try {
        var m = chq.Drive_Original.match(/\/d\/(.*?)\/(view|edit)/) || chq.Drive_Original.match(/id=([a-zA-Z0-9_-]+)/);
        if (m && m[1]) attachments.push(DriveApp.getFileById(m[1]).getBlob());
      } catch(e) {}
    }
  });
} else {
  htmlBody += "<tr><td colspan='7' align='center'>لا توجد معاملات معلقة حالياً.</td></tr>";
}
htmlBody += "</table></div>";

GmailApp.sendEmail(recipient, subject, "تقرير المعاملات", { htmlBody: htmlBody, attachments: attachments });
return { success:true };
} catch (error) { return { success:false, error: error.toString() }; }
}

function checkManagerAccess() {
var userEmail = "";
try { userEmail = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e) { userEmail = ""; }
if (userEmail && ALLOWED_MANAGERS && ALLOWED_MANAGERS.indexOf(userEmail) !== -1) {
return { authorized: true, email: userEmail };
} else { return { authorized: false, email: userEmail }; }
}

// ====================================================================================
// Phase 3: MTF Smart Treasury Engine (Active Cash Positioning & Routing Logic)
// ====================================================================================

function getNextWorkingDay(dateObj) {
  var d = new Date(dateObj.getTime());
  d.setDate(d.getDate() + 1);
  while (d.getDay() === 5 || d.getDay() === 6) {
      d.setDate(d.getDate() + 1);
  }
  return d;
}

function isMTFSessionOpen() {
  var now = new Date();
  var day = now.getDay();
  var hour = now.getHours();
  if (day >= 0 && day <= 4) {
      if (hour >= 9 && hour < 11) {
          return true;
      }
  }
  return false;
}

function getNextSessionString() {
  var now = new Date();
  var day = now.getDay();
  var hour = now.getHours();
  if (day === 5 || day === 6 || (day === 4 && hour >= 11)) return "الأحد 09:00 صباحاً";
  if (hour < 9) return "اليوم 09:00 صباحاً";
  return "غداً 09:00 صباحاً";
}

function formatCurrency_local(amount) {
  return Number(amount).toLocaleString('ar-EG');
}

function calculateSafeInvestmentHorizon(currentLiquidBalance, immediateObligations, chequesData, headers) {
  var today = new Date(); today.setHours(0,0,0,0);
  var simBalance = currentLiquidBalance - immediateObligations;
  
  var typeIdx = headers.indexOf("Type"); if(typeIdx===-1) typeIdx=1;
  var amtIdx = headers.indexOf("Amount"); if(amtIdx===-1) amtIdx=5;
  var dueIdx = headers.indexOf("Due_Date"); if(dueIdx===-1) dueIdx=7;
  var statusIdx = headers.indexOf("Is_Settled"); if(statusIdx===-1) statusIdx=18;

  var dailyNet = {};
  for (var i = 1; i < chequesData.length; i++) {
      var row = chequesData[i];
      if (isItemCleared(row[statusIdx] ? row[statusIdx].toString().trim() : "")) continue;
      
      var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
      var amt = parseFloat(row[amtIdx]) || 0;
      var dueStr = formatDateStr(row[dueIdx]);
      var dueDateObj = parseLocalDate(dueStr);
      if (!dueDateObj) continue;
      dueDateObj.setHours(0,0,0,0);

      if (dueDateObj.getTime() > today.getTime()) {
          var dStr = dueDateObj.getTime().toString();
          if (!dailyNet[dStr]) dailyNet[dStr] = { in: 0, out: 0 };
          if (type === "OUTGOING") dailyNet[dStr].out += amt;
      }
  }

  var loans = [];
  var loanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOANS");
  if (loanSheet && loanSheet.getLastRow() > 1) {
      var lData = loanSheet.getDataRange().getValues();
      var lHeaders = lData[0].map(function(h){ return h.toString().trim(); });
      var lAmtIdx = lHeaders.indexOf("Monthly_Amount"); if (lAmtIdx===-1) lAmtIdx=3;
      var lDayIdx = lHeaders.indexOf("Day_Of_Month"); if (lDayIdx===-1) lDayIdx=7;
      var lActIdx = lHeaders.indexOf("Active"); if (lActIdx===-1) lActIdx=8;
      for (var l = 1; l < lData.length; l++) {
          if (lData[l][lActIdx] && lData[l][lActIdx].toString().trim() === "Y") {
              loans.push({
                  amt: parseFloat(lData[l][lAmtIdx]) || 0,
                  day: parseInt(lData[l][lDayIdx], 10) || 1
              });
          }
      }
  }

  var safeDays = 0;
  var checkDate = new Date(today.getTime());
  
  for (var dayOffset = 1; dayOffset <= 60; dayOffset++) {
      checkDate = getNextWorkingDay(checkDate);
      var checkTime = checkDate.getTime();
      var dStr = checkTime.toString();
      
      if (dailyNet[dStr]) { simBalance -= dailyNet[dStr].out; }
      
      var currentDom = checkDate.getDate();
      for (var li=0; li<loans.length; li++) {
          var lastDayOfMonth = new Date(checkDate.getFullYear(), checkDate.getMonth() + 1, 0).getDate();
          var effectiveDay = loans[li].day > lastDayOfMonth ? lastDayOfMonth : loans[li].day;
          if (currentDom === effectiveDay) {
              simBalance -= loans[li].amt;
          }
      }

      if (simBalance < CASH_FLOOR_RESERVE) {
          safeDays = dayOffset - 1;
          if (safeDays < 0) safeDays = 0;
          return { days: safeDays, dateStr: Utilities.formatDate(checkDate, TZ, "yyyy-MM-dd") };
      }
  }
  return { days: 60, dateStr: "> 60 Days" };
}

function getSmartTreasuryDecision() {
  var panels = getBankAccountsPanel(); 
  var cibBalance = 0;
  var mtfBalance = 0;
  var totalBanksActual = 0;
  var tPlus1Pending = 0;

  panels.forEach(function(p) {
      if (p.bank === "CIB") cibBalance = p.estimatedActualBalance;
      if (p.bank === "MTF") mtfBalance = p.estimatedActualBalance;
      if (p.bank !== "MTF") totalBanksActual += p.estimatedActualBalance; 
  });

  var totalLiquidity = totalBanksActual + mtfBalance; 

  var mtfOpening = 0;
  var sData = getSheet("SETTINGS").getDataRange().getValues();
  for(var idx=1; idx<sData.length; idx++) {
      if(sData[idx][0].toString().trim() === "MTF") mtfOpening = parseFloat(sData[idx][1]) || 0;
  }
  
  var mtfIn = 0; var mtfOut = 0;
  var lData = getSheet("LOG").getDataRange().getValues();
  for(var k=1; k<lData.length; k++) {
      if(lData[k][2] && lData[k][2].toString().trim() === "MTF") {
          if(lData[k][4] === "IN") mtfIn += (parseFloat(lData[k][3]) || 0);
          if(lData[k][4] === "OUT") mtfOut += (parseFloat(lData[k][3]) || 0);
      }
  }
  var mtfProfit = mtfBalance - (mtfOpening + mtfIn - mtfOut);

  var sheet = getSheet("CHEQUES");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var typeIdx = headers.indexOf("Type"); if(typeIdx===-1) typeIdx=1;
  var bankIdx = headers.indexOf("Bank_Name"); if(bankIdx===-1) bankIdx=3;
  var amtIdx = headers.indexOf("Amount"); if(amtIdx===-1) amtIdx=5;
  var dueIdx = headers.indexOf("Due_Date"); if(dueIdx===-1) dueIdx=7;
  var statusIdx = headers.indexOf("Is_Settled"); if(statusIdx===-1) statusIdx=18;
  var notesIdx = headers.indexOf("Notes"); if(notesIdx===-1) notesIdx=14;

  var today = new Date(); today.setHours(0,0,0,0);
  var tomorrow = getNextWorkingDay(today);
  var immediateObligations = 0;
  var tomorrowObligations = 0;

  for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = row[statusIdx] ? row[statusIdx].toString().trim() : "";
      var type = row[typeIdx] ? row[typeIdx].toString().trim() : "";
      var bank = row[bankIdx] ? row[bankIdx].toString().trim() : "";
      var amt = parseFloat(row[amtIdx]) || 0;
      var notes = row[notesIdx] ? row[notesIdx].toString().trim() : "";

      if (isItemCleared(status)) continue;

      var dueStr = formatDateStr(row[dueIdx]);
      var dueDateObj = parseLocalDate(dueStr);
      if (!dueDateObj) continue;
      dueDateObj.setHours(0,0,0,0);

      if (type === "INCOMING" && bank === "CIB" && notes.indexOf("[MTF_REDEMPTION]") !== -1) {
          tPlus1Pending += amt;
      }

      if (type === "OUTGOING" && dueDateObj.getTime() <= today.getTime()) {
          immediateObligations += amt;
      }
      if (type === "OUTGOING" && dueDateObj.getTime() === tomorrow.getTime()) {
          tomorrowObligations += amt;
      }
  }

  var purchasingPower = totalBanksActual - CASH_FLOOR_RESERVE - immediateObligations;
  if (purchasingPower < 0) purchasingPower = 0;

  var horizon = calculateSafeInvestmentHorizon(totalBanksActual, immediateObligations, data, headers);

  return {
      sessionOpen: isMTFSessionOpen(),
      nextSessionStr: getNextSessionString(),
      purchasingPower: purchasingPower,
      cibBalance: cibBalance,
      mtfAvailable: mtfBalance, 
      tPlus1Pending: tPlus1Pending,
      safeHorizonDays: horizon.days,
      horizonDateStr: horizon.dateStr,
      tomorrowObligations: tomorrowObligations,
      immediateObligations: immediateObligations,
      currentLiquidBalance: totalBanksActual,
      totalLiquidity: totalLiquidity,
      mtfProfit: mtfProfit
  };
}

function executeLiquidityPoolingToCIB() {
  var lock = LockService.getScriptLock();
  try {
      lock.waitLock(15000);
      var panels = getBankAccountsPanel();
      var logSheet = getSheet("LOG");
      var now = new Date();
      var dateStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
      var moved = 0;
      
      panels.forEach(function(p) {
          if (p.bank === "NBE" || p.bank === "BM") {
              var surplus = p.estimatedActualBalance - 50000; 
              if (surplus > 10000) {
                  var rNum = Math.floor(Math.random() * 1000);
                  logSheet.appendRow(["LOG-" + now.getTime() + rNum + "A", "MANUAL", p.bank, surplus, "OUT", "⚡ كنس سريع إلى CIB", dateStr, now]);
                  logSheet.appendRow(["LOG-" + now.getTime() + rNum + "B", "MANUAL", "CIB", surplus, "IN", "⚡ كنس سريع من " + p.bank, dateStr, now]);
                  moved += surplus;
              }
          }
      });
      
      if(moved > 0) {
          clearAppCache_();
          return { success: true, moved: moved };
      } else {
          return { success: false, error: "لا يوجد فائض كافٍ في NBE أو BM لتسحبه (النظام يحافظ على 50 ألف كحد أدنى تشغيلي في كل بنك)." };
      }
  } catch (e) {
      return { success: false, error: e.toString() };
  } finally {
      lock.releaseLock();
  }
}

function executeMTFTransaction(actionType, amount) {
  var lock = LockService.getScriptLock();
  try {
      lock.waitLock(15000);
      var amt = parseFloat(amount);
      if (isNaN(amt) || amt <= 0) return { success: false, error: "المبلغ غير صحيح." };

      var dec = getSmartTreasuryDecision();
      
      if (actionType === "BUY") {
          if (!dec.sessionOpen) return { success: false, error: "جلسة الصندوق مغلقة الآن." };
          if (amt > dec.cibBalance) return { success: false, error: "رصيد CIB الجاهز لا يكفي. يرجى سحب الفائض لـ CIB أولاً." };
          
          var logSheet = getSheet("LOG");
          var now = new Date();
          var dateStr = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
          
          logSheet.appendRow(["MTF-" + now.getTime() + "A", "MANUAL", "CIB", amt, "OUT", "شراء وثائق MTF", dateStr, now]);
          logSheet.appendRow(["MTF-" + now.getTime() + "B", "MANUAL", "MTF", amt, "IN", "إيداع من CIB", dateStr, now]);
          
          clearAppCache_();
          return { success: true, message: "تم الشراء بنجاح وخصم المبلغ من CIB." };
      } 
      else if (actionType === "SELL") {
          if (!dec.sessionOpen) return { success: false, error: "جلسة الصندوق مغلقة الآن." };
          if (amt > dec.mtfAvailable) return { success: false, error: "المبلغ يتجاوز رصيد الصندوق المتاح للاسترداد." };
          
          var logSheet2 = getSheet("LOG");
          var chequesSheet = getSheet("CHEQUES");
          var now2 = new Date();
          var todayStr = Utilities.formatDate(now2, TZ, "yyyy-MM-dd");
          var tPlus1Date = getNextWorkingDay(now2);
          var tPlus1Str = Utilities.formatDate(tPlus1Date, TZ, "yyyy-MM-dd");
          
          logSheet2.appendRow(["MTF-SELL-" + now2.getTime(), "MANUAL", "MTF", amt, "OUT", "استرداد وثائق MTF (T+1)", todayStr, now2]);
          
          chequesSheet.appendRow([
              "TRF-MTF-" + now2.getTime(),
              "INCOMING",
              "صندوق MTF",
              "CIB",
              "MTF-" + now2.getTime(),
              amt,
              todayStr, 
              tPlus1Str, 
              "مستلم", 
              todayStr, 
              "", "", "", "", 
              "[MTF_REDEMPTION] تحويل استرداد وثائق",
              now2, now2,
              "BANK_TRANSFER",
              "N", 
              "" 
          ]);

          clearAppCache_();
          return { success: true, message: "تم أمر الاسترداد بنجاح. الأموال قيد التحويل وستصل حساب CIB يوم " + tPlus1Str };
      }
      
      return { success: false, error: "نوع العملية غير معروف." };
  } catch (error) {
      return { success: false, error: error.toString() };
  } finally {
      lock.releaseLock();
  }
}

function dailyMTFTriggerAlert() {
  try {
      var now = new Date();
      var day = now.getDay();
      if (day === 5 || day === 6) return; 

      var dec = getSmartTreasuryDecision();
      var neededLiquid = CASH_FLOOR_RESERVE + dec.tomorrowObligations;
      var shortfall = neededLiquid - dec.currentLiquidBalance;

      if (shortfall > 0 && dec.mtfAvailable > 0) {
          var suggestedWithdrawal = Math.min(shortfall, dec.mtfAvailable);
          var recipient = "adhamamro55@gmail.com";
          var subject = "⚠️ تنبيه سحب حتمي من صندوق MTF";
          var body = "مرحباً،\n\nنظام الخزانة الآلي يحذرك:\n";
          body += "لدينا التزامات غداً تستدعي كسر الاحتياطي الحديدي (السيولة الفعلية لا تكفي).\n";
          body += "المطلوب استرداده فوراً اليوم قبل الساعة 11:00 صباحاً: " + formatCurrency_local(suggestedWithdrawal) + " ج.م\n\n";
          body += "يرجى الدخول للنظام وتنفيذ الاسترداد ليسمع في بنك CIB غداً ويغطي الالتزامات.\n\nنظام القيادة المالية.";
          
          GmailApp.sendEmail(recipient, subject, body);
      }
  } catch(e) {}
}