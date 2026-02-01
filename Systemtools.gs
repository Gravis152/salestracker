/**
 * =================================================================================
 * ENHANCED GLOBAL UTILITIES (v4.6 - METADATA FIX)
 * Central commands matched to the clean menu structure
 * =================================================================================
 */

function refreshAllReports(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!silent) ss.toast("🔄 Syncing all data...", "System Update", 5);
  
  try {
    UnifiedDataAccess.clearCache();
    createRestoredMTDDashboard();
    createYTDReport();
    updateClientList();
    if (!silent) ss.toast("✅ All Reports Updated", "Success", 3);
  } catch (err) {
    console.error("Refresh Error: " + err.message);
    if (!silent) ss.toast("❌ Error: " + err.message, "Error", 5);
  }
}

function runYearEndReset() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "⚠️ CRITICAL ACTION: YEAR-END RESET", 
    "This will:\n1. Archive current YTD data\n2. Clear all monthly sales logs\n3. Reset all caches\n\nThis cannot be undone. Continue?", 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var currentYear = new Date().getFullYear();
    
    // 1. Archive
    var ytd = ss.getSheetByName("YTD Report");
    if (ytd) {
      var archName = "Archive_" + currentYear;
      var arch = ss.insertSheet(archName);
      // Simple value copy to avoid metadata errors
      var data = ytd.getDataRange().getValues();
      arch.getRange(1, 1, data.length, data[0].length).setValues(data);
      // Basic styling copy
      ytd.getDataRange().copyTo(arch.getRange(1,1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      ss.toast("Archive Created: " + archName);
    }
    
    // 2. Clear Data
    CONFIG.MONTH_NAMES.forEach(function(m) {
      var s = ss.getSheetByName(m);
      if (s) {
        var lastRow = s.getLastRow();
        if (lastRow > 1) {
          try {
            // Try standard clear first
            s.getRange(2, 1, lastRow - 1, 20).clearContent();
            s.getRange(2, 1, lastRow - 1, 20).clearDataValidations();
          } catch(e) {
            // Fallback for Smart Tables: Delete rows instead of clearing
            if (e.message.includes("typed")) {
              s.deleteRows(2, lastRow - 1);
            }
          }
        }
      }
    });
    
    // 3. Reset System
    ADVANCED_CACHE.clearAll();
    UnifiedDataAccess.clearCache();
    setupDataValidationRules(); 
    
    refreshAllReports(true);
    ui.alert("✅ Year-End Reset Complete");
    
  } catch (e) {
    ui.alert("❌ Error: " + e.message);
  }
}

function convertSmartTablesToRegularSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var count = 0;
  
  CONFIG.MONTH_NAMES.forEach(function(m) {
    var s = ss.getSheetByName(m);
    if (!s) return;
    
    var isSmart = false;
    try { s.getRange("A1").getDataValidation(); } catch(e) { isSmart = true; }
    
    if (isSmart) {
      try {
        var data = UnifiedDataAccess.getSheetData(m, {useCache: false});
        var newS = ss.insertSheet(m + "_Temp");
        if (data.length > 0) {
          newS.getRange(2, 1, data.length, data[0].length).setValues(data);
        }
        ss.deleteSheet(s);
        newS.setName(m);
        count++;
      } catch(e) {
        console.error("Conversion failed for " + m);
      }
    }
  });
  
  var msg = count > 0 
    ? "Conversion Complete: " + count + " sheets processed."
    : "No Smart Tables found requiring conversion.";
    
  SpreadsheetApp.getUi().alert(msg);
}

function clearAllSystemCaches() {
  ADVANCED_CACHE.clearAll();
  UnifiedDataAccess.clearCache();
  SpreadsheetApp.getActiveSpreadsheet().toast("System Caches Cleared", "Success", 3);
}

function verifyYearEndArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = new Date().getFullYear();
  var found = ss.getSheets().some(function(sheet) {
    return sheet.getName().includes("Archive") && sheet.getName().includes(year.toString());
  });
  
  if (found) {
    SpreadsheetApp.getUi().alert("✅ Archive Verified", "An archive sheet for " + year + " was found.", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ No Archive Found", "No archive found for " + year + ".", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showSmartUsageTips() {
  var tips = "🎯 Smart Usage Tips:\n\n";
  tips += "🚀 Performance:\n• Dashboard auto-refreshes.\n• Use 'Clear Cache' if data looks old.\n\n";
  tips += "📊 Data:\n• Smart Tables are auto-detected.\n• Future months ignored.\n";
  SpreadsheetApp.getUi().alert(tips);
}

function showEnhancedSystemOverview() {
  var overview = "🚀 System Overview (v4.6):\n\n";
  try {
    var cache = ADVANCED_CACHE.getStats();
    overview += "🏥 Health Status:\n• Cache: " + (cache.memoryKeys >= 0 ? "Active ✅" : "Inactive ❌") + "\n";
    var uda = UnifiedDataAccess.getStats();
    overview += "• Methods: " + (uda.accessMethodsAvailable ? uda.accessMethodsAvailable.length : 3) + "\n";
  } catch (e) {
    overview += "⚠️ Stats unavailable.\n";
  }
  SpreadsheetApp.getUi().alert(overview);
}
