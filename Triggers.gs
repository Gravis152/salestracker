/**
 * =================================================================================
 * TRIGGERS AND MENU SYSTEM (v5.4 - FIXED SELLING DAYS & PERCENTAGE)
 * =================================================================================
 */

function getQuickHealthStatus() {
  try {
    var score = 100;
    var stats = ADVANCED_CACHE.getStats();
    if (stats.memoryKeys > stats.maxKeys * 0.9) score -= 20;
    try {
      if (ScriptApp.getRemainingDailyTriggers() < CONFIG.EXECUTION_LIMITS.MIN_QUOTA_TRIGGERS) {
        score -= 30;
      }
    } catch (e) {}
    return { status: score >= 70 ? "Good" : "Poor", score: score };
  } catch (e) {
    return { status: "Unknown", score: 0 };
  }
}

function performBackgroundHealthCheck() {
  try {
    var remaining = ScriptApp.getRemainingDailyTriggers();
    if (remaining < CONFIG.EXECUTION_LIMITS.MIN_QUOTA_TRIGGERS) {
      console.warn("Low execution quota: " + remaining + " triggers");
    }
  } catch (e) {}
}

function createMainMenu() {
  var ui = SpreadsheetApp.getUi();
  var health = getQuickHealthStatus();
  var menuTitle = health.status === 'Poor' ? '🚀 Sheet Tools ⚠️' : '🚀 Sheet Tools';
  
  ui.createMenu(menuTitle)
    .addItem('🔄 Refresh All Reports', 'refreshAllReports')
    .addItem('🎯 Check Goal Progress', 'checkGoalProgress')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Individual Reports')
      .addItem('Dashboard', 'createRestoredMTDDashboard')
      .addItem('YTD Report', 'createYTDReport')
      .addItem('Client List', 'updateClientList'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ Maintenance')
      .addItem('🏥 System Health Check', 'runSystemHealthCheck')
      .addItem('🧹 Clear System Cache', 'clearAllSystemCaches')
      .addItem('✅ Re-Apply Validation', 'setupDataValidationRules')
      .addItem('🔄 Convert Smart Tables', 'convertSmartTablesToRegularSheets')
      .addItem('🔍 Analyze Sheet Types', 'analyzeSheetTypes'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🗓️ Year-End')
      .addItem('📁 Verify Archive', 'verifyYearEndArchive')
      .addItem('🏁 Run Year-End Reset', 'runYearEndReset'))
    .addSeparator()
    .addSubMenu(ui.createMenu('ℹ️ Help')
      .addItem('🎯 Usage Tips', 'showUsageTips')
      .addItem('📖 System Overview', 'showSystemOverview'))
    .addToUi();
}

function onOpen() {
  try {
    createMainMenu();
    performBackgroundHealthCheck();
  } catch (e) {
    try {
      SpreadsheetApp.getUi()
        .createMenu('🚀 Sheet Tools')
        .addItem('🔄 Refresh All', 'refreshAllReports')
        .addToUi();
    } catch (e2) {}
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var cellRef = range.getA1Notation();
  var row = range.getRow();
  
  if (sheetName === "Dashboard" && cellRef === "M2") {
    createRestoredMTDDashboard();
    return;
  }
  
  if (CONFIG.MONTH_NAMES.includes(sheetName)) {
    if (row > 1) {
      validateDataEntry(e);
      UnifiedDataAccess.clearCache(sheetName);
    }
    if (CONFIG.MONITORED_CELLS.includes(cellRef)) {
      createRestoredMTDDashboard();
    }
  }
}

function onInstall(e) {
  onOpen(e);
  ADVANCED_CACHE.clearAll();
  UnifiedDataAccess.clearCache();
}

function refreshAllReports(silent) {
  var startTime = Date.now();
  var results = { dashboard: false, ytd: false, client: false };
  
  try {
    if (!silent) {
      SpreadsheetApp.getActiveSpreadsheet().toast("🔄 Refreshing all reports...", "Please Wait", 30);
    }
    
    UnifiedDataAccess.clearCache();
    
    try { createRestoredMTDDashboard(); results.dashboard = true; } catch (e) {}
    try { createYTDReport(); results.ytd = true; } catch (e) {}
    try { updateClientList(); results.client = true; } catch (e) {}
    
    var elapsed = Math.round((Date.now() - startTime) / 1000);
    var successCount = Object.values(results).filter(Boolean).length;
    
    if (!silent) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "✅ Refresh Complete (" + successCount + "/3 in " + elapsed + "s)", "Success", 5
      );
    }
  } catch (e) {
    if (!silent) {
      SpreadsheetApp.getActiveSpreadsheet().toast("❌ Error: " + e.message, "Error", 5);
    }
  }
  
  return results;
}

function runSystemHealthCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checks = { sheets: 0, mappings: 0, cacheStatus: "Unknown" };
  
  CONFIG.MONTH_NAMES.forEach(function(m) {
    if (ss.getSheetByName(m)) {
      checks.sheets++;
      try {
        var map = UnifiedDataAccess.getColumnMapping(m);
        if (map.DATE && map.CUSTOMER) checks.mappings++;
      } catch (e) {}
    }
  });
  
  var cacheStats = ADVANCED_CACHE.getStats();
  checks.cacheStatus = cacheStats.cacheEnabled ? "Active ✅" : "Disabled ❌";
  
  var msg = "🏥 System Health Check:\n\n";
  msg += "📋 Sheets Found: " + checks.sheets + "/" + CONFIG.MONTH_NAMES.length + "\n";
  msg += "🗂️ Column Mappings: " + checks.mappings + "/" + checks.sheets + "\n";
  msg += "💾 Cache: " + checks.cacheStatus + " (" + cacheStats.memoryKeys + " keys)\n";
  msg += "⚡ Status: " + (checks.sheets === CONFIG.MONTH_NAMES.length ? "All Good ✅" : "Issues Found ⚠️");
  
  SpreadsheetApp.getUi().alert("System Health", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return checks;
}

function showUsageTips() {
  var tips = "🎯 Usage Tips:\n\n";
  tips += "🚀 Performance:\n";
  tips += "• Dashboard auto-refreshes when you change goals or month\n";
  tips += "• Use 'Clear Cache' if data looks stale\n\n";
  tips += "📊 Data Entry:\n";
  tips += "• Type, Device, Plan have dropdowns for valid values\n";
  tips += "• Validation will suggest corrections for typos\n\n";
  tips += "🔧 Maintenance:\n";
  tips += "• Run Health Check weekly\n";
  tips += "• Year-End Reset archives everything before clearing";
  
  SpreadsheetApp.getUi().alert("Usage Tips", tips, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showSystemOverview() {
  var overview = "📖 System Overview (v5.4):\n\n";
  
  try {
    var cache = ADVANCED_CACHE.getStats();
    var uda = UnifiedDataAccess.getStats();
    
    overview += "💾 Cache:\n";
    overview += "• Keys: " + cache.memoryKeys + "/" + cache.maxKeys + "\n";
    overview += "• Enabled: " + (cache.cacheEnabled ? "Yes" : "No") + "\n\n";
    overview += "🔧 Data Access:\n";
    overview += "• Methods: " + uda.accessMethodsAvailable.join(", ") + "\n";
    overview += "• Max Rows: " + uda.defaultMaxRows.toLocaleString() + "\n";
  } catch (e) {
    overview += "⚠️ Could not retrieve stats\n";
  }
  
  SpreadsheetApp.getUi().alert("System Overview", overview, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Goal Progress - v5.4 FIXED selling days & percentage display
 * TODAY IS COUNTED AS REMAINING (you still have today to sell)
 */
function checkGoalProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var today = new Date();
  var currentMonth = CONFIG.MONTH_NAMES[today.getMonth()];
  
  var sheet = ss.getSheetByName(currentMonth);
  if (!sheet) {
    ui.alert("Sheet not found: " + currentMonth);
    return;
  }
  
  // Get goals
  var ppvgaGoal = sheet.getRange("I2").getValue() || 0;
  var aiaGoal = sheet.getRange("I5").getValue() || 0;
  var accGoal = sheet.getRange("I8").getValue() || 0;
  
  // Get actuals
  var counts = getMonthTypeCounts(ss, currentMonth);
  var ppvgaActual = counts.PPVGA;
  var aiaActual = counts.AIA + counts.AIAC + counts.AIAB;
  var upgActual = counts.UPG;
  var plus1Actual = counts.Plus1;
  
  // Get accessory actual
  var accActual = 0;
  try {
    var rawAcc = sheet.getRange("I9").getValue();
    if (rawAcc !== "" && rawAcc !== null) {
      accActual = (typeof rawAcc === 'number') ? rawAcc : parseFloat(String(rawAcc).replace(/[^0-9.-]+/g, "")) || 0;
    }
  } catch (e) {}
  
  // Get date info
  var year = today.getFullYear();
  var month = today.getMonth();
  var currentDay = today.getDate();
  var daysInMonth = new Date(year, month + 1, 0).getDate();
  
  // Count selling days manually
  var totalSellingDays = 0;
  var currentSellingDay = 0;
  var sellingDaysRemaining = 0;
  
  for (var d = 1; d <= daysInMonth; d++) {
    var checkDate = new Date(year, month, d);
    var dayOfWeek = checkDate.getDay();
    
    // Skip Thursday (4) and Friday (5)
    if (dayOfWeek === 4 || dayOfWeek === 5) {
      continue;
    }
    
    // This is a selling day
    totalSellingDays++;
    
    if (d < currentDay) {
      // Past selling day (before today)
      // Don't count yet
    } else if (d === currentDay) {
      // Today - this is the current selling day number
      currentSellingDay = totalSellingDays;
    }
    
    if (d >= currentDay) {
      // Today and future selling days
      sellingDaysRemaining++;
    }
  }
  
  // If today is Thu/Fri (not a selling day), find the next selling day
  if (currentSellingDay === 0) {
    // Today is not a selling day, so we're between selling days
    // Count how many selling days have passed
    var passedCount = 0;
    for (var d = 1; d < currentDay; d++) {
      var checkDate = new Date(year, month, d);
      var dayOfWeek = checkDate.getDay();
      if (dayOfWeek !== 4 && dayOfWeek !== 5) {
        passedCount++;
      }
    }
    currentSellingDay = passedCount; // Last completed selling day
  }
  
  // Ensure minimum values
  totalSellingDays = Math.max(1, totalSellingDays);
  currentSellingDay = Math.max(1, currentSellingDay);
  sellingDaysRemaining = Math.max(1, sellingDaysRemaining);
  
  // Expected percentage based on where you SHOULD be by end of today
  var expectedPct = (currentSellingDay / totalSellingDays * 100).toFixed(1);
  
  // Progress percentages
  var ppvgaPct = ppvgaGoal > 0 ? (ppvgaActual / ppvgaGoal * 100) : 0;
  var aiaPct = aiaGoal > 0 ? (aiaActual / aiaGoal * 100) : 0;
  var accPct = accGoal > 0 ? (accActual / accGoal * 100) : 0;
  
  // What's needed
  var ppvgaNeeded = Math.max(0, ppvgaGoal - ppvgaActual);
  var aiaNeeded = Math.max(0, aiaGoal - aiaActual);
  var accNeeded = Math.max(0, accGoal - accActual);
  
  // Daily pace
  var daysForPace = Math.max(1, sellingDaysRemaining);
  var ppvgaPace = (ppvgaNeeded / daysForPace).toFixed(1);
  var aiaPace = (aiaNeeded / daysForPace).toFixed(1);
  var accPace = (accNeeded / daysForPace).toFixed(0);
  
  // Stretch goals
  var ppvgaStretch = Math.ceil(ppvgaGoal * 1.25);
  var aiaStretch = Math.ceil(aiaGoal * 1.25);
  var accStretch = Math.ceil(accGoal * 1.25);
  
  function getStatus(actualPct, expectedPct) {
    if (actualPct >= expectedPct) return "✅ ON TRACK";
    if (actualPct >= expectedPct * 0.8) return "⚠️ SLIGHTLY BEHIND";
    return "🔴 BEHIND";
  }
  
  // Build message
  var msg = "📅 Selling Day " + currentSellingDay + " of " + totalSellingDays + "\n";
  msg += "⏳ " + sellingDaysRemaining + " selling days left (including today)\n";
  msg += "📆 Excludes Thu/Fri - your days off\n";
  msg += "🎯 You should be at " + expectedPct + "% by end of today\n";
  msg += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";
  
  msg += "📊 PPVGA\n";
  msg += "   Base Goal: " + ppvgaActual + " / " + ppvgaGoal + " (" + ppvgaPct.toFixed(1) + "%)\n";
  msg += "   Stretch (125%): " + ppvgaActual + " / " + ppvgaStretch + "\n";
  msg += "   Status: " + getStatus(ppvgaPct, parseFloat(expectedPct)) + "\n";
  msg += "   Need: " + ppvgaNeeded + " more (" + ppvgaPace + "/day)\n\n";
  
  msg += "💜 AIA/AC\n";
  msg += "   Base Goal: " + aiaActual + " / " + aiaGoal + " (" + aiaPct.toFixed(1) + "%)\n";
  msg += "   Stretch (125%): " + aiaActual + " / " + aiaStretch + "\n";
  msg += "   Status: " + getStatus(aiaPct, parseFloat(expectedPct)) + "\n";
  msg += "   Need: " + aiaNeeded + " more (" + aiaPace + "/day)\n\n";
  
  msg += "💰 ACCESSORIES\n";
  msg += "   Base Goal: $" + accActual.toLocaleString() + " / $" + accGoal.toLocaleString() + " (" + accPct.toFixed(1) + "%)\n";
  msg += "   Stretch (125%): $" + accActual.toLocaleString() + " / $" + accStretch.toLocaleString() + "\n";
  msg += "   Status: " + getStatus(accPct, parseFloat(expectedPct)) + "\n";
  msg += "   Need: $" + accNeeded.toLocaleString() + " more ($" + accPace + "/day)\n\n";
  
  msg += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += "📈 OTHER METRICS (MTD)\n";
  msg += "   UPG: " + upgActual + "\n";
  msg += "   Plus1: " + plus1Actual + "\n";
  
  ui.alert("🎯 " + currentMonth.toUpperCase() + " GOAL PROGRESS", msg, ui.ButtonSet.OK);
}
