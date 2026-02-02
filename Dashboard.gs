/**
 * =================================================================================
 * MTD DASHBOARD GENERATOR (v5.5 - PRODUCTION)
 * Full month-over-month comparison with clean output
 * FIXED: Month selector now properly handles Auto/Manual selection
 * =================================================================================
 */

function createRestoredMTDDashboard() {
  if (!checkExecutionQuota()) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard", 0);
  var trendSheet = ss.getSheetByName("TrendData_DEV") || ss.insertSheet("TrendData_DEV");
  var style = CONFIG.STYLES;
  
  // 🆕 FIXED: Improved month selector logic
  var m2Range = dashboard.getRange("M2");
  var selectedValue = m2Range.getValue();
  
  // Initialize month selector dropdown if not already set
  if (!selectedValue || selectedValue === "") {
    m2Range.setValue("Auto");
    selectedValue = "Auto";
  }
  
  // Determine current month
  var today = new Date();
  var currentMonth;
  
  if (selectedValue === "Auto") {
    // Use current month
    currentMonth = CONFIG.MONTH_NAMES[today.getMonth()];
  } else if (CONFIG.MONTH_NAMES.indexOf(selectedValue) >= 0) {
    // Use selected month from dropdown
    currentMonth = selectedValue;
  } else {
    // Fallback to current month if invalid selection
    currentMonth = CONFIG.MONTH_NAMES[today.getMonth()];
    m2Range.setValue("Auto");
  }
  
  var monthSheet = ss.getSheetByName(currentMonth);
  if (!monthSheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + currentMonth + '" not found. Please create it first.');
    return;
  }
  
  // Calculate date metrics
  var monthIdx = CONFIG.MONTH_NAMES.indexOf(currentMonth);
  var daysInMonth = new Date(today.getFullYear(), monthIdx + 1, 0).getDate();
  var isCurrentMonth = (monthIdx === today.getMonth());
  var daysPassed = isCurrentMonth ? today.getDate() : daysInMonth;
  var runRateDivisor = Math.max(1, daysPassed);
  var prevMonthIdx = monthIdx - 1;
  var prevMonth = (prevMonthIdx >= 0) ? CONFIG.MONTH_NAMES[prevMonthIdx] : "Dec";
  var workingDaysLeft = calculateWorkingDaysLeft(currentMonth, today);
  
  // Get raw data for header alert check
  var rawData;
  try {
    rawData = monthSheet.getRange("A2:B" + Math.max(2, monthSheet.getLastRow())).getValues();
  } catch (e) {
    try {
      rawData = UnifiedDataAccess.getSheetData(currentMonth, { maxRows: 1000, maxCols: 2, useCache: false });
    } catch (e2) {
      rawData = [];
    }
  }
  
  // Build dashboard
  cleanupDashboard(dashboard, trendSheet);
  buildMonthSelector(dashboard, style, currentMonth); // 🆕 NEW: Add month selector
  buildRefreshSection(dashboard, style);
  buildHeader(dashboard, style, currentMonth, rawData, runRateDivisor, daysInMonth, monthSheet);
  buildKPICards(dashboard, style, currentMonth, prevMonth, ss);
  buildTrackers(dashboard, style, currentMonth);
  buildPaceSection(dashboard, style, currentMonth, workingDaysLeft, daysInMonth, daysPassed, runRateDivisor);
  buildCharts(dashboard, trendSheet, style);
  
  dashboard.setHiddenGridlines(true);
}

/**
 * 🆕 BUILD MONTH SELECTOR DROPDOWN
 */
function buildMonthSelector(dashboard, style, currentMonth) {
  // Label for month selector
  dashboard.getRange("M5").setValue("Select Month:")
    .setFontSize(9)
    .setFontWeight("bold")
    .setFontColor(style.accent)
    .setHorizontalAlignment("center");
  
  // Month selector dropdown
  var m2Range = dashboard.getRange("M2");
  
  // Create dropdown with Auto + all month names
  var monthOptions = ["Auto"].concat(CONFIG.MONTH_NAMES);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(monthOptions, true)
    .setAllowInvalid(false)
    .build();
  
  m2Range.setDataValidation(rule);
  
  // Style the dropdown cell
  m2Range
    .setBackground("#ffffff")
    .setFontWeight("bold")
    .setFontColor(style.main)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, null, null, style.main, SpreadsheetApp.BorderStyle.SOLID);
  
  // Set current value if not already set
  if (!m2Range.getValue() || m2Range.getValue() === "") {
    m2Range.setValue("Auto");
  }
  
  // Add note to explain Auto mode
  m2Range.setNote(
    "Auto: Shows current month\n" +
    "Or select a specific month to view"
  );
}

/**
 * Cleanup dashboard and trend sheet
 */
function cleanupDashboard(dashboard, trendSheet) {
  try {
    dashboard.getRange("A1:L100").clear();
    // 🆕 DON'T clear M2 (month selector) or M5 (label)
    dashboard.getRange("M1:M1").clear();
    dashboard.getRange("M3:M4").clear();
    dashboard.getRange("M6:M100").clear();
    dashboard.getRange("Z1:Z20").clear();
    dashboard.clearFormats();
    dashboard.clearConditionalFormatRules();
  } catch (e) {
    try { 
      dashboard.getRange("A1:L100").clearContent();
      dashboard.getRange("M1:M1").clearContent();
      dashboard.getRange("M3:M4").clearContent();
      dashboard.getRange("M6:M100").clearContent();
      dashboard.getRange("Z1:Z20").clearContent();
    } catch (e2) {}
  }
  
  try { trendSheet.clear().hideSheet(); } catch (e) {}
  
  var charts = dashboard.getCharts();
  charts.forEach(function(chart) { dashboard.removeChart(chart); });
}

/**
 * Build refresh button and timestamp
 */
function buildRefreshSection(dashboard, style) {
  dashboard.getRange("M1").setValue("🔄 REFRESH")
    .setBackground(style.main).setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
  
  dashboard.getRange("M3").setValue("Last Update:")
    .setFontSize(8).setFontWeight("bold").setFontColor(style.accent).setHorizontalAlignment("center");
  
  dashboard.getRange("M4").setValue(Utilities.formatDate(new Date(), "GMT-5", "MMM d, h:mm a"))
    .setFontSize(9).setFontColor(style.main).setHorizontalAlignment("center");
}

/**
 * Build dynamic header with alert styling
 */
function buildHeader(dashboard, style, currentMonth, rawData, runRateDivisor, daysInMonth, monthSheet) {
  var actPpvga = rawData.filter(function(r) { return r[1] === "PPVGA"; }).length;
  var goalPpvga = monthSheet.getRange("I2").getValue() || 0;
  var isAlert = (actPpvga / runRateDivisor * daysInMonth) < (goalPpvga * 0.9);
  
  dashboard.getRange("A1:L1").merge()
    .setValue("🏆 " + currentMonth.toUpperCase() + " PERFORMANCE")
    .setFontFamily("Trebuchet MS").setFontSize(28).setFontColor("#FFFFFF")
    .setBackground(isAlert ? style.alert : style.main)
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
  
  dashboard.setRowHeight(1, 70);
}

/**
 * Build MTD KPI cards - compares current month to FULL previous month
 */
function buildKPICards(dashboard, style, currentMonth, prevMonth, ss) {
  var currentCounts = getMonthTypeCounts(ss, currentMonth);
  var prevCounts = getMonthTypeCounts(ss, prevMonth);
  
  var cards = [
    { 
      range: "A3:C5", 
      title: "PPVGA (MTD)", 
      type: "PPVGA",
      color: style.main,
      currentCount: currentCounts.PPVGA || 0,
      prevCount: prevCounts.PPVGA || 0
    },
    { 
      range: "D3:F5", 
      title: "AIA/AC (MTD)", 
      type: "AIA",
      color: style.accent,
      currentCount: (currentCounts.AIA || 0) + (currentCounts.AIAC || 0) + (currentCounts.AIAB || 0),
      prevCount: (prevCounts.AIA || 0) + (prevCounts.AIAC || 0) + (prevCounts.AIAB || 0)
    },
    { 
      range: "G3:I5", 
      title: "UPG (MTD)", 
      type: "UPG",
      color: "#475569",
      currentCount: currentCounts.UPG || 0,
      prevCount: prevCounts.UPG || 0
    },
    { 
      range: "J3:L5", 
      title: "PLUS1 (MTD)", 
      type: "Plus1",
      color: "#94A3B8",
      currentCount: currentCounts.Plus1 || 0,
      prevCount: prevCounts.Plus1 || 0
    }
  ];
  
  cards.forEach(function(card) {
    buildKPICard(dashboard, card, currentMonth, style);
  });
}

/**
 * Get type counts for entire month
 */
function getMonthTypeCounts(ss, monthName) {
  var counts = { PPVGA: 0, AIA: 0, AIAC: 0, AIAB: 0, UPG: 0, Plus1: 0 };
  
  try {
    var sheet = ss.getSheetByName(monthName);
    if (!sheet) return counts;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return counts;
    
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    data.forEach(function(row) {
      var type = String(row[1] || '').trim();
      if (type && counts.hasOwnProperty(type)) {
        counts[type]++;
      }
    });
  } catch (e) {}
  
  return counts;
}

/**
 * Build single KPI card with direct value comparison
 */
function buildKPICard(dashboard, config, currentMonth, style) {
  var range = dashboard.getRange(config.range);
  var row = range.getRow();
  var col = range.getColumn();
  
  range.setBorder(true, true, true, true, null, null, "#BDC3C7", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground("#FFFFFF");
  
  dashboard.getRange(row, col).setValue(config.title)
    .setFontSize(9).setFontWeight("bold").setFontColor("#64748B").setFontLine("underline");
  
  var currentFormula;
  if (config.type === "AIA") {
    currentFormula = '=COUNTIF(\'' + currentMonth + '\'!B:B,"AIA")+COUNTIF(\'' + currentMonth + '\'!B:B,"AIAC")+COUNTIF(\'' + currentMonth + '\'!B:B,"AIAB")';
  } else {
    currentFormula = '=COUNTIF(\'' + currentMonth + '\'!B:B,"' + config.type + '")';
  }
  
  var mainCell = dashboard.getRange(row + 1, col, 1, 3).merge()
    .setFormula(currentFormula)
    .setFontSize(26).setFontWeight("bold").setFontColor(config.color).setHorizontalAlignment("center");
  
  var arrowText;
  var arrowColor;
  var curr = config.currentCount;
  var prev = config.prevCount;
  
  if (curr > prev) {
    arrowText = "▲ beat last mo (" + prev + ")";
    arrowColor = style.success;
  } else if (curr < prev) {
    arrowText = "▼ behind last mo (" + prev + ")";
    arrowColor = style.alert;
  } else {
    arrowText = "— tied last mo (" + prev + ")";
    arrowColor = "#64748B";
  }
  
  dashboard.getRange(row + 2, col, 1, 3).merge()
    .setValue(arrowText)
    .setFontSize(9).setFontWeight("bold").setFontColor(arrowColor)
    .setHorizontalAlignment("center").setVerticalAlignment("top");
}

/**
 * Build tracker boxes
 */
function buildTrackers(dashboard, style, currentMonth) {
  dashboard.setColumnWidths(1, 12, 110);
  
  var trackers = [
    { 
      title: "📊 PPVGA TRACKING", 
      range: "A8:F11", 
      color: style.main, 
      bg: style.bg, 
      target: "I2",
      actual: "=COUNTIF('" + currentMonth + "'!B:B,\"PPVGA\")", 
      isCurrency: false 
    },
    { 
      title: "💜 AIA TRACKING", 
      range: "H8:M11", 
      color: style.accent, 
      bg: "#F1F5F9", 
      target: "I5",
      actual: "=COUNTIF('" + currentMonth + "'!B:B,\"AIA\")+COUNTIF('" + currentMonth + "'!B:B,\"AIAC\")+COUNTIF('" + currentMonth + "'!B:B,\"AIAB\")", 
      isCurrency: false 
    },
    { 
      title: "🎯 OPS TRACKING", 
      range: "A14:F17", 
      color: "#475569", 
      bg: "#F8FAFC", 
      target: "I7",
      actual: "=IFERROR(VALUE(REGEXEXTRACT(TO_TEXT('" + currentMonth + "'!H7),\"[0-9.]+\")),0)", 
      isCurrency: false 
    },
    { 
      title: "💰 ACCESSORY TRACKING", 
      range: "H14:M17", 
      color: "#334155", 
      bg: "#F1F5F9", 
      target: "I8",
      actual: "=IFERROR(N('" + currentMonth + "'!I9),0)", 
      isCurrency: true 
    }
  ];
  
  trackers.forEach(function(tracker) {
    buildTracker(dashboard, tracker, currentMonth);
  });
}

/**
 * Build single tracker box
 */
function buildTracker(dashboard, config, currentMonth) {
  var r = dashboard.getRange(config.range);
  var row = r.getRow();
  var col = r.getColumn();
  
  dashboard.getRange(row - 1, col, 1, 6).merge()
    .setValue(config.title)
    .setFontWeight("bold").setFontSize(12).setFontColor(config.color)
    .setFontLine("underline").setHorizontalAlignment("left");
  
  r.setBorder(true, true, true, true, null, null, config.color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground(config.bg);
  
  dashboard.getRange(row, col, 1, 6)
    .setValues([["", "GOAL TYPE", "VALUE", "PROGRESS %", "REMAINING", ""]])
    .setFontWeight("bold").setFontColor("#334155").setHorizontalAlignment("center");
  
  var baseVal = dashboard.getRange(row + 1, col + 2)
    .setFormula("=IFERROR(N('" + currentMonth + "'!" + config.target + "),0)");
  
  var stretchVal = dashboard.getRange(row + 2, col + 2)
    .setFormula("=ROUNDUP(" + baseVal.getA1Notation() + "*1.25)");
  
  var actVal = dashboard.getRange(row + 3, col + 2)
    .setFormula(config.actual).setFontWeight("bold");
  
  dashboard.getRange(row + 1, col + 1).setValue("Base (100%)");
  dashboard.getRange(row + 2, col + 1).setValue("Stretch (125%)");
  dashboard.getRange(row + 3, col + 1).setValue("Current Actual").setFontWeight("bold");
  
  dashboard.getRange(row + 1, col + 3)
    .setFormula("=IFERROR(" + actVal.getA1Notation() + "/" + baseVal.getA1Notation() + ",0)")
    .setNumberFormat("0%").setFontWeight("bold");
  
  dashboard.getRange(row + 2, col + 3)
    .setFormula("=IFERROR(" + actVal.getA1Notation() + "/" + stretchVal.getA1Notation() + ",0)")
    .setNumberFormat("0%").setFontWeight("bold");
  
  dashboard.getRange(row + 2, col + 4)
    .setFormula("=MAX(0," + stretchVal.getA1Notation() + "-" + actVal.getA1Notation() + ")")
    .setFontColor("#C0392B").setFontWeight("bold");
  
  dashboard.getRange(row + 1, col, 3, 6).setHorizontalAlignment("center");
  
  if (config.isCurrency) {
    baseVal.setNumberFormat("$#,##0").setFontWeight("bold");
    stretchVal.setNumberFormat("$#,##0").setFontWeight("bold");
    actVal.setNumberFormat("$#,##0").setFontWeight("bold");
    dashboard.getRange(row + 2, col + 4).setNumberFormat("$#,##0");
  }
}

/**
 * Build pace and projection section
 */
function buildPaceSection(dashboard, style, currentMonth, workingDaysLeft, daysInMonth, daysPassed, runRateDivisor) {
  dashboard.getRange("A19:M19").merge()
    .setValue("📅 PACE & PROJECTION (Targeting 125% Stretch)")
    .setFontSize(12).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground(style.pace).setHorizontalAlignment("center");
  
  dashboard.getRange("A20:M22").setBackground(style.bg)
    .setBorder(true, true, true, true, null, null, style.pace, SpreadsheetApp.BorderStyle.SOLID);
  
  dashboard.getRange("A20:M20").setFontWeight("bold");
  dashboard.getRange("A20").setValue("Selling Days Left").setHorizontalAlignment("center");
  dashboard.getRange("C20:D20").merge().setValue("PPVGA (125%)").setHorizontalAlignment("center");
  dashboard.getRange("F20:G20").merge().setValue("AIA (125%)").setHorizontalAlignment("center");
  dashboard.getRange("I20:J20").merge().setValue("ACC $ (125%)").setHorizontalAlignment("center");
  
  dashboard.getRange("A21:A22").merge()
    .setValue(workingDaysLeft)
    .setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  dashboard.getRange("B21").setValue("Daily Pace Needed:")
    .setFontSize(8).setFontColor("#64748B").setFontWeight("bold").setVerticalAlignment("middle");
  dashboard.getRange("B22").setValue("MTD Projection:")
    .setFontSize(8).setFontColor("#64748B").setFontWeight("bold").setVerticalAlignment("middle");
  
  var div = Math.max(1, workingDaysLeft);
  
  var ppvgaPace = dashboard.getRange("C21:D21").merge()
    .setFormula('=ROUNDUP(MAX(0,(C10-C11)/' + div + '),1)')
    .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  
  var aiaPace = dashboard.getRange("F21:G21").merge()
    .setFormula('=ROUNDUP(MAX(0,(J10-J11)/' + div + '),1)')
    .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  
  var accPace = dashboard.getRange("I21:J21").merge()
    .setFormula('=MAX(0,(J16-J17)/' + div + ')')
    .setNumberFormat("$#,##0").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  
  dashboard.getRange("C22:D22").merge()
    .setFormula(buildGapFormula("C11", "C10", runRateDivisor, daysInMonth, daysPassed, false))
    .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");
  
  dashboard.getRange("F22:G22").merge()
    .setFormula(buildGapFormula("J11", "J10", runRateDivisor, daysInMonth, daysPassed, false))
    .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");
  
  dashboard.getRange("I22:J22").merge()
    .setFormula(buildGapFormula("J17", "J16", runRateDivisor, daysInMonth, daysPassed, true))
    .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");
  
  applyPaceConditionalFormatting(dashboard, ppvgaPace, aiaPace, accPace, daysInMonth);
}

/**
 * Build gap/projection formula
 */
function buildGapFormula(actualCell, baseCell, divisor, daysInMonth, daysPassed, isCurrency) {
  if (isCurrency) {
    return '=TEXT((' + actualCell + '/' + divisor + ')*' + daysInMonth + ',"$#,##0")&" ("&IF((' + actualCell + '-((' + baseCell + '/' + daysInMonth + ')*' + daysPassed + '))>=0,"+$"&TEXT(' + actualCell + '-((' + baseCell + '/' + daysInMonth + ')*' + daysPassed + '),"#,##0"),"-$"&TEXT(ABS(' + actualCell + '-((' + baseCell + '/' + daysInMonth + ')*' + daysPassed + ')),"#,##0"))&")"';
  } else {
    return '=ROUND(((' + actualCell + '/' + divisor + ')*' + daysInMonth + '),0)&" ("&IF(ROUND(' + actualCell + '-(' + baseCell + '/' + daysInMonth + ')*' + daysPassed + ',1)>=0,"+"&ROUND(' + actualCell + '-(' + baseCell + '/' + daysInMonth + ')*' + daysPassed + ',1),ROUND(' + actualCell + '-(' + baseCell + '/' + daysInMonth + ')*' + daysPassed + ',1))&" units)"';
  }
}

/**
 * Apply pace heatmap conditional formatting
 */
function applyPaceConditionalFormatting(dashboard, ppvgaPace, aiaPace, accPace, daysInMonth) {
  var rules = [];
  
  var paceConfigs = [
    { range: ppvgaPace, target: "C10" },
    { range: aiaPace, target: "J10" },
    { range: accPace, target: "J16" }
  ];
  
  paceConfigs.forEach(function(cfg) {
    var idealPace = '(' + cfg.target + '/' + daysInMonth + ')';
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + cfg.range.getA1Notation() + '<=' + idealPace)
      .setFontColor("#27AE60")
      .setRanges([cfg.range])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(' + cfg.range.getA1Notation() + '>' + idealPace + ',' + cfg.range.getA1Notation() + '<=(' + idealPace + '*1.5))')
      .setFontColor("#F39C12")
      .setRanges([cfg.range])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + cfg.range.getA1Notation() + '>(' + idealPace + '*1.5)')
      .setFontColor("#C0392B")
      .setRanges([cfg.range])
      .build());
  });
  
  dashboard.setConditionalFormatRules(rules);
}

/**
 * Build charts
 */
function buildCharts(dashboard, trendSheet, style) {
  trendSheet.getRange("A1:C4").setValues([
    ["Metric", "Actual", "Gap"],
    ["PPVGA", "='Dashboard'!C11", "='Dashboard'!E10"],
    ["AIA", "='Dashboard'!J11", "='Dashboard'!L10"],
    ["Ops", "='Dashboard'!C17", "='Dashboard'!E16"]
  ]);
  
  trendSheet.getRange("A6:C7").setValues([
    ["Metric", "Actual", "Gap"],
    ["ACC $", "='Dashboard'!J17", "='Dashboard'!L16"]
  ]);
  
  var unitChart = dashboard.newChart()
    .asBarChart()
    .addRange(trendSheet.getRange("A1:C4"))
    .setStacked()
    .setTitle("Unit Stretch Progress")
    .setOption("series", { 
      0: { color: style.main, dataLabel: 'value' }, 
      1: { color: '#E2E8F0', dataLabel: 'value' } 
    })
    .setOption("width", 750)
    .setOption("height", 240)
    .setPosition(25, 1, 0, 0)
    .build();
  dashboard.insertChart(unitChart);
  
  var accChart = dashboard.newChart()
    .asColumnChart()
    .addRange(trendSheet.getRange("A6:C7"))
    .setStacked()
    .setTitle("Acc $ Progress")
    .setOption("series", { 
      0: { color: style.accent, dataLabel: 'value' }, 
      1: { color: '#F1F5F9', dataLabel: 'value' } 
    })
    .setOption("width", 250)
    .setOption("height", 350)
    .setPosition(25, 10, 0, 0)
    .build();
  dashboard.insertChart(accChart);
}
