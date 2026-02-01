/**
 * =================================================================================
 * ENHANCED YTD REPORT GENERATOR (v4.9 - EXACT STYLE + ACC FIX)
 * Uses UnifiedDataAccess with unlimited row processing + Robust Acc Parsing
 * =================================================================================
 */

function createYTDReport() {
  if (!checkExecutionQuota()) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName("YTD Report") || ss.insertSheet("YTD Report");
  var monthNames = CONFIG.MONTH_NAMES;
  
  // Clear
  reportSheet.clear();
  var charts = reportSheet.getCharts();
  for (var i = 0; i < charts.length; i++) reportSheet.removeChart(charts[i]);
  
  // Init Totals
  var totals = { 
    "Type": { "PPVGA":0, "UPG":0, "Plus1":0, "AIA":0 }, 
    "Device": {}, 
    "Plan": {}, 
    "Accessories": 0 
  };
  
  var mData = { "PPVGA": [], "AIA": [], "UPG": [], "Plus1": [], "Accessories": [] };
  var monthlyTable = [["Metric"].concat(monthNames)];
  
  // Process Each Month
  for (var m = 0; m < monthNames.length; m++) {
    var mName = monthNames[m];
    
    // Init month stats
    var mStats = { PPVGA:0, AIA:0, UPG:0, Plus1:0, Acc:0 };
    
    try {
      // 1. Get Accessories Goal (I9) - Safe Read & Parsing
      try {
        var sheet = ss.getSheetByName(mName);
        if (sheet) {
          var rawAcc = sheet.getRange("I9").getValue();
          // 🆕 ROBUST PARSING FIX
          if (rawAcc !== "" && rawAcc !== null) {
            if (typeof rawAcc === 'number') {
              mStats.Acc = rawAcc;
            } else {
              // Strip everything except digits, decimal point, and negative sign
              mStats.Acc = parseFloat(String(rawAcc).replace(/[^0-9.-]+/g, "")) || 0;
            }
          }
        }
      } catch (e) {}
      
      // 2. Get Row Data via Unified Access
      var result = getSheetDataWithMapping(mName, { useCache: true });
      var data = result.data;
      var map = result.columnMapping;
      
      // 3. Process Rows
      data.forEach(function(row) {
        // Safe value extraction
        var type = String(row[map.TYPE_INDEX] || '').trim().toLowerCase();
        var dev = String(row[map.DEVICE_INDEX] || '').trim();
        var plan = String(row[map.PLAN_INDEX] || '').trim();
        
        // Skip empty rows
        if (!type && !dev && !plan) return;
        
        // Count Types
        if (type === 'ppvga') { mStats.PPVGA++; totals.Type.PPVGA++; }
        else if (type === 'upg') { mStats.UPG++; totals.Type.UPG++; }
        else if (type === 'plus1') { mStats.Plus1++; totals.Type.Plus1++; }
        else if (type.includes('aia')) { mStats.AIA++; totals.Type.AIA++; }
        
        // Count Devices
        if (dev) {
          var dKey = dev.charAt(0).toUpperCase() + dev.slice(1).toLowerCase();
          totals.Device[dKey] = (totals.Device[dKey] || 0) + 1;
        }
        
        // Count Plans
        if (plan) {
          var pKey = plan.charAt(0).toUpperCase() + plan.slice(1).toLowerCase();
          totals.Plan[pKey] = (totals.Plan[pKey] || 0) + 1;
        }
      });
      
    } catch (e) {
      // Month likely missing or empty, zeros remain
    }
    
    // Store Month Data
    mData.PPVGA.push(mStats.PPVGA);
    mData.AIA.push(mStats.AIA);
    mData.UPG.push(mStats.UPG);
    mData.Plus1.push(mStats.Plus1);
    mData.Accessories.push(mStats.Acc);
    totals.Accessories += mStats.Acc;
  }
  
  // Build Main Table
  monthlyTable.push(["PPVGA"].concat(mData.PPVGA));
  monthlyTable.push(["AIA"].concat(mData.AIA));
  monthlyTable.push(["UPG"].concat(mData.UPG));
  monthlyTable.push(["Plus1"].concat(mData.Plus1));
  monthlyTable.push(["Accessories $"].concat(mData.Accessories));
  
  // Render Header
  var style = CONFIG.STYLES;
  reportSheet.getRange("A1:M1").merge().setValue("📊 YTD PERFORMANCE SUMMARY")
    .setBackground(style.header).setFontColor("white").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
    
  reportSheet.getRange("A2").setValue("Generated: " + new Date().toLocaleString());
  
  // Render Monthly Table
  reportSheet.getRange("A4:M4").merge().setValue("📈 MONTHLY TRENDS").setBackground(style.sub).setFontColor("white").setFontWeight("bold");
  
  var tableRange = reportSheet.getRange(5, 1, 6, 13);
  tableRange.setValues(monthlyTable);
  tableRange.setBorder(true, true, true, true, true, true, style.border, SpreadsheetApp.BorderStyle.SOLID);
  reportSheet.getRange(5, 1, 1, 13).setFontWeight("bold").setBackground("#F1F5F9");
  
  // Currency Format
  reportSheet.getRange(10, 2, 1, 12).setNumberFormat("$#,##0");
  
  // Render Summaries
  var row = 12;
  
  // Type Summary
  renderSummaryTable(reportSheet, row, 1, "Sales Type", totals.Type);
  
  // Device Summary
  renderSummaryTable(reportSheet, row, 4, "Device Brand", totals.Device);
  
  // Plan Summary
  renderSummaryTable(reportSheet, row, 7, "Rate Plan", totals.Plan);
  
  // Total Accessories Card
  reportSheet.getRange(row, 10, 1, 2).merge().setValue("💰 Total Accessories")
    .setBackground(style.header).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange(row+1, 10, 1, 2).merge().setValue(totals.Accessories)
    .setNumberFormat("$#,##0").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
    
  // Formatting
  reportSheet.setColumnWidths(1, 13, 100);
  reportSheet.setHiddenGridlines(true);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ YTD Report Generated", "Success", 3);
}

/**
 * Helper to render summary tables
 */
function renderSummaryTable(sheet, startRow, startCol, title, dataObj) {
  var keys = Object.keys(dataObj).sort();
  var values = [[title, "Total"]];
  
  keys.forEach(function(k) {
    if (dataObj[k] > 0 || title !== "Sales Type") { // Filter 0 types, keep others
      values.push([k, dataObj[k]]);
    }
  });
  
  if (values.length > 1) {
    var range = sheet.getRange(startRow, startCol, values.length, 2);
    range.setValues(values);
    range.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(startRow, startCol, 1, 2).setBackground("#475569").setFontColor("white").setFontWeight("bold");
  }
}
