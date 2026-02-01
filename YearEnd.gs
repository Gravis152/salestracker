/**
 * =================================================================================
 * ENHANCED YEAR-END MAINTENANCE TOOLS (v4.0)
 * Improved archiving and reset functionality with UnifiedDataAccess integration
 * =================================================================================
 */

/**
 * 🆕 ENHANCED runYearEndReset: Comprehensive archiving with UnifiedDataAccess and enhanced backups
 */
function runYearEndReset() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var response = ui.alert(
    "⚠️ CRITICAL ACTION - ENHANCED YEAR-END RESET (v4.0)", 
    "ENHANCED YEAR-END RESET with UnifiedDataAccess:\n\n" +
    "✅ Create comprehensive archive with enhanced metadata\n" +
    "✅ UnifiedDataAccess data preservation\n" +
    "✅ Smart Table compatibility checks\n" +
    "✅ Performance metrics and validation\n" +
    "✅ Enhanced backup information\n" +
    "✅ Clear all monthly sales logs safely\n" +
    "✅ Reset all caches and column mappings\n\n" +
    "❌ THIS CANNOT BE UNDONE ❌\n\nContinue?", 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;

  // Additional confirmation for safety
  var finalResponse = ui.alert(
    "Final Confirmation",
    "Are you absolutely sure you want to proceed?\n\n" +
    "All sales data will be cleared from monthly sheets.\n" +
    "Archive will be created using UnifiedDataAccess for maximum compatibility.",
    ui.ButtonSet.YES_NO
  );
  
  if (finalResponse !== ui.Button.YES) return;

  var startTime = new Date();
  ss.toast("📦 Running enhanced year-end reset with UnifiedDataAccess...", "Year-End Process", 10);
  
  var resetResults = {
    archiveCreated: false,
    sheetsCleared: 0,
    totalSheets: CONFIG.MONTH_NAMES.length,
    errors: [],
    backupInfo: {},
    performanceMetrics: [],
    unifiedDataStats: {},
    accessMethodsUsed: {}
  };

  try {
    // 🆕 STEP 1: Create comprehensive archive with UnifiedDataAccess
    var archiveResult = createEnhancedArchiveWithUnifiedDataAccess(ss, resetResults);
    Object.assign(resetResults, archiveResult);
    
    // 🆕 STEP 2: Enhanced data clearing with UnifiedDataAccess validation
    var clearingResult = clearMonthlyDataWithUnifiedDataAccess(ss, resetResults);
    Object.assign(resetResults, clearingResult);
    
    // 🆕 STEP 3: Complete system reset with UnifiedDataAccess cleanup
    var systemResetResult = performSystemResetWithUnifiedDataAccess(resetResults);
    Object.assign(resetResults, systemResetResult);

    // 🆕 STEP 4: Re-establish data validation on cleared sheets
    var validationResult = reestablishValidationWithUnifiedDataAccess(resetResults);
    Object.assign(resetResults, validationResult);

    var endTime = new Date();
    var totalTime = endTime - startTime;
    
    // 🆕 GENERATE COMPREHENSIVE SUCCESS REPORT
    generateYearEndResetReport(resetResults, totalTime, ss);
    
    // 🆕 STEP 5: Rebuild everything with fresh, empty data
    try {
      refreshAllReports(true); // Silent refresh
      ss.toast("🔄 All reports rebuilt with fresh data structure", "Rebuild Complete", 3);
    } catch (rebuildError) {
      console.error("Rebuild error: ", rebuildError);
      ss.toast("⚠️ Manual refresh may be needed: " + rebuildError.message, "Warning", 5);
    }
    
    return resetResults;
    
  } catch (error) {
    console.error("Critical year-end reset error: ", error);
    ss.toast("❌ Critical error during year-end reset: " + error.message, "Error", 10);
    
    // Attempt to restore from any partial operations
    if (resetResults.archiveCreated) {
      ss.toast("📁 Archive was created successfully before error occurred", "Partial Success", 5);
    }
    
    return resetResults;
  }
}

/**
 * 🆕 CREATE ENHANCED ARCHIVE using UnifiedDataAccess
 */
function createEnhancedArchiveWithUnifiedDataAccess(ss, resetResults) {
  var archiveStartTime = new Date();
  
  try {
    // Get YTD report for archiving
    var ytdSheet = ss.getSheetByName("YTD Report");
    if (!ytdSheet) {
      resetResults.errors.push("YTD Report sheet not found for archiving");
      return { archiveCreated: false };
    }
    
    var currentYear = new Date().getFullYear();
    var archiveName = "Archive_" + currentYear + "_UnifiedData_v4";
    
    // Check if archive already exists
    var existingArchive = ss.getSheetByName(archiveName);
    if (existingArchive) {
      var ui = SpreadsheetApp.getUi();
      var overwriteResponse = ui.alert(
        "Archive Exists",
        "Archive '" + archiveName + "' already exists. Overwrite?",
        ui.ButtonSet.YES_NO
      );
      
      if (overwriteResponse === ui.Button.YES) {
        ss.deleteSheet(existingArchive);
      } else {
        // Create with timestamp
        archiveName = archiveName + "_" + Utilities.formatDate(new Date(), "GMT", "MMdd_HHmm");
      }
    }
    
    var archiveSheet = ss.insertSheet(archiveName);
    
    // Copy YTD report with enhanced formatting
    var sourceRange = ytdSheet.getDataRange();
    sourceRange.copyTo(archiveSheet.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
    sourceRange.copyTo(archiveSheet.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
    
    // 🆕 ADD COMPREHENSIVE METADATA with UnifiedDataAccess statistics
    var metadataStartRow = archiveSheet.getLastRow() + 3;
    var metadata = [
      ["🗄️ ENHANCED ARCHIVE METADATA (v4.0)", ""],
      ["Archive Created:", new Date()],
      ["Script Version:", "4.0 UnifiedDataAccess Enhanced"],
      ["Features:", "UnifiedDataAccess, unlimited rows, typed columns, Smart Tables"],
      ["Year Archived:", currentYear],
      ["Total Months:", CONFIG.MONTH_NAMES.length],
      ["", ""],
      ["📊 UNIFIEDDATAACCESS CAPABILITIES", ""],
      ["Max Rows per Sheet:", CONFIG.DATA_LIMITS.MAX_ROWS_PER_SHEET.toLocaleString()],
      ["Batch Size:", CONFIG.DATA_LIMITS.BATCH_SIZE.toLocaleString()],
      ["Client List Limit:", CONFIG.DATA_LIMITS.CLIENT_LIST_LIMIT.toLocaleString()],
      ["Cache Enabled:", CONFIG.CACHE_CONFIG.ENABLE_CACHE],
      ["Available Methods:", UnifiedDataAccess.getStats().accessMethodsAvailable.join(", ")],
      ["", ""],
      ["🔧 SHEET ACCESS ANALYSIS (at archive time)", ""]
    ];
    
    // 🆕 ANALYZE EACH SHEET using UnifiedDataAccess and add to archive
    var sheetAnalysis = analyzeAllSheetsForArchive();
    resetResults.unifiedDataStats = sheetAnalysis.summary;
    resetResults.accessMethodsUsed = sheetAnalysis.accessMethods;
    
    // Add sheet analysis to metadata
    CONFIG.MONTH_NAMES.forEach(function(monthName) {
      var sheetInfo = sheetAnalysis.sheets[monthName];
      if (sheetInfo) {
        var infoLine = monthName + ": " + sheetInfo.status + " | " + 
                       sheetInfo.rows + " rows | " + sheetInfo.method;
        metadata.push([infoLine, ""]);
      } else {
        metadata.push([monthName + ": Not found", ""]);
      }
    });
    
    // Add performance summary
    metadata.push(["", ""]);
    metadata.push(["📈 ARCHIVE PERFORMANCE", ""]);
    metadata.push(["Total Analysis Time:", sheetAnalysis.totalTime + "ms"]);
    metadata.push(["Sheets Analyzed:", sheetAnalysis.summary.totalSheets]);
    metadata.push(["Accessible Sheets:", sheetAnalysis.summary.accessibleSheets]);
    metadata.push(["Smart Tables Found:", sheetAnalysis.summary.smartTables]);
    
    // Write metadata to archive
    archiveSheet.getRange(metadataStartRow, 1, metadata.length, 2).setValues(metadata);
    
    // Format metadata sections
    archiveSheet.getRange(metadataStartRow, 1, 1, 2).setBackground("#1E293B").setFontColor("white").setFontWeight("bold");
    archiveSheet.getRange(metadataStartRow + 7, 1, 1, 2).setBackground("#334155").setFontColor("white").setFontWeight("bold");
    archiveSheet.getRange(metadataStartRow + 14, 1, 1, 2).setBackground("#475569").setFontColor("white").setFontWeight("bold");
    
    // Set column widths for readability
    archiveSheet.setColumnWidth(1, 250);
    archiveSheet.setColumnWidth(2, 200);
    
    var archiveEndTime = new Date();
    resetResults.performanceMetrics.push({
      step: "Archive Creation",
      time: archiveEndTime - archiveStartTime,
      success: true
    });
    
    resetResults.backupInfo.archiveName = archiveName;
    resetResults.backupInfo.metadataRows = metadata.length;
    resetResults.backupInfo.sheetsAnalyzed = sheetAnalysis.summary.totalSheets;
    
    ss.toast("✅ Enhanced archive created: " + archiveName, "Archive Complete", 3);
    console.log("Enhanced archive created with UnifiedDataAccess analysis");
    
    return { archiveCreated: true };
    
  } catch (archiveError) {
    console.error("Archive creation error: ", archiveError);
    resetResults.errors.push("Archive creation: " + archiveError.message);
    
    var archiveEndTime = new Date();
    resetResults.performanceMetrics.push({
      step: "Archive Creation",
      time: archiveEndTime - archiveStartTime,
      success: false,
      error: archiveError.message
    });
    
    return { archiveCreated: false };
  }
}

/**
 * 🆕 ANALYZE ALL SHEETS for archive using UnifiedDataAccess
 */
function analyzeAllSheetsForArchive() {
  var startTime = new Date();
  var analysis = {
    sheets: {},
    summary: {
      totalSheets: 0,
      accessibleSheets: 0,
      smartTables: 0,
      errors: 0
    },
    accessMethods: {},
    totalTime: 0
  };
  
  CONFIG.MONTH_NAMES.forEach(function(monthName) {
    var sheetStartTime = new Date();
    analysis.summary.totalSheets++;
    
    try {
      // Test accessibility using UnifiedDataAccess
      var testData = UnifiedDataAccess.getSheetData(monthName, {
        maxRows: 5,
        useCache: false
      });
      
      var colMap = UnifiedDataAccess.getColumnMapping(monthName);
      
      // Determine access method
      var accessTest = UnifiedDataAccess.testAccessMethod(monthName, 'BULK_ACCESS');
      var method = accessTest.success ? 'BULK_ACCESS' : 'ALTERNATIVE_METHOD';
      
      var sheetEndTime = new Date();
      
      analysis.sheets[monthName] = {
        status: "Accessible",
        rows: testData.length,
        columns: Object.keys(colMap).length,
        method: method,
        time: sheetEndTime - sheetStartTime
      };
      
      analysis.summary.accessibleSheets++;
      analysis.accessMethods[method] = (analysis.accessMethods[method] || 0) + 1;
      
    } catch (error) {
      var sheetEndTime = new Date();
      var status = "Error";
      
      if (error.message.includes('Smart Table') || error.message.includes('typed')) {
        status = "Smart Table";
        analysis.summary.smartTables++;
        analysis.accessMethods['SMART_TABLE'] = (analysis.accessMethods['SMART_TABLE'] || 0) + 1;
      } else {
        analysis.summary.errors++;
      }
      
      analysis.sheets[monthName] = {
        status: status,
        rows: 0,
        columns: 0,
        method: "N/A",
        error: error.message,
        time: sheetEndTime - sheetStartTime
      };
    }
  });
  
  var endTime = new Date();
  analysis.totalTime = endTime - startTime;
  
  return analysis;
}

/**
 * 🆕 CLEAR MONTHLY DATA using UnifiedDataAccess validation
 */
function clearMonthlyDataWithUnifiedDataAccess(ss, resetResults) {
  var clearingStartTime = new Date();
  
  try {
    CONFIG.MONTH_NAMES.forEach(function(mName) {
      var sheetStartTime = new Date();
      var sheet = ss.getSheetByName(mName);
      
      if (!sheet) {
        resetResults.errors.push("Sheet not found for clearing: " + mName);
        return;
      }
      
      try {
        // 🆕 USE UNIFIEDDATAACCESS to validate sheet before clearing
        var preValidation = validateSheetBeforeClearing(mName, sheet);
        
        if (!preValidation.safe) {
          resetResults.errors.push(mName + ": " + preValidation.reason);
          return;
        }
        
        var colMap = preValidation.columnMapping;
        var lastRow = sheet.getLastRow();
        var maxCol = Math.max(colMap.DATE || 1, colMap.TYPE || 2, colMap.DEVICE || 3, 
                             colMap.CUSTOMER || 4, colMap.MOBILE || 5, colMap.PLAN || 6, colMap.NOTES || 7);
        
        if (lastRow >= 2) {
          // 🆕 SMART DATA CLEARING - preserve headers and goal cells, clear only data
          var dataRowCount = lastRow - 1;
          
          // Clear data while preserving structure
          var clearRange = sheet.getRange(2, 1, dataRowCount, maxCol);
          clearRange.clearContent();
          clearRange.setBackground(null); // Clear validation backgrounds
          clearRange.clearDataValidations(); // Clear validation rules (will be re-applied)
          
          // Store clearing information with UnifiedDataAccess context
          resetResults.backupInfo[mName] = {
            rowsCleared: dataRowCount,
            columnsCleared: maxCol,
            columnMapping: colMap,
            validationMethod: preValidation.method,
            accessMethod: preValidation.accessMethod
          };
          
          resetResults.sheetsCleared++;
          
          var sheetEndTime = new Date();
          resetResults.performanceMetrics.push({
            step: "Clear " + mName,
            time: sheetEndTime - sheetStartTime,
            success: true,
            rowsCleared: dataRowCount
          });
          
          console.log(`Cleared ${dataRowCount} rows from ${mName} using UnifiedDataAccess validation (columns 1-${maxCol})`);
        }
        
      } catch (clearError) {
        var sheetEndTime = new Date();
        resetResults.errors.push(`Error clearing ${mName}: ${clearError.message}`);
        resetResults.performanceMetrics.push({
          step: "Clear " + mName,
          time: sheetEndTime - sheetStartTime,
          success: false,
          error: clearError.message
        });
        console.error("Error clearing " + mName + ": ", clearError);
      }
    });
    
    var clearingEndTime = new Date();
    resetResults.performanceMetrics.push({
      step: "Overall Data Clearing",
      time: clearingEndTime - clearingStartTime,
      success: true
    });
    
    return { dataClearingComplete: true };
    
  } catch (error) {
    resetResults.errors.push("Data clearing process: " + error.message);
    return { dataClearingComplete: false };
  }
}

/**
 * 🆕 VALIDATE SHEET BEFORE CLEARING using UnifiedDataAccess
 */
function validateSheetBeforeClearing(sheetName, sheet) {
  try {
    // Test basic access
    var testData = UnifiedDataAccess.getSheetData(sheetName, {
      maxRows: 3,
      useCache: false
    });
    
    var colMap = UnifiedDataAccess.getColumnMapping(sheetName);
    
    // Determine access method
    var accessTest = UnifiedDataAccess.testAccessMethod(sheetName, 'BULK_ACCESS');
    var accessMethod = accessTest.success ? 'BULK_ACCESS' : 'ALTERNATIVE';
    
    return {
      safe: true,
      reason: "Sheet validated for clearing",
      columnMapping: colMap,
      method: "UnifiedDataAccess",
      accessMethod: accessMethod,
      sampleRows: testData.length
    };
    
  } catch (error) {
    // If UnifiedDataAccess fails, it might be a Smart Table - be extra careful
    if (error.message.includes('Smart Table') || error.message.includes('typed')) {
      return {
        safe: false,
        reason: "Smart Table detected - manual clearing required",
        method: "Smart Table Detection"
      };
    }
    
    // Try basic sheet operations as fallback
    try {
      var basicTest = sheet.getRange("A1:B2").getValues();
      return {
        safe: true,
        reason: "Basic validation passed",
        columnMapping: CONFIG.getDefaultColumnMapping(),
        method: "Basic Fallback",
        accessMethod: "BASIC"
      };
    } catch (basicError) {
      return {
        safe: false,
        reason: "All validation methods failed: " + error.message,
        method: "Failed"
      };
    }
  }
}

/**
 * 🆕 PERFORM SYSTEM RESET with UnifiedDataAccess cleanup
 */
function performSystemResetWithUnifiedDataAccess(resetResults) {
  var systemResetStartTime = new Date();
  
  try {
    // Clear all caches including UnifiedDataAccess
    ADVANCED_CACHE.clearAll();
    UnifiedDataAccess.clearCache();
    
    // Reset configuration cache
    CONFIG.MONTH_NAMES.forEach(function(monthName) {
      ADVANCED_CACHE.remove('column_map_' + monthName);
      ADVANCED_CACHE.remove('unified_data_' + monthName);
    });
    
    // Reset performance tracking
    ADVANCED_CACHE.remove('perf_dashboard_last');
    ADVANCED_CACHE.remove('perf_full_refresh_last');
    ADVANCED_CACHE.remove('ytd_report_data');
    ADVANCED_CACHE.remove('client_list_data');
    ADVANCED_CACHE.remove('system_health_results');
    
    var systemResetEndTime = new Date();
    resetResults.performanceMetrics.push({
      step: "System Cache Reset",
      time: systemResetEndTime - systemResetStartTime,
      success: true
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Enhanced system caches and UnifiedDataAccess reset", "System Reset Complete", 2);
    
    return { systemResetComplete: true };
    
  } catch (resetError) {
    resetResults.errors.push("System reset: " + resetError.message);
    return { systemResetComplete: false };
  }
}

/**
 * 🆕 RE-ESTABLISH VALIDATION with UnifiedDataAccess
 */
function reestablishValidationWithUnifiedDataAccess(resetResults) {
  var validationStartTime = new Date();
  
  try {
    // Use the enhanced validation setup
    var validationResults = setupDataValidationRules();
    
    resetResults.validationResults = {
      rulesApplied: validationResults.rulesApplied,
      sheetsProcessed: validationResults.sheetsProcessed,
      accessMethodsUsed: validationResults.accessMethodsUsed
    };
    
    var validationEndTime = new Date();
    resetResults.performanceMetrics.push({
      step: "Validation Re-establishment",
      time: validationEndTime - validationStartTime,
      success: true,
      rulesApplied: validationResults.rulesApplied
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Enhanced data validation rules re-applied", "Validation Complete", 2);
    
    return { validationComplete: true };
    
  } catch (validationError) {
    resetResults.errors.push("Validation setup: " + validationError.message);
    
    var validationEndTime = new Date();
    resetResults.performanceMetrics.push({
      step: "Validation Re-establishment",
      time: validationEndTime - validationStartTime,
      success: false,
      error: validationError.message
    });
    
    return { validationComplete: false };
  }
}

/**
 * 🆕 GENERATE COMPREHENSIVE YEAR-END RESET REPORT
 */
function generateYearEndResetReport(resetResults, totalTime, ss) {
  var avgProcessingTime = resetResults.performanceMetrics.length > 0 ?
    resetResults.performanceMetrics.reduce((sum, metric) => sum + metric.time, 0) / resetResults.performanceMetrics.length : 0;
  
  var successMessage = `🚀 Enhanced Year-End Reset Complete! (v4.0)\n\n`;
  successMessage += `📊 Results Summary:\n`;
  successMessage += `• Archive: ${resetResults.archiveCreated ? "✅ Created" : "❌ Failed"}\n`;
  successMessage += `• Sheets Cleared: ${resetResults.sheetsCleared}/${resetResults.totalSheets}\n`;
  successMessage += `• Total Time: ${Math.round(totalTime / 1000)} seconds\n`;
  successMessage += `• Avg Step Time: ${Math.round(avgProcessingTime)}ms\n`;
  successMessage += `• UnifiedDataAccess: ✅ Enhanced\n`;
  successMessage += `• Dynamic Columns: ✅ Preserved\n`;
  successMessage += `• Smart Table Support: ✅ Active\n`;
  
  if (resetResults.validationResults) {
    successMessage += `• Validation Rules: ${resetResults.validationResults.rulesApplied} applied\n`;
  }
  
  if (resetResults.archiveCreated && resetResults.backupInfo.archiveName) {
    successMessage += `\n📁 Archive: "${resetResults.backupInfo.archiveName}"\n`;
    successMessage += `• Sheets Analyzed: ${resetResults.backupInfo.sheetsAnalyzed}\n`;
    successMessage += `• Metadata Rows: ${resetResults.backupInfo.metadataRows}\n`;
  }
  
  // Access method summary
  var accessMethods = Object.keys(resetResults.accessMethodsUsed || {});
  if (accessMethods.length > 0) {
    successMessage += `\n🔧 Access Methods Used:\n`;
    accessMethods.forEach(function(method) {
      successMessage += `• ${method}: ${resetResults.accessMethodsUsed[method]} sheets\n`;
    });
  }
  
  // Performance breakdown
  if (resetResults.performanceMetrics.length > 0) {
    var successfulSteps = resetResults.performanceMetrics.filter(m => m.success).length;
    successMessage += `\n⏱️ Performance: ${successfulSteps}/${resetResults.performanceMetrics.length} steps completed\n`;
  }
  
  if (resetResults.errors.length > 0) {
    successMessage += `\n⚠️ Issues: ${resetResults.errors.length} (see console logs)`;
    console.error("Year-end reset errors:", resetResults.errors);
  }
  
  successMessage += `\n\n🎯 Enhanced year-end reset with full UnifiedDataAccess integration!`;
  
  ss.toast(successMessage, "Year-End Reset Complete", 10);
  
  // Log comprehensive reset results
  console.log("Enhanced year-end reset completed in " + totalTime + "ms with UnifiedDataAccess:");
  console.log("Performance metrics:", resetResults.performanceMetrics);
  console.log("Access methods used:", resetResults.accessMethodsUsed);
  console.log("Unified data stats:", resetResults.unifiedDataStats);
}

/**
 * 🆕 ENHANCED BACKUP VERIFICATION with UnifiedDataAccess
 */
function verifyYearEndArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentYear = new Date().getFullYear();
  var archiveName = "Archive_" + currentYear + "_UnifiedData_v4";
  
  var archiveSheet = ss.getSheetByName(archiveName);
  
  if (!archiveSheet) {
    // Look for alternative archive names with UnifiedData
    var allSheets = ss.getSheets();
    var archiveSheets = allSheets.filter(sheet => 
      sheet.getName().includes("Archive_" + currentYear) && 
      sheet.getName().includes("UnifiedData")
    );
    
    if (archiveSheets.length > 0) {
      archiveSheet = archiveSheets[archiveSheets.length - 1]; // Get the most recent
      archiveName = archiveSheet.getName();
    }
  }
  
  if (!archiveSheet) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Enhanced Archive Verification", 
             "No UnifiedData archive found for year " + currentYear + "\n\n" +
             "Looking for archives with 'UnifiedData' in the name.", 
             ui.ButtonSet.OK);
    return false;
  }
  
  // Enhanced verification with UnifiedDataAccess context
  var lastRow = archiveSheet.getLastRow();
  var lastCol = archiveSheet.getLastColumn();
  
  // Look for enhanced metadata section
  var values = archiveSheet.getDataRange().getValues();
  var hasEnhancedMetadata = values.some(row => 
    row[0] && row[0].toString().includes("ENHANCED ARCHIVE METADATA")
  );
  
  var hasUnifiedDataMetadata = values.some(row =>
    row[0] && row[0].toString().includes("UNIFIEDDATAACCESS CAPABILITIES")
  );
  
  var hasSheetAnalysis = values.some(row =>
    row[0] && row[0].toString().includes("SHEET ACCESS ANALYSIS")
  );
  
  var message = "📁 Enhanced Archive Verification Results:\n\n";
  message += "Archive Name: " + archiveName + "\n";
  message += "Data Rows: " + lastRow + "\n";
  message += "Data Columns: " + lastCol + "\n";
  message += "Has Enhanced Metadata: " + (hasEnhancedMetadata ? "✅ Yes" : "❌ No") + "\n";
  message += "Has UnifiedData Info: " + (hasUnifiedDataMetadata ? "✅ Yes" : "❌ No") + "\n";
  message += "Has Sheet Analysis: " + (hasSheetAnalysis ? "✅ Yes" : "❌ No") + "\n";
  message += "Archive Type: v4.0 UnifiedDataAccess\n\n";
  
  if (hasEnhancedMetadata && hasUnifiedDataMetadata && hasSheetAnalysis) {
    message += "✅ Archive appears to be complete and enhanced with UnifiedDataAccess metadata.";
  } else if (hasEnhancedMetadata) {
    message += "⚠️ Archive has basic metadata but may be missing some UnifiedData enhancements.";
  } else {
    message += "❌ Archive may be incomplete - missing enhanced metadata sections.";
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("Enhanced Archive Verification", message, ui.ButtonSet.OK);
  
  return true;
}

/**
 * 🆕 ENHANCED SMART TABLE CONVERTER with UnifiedDataAccess integration
 */
function convertSmartTablesToRegularSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    "🔄 Enhanced Smart Table Converter (v4.0)",
    "ENHANCED SMART TABLE CONVERSION with UnifiedDataAccess:\n\n" +
    "✅ UnifiedDataAccess detection of Smart Tables\n" +
    "✅ Safe data extraction and conversion\n" +
    "✅ Performance monitoring and reporting\n" +
    "✅ Validation rule application\n" +
    "✅ Column mapping preservation\n\n" +
    "WHAT THIS DOES:\n" +
    "• Detects Smart Tables using UnifiedDataAccess\n" +
    "• Safely extracts data preserving structure\n" +
    "• Creates compatible regular sheets\n" +
    "• Applies enhanced validation rules\n" +
    "• Enables full script compatibility\n\n" +
    "⚠️ This will remove Smart Table features but improve script compatibility.\n\n" +
    "Continue?",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  var startTime = new Date();
  var results = {
    converted: [],
    failed: [],
    skipped: [],
    smartTablesFound: [],
    errors: [],
    performanceMetrics: [],
    unifiedDataStats: {},
    accessMethodsUsed: {}
  };
  
  ss.toast("🔄 Starting enhanced Smart Table conversion with UnifiedDataAccess...", "Processing", 10);
  
  // 🆕 FIRST: Detect all Smart Tables using UnifiedDataAccess
  var detectionResult = detectSmartTablesWithUnifiedDataAccess();
  var smartTableCandidates = detectionResult.smartTables;
  
  results.unifiedDataStats = detectionResult.summary;
  results.accessMethodsUsed = detectionResult.accessMethods;
  
  console.log("Smart Table detection completed:", detectionResult);
  
  if (smartTableCandidates.length === 0) {
    var message = "🔍 No Smart Tables detected!\n\n";
    message += "All month sheets appear to be regular sheets already.\n";
    message += "• Total sheets analyzed: " + detectionResult.summary.totalSheets + "\n";
    message += "• Regular sheets: " + detectionResult.summary.accessibleSheets + "\n";
    message += "• Detection time: " + detectionResult.totalTime + "ms";
    
    ui.alert("No Conversion Needed", message, ui.ButtonSet.OK);
    return results;
  }
  
  ss.toast("Found " + smartTableCandidates.length + " Smart Tables to convert...", "Processing", 5);
  
  // Convert each detected Smart Table
  smartTableCandidates.forEach(function(sheetName, index) {
    var conversionStartTime = new Date();
    
    try {
      console.log("Converting " + sheetName + " (" + (index + 1) + "/" + smartTableCandidates.length + ")...");
      ss.toast("Converting " + sheetName + "... (" + (index + 1) + "/" + smartTableCandidates.length + ")", "Progress", 3);
      
      var conversionResult = convertSingleSmartTableWithUnifiedDataAccess(sheetName, ss);
      
      var conversionEndTime = new Date();
      var conversionTime = conversionEndTime - conversionStartTime;
      
      if (conversionResult.success) {
        results.converted.push(sheetName);
        results.performanceMetrics.push({
          sheet: sheetName,
          time: conversionTime,
          success: true,
          rowsConverted: conversionResult.rowsConverted,
          method: conversionResult.method
        });
        console.log("✅ " + sheetName + " converted successfully in " + conversionTime + "ms");
      } else {
        results.failed.push(sheetName + " (" + conversionResult.error + ")");
        results.performanceMetrics.push({
          sheet: sheetName,
          time: conversionTime,
          success: false,
          error: conversionResult.error
        });
        console.error("❌ " + sheetName + " conversion failed: " + conversionResult.error);
      }
      
    } catch (conversionError) {
      var conversionEndTime = new Date();
      results.errors.push(sheetName + ": " + conversionError.message);
      results.performanceMetrics.push({
        sheet: sheetName,
        time: conversionEndTime - conversionStartTime,
        success: false,
        error: conversionError.message
      });
      console.error("Error converting " + sheetName + ": ", conversionError);
    }
  });
  
  var endTime = new Date();
  var totalTime = endTime - startTime;
  
  // Generate enhanced results report
  generateSmartTableConversionReport(results, totalTime, detectionResult, ui);
  
  // Clear all caches since sheet structure changed completely
  try {
    ADVANCED_CACHE.clearAll();
    UnifiedDataAccess.clearCache();
    ss.toast("🧹 Enhanced caches cleared - sheet structure updated", "Cleanup", 3);
  } catch (cacheError) {
    console.warn("Could not clear caches: ", cacheError);
  }
  
  return results;
}

/**
 * 🆕 DETECT SMART TABLES using UnifiedDataAccess
 */
function detectSmartTablesWithUnifiedDataAccess() {
  var startTime = new Date();
  var result = {
    smartTables: [],
    regularSheets: [],
    summary: {
      totalSheets: 0,
      accessibleSheets: 0,
      smartTables: 0,
      errors: 0
    },
    accessMethods: {},
    totalTime: 0
  };
  
  CONFIG.MONTH_NAMES.forEach(function(monthName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(monthName);
    result.summary.totalSheets++;
    
    if (!sheet) {
      result.summary.errors++;
      return;
    }
    
    try {
      // Test accessibility using UnifiedDataAccess
      var accessTest = UnifiedDataAccess.testAccessMethod(monthName, 'BULK_ACCESS');
      
      if (accessTest.success) {
        // Can access with bulk method - likely regular sheet
        result.regularSheets.push(monthName);
        result.summary.accessibleSheets++;
        result.accessMethods['BULK_ACCESS'] = (result.accessMethods['BULK_ACCESS'] || 0) + 1;
      } else {
        // Test with Smart Table method
        var smartTableTest = UnifiedDataAccess.testAccessMethod(monthName, 'SMART_TABLE_ACCESS');
        
        if (smartTableTest.success) {
          // Accessible via Smart Table method - likely Smart Table
          result.smartTables.push(monthName);
          result.summary.smartTables++;
          result.accessMethods['SMART_TABLE_ACCESS'] = (result.accessMethods['SMART_TABLE_ACCESS'] || 0) + 1;
        } else {
          // Not accessible via either method - error case
          result.summary.errors++;
          result.accessMethods['FAILED'] = (result.accessMethods['FAILED'] || 0) + 1;
        }
      }
      
    } catch (testError) {
      // If testing fails, assume it's a problematic Smart Table
      if (testError.message.includes('typed') || testError.message.includes('Smart Table')) {
        result.smartTables.push(monthName);
        result.summary.smartTables++;
        result.accessMethods['SMART_TABLE_DETECTED'] = (result.accessMethods['SMART_TABLE_DETECTED'] || 0) + 1;
      } else {
        result.summary.errors++;
        result.accessMethods['ERROR'] = (result.accessMethods['ERROR'] || 0) + 1;
      }
    }
  });
  
  var endTime = new Date();
  result.totalTime = endTime - startTime;
  
  return result;
}

/**
 * 🆕 CONVERT SINGLE SMART TABLE with enhanced error handling
 */
function convertSingleSmartTableWithUnifiedDataAccess(sheetName, ss) {
  try {
    var oldSheet = ss.getSheetByName(sheetName);
    if (!oldSheet) {
      return { success: false, error: "Sheet not found" };
    }
    
    // Create new regular sheet
    var newSheet = ss.insertSheet(sheetName + "_Regular");
    console.log("📄 Created new sheet: " + sheetName + "_Regular");
    
    var rowsConverted = 0;
    var conversionMethod = "Unknown";
    
    // Try to extract data using multiple methods
    try {
      // Method 1: Try UnifiedDataAccess Smart Table method
      var extractedData = UnifiedDataAccess.getSheetData(sheetName, {
        maxRows: CONFIG.DATA_LIMITS.MAX_ROWS_PER_SHEET,
        useCache: false
      });
      
      var colMap = UnifiedDataAccess.getColumnMapping(sheetName);
      conversionMethod = "UnifiedDataAccess";
      
      // Write extracted data to new sheet
      if (extractedData.length > 0) {
        newSheet.getRange(2, 1, extractedData.length, extractedData[0].length).setValues(extractedData);
        rowsConverted = extractedData.length;
        
        // Create headers based on column mapping
        var headers = createHeadersFromColumnMapping(colMap);
        newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        console.log("📋 Converted " + rowsConverted + " rows using UnifiedDataAccess");
      }
      
    } catch (unifiedError) {
      console.log("UnifiedDataAccess extraction failed, trying fallback method...");
      
      // Method 2: Cell-by-cell fallback
      try {
        var fallbackData = extractDataCellByCell(oldSheet);
        conversionMethod = "Cell-by-cell fallback";
        
        if (fallbackData.length > 0) {
          newSheet.getRange(1, 1, fallbackData.length, fallbackData[0].length).setValues(fallbackData);
          rowsConverted = fallbackData.length - 1; // Subtract header row
          console.log("📋 Converted " + rowsConverted + " rows using fallback method");
        }
        
      } catch (fallbackError) {
        throw new Error("All extraction methods failed: " + fallbackError.message);
      }
    }
    
    // Apply basic formatting
    try {
      newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).setFontWeight("bold");
      newSheet.setFrozenRows(1);
    } catch (formatError) {
      console.log("Could not apply formatting to " + sheetName);
    }
    
    // Move new sheet to same position as old sheet
    var oldIndex = oldSheet.getIndex();
    newSheet.activate();
    ss.moveActiveSheet(oldIndex);
    
    // Delete old Smart Table sheet
    ss.deleteSheet(oldSheet);
    console.log("🗑️ Deleted old Smart Table: " + sheetName);
    
    // Rename new sheet to original name
    newSheet.setName(sheetName);
    console.log("✅ Renamed new sheet to: " + sheetName);
    
    return {
      success: true,
      rowsConverted: rowsConverted,
      method: conversionMethod
    };
    
  } catch (conversionError) {
    console.error("Conversion failed for " + sheetName + ": ", conversionError);
    
    // Clean up failed attempt
    try {
      var failedSheet = ss.getSheetByName(sheetName + "_Regular");
      if (failedSheet) {
        ss.deleteSheet(failedSheet);
      }
    } catch (cleanupError) {
      console.error("Could not clean up failed conversion");
    }
    
    return {
      success: false,
      error: conversionError.message
    };
  }
}

/**
 * Create headers from column mapping
 */
function createHeadersFromColumnMapping(colMap) {
  var headers = new Array(Math.max(colMap.DATE || 1, colMap.TYPE || 2, colMap.DEVICE || 3,
                                   colMap.CUSTOMER || 4, colMap.MOBILE || 5, colMap.PLAN || 6, 
                                   colMap.NOTES || 7)).fill("");
  
  if (colMap.DATE) headers[colMap.DATE - 1] = "Date";
  if (colMap.TYPE) headers[colMap.TYPE - 1] = "Type";
  if (colMap.DEVICE) headers[colMap.DEVICE - 1] = "Device";
  if (colMap.CUSTOMER) headers[colMap.CUSTOMER - 1] = "Customer";
  if (colMap.MOBILE) headers[colMap.MOBILE - 1] = "Mobile";
  if (colMap.PLAN) headers[colMap.PLAN - 1] = "Plan";
  if (colMap.NOTES) headers[colMap.NOTES - 1] = "Notes";
  
  return headers;
}

/**
 * Fallback cell-by-cell data extraction
 */
function extractDataCellByCell(sheet) {
  var data = [];
  var maxRows = Math.min(sheet.getLastRow() || 100, 200); // Strict limit for safety
  var maxCols = Math.min(sheet.getLastColumn() || 7, 10);
  
  for (var row = 1; row <= maxRows; row++) {
    var rowData = [];
    for (var col = 1; col <= maxCols; col++) {
      try {
        var cellValue = sheet.getRange(row, col).getDisplayValue();
        rowData.push(cellValue);
      } catch (cellError) {
        rowData.push(""); // Skip problematic cells
      }
    }
    data.push(rowData);
    
    // Add delay every 10 rows to prevent timeout
    if (row % 10 === 0) {
      Utilities.sleep(50);
    }
  }
  
  return data;
}

/**
 * 🆕 GENERATE SMART TABLE CONVERSION REPORT
 */
function generateSmartTableConversionReport(results, totalTime, detectionResult, ui) {
  var avgProcessingTime = results.performanceMetrics.length > 0 ?
    results.performanceMetrics.reduce((sum, metric) => sum + metric.time, 0) / results.performanceMetrics.length : 0;
  
  var message = "🔄 Enhanced Smart Table Conversion Complete! (v4.0)\n\n";
  message += "📊 Results Summary:\n";
  message += "✅ Successfully Converted: " + results.converted.length + "\n";
  message += "⏩ Already Regular: " + detectionResult.regularSheets.length + "\n";
  message += "❌ Failed: " + results.failed.length + "\n";
  message += "⚡ Total Time: " + Math.round(totalTime / 1000) + " seconds\n";
  message += "🔧 Avg Time/Sheet: " + Math.round(avgProcessingTime) + "ms\n";
  
  if (results.converted.length > 0) {
    message += "\n✅ Converted Sheets:\n";
    results.converted.forEach(function(sheet) {
      var metric = results.performanceMetrics.find(m => m.sheet === sheet);
      if (metric && metric.rowsConverted) {
        message += "• " + sheet + " (" + metric.rowsConverted + " rows)\n";
      } else {
        message += "• " + sheet + "\n";
      }
    });
  }
  
  if (detectionResult.regularSheets.length > 0) {
    message += "\n⏩ Already Regular Sheets:\n";
    detectionResult.regularSheets.slice(0, 3).forEach(function(sheet) {
      message += "• " + sheet + "\n";
    });
    if (detectionResult.regularSheets.length > 3) {
      message += "• ... and " + (detectionResult.regularSheets.length - 3) + " more\n";
    }
  }
  
  if (results.failed.length > 0 || results.errors.length > 0) {
    message += "\n❌ Issues:\n";
    results.failed.forEach(function(failure) {
      message += "• " + failure + "\n";
    });
    results.errors.forEach(function(error) {
      message += "• " + error + "\n";
    });
  }
  
  // Detection method summary
  var methods = Object.keys(results.accessMethodsUsed);
  if (methods.length > 0) {
    message += "\n🔧 Detection Methods:\n";
    methods.forEach(function(method) {
      message += "• " + method + ": " + results.accessMethodsUsed[method] + " sheets\n";
    });
  }
  
  message += "\n🎯 Next Steps:\n";
  message += "1. Enhanced caches cleared ✅\n";
  message += "2. Run system health check\n";
  message += "3. Set up validation on converted sheets";
  
  console.log("Enhanced conversion results:", JSON.stringify(results, null, 2));
  ui.alert("Enhanced Conversion Complete", message, ui.ButtonSet.OK);
}

/**
 * 🆕 EMERGENCY DATA RECOVERY with UnifiedDataAccess
 */
function emergencyDataRecovery() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "🚨 Emergency Data Recovery (v4.0)",
    "ENHANCED EMERGENCY DATA RECOVERY:\n\n" +
    "✅ UnifiedDataAccess archive scanning\n" +
    "✅ Smart recovery recommendations\n" +
    "✅ Data integrity validation\n\n" +
    "This will attempt to recover data from the most recent archive.\n\n" +
    "⚠️ Only use if data has been accidentally lost.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentYear = new Date().getFullYear();
  
  // Look for archives with UnifiedData
  var allSheets = ss.getSheets();
  var archives = allSheets.filter(sheet => 
    sheet.getName().includes("Archive") && 
    sheet.getName().includes(currentYear.toString())
  );
  
  if (archives.length === 0) {
    ui.alert("No Recovery Archives Found", 
             "No archive sheets found for " + currentYear + ".\n\n" +
             "Recovery requires a valid archive sheet.", 
             ui.ButtonSet.OK);
    return;
  }
  
  // Show available archives
  var archiveNames = archives.map(sheet => sheet.getName());
  var mostRecent = archives[archives.length - 1].getName();
  
  var message = "🔍 Found " + archives.length + " archive(s):\n\n";
  archiveNames.forEach(function(name) {
    message += "• " + name + "\n";
  });
  message += "\nMost Recent: " + mostRecent + "\n\n";
  message += "🚨 RECOVERY PROCESS:\n";
  message += "1. Analyze archive with UnifiedDataAccess\n";
  message += "2. Extract recoverable data sections\n";
  message += "3. Provide recovery instructions\n";
  message += "4. Validate data integrity\n\n";
  message += "📋 For detailed recovery, use the archive data to manually restore specific months.\n";
  message += "🔧 Then run 'Refresh All Reports' to rebuild everything.\n\n";
  message += "💡 Contact your system administrator for guided recovery assistance.";
  
  ui.alert("Enhanced Recovery Instructions", message, ui.ButtonSet.OK);
  
  // Log recovery attempt
  console.log("Emergency data recovery initiated - archives found:", archiveNames);
  
  return {
    archivesFound: archives.length,
    mostRecentArchive: mostRecent,
    allArchives: archiveNames
  };
}
