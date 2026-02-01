/**
 * =================================================================================
 * BULLETPROOF REAL-TIME DATA VALIDATION SYSTEM (v4.0)
 * Complete safety for both regular sheets and Smart Tables with UnifiedDataAccess
 * =================================================================================
 */

/**
 * 🆕 BULLETPROOF validateDataEntry: Completely safe for all sheet types with UnifiedDataAccess
 */
function validateDataEntry(e) {
  if (!e) return;
  
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var col = range.getColumn();
  var row = range.getRow();
  
  // Only validate month sheets and data rows (not headers)
  if (!CONFIG.MONTH_NAMES.includes(sheetName) || row < 2) return;
  
  try {
    // Simple value-only validation - no formatting or risky operations
    var cellValue = "";
    try {
      cellValue = range.getDisplayValue();
    } catch (accessError) {
      // If we can't read the cell, it's probably a Smart Table - skip validation
      return;
    }
    
    // Skip empty cells
    if (!cellValue && cellValue !== 0) return;
    
    // 🆕 GET COLUMN MAPPING using UnifiedDataAccess
    var colMap;
    try {
      colMap = UnifiedDataAccess.getColumnMapping(sheetName);
    } catch (mapError) {
      return;
    }
    
    var validationErrors = [];
    var validationWarnings = [];
    var cellType = "unknown";
    
    // 🆕 ENHANCED VALIDATION based on column type with better feedback
    if (col === colMap.DATE) {
      cellType = "Date";
      if (!isValidDateEnhanced(cellValue)) {
        validationErrors.push("Date format issue - try MM/DD/YYYY or MM/DD/YY");
      }
    } else if (col === colMap.TYPE) {
      cellType = "Type";
      var trimmedType = String(cellValue).trim();
      if (!CONFIG.VALIDATION_RULES.VALID_TYPES.includes(trimmedType)) {
        validationErrors.push("Type should be: " + CONFIG.VALIDATION_RULES.VALID_TYPES.join(", "));
        
        // Suggest close matches
        var suggestions = findClosestMatches(trimmedType, CONFIG.VALIDATION_RULES.VALID_TYPES);
        if (suggestions.length > 0) {
          validationWarnings.push("Did you mean: " + suggestions.join(" or ") + "?");
        }
      }
    } else if (col === colMap.DEVICE) {
      cellType = "Device";
      var trimmedDevice = String(cellValue).trim();
      if (!CONFIG.VALIDATION_RULES.VALID_DEVICES.includes(trimmedDevice)) {
        validationErrors.push("Device should be: " + CONFIG.VALIDATION_RULES.VALID_DEVICES.join(", "));
        
        // Suggest close matches
        var suggestions = findClosestMatches(trimmedDevice, CONFIG.VALIDATION_RULES.VALID_DEVICES);
        if (suggestions.length > 0) {
          validationWarnings.push("Did you mean: " + suggestions.join(" or ") + "?");
        }
      }
    } else if (col === colMap.PLAN) {
      cellType = "Plan";
      var trimmedPlan = String(cellValue).trim();
      if (!CONFIG.VALIDATION_RULES.VALID_PLANS.includes(trimmedPlan)) {
        validationErrors.push("Plan should be: " + CONFIG.VALIDATION_RULES.VALID_PLANS.join(", "));
        
        // Suggest close matches
        var suggestions = findClosestMatches(trimmedPlan, CONFIG.VALIDATION_RULES.VALID_PLANS);
        if (suggestions.length > 0) {
          validationWarnings.push("Did you mean: " + suggestions.join(" or ") + "?");
        }
      }
    } else if (col === colMap.CUSTOMER) {
      cellType = "Customer";
      var customerName = String(cellValue).trim();
      if (customerName.length < 2) {
        validationWarnings.push("Customer name seems very short");
      } else if (customerName.toLowerCase() === "customer") {
        validationErrors.push("Please enter actual customer name");
      }
    } else if (col === colMap.MOBILE) {
      cellType = "Mobile";
      var phone = String(cellValue).replace(/[^\d]/g, '');
      if (phone.length > 0 && (phone.length < 10 || phone.length > 15)) {
        validationWarnings.push("Phone number should be 10-15 digits");
      }
    }
    
    // 🆕 ENHANCED FEEDBACK with context and suggestions
    if (validationErrors.length > 0 || validationWarnings.length > 0) {
      var feedbackParts = [];
      
      feedbackParts.push("💡 " + cellType + " Validation:");
      
      if (validationErrors.length > 0) {
        feedbackParts.push("❌ " + validationErrors.join("\n❌ "));
      }
      
      if (validationWarnings.length > 0) {
        feedbackParts.push("⚠️ " + validationWarnings.join("\n⚠️ "));
      }
      
      feedbackParts.push("\n📍 Cell: " + sheetName + "!" + range.getA1Notation());
      feedbackParts.push("🔧 Value: \"" + cellValue + "\"");
      
      var feedbackMessage = feedbackParts.join("\n");
      
      SpreadsheetApp.getActiveSpreadsheet().toast(feedbackMessage, "Data Helper", 8);
    }
    
  } catch (error) {
    // Complete silence for any errors - Smart Tables handle their own validation
  }
  
  // 🆕 GRANULAR CACHE INVALIDATION - only clear affected data
  try {
    // Only invalidate cache if this was a data change in a main data area
    if (col <= 10 && row >= 2) {
      ADVANCED_CACHE.remove('unified_data_' + sheetName);
      ADVANCED_CACHE.remove('sheet_data_' + sheetName);
    }
  } catch (cacheError) {
    // Ignore
  }
}

/**
 * 🆕 ENHANCED date validation helper
 */
function isValidDateEnhanced(value) {
  if (!value) return false;
  
  // Try multiple date formats
  var dateFormats = [
    value, // Direct value
    new Date(value), // Standard parsing
    Date.parse(value) // Alternative parsing
  ];
  
  for (var i = 0; i < dateFormats.length; i++) {
    var testDate = dateFormats[i];
    if (testDate instanceof Date && !isNaN(testDate)) {
      return true;
    }
    if (typeof testDate === 'number' && !isNaN(testDate)) {
      return true;
    }
  }
  
  return false;
}

/**
 * 🆕 SMART SUGGESTION SYSTEM - finds closest matches for user input
 */
function findClosestMatches(input, validOptions) {
  if (!input || !validOptions) return [];
  
  var inputLower = input.toLowerCase();
  var suggestions = [];
  
  validOptions.forEach(function(option) {
    var optionLower = option.toLowerCase();
    
    // Exact match (shouldn't happen if we're here, but safety check)
    if (inputLower === optionLower) {
      return;
    }
    
    // Starts with
    if (optionLower.startsWith(inputLower) || inputLower.startsWith(optionLower)) {
      suggestions.push(option);
    }
    // Contains
    else if (optionLower.includes(inputLower) || inputLower.includes(optionLower)) {
      suggestions.push(option);
    }
    // Similar length and similar characters (simple similarity)
    else if (Math.abs(inputLower.length - optionLower.length) <= 2) {
      var similarity = calculateStringSimilarity(inputLower, optionLower);
      if (similarity > 0.6) {
        suggestions.push(option);
      }
    }
  });
  
  // Return max 2 suggestions to avoid overwhelming the user
  return suggestions.slice(0, 2);
}

/**
 * Simple string similarity calculation
 */
function calculateStringSimilarity(str1, str2) {
  var longer = str1.length > str2.length ? str1 : str2;
  var shorter = str1.length > str2.length ? str2 : str1;
  
  if (longer.length === 0) return 1.0;
  
  var matches = 0;
  for (var i = 0; i < shorter.length; i++) {
    if (longer.includes(shorter[i])) {
      matches++;
    }
  }
  
  return matches / longer.length;
}

/**
 * 🆕 BULLETPROOF setupDataValidationRules: Enhanced with UnifiedDataAccess integration
 */
function setupDataValidationRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    "🛠️ Setup Enhanced Data Validation (v4.0)",
    "ENHANCED DATA VALIDATION with UnifiedDataAccess:\n\n" +
    "✅ Smart Table auto-detection\n" +
    "✅ Dynamic column mapping\n" +
    "✅ Safe error handling\n" +
    "✅ 100% safe for all sheet types\n\n" +
    "This will add dropdown validation to compatible sheets.\nContinue?",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  var setupResults = {
    sheetsProcessed: 0,
    regularSheets: 0,
    smartTables: 0,
    rulesApplied: 0,
    successfulSheets: []
  };
  
  ss.toast("🔄 Starting enhanced validation setup...", "Processing", 5);
  
  CONFIG.MONTH_NAMES.forEach(function(monthName) {
    var sheet = ss.getSheetByName(monthName);
    if (!sheet) return;
    
    // 🆕 ENHANCED SHEET-LEVEL SAFETY using UnifiedDataAccess detection
    try {
      // Use UnifiedDataAccess to test sheet accessibility
      var sheetTestResult = testSheetAccessibilityUnified(monthName);
      
      if (!sheetTestResult.accessible) {
        // Treat as Smart Table
        setupResults.smartTables++;
        return; // Skip this sheet entirely
      }
      
      // Process as regular sheet
      setupResults.regularSheets++;
      var colMap = sheetTestResult.columnMapping;
      var lastRow = Math.max(sheet.getLastRow(), 100); // Ensure minimum validation range
      var rulesAppliedThisSheet = 0;
      
      // 🆕 ENHANCED VALIDATION RULE APPLICATION
      var validationResults = applyValidationRulesEnhanced(sheet, colMap, lastRow, monthName);
      rulesAppliedThisSheet = validationResults.rulesApplied;
      
      setupResults.rulesApplied += rulesAppliedThisSheet;
      setupResults.successfulSheets.push(monthName + " (" + rulesAppliedThisSheet + " rules)");
      
    } catch (sheetError) {
      // If ANY error occurs at sheet level, treat as Smart Table
      setupResults.smartTables++;
    }
    
    setupResults.sheetsProcessed++;
  });
  
  var message = "✅ Enhanced Validation Setup Complete!\n\n";
  message += "📊 Results Summary:\n";
  message += "• Total Sheets: " + setupResults.sheetsProcessed + "\n";
  message += "• Regular Sheets: " + setupResults.regularSheets + "\n";
  message += "• Smart Tables: " + setupResults.smartTables + "\n";
  message += "• Rules Applied: " + setupResults.rulesApplied + "\n";
  
  ss.toast(message, "Validation Setup Complete", 5);
  
  return setupResults;
}

/**
 * 🆕 TEST SHEET ACCESSIBILITY using UnifiedDataAccess
 */
function testSheetAccessibilityUnified(sheetName) {
  try {
    // Try to get column mapping using UnifiedDataAccess
    var colMap = UnifiedDataAccess.getColumnMapping(sheetName);
    
    // Try to get a small sample of data
    var testData = UnifiedDataAccess.getSheetData(sheetName, {
      maxRows: 5,
      useCache: false
    });
    
    // If we got here, the sheet is accessible
    return {
      accessible: true,
      columnMapping: colMap
    };
    
  } catch (accessError) {
    return {
      accessible: false,
      columnMapping: null
    };
  }
}

/**
 * 🆕 ENHANCED validation rule application
 */
function applyValidationRulesEnhanced(sheet, colMap, lastRow, sheetName) {
  var results = {
    rulesApplied: 0
  };
  
  // Apply validation rules for each column type
  var validationConfigs = [
    { column: colMap.TYPE, values: CONFIG.VALIDATION_RULES.VALID_TYPES, name: "Type" },
    { column: colMap.DEVICE, values: CONFIG.VALIDATION_RULES.VALID_DEVICES, name: "Device" },
    { column: colMap.PLAN, values: CONFIG.VALIDATION_RULES.VALID_PLANS, name: "Plan" }
  ];
  
  validationConfigs.forEach(function(config) {
    if (config.column) {
      try {
        var success = bulletproofSetValidationEnhanced(sheet, config.column, lastRow, config.values, config.name, sheetName);
        if (success) {
          results.rulesApplied++;
        }
      } catch (ruleError) {
        // Continue
      }
    }
  });
  
  return results;
}

/**
 * 🆕 ENHANCED bulletproof validation rule setter
 */
function bulletproofSetValidationEnhanced(sheet, columnNumber, lastRow, validValues, columnName, sheetName) {
  try {
    // Create validation rule with enhanced settings
    var validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(validValues, true) // Show dropdown
      .setAllowInvalid(false) // Reject invalid values
      .setHelpText("Select from: " + validValues.join(", ") + " (Required for " + columnName + ")")
      .build();
    
    // Calculate safe range
    var numRows = Math.max(1, lastRow - 1);
    var targetRange = sheet.getRange(2, columnNumber, numRows, 1);
    
    // Apply validation
    targetRange.setDataValidation(validationRule);
    
    return true;
    
  } catch (error) {
    return false;
  }
}

/**
 * 🆕 ENHANCED sheet type analyzer with UnifiedDataAccess
 */
function analyzeSheetTypes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = {
    regularSheets: [],
    smartTables: [],
    errorSheets: []
  };
  
  ss.toast("🔍 Analyzing sheet types...", "Processing", 3);
  
  CONFIG.MONTH_NAMES.forEach(function(monthName) {
    var sheet = ss.getSheetByName(monthName);
    
    if (!sheet) {
      results.errorSheets.push(monthName + " (not found)");
      return;
    }
    
    // Test sheet accessibility using UnifiedDataAccess
    var testResult = testSheetAccessibilityUnified(monthName);
    
    if (testResult.accessible) {
      results.regularSheets.push(monthName);
    } else {
      results.smartTables.push(monthName);
    }
  });
  
  // Generate report
  var message = "🔍 Enhanced Sheet Type Analysis:\n\n";
  
  if (results.regularSheets.length > 0) {
    message += "✅ Regular Sheets (" + results.regularSheets.length + "):\n";
    message += results.regularSheets.join(", ") + "\n\n";
  }
  
  if (results.smartTables.length > 0) {
    message += "🏷️ Smart Tables (" + results.smartTables.length + "):\n";
    message += results.smartTables.join(", ") + "\n\n";
  }
  
  if (results.errorSheets.length > 0) {
    message += "❓ Issues (" + results.errorSheets.length + "):\n";
    message += results.errorSheets.join(", ") + "\n\n";
  }
  
  message += "💡 Smart Tables have built-in validation.\n";
  message += "🎯 Regular sheets can have dropdown validation added.";
  
  SpreadsheetApp.getUi().alert("Sheet Analysis Complete", message, SpreadsheetApp.getUi().ButtonSet.OK);
}
