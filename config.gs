/**
 * =================================================================================
 * ENHANCED CORE CONFIGURATION AND UTILITIES (v4.0)
 * Foundation file with UnifiedDataAccess integration
 * =================================================================================
 */

var CONFIG = {
  // Core configuration
  MONTH_NAMES: ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"],
  WEEKEND_DAYS: [4, 5], // Thursday=4, Friday=5
  MONITORED_CELLS: ["I2", "I5", "I7", "I8", "I9"],
  
  // FORCED COLUMN MAPPINGS - Override auto-detection for specific sheets
  FORCED_COLUMN_MAPPINGS: {
    "Jan": {
      DATE: 1, DATE_INDEX: 0,
      TYPE: 2, TYPE_INDEX: 1,
      DEVICE: 3, DEVICE_INDEX: 2,
      CUSTOMER: 4, CUSTOMER_INDEX: 3,
      MOBILE: 5, MOBILE_INDEX: 4,
      PLAN: 6, PLAN_INDEX: 5,
      NOTES: 7, NOTES_INDEX: 6,
      _detectionMethod: 'forced_override'
    }
    // Add other sheets here if they have detection issues
  },
  
  // 🆕 ENHANCED DYNAMIC COLUMN MAPPING with UnifiedDataAccess integration
  getColumnMap: function(sheet) {
    var sheetName = sheet.getName();
    
    // CHECK FOR FORCED MAPPING FIRST
    if (this.FORCED_COLUMN_MAPPINGS && this.FORCED_COLUMN_MAPPINGS[sheetName]) {
      return this.FORCED_COLUMN_MAPPINGS[sheetName];
    }
    
    var cacheKey = 'enhanced_column_map_' + sheetName;
    var cached = ADVANCED_CACHE.get(cacheKey);
    
    if (cached && cached.timestamp && (new Date() - new Date(cached.timestamp)) < 3600000) { // 1 hour TTL
      return cached.columnMap;
    }
    
    try {
      // Enhanced column detection
      var columnMap = this.detectColumnsWithUnifiedDataAccessIntegration(sheet);
      
      // Enhanced cache with metadata
      ADVANCED_CACHE.set(cacheKey, {
        columnMap: columnMap,
        timestamp: new Date(),
        version: "4.0"
      }, 3600);
      
      return columnMap;
      
    } catch (error) {
      // Return safe default mapping with Smart Table detection
      if (error.message.includes('typed') || error.message.includes('Smart Table')) {
        return this.getSmartTableColumnMapping(sheetName);
      } else {
        return this.getDefaultColumnMapping();
      }
    }
  },
  
  // 🆕 ENHANCED COLUMN DETECTION with UnifiedDataAccess integration
  detectColumnsWithUnifiedDataAccessIntegration: function(sheet) {
    var sheetName = sheet.getName();
    
    // CHECK FORCED MAPPINGS AGAIN (in case called directly)
    if (this.FORCED_COLUMN_MAPPINGS && this.FORCED_COLUMN_MAPPINGS[sheetName]) {
      return this.FORCED_COLUMN_MAPPINGS[sheetName];
    }
    
    // Method 1: Try UnifiedDataAccess header detection (safest and most reliable)
    try {
      var detectionResult = this.detectColumnsViaUnifiedDataAccess(sheetName);
      if (detectionResult.confidence > 0.8) {
        return this.applyDefaultMappings(detectionResult.columnMap);
      }
    } catch (unifiedError) {
      // Silent fail, proceed to fallback
    }
    
    // Method 2: Try direct header reading (for regular sheets)
    try {
      var maxCols = Math.min(sheet.getLastColumn() || 10, 15);
      var headers = sheet.getRange(1, 1, 1, maxCols).getDisplayValues()[0];
      var headerColumnMap = this.mapHeadersToColumns(headers);
      
      if (this.validateColumnMapping(headerColumnMap)) {
        return this.applyDefaultMappings(headerColumnMap);
      }
    } catch (directError) {
      // Silent fail
    }
    
    // Method 3: Smart Table specific detection
    if (this.isLikelySmartTable(sheet, sheetName)) {
      return this.getSmartTableColumnMapping(sheetName);
    }
    
    // Final fallback
    return this.getDefaultColumnMapping();
  },
  
  // 🆕 DETECT COLUMNS via UnifiedDataAccess
  detectColumnsViaUnifiedDataAccess: function(sheetName) {
    try {
      // Get sample data to analyze headers and patterns
      var sampleData = UnifiedDataAccess.getSheetData(sheetName, {
        maxRows: 5,
        useCache: false,
        includeHeaders: true
      });
      
      if (sampleData.length === 0) {
        throw new Error("No sample data available");
      }
      
      // Analyze headers (first row) and data patterns
      var headers = sampleData[0] || [];
      var dataRows = sampleData.slice(1);
      
      var headerColumnMap = this.mapHeadersToColumns(headers);
      var patternColumnMap = this.analyzeDataPatterns(dataRows);
      
      // Combine header and pattern detection
      var combinedColumnMap = this.combineColumnMappings(headerColumnMap, patternColumnMap);
      
      // Calculate confidence based on matches
      var confidence = this.calculateMappingConfidence(combinedColumnMap, headers, dataRows);
      
      return {
        columnMap: combinedColumnMap,
        confidence: confidence
      };
      
    } catch (unifiedError) {
      throw new Error("UnifiedDataAccess detection failed: " + unifiedError.message);
    }
  },
  
  // 🆕 ANALYZE DATA PATTERNS in sample data
  analyzeDataPatterns: function(dataRows) {
    var columnMap = {};
    
    if (!dataRows || dataRows.length === 0) return columnMap;
    
    var maxCols = Math.min(dataRows[0].length, 10);
    
    for (var col = 0; col < maxCols; col++) {
      var colData = dataRows.map(row => row[col]).filter(val => val && val !== "");
      
      if (colData.length === 0) continue;
      
      var columnIndex = col + 1; // 1-indexed for getRange
      
      // Enhanced pattern detection
      if (this.seemsLikeDateEnhanced(colData)) {
        columnMap.DATE = columnIndex;
        columnMap.DATE_INDEX = col;
      }
      else if (this.seemsLikeTypeEnhanced(colData)) {
        columnMap.TYPE = columnIndex;
        columnMap.TYPE_INDEX = col;
      }
      else if (this.seemsLikeDeviceEnhanced(colData)) {
        columnMap.DEVICE = columnIndex;
        columnMap.DEVICE_INDEX = col;
      }
      else if (this.seemsLikeNameEnhanced(colData)) {
        columnMap.CUSTOMER = columnIndex;
        columnMap.CUSTOMER_INDEX = col;
      }
      else if (this.seemsLikePhoneEnhanced(colData)) {
        columnMap.MOBILE = columnIndex;
        columnMap.MOBILE_INDEX = col;
      }
      else if (this.seemsLikePlanEnhanced(colData)) {
        columnMap.PLAN = columnIndex;
        columnMap.PLAN_INDEX = col;
      }
      else if (this.seemsLikeNotesEnhanced(colData)) {
        columnMap.NOTES = columnIndex;
        columnMap.NOTES_INDEX = col;
      }
    }
    
    return columnMap;
  },
  
  // 🆕 ENHANCED PATTERN HELPERS with better detection
  seemsLikeDateEnhanced: function(data) {
    var dateCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      if (this.isValidDateValue(data[i])) dateCount++;
    }
    return (dateCount / Math.min(data.length, 5)) > 0.6;
  },
  
  seemsLikeTypeEnhanced: function(data) {
    var typePatterns = this.VALIDATION_RULES.VALID_TYPES;
    var matchCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      var val = String(data[i]).trim();
      if (typePatterns.some(pattern => val.toLowerCase() === pattern.toLowerCase())) matchCount++;
    }
    return (matchCount / Math.min(data.length, 5)) > 0.5;
  },
  
  seemsLikeDeviceEnhanced: function(data) {
    var devicePatterns = this.VALIDATION_RULES.VALID_DEVICES;
    var matchCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      var val = String(data[i]).trim();
      if (devicePatterns.some(pattern => val.toLowerCase() === pattern.toLowerCase())) matchCount++;
    }
    return (matchCount / Math.min(data.length, 5)) > 0.4;
  },
  
  seemsLikeNameEnhanced: function(data) {
    var nameCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      var str = String(data[i]).trim();
      if (str.length > 3 && str.includes(' ') && !/^\d+$/.test(str) && !str.includes('@')) nameCount++;
    }
    return (nameCount / Math.min(data.length, 5)) > 0.6;
  },
  
  seemsLikePhoneEnhanced: function(data) {
    var phoneCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      var str = String(data[i]).replace(/[^\d]/g, '');
      if (str.length >= 10 && str.length <= 15) phoneCount++;
    }
    return (phoneCount / Math.min(data.length, 5)) > 0.5;
  },
  
  seemsLikePlanEnhanced: function(data) {
    var planPatterns = this.VALIDATION_RULES.VALID_PLANS;
    var matchCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      var val = String(data[i]).trim();
      if (planPatterns.some(pattern => val.toLowerCase() === pattern.toLowerCase())) matchCount++;
    }
    return (matchCount / Math.min(data.length, 5)) > 0.4;
  },
  
  seemsLikeNotesEnhanced: function(data) {
    var notesCount = 0;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      var str = String(data[i]).trim();
      if (str.length > 10 && (str.includes(' ') || str.includes(',') || str.includes('.'))) notesCount++;
    }
    return (notesCount / Math.min(data.length, 5)) > 0.3;
  },
  
  // 🆕 ENHANCED DATE VALIDATION
  isValidDateValue: function(value) {
    if (!value) return false;
    try {
      var testDate1 = new Date(value);
      if (testDate1 instanceof Date && !isNaN(testDate1) && testDate1.getFullYear() > 1900) return true;
      var testDate2 = Date.parse(value);
      if (!isNaN(testDate2)) {
        var parsedDate = new Date(testDate2);
        if (parsedDate.getFullYear() > 1900 && parsedDate.getFullYear() < 2100) return true;
      }
      return false;
    } catch (error) {
      return false;
    }
  },
  
  // 🆕 COMBINE COLUMN MAPPINGS from different detection methods
  combineColumnMappings: function(headerMap, patternMap) {
    var combined = Object.assign({}, headerMap);
    Object.keys(patternMap).forEach(function(key) {
      if (!combined[key]) combined[key] = patternMap[key];
    });
    return combined;
  },
  
  // 🆕 CALCULATE MAPPING CONFIDENCE
  calculateMappingConfidence: function(columnMap, headers, dataRows) {
    var totalPossibleColumns = 7; 
    var mappedColumns = Object.keys(columnMap).filter(key => !key.includes('_INDEX')).length;
    var essentialColumns = ['DATE', 'TYPE', 'CUSTOMER'].filter(col => columnMap[col]);
    
    var baseConfidence = mappedColumns / totalPossibleColumns;
    var essentialBonus = essentialColumns.length / 3 * 0.3;
    
    return Math.min(baseConfidence + essentialBonus, 1.0);
  },
  
  // 🆕 VALIDATE COLUMN MAPPING quality
  validateColumnMapping: function(columnMap) {
    var essentialColumns = ['DATE', 'TYPE', 'CUSTOMER'];
    var foundEssential = essentialColumns.filter(col => columnMap[col]).length;
    return foundEssential >= 2; // At least 2 essential columns found
  },
  
  // 🆕 DETECT if sheet is likely a Smart Table
  isLikelySmartTable: function(sheet, sheetName) {
    try {
      var testRange = sheet.getRange("A1:B2");
      testRange.getValues();
      return false; 
    } catch (error) {
      return true;
    }
  },
  
  // 🆕 SMART TABLE COLUMN MAPPING
  getSmartTableColumnMapping: function(sheetName) {
    return {
      DATE: 1, DATE_INDEX: 0,
      TYPE: 2, TYPE_INDEX: 1,
      DEVICE: 3, DEVICE_INDEX: 2,
      CUSTOMER: 4, CUSTOMER_INDEX: 3,
      MOBILE: 5, MOBILE_INDEX: 4,
      PLAN: 6, PLAN_INDEX: 5,
      NOTES: 7, NOTES_INDEX: 6,
      _isSmartTable: true,
      _detectionMethod: 'smart_table_default'
    };
  },
  
  // Original functions enhanced
  mapHeadersToColumns: function(headers) {
    var columnMap = {};
    var columnPatterns = {
      DATE: ['date', 'entry date', 'sale date', 'day'],
      TYPE: ['type', 'sale type', 'category'],
      DEVICE: ['device', 'brand', 'model'],
      CUSTOMER: ['customer', 'name', 'client'],
      MOBILE: ['mobile', 'phone', 'number'],
      PLAN: ['plan', 'rate plan', 'package'],
      NOTES: ['notes', 'comments', 'description']
    };
    
    for (var key in columnPatterns) {
      var patterns = columnPatterns[key];
      for (var i = 0; i < headers.length; i++) {
        var header = String(headers[i] || '').toLowerCase().trim();
        if (header && patterns.some(pattern => header.includes(pattern))) {
          columnMap[key] = i + 1; 
          columnMap[key + '_INDEX'] = i; 
          break;
        }
      }
    }
    return columnMap;
  },
  
  applyDefaultMappings: function(columnMap) {
    if (!columnMap.DATE) { columnMap.DATE = 1; columnMap.DATE_INDEX = 0; }
    if (!columnMap.TYPE) { columnMap.TYPE = 2; columnMap.TYPE_INDEX = 1; }
    if (!columnMap.DEVICE) { columnMap.DEVICE = 3; columnMap.DEVICE_INDEX = 2; }
    if (!columnMap.CUSTOMER) { columnMap.CUSTOMER = 4; columnMap.CUSTOMER_INDEX = 3; }
    if (!columnMap.MOBILE) { columnMap.MOBILE = 5; columnMap.MOBILE_INDEX = 4; }
    if (!columnMap.PLAN) { columnMap.PLAN = 6; columnMap.PLAN_INDEX = 5; }
    if (!columnMap.NOTES) { columnMap.NOTES = 7; columnMap.NOTES_INDEX = 6; }
    return columnMap;
  },
  
  getDefaultColumnMapping: function() {
    return {
      DATE: 1, DATE_INDEX: 0,
      TYPE: 2, TYPE_INDEX: 1,
      DEVICE: 3, DEVICE_INDEX: 2,
      CUSTOMER: 4, CUSTOMER_INDEX: 3,
      MOBILE: 5, MOBILE_INDEX: 4,
      PLAN: 6, PLAN_INDEX: 5,
      NOTES: 7, NOTES_INDEX: 6,
      _detectionMethod: 'default_fallback'
    };
  },
  
  // Enhanced validation rules
  VALIDATION_RULES: {
    VALID_TYPES: ["PPVGA", "AIA", "AIAC", "AIAB", "UPG", "Plus1"],
    VALID_DEVICES: ["Apple", "Samsung", "Google", "Motorola", "BYOD"],
    VALID_PLANS: ["Premium", "Extra", "Starter"]
  },
  
  // 🆕 ENHANCED DATA LIMITS with dynamic adjustment
  DATA_LIMITS: {
    MAX_ROWS_PER_SHEET: 10000,
    BATCH_SIZE: 1000,
    CHART_DATA_POINTS: 500,
    CLIENT_LIST_LIMIT: 5000,
    COLUMN_DETECTION_TIMEOUT: 5000
  },
  
  // Style Configuration
  STYLES: {
    main: "#334155", accent: "#64748B", pace: "#1E293B", bg: "#F8FAFC",
    alert: "#991B1B", success: "#15803D", header: "#1E293B", text: "#FFFFFF",
    border: "#CBD5E1", sub: "#475569", zebraLight: "#FFFFFF", zebraDark: "#E2E8F0"
  },
  
  CHART_CONFIGS: {
    unitChart: {
      type: 'BAR',
      options: {
        series: { 
          0: {color: '#1E293B', dataLabel: 'value'},
          1: {color: '#94A3B8', dataLabel: 'value'}
        },
        width: 750, height: 240
      }
    },
    accChart: {
      type: 'COLUMN',
      options: {
        series: { 
          0: {color: '#475569', dataLabel: 'value'},
          1: {color: '#CBD5E1', dataLabel: 'value'}
        },
        width: 250, height: 350
      }
    }
  },
  
  CACHE_CONFIG: {
    DEFAULT_TTL: 300,
    SHEET_DATA_TTL: 180,
    CHART_DATA_TTL: 600,
    COLUMN_MAP_TTL: 3600,
    MAX_CACHE_KEYS: 100,
    ENABLE_CACHE: true
  },
  
  EXECUTION_LIMITS: {
    MAX_TIME_MS: 300000,
    MIN_QUOTA_TRIGGERS: 10,
    MAX_API_CALLS_PER_RUN: 5000
  }
};

// =================================================================================
// ENHANCED FORMULA BUILDER UTILITIES
// =================================================================================

var FormulaBuilder = {
  buildFormula: function(template, params) {
    return template.replace(/\{(\w+)\}/g, function(match, key) {
      return params[key] !== undefined ? params[key] : match;
    });
  },
  
  // 🆕 ENHANCED FORMULA TEMPLATES with UnifiedDataAccess awareness
  templates: {
    COUNTIF_MONTH: '=COUNTIF(\'{sheetName}\'!{column}:{column},"{criteria}")',
    COUNTIFS_DATE: '=COUNTIFS(\'{sheetName}\'!{typeColumn}:{typeColumn},"{criteria}",\'{sheetName}\'!{dateColumn}:{dateColumn},"<="&{dayValue})',
    GAP_PROJECTION: '=ROUND((({actualCell}/{divisor})*{daysInMonth}), 0) & " (" & IF(({actualCell}-({baseCell}/{daysInMonth})*{daysPassed})>=0, "+" & ROUND({actualCell}-({baseCell}/{daysInMonth})*{daysPassed}, 1), ROUND({actualCell}-({baseCell}/{daysInMonth})*{daysPassed}, 1)) & " units)"',
    IFERROR_VALUE: '=IFERROR(N(\'{sheetName}\'!{cell}), 0)',
    PERCENTAGE: '=IFERROR({numerator}/{denominator}, 0)',
    COMPARISON_ARROW: '=IF({mainCell}>{prevCell}, "▲ vs Last Mo", IF({mainCell}<{prevCell}, "▼ vs Last Mo", "— vs Last Mo"))',
    ACCESSORY_PROJECTION: '=TEXT(({actualValue}/{divisor})*{daysInMonth}, "$#,##0") & " (" & IF(({actualValue}-(({goalValue}/{daysInMonth})*{daysPassed}))>=0, "+$" & TEXT({actualValue}-(({goalValue}/{daysInMonth})*{daysPassed}), "#,##0"), "-$" & TEXT(ABS({actualValue}-(({goalValue}/{daysInMonth})*{daysPassed})), "#,##0")) & ")"'
  },
  
  // Enhanced formula building functions
  buildComparisonFormula: function(mainCell, prevCell) {
    return this.buildFormula(this.templates.COMPARISON_ARROW, {
      mainCell: mainCell,
      prevCell: prevCell
    });
  }
};

// =================================================================================
// ENHANCED CORE UTILITY FUNCTIONS
// =================================================================================

/**
 * 🆕 ENHANCED checkExecutionQuota
 */
function checkExecutionQuota() {
  try {
    var remainingTriggers = ScriptApp.getRemainingDailyTriggers();
    var isLow = remainingTriggers < CONFIG.EXECUTION_LIMITS.MIN_QUOTA_TRIGGERS;
    
    if (isLow) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "⚠️ Low daily quota: " + remainingTriggers + " triggers remaining", 
        "Quota Warning", 6
      );
    }
    
    return !isLow;
    
  } catch (quotaError) {
    return true; // Continue if quota check fails
  }
}

/**
 * Enhanced working days calculation
 */
function calculateWorkingDaysLeft(currentMonth, today) {
  try {
    var targetIdx = CONFIG.MONTH_NAMES.indexOf(currentMonth);
    var nowIdx = today.getMonth();
    var lastDay = new Date(today.getFullYear(), targetIdx + 1, 0).getDate();
    
    if (targetIdx < nowIdx) return 0;
    
    var startDay = (targetIdx === nowIdx) ? today.getDate() : 1;
    var count = 0;
    
    for (var d = startDay; d <= lastDay; d++) {
      var date = new Date(today.getFullYear(), targetIdx, d);
      var dayOfWeek = date.getDay();
      
      // Skip weekends
      if (!CONFIG.WEEKEND_DAYS.includes(dayOfWeek)) {
        count++;
      }
    }
    
    return Math.max(1, count);
    
  } catch (error) {
    return 1; // Safe fallback
  }
}
