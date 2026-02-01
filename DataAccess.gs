/**
 * =================================================================================
 * UNIFIED DATA ACCESS LAYER (v4.8 - RANGE SAFETY FIX)
 * =================================================================================
 */

var UnifiedDataAccess = {
  getSheetData: function(sheetName, options = {}) {
    var config = this.buildConfig(options);
    var cacheKey = this.buildCacheKey(sheetName, config);
    
    if (config.useCache) {
      var cached = ADVANCED_CACHE.get(cacheKey);
      if (cached) return cached;
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet not found: " + sheetName);
    
    var method = this.detectAccessMethod(sheet, config);
    var data = this.accessMethods[method].execute(sheet, config);
    
    if (config.useCache && data.length > 0) {
      ADVANCED_CACHE.set(cacheKey, data, config.cacheTTL);
    }
    return data;
  },
  
  buildConfig: function(options) {
    return {
      maxRows: options.maxRows || CONFIG.DATA_LIMITS.MAX_ROWS_PER_SHEET,
      maxCols: options.maxCols || 20,
      startRow: options.startRow || 2,
      batchSize: options.batchSize || CONFIG.DATA_LIMITS.BATCH_SIZE,
      useCache: options.useCache === true,
      cacheTTL: options.cacheTTL || CONFIG.CACHE_CONFIG.SHEET_DATA_TTL,
      includeHeaders: options.includeHeaders || false
    };
  },
  
  detectAccessMethod: function(sheet, config) {
    try {
      sheet.getRange("A1:C1").getValues();
      var lastRow = sheet.getLastRow();
      return (lastRow <= config.batchSize + config.startRow) ? 'BULK_ACCESS' : 'BATCH_PROCESSING';
    } catch (e) {
      return 'SMART_TABLE_ACCESS';
    }
  },
  
  accessMethods: {
    BULK_ACCESS: {
      execute: function(sheet, config) {
        var lastRow = Math.min(sheet.getLastRow(), config.maxRows + config.startRow - 1);
        
        // SAFETY CHECK: Ensure we have valid dimensions
        if (lastRow < config.startRow) return [];
        
        var start = config.includeHeaders ? 1 : config.startRow;
        var rows = config.includeHeaders ? (lastRow - start + 1) : (lastRow - config.startRow + 1);
        
        if (rows <= 0) return []; // Prevent invalid range error
        
        return sheet.getRange(start, 1, rows, config.maxCols).getDisplayValues();
      }
    },
    BATCH_PROCESSING: {
      execute: function(sheet, config) {
        var allData = [];
        var lastRow = Math.min(sheet.getLastRow(), config.maxRows + config.startRow - 1);
        var start = config.includeHeaders ? 1 : config.startRow;
        
        for (var r = start; r <= lastRow; r += config.batchSize) {
          var num = Math.min(r + config.batchSize - 1, lastRow) - r + 1;
          if (num > 0) {
            allData = allData.concat(sheet.getRange(r, 1, num, config.maxCols).getDisplayValues());
          }
        }
        return allData;
      }
    },
    SMART_TABLE_ACCESS: {
      execute: function(sheet, config) {
        var data = [];
        var lastRow = Math.min(sheet.getLastRow(), config.maxRows + config.startRow - 1);
        var start = config.includeHeaders ? 1 : config.startRow;
        for (var r = start; r <= lastRow; r++) {
          var row = [];
          for (var c = 1; c <= config.maxCols; c++) {
            try { row.push(sheet.getRange(r, c).getDisplayValue()); } catch(e) { row.push(""); }
          }
          data.push(row);
        }
        return data;
      }
    }
  },
  
  buildCacheKey: function(sheetName, config) {
    return 'unified_data_' + sheetName + '_' + config.maxRows + '_' + (config.includeHeaders ? 'h' : 'nh');
  },
  
  getColumnMapping: function(sheetName) {
    var cacheKey = 'unified_column_map_' + sheetName;
    var cached = ADVANCED_CACHE.get(cacheKey);
    if (cached) return cached;
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var mapping = sheet ? CONFIG.getColumnMap(sheet) : CONFIG.getDefaultColumnMapping();
    ADVANCED_CACHE.set(cacheKey, mapping, 3600);
    return mapping;
  },
  
  clearCache: function(sheetName) {
    if (sheetName) {
      ADVANCED_CACHE.remove('unified_column_map_' + sheetName);
    } else {
      ADVANCED_CACHE.invalidateGroup('SHEET_DATA');
    }
  },
  
  getStats: function() {
    return {
      accessMethodsAvailable: Object.keys(this.accessMethods),
      defaultMaxRows: CONFIG.DATA_LIMITS.MAX_ROWS_PER_SHEET
    };
  }
};

function getSheetDataWithMapping(sheetName, options = {}) {
  var data = UnifiedDataAccess.getSheetData(sheetName, options);
  var colMap = UnifiedDataAccess.getColumnMapping(sheetName);
  return { data: data, columnMapping: colMap };
}
