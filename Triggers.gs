/**
 * Triggers.gs
 * Menu system, onOpen/onEdit handlers, and goal tracking
 * 🆕 v2.0 - Added Dashboard auto-refresh on month selection
 */

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 Sales Tracker')
    .addItem('🔄 Refresh All Data', 'refreshAllData')
    .addItem('📈 Generate MTD Dashboard', 'generateMTDDashboard')
    .addItem('📊 Generate YTD Report', 'generateYTDReport')
    .addItem('👥 Update Client List', 'updateClientList')
    .addSeparator()
    .addItem('📱 Quick Entry Form', 'showQuickEntryForm')
    .addSeparator()
    .addItem('🎯 Show Goal Progress', 'showGoalProgress')
    .addItem('⚙️ System Tools', 'showSystemTools')
    .addSeparator()
    .addItem('🗄️ Year-End Archive', 'showYearEndDialog')
    .addToUi();
  
  // Clean up old triggers on open
  cleanupOldTriggers();
}

/**
 * Handles edit events for real-time validation
 * Delegates to DataValidation.gs
 * 🆕 AUTO-REFRESHES Dashboard when month selector changes
 */
function onEdit(e) {
  if (!e) return;
  
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    
    // 🆕 AUTO-REFRESH DASHBOARD when month selector (M2) changes
    if (sheetName === "Dashboard" && range.getA1Notation() === "M2") {
      // Small delay to ensure value is set
      Utilities.sleep(100);
      
      // Regenerate dashboard with new month
      createRestoredMTDDashboard();
      
      // Show toast notification
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Dashboard updated to ' + range.getValue(),
        '📊 Refreshed',
        2
      );
      
      return; // Skip validation for dashboard edits
    }
    
    // Call validation handler from DataValidation.gs for data sheets
    if (typeof validateOnEdit === 'function') {
      validateOnEdit(e);
    }
  } catch (error) {
    console.error('onEdit error:', error);
    // Don't show UI alerts in onEdit - can be disruptive
  }
}

/**
 * Cleanup old/duplicate triggers
 * Prevents trigger accumulation over time
 */
function cleanupOldTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const seen = {};
    
    triggers.forEach(function(trigger) {
      const fn = trigger.getHandlerFunction();
      
      // Only clean up onOpen/onEdit duplicates (keep first instance)
      if (['onOpen', 'onEdit'].includes(fn)) {
        if (seen[fn]) {
          ScriptApp.deleteTrigger(trigger); // Delete duplicate
        } else {
          seen[fn] = true; // Mark first as seen
        }
      }
    });
  } catch (error) {
    console.error('cleanupOldTriggers error:', error);
  }
}

/**
 * Shows goal progress dialog
 * Displays MTD/YTD progress vs goals
 */
function showGoalProgress() {
  try {
    const config = getConfig();
    const progress = calculateGoalProgress();
    
    const ui = SpreadsheetApp.getUi();
    
    const message = `
📊 GOAL PROGRESS

Monthly Goal: ${formatCurrency(config.MONTHLY_GOAL)}
MTD Actual:   ${formatCurrency(progress.mtd.actual)}
MTD Progress: ${progress.mtd.percentage}%
Status:       ${progress.mtd.status}

─────────────────────────

Yearly Goal:  ${formatCurrency(config.YEARLY_GOAL)}
YTD Actual:   ${formatCurrency(progress.ytd.actual)}
YTD Progress: ${progress.ytd.percentage}%
Status:       ${progress.ytd.status}

─────────────────────────

${progress.message}
    `.trim();
    
    ui.alert('🎯 Goal Tracking', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error loading goal progress: ' + error.message);
    console.error('showGoalProgress error:', error);
  }
}

/**
 * Calculates goal progress for MTD and YTD
 * @returns {Object} Progress data
 */
function calculateGoalProgress() {
  try {
    const config = getConfig();
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    // Get MTD data
    const mtdData = getMTDSalesData(currentMonth, currentYear);
    const mtdRevenue = mtdData.reduce(function(sum, row) {
      const amount = parseFloat(row[config.COLUMNS.AMOUNT]) || 0;
      return sum + amount;
    }, 0);
    
    // Get YTD data
    const ytdData = getYTDSalesData(currentYear);
    const ytdRevenue = ytdData.reduce(function(sum, row) {
      const amount = parseFloat(row[config.COLUMNS.AMOUNT]) || 0;
      return sum + amount;
    }, 0);
    
    // Calculate percentages
    const mtdPercentage = Math.round((mtdRevenue / config.MONTHLY_GOAL) * 100);
    const ytdPercentage = Math.round((ytdRevenue / config.YEARLY_GOAL) * 100);
    
    // Determine status
    const mtdStatus = mtdPercentage >= 100 ? '🎉 GOAL MET!' : 
                      mtdPercentage >= 75 ? '💪 On Track' :
                      mtdPercentage >= 50 ? '⚠️ Needs Attention' :
                      '🚨 Behind Goal';
    
    const ytdStatus = ytdPercentage >= 100 ? '🎉 GOAL MET!' :
                      ytdPercentage >= 75 ? '💪 On Track' :
                      ytdPercentage >= 50 ? '⚠️ Needs Attention' :
                      '🚨 Behind Goal';
    
    // Calculate what's needed
    const mtdRemaining = config.MONTHLY_GOAL - mtdRevenue;
    const ytdRemaining = config.YEARLY_GOAL - ytdRevenue;
    
    let message = '';
    if (mtdRemaining > 0) {
      const daysLeft = new Date(currentYear, currentMonth + 1, 0).getDate() - now.getDate();
      const dailyTarget = mtdRemaining / Math.max(daysLeft, 1);
      message = `Need ${formatCurrency(dailyTarget)}/day for ${daysLeft} days to hit monthly goal.`;
    } else {
      message = `Monthly goal exceeded by ${formatCurrency(Math.abs(mtdRemaining))}! 🎉`;
    }
    
    return {
      mtd: {
        actual: mtdRevenue,
        goal: config.MONTHLY_GOAL,
        percentage: mtdPercentage,
        status: mtdStatus,
        remaining: mtdRemaining
      },
      ytd: {
        actual: ytdRevenue,
        goal: config.YEARLY_GOAL,
        percentage: ytdPercentage,
        status: ytdStatus,
        remaining: ytdRemaining
      },
      message: message
    };
    
  } catch (error) {
    console.error('calculateGoalProgress error:', error);
    return {
      mtd: { actual: 0, goal: 0, percentage: 0, status: 'Error', remaining: 0 },
      ytd: { actual: 0, goal: 0, percentage: 0, status: 'Error', remaining: 0 },
      message: 'Error calculating progress'
    };
  }
}

/**
 * Shows system tools dialog
 * Provides access to maintenance functions
 */
function showSystemTools() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      '⚙️ System Tools',
      'Choose an action:\n\n' +
      '1️⃣ Clear All Caches\n' +
      '2️⃣ Rebuild Column Mappings\n' +
      '3️⃣ Validate All Data\n' +
      '4️⃣ Export System Config\n' +
      '5️⃣ View Cache Statistics\n\n' +
      'Enter number (1-5):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    const choice = response.getResponseText().trim();
    
    switch(choice) {
      case '1':
        clearAllCaches();
        ui.alert('✅ All caches cleared');
        break;
      case '2':
        rebuildColumnMappings();
        ui.alert('✅ Column mappings rebuilt');
        break;
      case '3':
        validateAllData();
        break;
      case '4':
        exportSystemConfig();
        break;
      case '5':
        showCacheStatistics();
        break;
      default:
        ui.alert('Invalid choice');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error in system tools: ' + error.message);
    console.error('showSystemTools error:', error);
  }
}

/**
 * Clears all cache layers
 */
function clearAllCaches() {
  try {
    if (typeof Cache !== 'undefined') {
      Cache.clearAll();
    }
    
    // Also clear script cache
    CacheService.getScriptCache().removeAll(['config', 'sales_data', 'dashboard', 'ytd', 'clients']);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('All caches cleared', '✅ Cache', 3);
    
  } catch (error) {
    console.error('clearAllCaches error:', error);
    throw error;
  }
}

/**
 * Rebuilds column mappings for all sheets
 */
function rebuildColumnMappings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let rebuilt = 0;
    
    sheets.forEach(function(sheet) {
      const sheetName = sheet.getName();
      
      // Skip utility sheets
      if (sheetName.startsWith('_') || 
          sheetName === 'Dashboard' || 
          sheetName === 'YTD Report' ||
          sheetName === 'Client List') {
        return;
      }
      
      try {
        const dataAccess = new UnifiedDataAccess(sheet);
        // Force column detection
        dataAccess.detectColumns();
        rebuilt++;
      } catch (e) {
        console.warn(`Could not rebuild mapping for ${sheetName}:`, e);
      }
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Rebuilt mappings for ${rebuilt} sheets`,
      '✅ Column Mappings',
      3
    );
    
  } catch (error) {
    console.error('rebuildColumnMappings error:', error);
    throw error;
  }
}

/**
 * Validates all data across all sheets
 */
function validateAllData() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      '⏳ Validation Running',
      'This may take a few minutes for large datasets.\n\n' +
      'Results will appear in a new sheet called "Validation Report".',
      ui.ButtonSet.OK
    );
    
    // Run validation (implementation in DataValidation.gs)
    if (typeof runFullValidation === 'function') {
      const results = runFullValidation();
      
      ui.alert(
        '✅ Validation Complete',
        `Found ${results.errorCount} errors in ${results.rowsChecked} rows.\n\n` +
        'See "Validation Report" sheet for details.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Validation function not available');
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Validation error: ' + error.message);
    console.error('validateAllData error:', error);
  }
}

/**
 * Exports system configuration
 */
function exportSystemConfig() {
  try {
    const config = getConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get config export sheet
    let sheet = ss.getSheetByName('_Config_Export');
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet('_Config_Export');
    }
    
    // Write config as JSON
    const configJson = JSON.stringify(config, null, 2);
    sheet.getRange(1, 1).setValue('System Configuration Export');
    sheet.getRange(2, 1).setValue(new Date().toISOString());
    sheet.getRange(4, 1).setValue(configJson);
    
    // Format
    sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    sheet.setColumnWidth(1, 800);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Config exported to "_Config_Export" sheet',
      '✅ Export Complete',
      3
    );
    
  } catch (error) {
    console.error('exportSystemConfig error:', error);
    throw error;
  }
}

/**
 * Shows cache statistics
 */
function showCacheStatistics() {
  try {
    const stats = getCacheStats();
    const ui = SpreadsheetApp.getUi();
    
    const message = `
📊 CACHE STATISTICS

Memory Cache:
  Size: ${stats.memory.size} items
  Hit Rate: ${stats.memory.hitRate}%

Script Cache:
  Size: ${stats.script.size} items
  Hit Rate: ${stats.script.hitRate}%

Cache Groups:
  ${Object.keys(stats.groups).map(g => `${g}: ${stats.groups[g]} items`).join('\n  ')}

Last Cleared: ${stats.lastCleared || 'Never'}
    `.trim();
    
    ui.alert('📈 Cache Statistics', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error loading cache stats: ' + error.message);
    console.error('showCacheStatistics error:', error);
  }
}

/**
 * Gets cache statistics
 * @returns {Object} Cache stats
 */
function getCacheStats() {
  try {
    if (typeof Cache !== 'undefined' && typeof Cache.getStats === 'function') {
      return Cache.getStats();
    }
    
    // Fallback if Cache.getStats doesn't exist
    return {
      memory: { size: 0, hitRate: 0 },
      script: { size: 0, hitRate: 0 },
      groups: {},
      lastCleared: null
    };
    
  } catch (error) {
    console.error('getCacheStats error:', error);
    return {
      memory: { size: 0, hitRate: 0 },
      script: { size: 0, hitRate: 0 },
      groups: {},
      lastCleared: null
    };
  }
}

/**
 * Helper to format currency
 */
function formatCurrency(value) {
  if (typeof value !== 'number') {
    value = parseFloat(value) || 0;
  }
  return '$' + value.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Helper functions that might be called from other files
 * These delegate to the appropriate modules
 */

function refreshAllData() {
  if (typeof refreshAll === 'function') {
    refreshAll();
  } else {
    clearAllCaches();
    SpreadsheetApp.getActiveSpreadsheet().toast('Data refreshed', '✅ Refresh', 3);
  }
}

function generateMTDDashboard() {
  if (typeof createRestoredMTDDashboard === 'function') {
    createRestoredMTDDashboard();
  }
}

function generateYTDReport() {
  if (typeof createYTDReport === 'function') {
    createYTDReport();
  }
}

function updateClientList() {
  if (typeof generateClientList === 'function') {
    generateClientList();
  }
}

function showYearEndDialog() {
  if (typeof showYearEndArchiveDialog === 'function') {
    showYearEndArchiveDialog();
  }
}

/**
 * Get MTD sales data
 * Helper function for goal tracking
 */
function getMTDSalesData(month, year) {
  try {
    // Your exact month names
    const monthNames = [
      'Jan', 'Feb', 'Mar', 'Apr', 'May', 'June',
      'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'
    ];
    
    const sheetName = monthNames[month];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn(`Sheet not found: ${sheetName}`);
      return [];
    }
    
    const dataAccess = new UnifiedDataAccess(sheet);
    return dataAccess.getAllData();
    
  } catch (error) {
    console.error('getMTDSalesData error:', error);
    return [];
  }
}

/**
 * Get YTD sales data
 * Helper function for goal tracking
 */
function getYTDSalesData(year) {
  try {
    // Use existing YTD function if available
    if (typeof getAllSalesDataForYear === 'function') {
      return getAllSalesDataForYear(year);
    }
    
    // Fallback: aggregate all month sheets
    const monthNames = [
      'Jan', 'Feb', 'Mar', 'Apr', 'May', 'June',
      'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'
    ];
    
    let allData = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentMonth = new Date().getMonth();
    
    // Only process months up to current month for current year
    const monthsToProcess = (year === new Date().getFullYear()) 
      ? monthNames.slice(0, currentMonth + 1)
      : monthNames;
    
    monthsToProcess.forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        try {
          const dataAccess = new UnifiedDataAccess(sheet);
          const monthData = dataAccess.getAllData();
          allData = allData.concat(monthData);
        } catch (e) {
          console.warn(`Could not read data from ${sheetName}:`, e);
        }
      }
    });
    
    return allData;
    
  } catch (error) {
    console.error('getYTDSalesData error:', error);
    return [];
  }
}

/**
 * Helper to get UnifiedDataAccess instance
 */
function getUnifiedDataAccess(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    return new UnifiedDataAccess(sheet);
    
  } catch (error) {
    console.error('getUnifiedDataAccess error:', error);
    return null;
  }
}
