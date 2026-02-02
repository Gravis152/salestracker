/**
 * SidebarForm.gs
 * Mobile-optimized quick entry form
 * Integrates with dynamic sheet routing and caching
 */

/**
 * Opens the quick entry sidebar (desktop)
 * Called from menu: Tools → Quick Entry Form
 */
function showQuickEntryForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SidebarForm')
      .setTitle('📱 Quick Entry')
      .setWidth(320);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error opening form: ' + error.message);
    console.error('showQuickEntryForm error:', error);
  }
}

/**
 * Alternative entry point for mobile users
 * Can be triggered from a button/image in the sheet
 */
function openQuickEntryMobile() {
  showQuickEntryForm();
}

/**
 * Web app entry point (for direct URL access)
 * Allows mobile users to bookmark and access form directly
 */
function doGet() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarForm')
    .setTitle('📱 Quick Entry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return html;
}

/**
 * Returns configuration data for sidebar form
 */
function getSidebarConfig() {
  try {
    const config = getConfig();
    
    return {
      dropdowns: {
        type: config.DROPDOWN_VALUES?.TYPE || ['PPVGA', 'UPG', 'AIA', 'PLUS1'],
        device: config.DROPDOWN_VALUES?.DEVICE || ['Apple', 'Samsung', 'AiAC', 'AiAB', 'BYOD', 'Google', 'Motorola'],
        ratePlan: config.DROPDOWN_VALUES?.RATE_PLAN || ['Premium', 'Extra', 'Starter']
      },
      
      requiredFields: ['date', 'type', 'device', 'customer', 'mobileNumber', 'ratePlan'],
      
      dateDefault: new Date().toISOString().split('T')[0]
    };
    
  } catch (error) {
    console.error('getSidebarConfig error:', error);
    throw new Error('Failed to load form configuration');
  }
}
/**
 * 🆕 Returns the spreadsheet URL for the "Open Spreadsheet" link
 */
function getSpreadsheetUrl() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getUrl();
  } catch (error) {
    console.error('getSpreadsheetUrl error:', error);
    return null;
  }
}
/**
 * Submits a new sale entry from the sidebar form
 */
function submitSaleEntry(formData) {
  try {
    // Validate required fields
    const config = getSidebarConfig();
    const missing = config.requiredFields.filter(function(field) {
      return !formData[field] || formData[field].toString().trim() === '';
    });
    
    if (missing.length > 0) {
      return {
        success: false,
        error: 'Missing required fields: ' + missing.join(', ')
      };
    }
    
    // Validate and normalize phone number
const digitsOnly = formData.mobileNumber.replace(/\D/g, '');
if (digitsOnly.length !== 10) {
  return {
    success: false,
    error: 'Phone number must be 10 digits'
  };
}
// 🆕 Store as plain number (let sheet formatting handle display)
formData.mobileNumber = digitsOnly;
    // PARSE DATE WITHOUT TIMEZONE CONVERSION
    const dateParts = formData.date.split('-');
    const year = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10) - 1;
    const day = parseInt(dateParts[2], 10);
    
    const entryDate = new Date(year, month, day);
    
    if (isNaN(entryDate.getTime())) {
      return {
        success: false,
        error: 'Invalid date format'
      };
    }
    
    // GET TARGET SHEET based on entry date's month
    const targetSheetName = getMonthSheetName(entryDate);
    
    if (!targetSheetName) {
      return {
        success: false,
        error: 'Could not determine target sheet for date: ' + formData.date
      };
    }
    
    // Get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(targetSheetName);
    
    if (!targetSheet) {
      return {
        success: false,
        error: `Sheet "${targetSheetName}" does not exist. Please create it first.`
      };
    }
    
    // Get column mappings
    const columns = detectSheetColumns(targetSheet);
    
    // BUILD ROW DATA - ONLY for columns A-G (data columns only)
    const dataColumns = ['Date', 'Type', 'Device', 'Customer', 'Mobile Number', 'Rate Plan', 'Notes'];
    const maxDataColIndex = Math.max(...dataColumns.map(col => columns[col] !== undefined ? columns[col] : -1));
    const numDataCols = maxDataColIndex + 1;
    
    const rowData = new Array(numDataCols).fill('');
    
    rowData[columns['Date']] = formData.date;
    rowData[columns['Type']] = formData.type;
    rowData[columns['Device']] = formData.device;
    rowData[columns['Customer']] = formData.customer;
    rowData[columns['Mobile Number']] = formData.mobileNumber;
    rowData[columns['Rate Plan']] = formData.ratePlan;
    if (columns['Notes'] !== undefined) {
      rowData[columns['Notes']] = formData.notes || '';
    }
    
    // Check for duplicates
    const isDuplicate = checkForDuplicate(targetSheet, formData, columns);
    if (isDuplicate) {
      return {
        success: false,
        error: '⚠️ Possible duplicate: This customer/phone was already entered today'
      };
    }
    
    // Find first empty row in data columns
    const firstEmptyRow = findFirstEmptyRowInDataColumns(targetSheet, columns);
    
    // WRITE ONLY TO DATA COLUMNS (A-G)
    targetSheet.getRange(firstEmptyRow, 1, 1, numDataCols).setValues([rowData]);
    
    // Invalidate caches
    if (typeof Cache !== 'undefined') {
      Cache.invalidate('sales_data');
      Cache.invalidate('dashboard');
      Cache.invalidate('clients');
    }
    
    return {
      success: true,
      message: `✅ Entry added to ${targetSheetName} (Row ${firstEmptyRow})`,
      details: `${formData.customer} - ${formData.type} - ${formData.device}`
    };
    
  } catch (error) {
    console.error('submitSaleEntry error:', error);
    return {
      success: false,
      error: 'Failed to submit entry: ' + error.message
    };
  }
}

/**
 * FINDS FIRST EMPTY ROW in data columns only
 */
function findFirstEmptyRowInDataColumns(sheet, columns) {
  try {
    const dataColumns = ['Date', 'Type', 'Device', 'Customer', 'Mobile Number', 'Rate Plan', 'Notes'];
    const dataColIndices = dataColumns
      .map(col => columns[col])
      .filter(idx => idx !== undefined);
    
    if (dataColIndices.length === 0) {
      throw new Error('No data columns detected');
    }
    
    const maxDataCol = Math.max(...dataColIndices) + 1;
    
    let currentRow = 2;
    const maxRowsToCheck = 1000;
    
    while (currentRow < maxRowsToCheck) {
      const rowData = sheet.getRange(currentRow, 1, 1, maxDataCol).getValues()[0];
      const isEmpty = rowData.every(cell => !cell || cell.toString().trim() === '');
      
      if (isEmpty) {
        return currentRow;
      }
      
      currentRow++;
    }
    
    return currentRow;
    
  } catch (error) {
    console.error('findFirstEmptyRowInDataColumns error:', error);
    return sheet.getLastRow() + 1;
  }
}

/**
 * Detects column positions in a sheet
 */
function detectSheetColumns(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      throw new Error('Sheet has no columns');
    }
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const columns = {};
    
    headers.forEach(function(header, index) {
      if (header) {
        const normalized = header.toString().trim();
        columns[normalized] = index;
      }
    });
    
    const required = ['Date', 'Type', 'Device', 'Customer', 'Mobile Number', 'Rate Plan'];
    const missing = required.filter(function(col) {
      return columns[col] === undefined;
    });
    
    if (missing.length > 0) {
      throw new Error('Missing required columns: ' + missing.join(', '));
    }
    
    return columns;
    
  } catch (error) {
    console.error('detectSheetColumns error:', error);
    throw error;
  }
}

/**
 * Checks for duplicate entries
 */
function checkForDuplicate(sheet, formData, columns) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    
    const startRow = Math.max(2, lastRow - 49);
    const numRows = lastRow - startRow + 1;
    
    const dataColumns = ['Date', 'Type', 'Device', 'Customer', 'Mobile Number', 'Rate Plan'];
    const maxDataCol = Math.max(...dataColumns.map(col => columns[col] || 0)) + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, maxDataCol).getValues();
    
    const formDate = new Date(formData.date).toDateString();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowDate = new Date(row[columns['Date']]).toDateString();
      const rowPhone = row[columns['Mobile Number']];
      const rowCustomer = row[columns['Customer']];
      
      if (rowDate === formDate && 
          rowPhone === formData.mobileNumber && 
          rowCustomer === formData.customer) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('checkForDuplicate error:', error);
    return false;
  }
}

/**
 * DETERMINES CORRECT MONTH SHEET based on entry date
 */
function getMonthSheetName(date) {
  try {
    const monthNames = [
      'Jan', 'Feb', 'Mar', 'Apr', 'May', 'June',
      'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'
    ];
    
    const monthIndex = date.getMonth();
    const sheetName = monthNames[monthIndex];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found. Please create it first.`);
    }
    
    return sheetName;
    
  } catch (error) {
    console.error('getMonthSheetName error:', error);
    throw error;
  }
}
