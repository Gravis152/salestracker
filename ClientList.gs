/**
 * =================================================================================
 * CLIENT LIST GENERATOR (v5.5 - PRODUCTION)
 * Processes all months, handles empty/zero dates gracefully
 * =================================================================================
 */

function updateClientList() {
  var startTime = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clientSheet = ss.getSheetByName("Client List") || ss.insertSheet("Client List");
  
  clientSheet.clear();
  
  var today = new Date();
  var currentYear = today.getFullYear();
  var clientMap = {}; 
  var totalTransactions = 0;
  var processedMonths = [];
  var errorMonths = [];
  
  // Process ALL Months
  CONFIG.MONTH_NAMES.forEach(function(mName, index) {
    try {
      var result = getSheetDataWithMapping(mName, { useCache: false });
      var data = result.data;
      var map = result.columnMapping;
      
      if (!data || data.length === 0) return;
      
      var monthTransactions = 0;
      
      data.forEach(function(row) {
        var rawDate = row[map.DATE_INDEX];
        var type = String(row[map.TYPE_INDEX] || '').trim();
        var name = String(row[map.CUSTOMER_INDEX] || '').trim();
        var phone = String(row[map.MOBILE_INDEX] || '').trim();
        var note = String(row[map.NOTES_INDEX] || '').trim();
        
        if (!name || name.toLowerCase() === 'customer' || name.toLowerCase() === 'undefined' || name === '') {
          return;
        }
        
        var dateObj = parseClientDate(rawDate, currentYear, index);
        
        if (!dateObj || isNaN(dateObj.getTime())) {
          dateObj = new Date(currentYear, index, 1, 12, 0, 0);
        }
        
        if (type.toLowerCase() === 'plus1') {
          phone += " ➕";
        }
        
        if (!clientMap[name]) { 
          clientMap[name] = { phones: [], date: dateObj, notes: [] }; 
        }
        
        if (dateObj < clientMap[name].date) {
          clientMap[name].date = dateObj;
        }
        
        if (phone && !clientMap[name].phones.includes(phone)) {
          clientMap[name].phones.push(phone);
        }
        
        if (note && !clientMap[name].notes.includes(note)) {
          clientMap[name].notes.push(note);
        }
        
        monthTransactions++;
        totalTransactions++;
      });
      
      if (monthTransactions > 0) {
        processedMonths.push(mName + " (" + monthTransactions + ")");
      }
      
    } catch (e) {
      errorMonths.push(mName);
    }
  });
  
  // Format Output
  var output = [];
  
  for (var name in clientMap) {
    var c = clientMap[name];
    var pStr = c.phones.length > 1 ? "• " + c.phones.join("\n• ") : (c.phones[0] || "");
    var nStr = c.notes.length > 1 ? "• " + c.notes.join("\n• ") : (c.notes[0] || "");
    output.push([c.date, name, pStr, nStr]);
  }
  
  output.sort(function(a, b) { 
    return String(a[1]).localeCompare(String(b[1]), undefined, {sensitivity: 'base'}); 
  });
  
  // Render
  clientSheet.getRange("A1:D1")
    .setValues([["Date Added", "Customer Name", "Mobile Numbers", "Notes (Combined)"]])
    .setBackground("#1E293B")
    .setFontColor("white")
    .setFontWeight("bold");
    
  if (output.length > 0) {
    var range = clientSheet.getRange(2, 1, output.length, 4);
    range.setValues(output);
    
    clientSheet.getRange(2, 1, output.length, 1).setNumberFormat("MM/dd/yyyy");
    
    range.setHorizontalAlignment("left")
      .setVerticalAlignment("top")
      .setWrap(true);
    
    range.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    
    var colors = output.map(function(_, i) {
      var bgColor = (i % 2 === 0) ? "#FFFFFF" : "#E2E8F0";
      return [bgColor, bgColor, bgColor, bgColor];
    });
    range.setBackgrounds(colors);
  }
  
  clientSheet.setColumnWidth(1, 110); 
  clientSheet.setColumnWidth(2, 220); 
  clientSheet.setColumnWidth(3, 180); 
  clientSheet.setColumnWidth(4, 450);
  clientSheet.setFrozenRows(1);
  
  var processingTime = new Date() - startTime;
  
  var summaryMsg = "✅ Client List Updated\n";
  summaryMsg += "• " + output.length + " unique clients\n";
  summaryMsg += "• " + totalTransactions + " transactions\n";
  summaryMsg += "• Months: " + processedMonths.length + " with data";
  
  if (errorMonths.length > 0) {
    summaryMsg += "\n⚠️ Errors in: " + errorMonths.join(", ");
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(summaryMsg, "Complete", 8);
  
  ADVANCED_CACHE.set('client_list_data', {
    count: output.length,
    transactions: totalTransactions,
    timestamp: new Date(),
    processingTime: processingTime,
    monthsProcessed: processedMonths,
    errors: errorMonths
  }, 3600);
  
  return {
    clients: output.length,
    transactions: totalTransactions,
    processed: processedMonths,
    errors: errorMonths
  };
}

/**
 * Parse date from sheet - handles multiple formats without timezone shift
 */
function parseClientDate(rawDate, currentYear, monthIndex) {
  currentYear = currentYear || new Date().getFullYear();
  
  if (rawDate === null || rawDate === undefined || rawDate === '' || rawDate === 0 || rawDate === '0') {
    return null;
  }
  
  if (rawDate instanceof Date) {
    if (isNaN(rawDate.getTime())) return null;
    if (rawDate.getFullYear() < 2010) return null;
    return new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate(), 12, 0, 0);
  }
  
  if (typeof rawDate === 'number') {
    if (rawDate <= 0) return null;
    
    if (rawDate > 40000 && rawDate < 100000) {
      var excelDate = new Date((rawDate - 25569) * 86400 * 1000);
      if (!isNaN(excelDate.getTime()) && excelDate.getFullYear() >= 2010) {
        return new Date(excelDate.getFullYear(), excelDate.getMonth(), excelDate.getDate(), 12, 0, 0);
      }
    }
    
    if (rawDate >= 1 && rawDate <= 31 && monthIndex !== undefined) {
      return new Date(currentYear, monthIndex, rawDate, 12, 0, 0);
    }
    
    return null;
  }
  
  var dateStr = String(rawDate).trim();
  
  if (!dateStr || dateStr === '0' || dateStr.startsWith('#')) {
    return null;
  }
  
  var slashMatch = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    var month = parseInt(slashMatch[1], 10) - 1;
    var day = parseInt(slashMatch[2], 10);
    var year = parseInt(slashMatch[3], 10);
    
    if (year < 100) year += (year < 50) ? 2000 : 1900;
    
    if (month >= 0 && month <= 11 && day >= 1 && day <= 31 && year >= 2010 && year <= 2100) {
      return new Date(year, month, day, 12, 0, 0);
    }
  }
  
  var dashMatch = dateStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (dashMatch) {
    var year = parseInt(dashMatch[1], 10);
    var month = parseInt(dashMatch[2], 10) - 1;
    var day = parseInt(dashMatch[3], 10);
    
    if (month >= 0 && month <= 11 && day >= 1 && day <= 31 && year >= 2010 && year <= 2100) {
      return new Date(year, month, day, 12, 0, 0);
    }
  }
  
  var dashMatch2 = dateStr.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dashMatch2) {
    var month = parseInt(dashMatch2[1], 10) - 1;
    var day = parseInt(dashMatch2[2], 10);
    var year = parseInt(dashMatch2[3], 10);
    
    if (month >= 0 && month <= 11 && day >= 1 && day <= 31 && year >= 2010 && year <= 2100) {
      return new Date(year, month, day, 12, 0, 0);
    }
  }
  
  var textMatch = dateStr.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s*(\d{4})$/);
  if (textMatch) {
    var monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    var monthStr = textMatch[1].toLowerCase().substring(0, 3);
    var monthNum = monthNames.indexOf(monthStr);
    var day = parseInt(textMatch[2], 10);
    var year = parseInt(textMatch[3], 10);
    
    if (monthNum >= 0 && day >= 1 && day <= 31 && year >= 2010 && year <= 2100) {
      return new Date(year, monthNum, day, 12, 0, 0);
    }
  }
  
  try {
    var fallback = new Date(dateStr);
    if (!isNaN(fallback.getTime()) && fallback.getFullYear() >= 2010) {
      return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate(), 12, 0, 0);
    }
  } catch (e) {}
  
  return null;
}
