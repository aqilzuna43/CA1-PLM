/**
 * Google Sheets API wrapper with caching
 */

class SheetService {
  constructor(spreadsheetId) {
    this.spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  }
  
  /**
   * Get sheet data with caching
   * @param {string} sheetName - Sheet name
   * @param {boolean} useCache - Whether to use cache
   * @returns {Array} Sheet data
   */
  getSheetData(sheetName, useCache = true) {
    const cacheKey = `sheet_${sheetName}`;
    
    // Try cache first
    if (useCache) {
      const cached = cacheManager.get(cacheKey);
      if (cached) return cached;
    }
    
    // Fetch from sheet
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Cache the result
    if (useCache) {
      cacheManager.set(cacheKey, data);
    }
    
    return data;
  }
  
  /**
   * Batch write to sheet (optimized)
   * @param {string} sheetName - Sheet name
   * @param {Array} data - 2D array of data
   * @param {number} startRow - Start row (1-based)
   * @param {number} startCol - Start column (1-based)
   */
  batchWrite(sheetName, data, startRow = 1, startCol = 1) {
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const numRows = data.length;
    const numCols = data[0].length;
    
    sheet.getRange(startRow, startCol, numRows, numCols).setValues(data);
    
    // Invalidate cache
    cacheManager.clear(`sheet_${sheetName}`);
  }
  
  /**
   * Find row by criteria
   * @param {string} sheetName - Sheet name
   * @param {number} colIndex - Column index to search (0-based)
   * @param {*} value - Value to find
   * @returns {number} Row index (0-based) or -1
   */
  findRow(sheetName, colIndex, value) {
    const data = this.getSheetData(sheetName);
    return data.findIndex(row => row[colIndex] === value);
  }
}

// Global instance
const sheetService = new SheetService(CONFIG.MASTER_SHEET_ID);