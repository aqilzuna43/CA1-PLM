/**
 * Google Sheets API wrapper with caching support and pre-write validation.
 * Provides batch read/write operations to minimize API calls.
 *
 * Usage:
 *   const svc = new SheetService('your-spreadsheet-id');
 *   const data = svc.getSheetData('MASTER');           // cached
 *   svc.batchWrite('MASTER', rows, 2, 1);              // writes + cache invalidation
 *   svc.validatedBatchWrite('MASTER', rows, 2, 1);     // validates + writes
 */

class SheetService {
  /**
   * @param {string} spreadsheetId The Google Sheets ID to operate on.
   */
  constructor(spreadsheetId) {
    this.spreadsheetId = spreadsheetId;
    this._spreadsheet = null; // Lazy-loaded
  }

  /** @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} */
  get spreadsheet() {
    if (!this._spreadsheet) {
      this._spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
    }
    return this._spreadsheet;
  }

  /**
   * Get sheet data with optional caching.
   * @param {string} sheetName Sheet name.
   * @param {boolean} [useCache=true] Whether to use CacheService.
   * @returns {Array<Array>} 2D array of sheet data.
   */
  getSheetData(sheetName, useCache = true) {
    const cacheKey = `sheet_${sheetName}`;

    if (useCache) {
      const cached = cacheManager.get(cacheKey);
      if (cached) return cached;
    }

    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: "${sheetName}" in spreadsheet ${this.spreadsheetId}`);
    }

    const data = sheet.getDataRange().getValues();

    if (useCache) {
      cacheManager.set(cacheKey, data);
    }

    return data;
  }

  /**
   * Batch write data to a sheet (single API call).
   * Automatically invalidates the cache for the target sheet.
   * @param {string} sheetName Sheet name.
   * @param {Array<Array>} data 2D array of data to write.
   * @param {number} [startRow=1] Start row (1-based).
   * @param {number} [startCol=1] Start column (1-based).
   */
  batchWrite(sheetName, data, startRow = 1, startCol = 1) {
    if (!data || data.length === 0) return;

    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: "${sheetName}" in spreadsheet ${this.spreadsheetId}`);
    }

    const numRows = data.length;
    const numCols = data[0].length;

    sheet.getRange(startRow, startCol, numRows, numCols).setValues(data);

    // Invalidate cache after write
    cacheManager.clear(`sheet_${sheetName}`);
  }

  /**
   * Validated batch write — runs pre-write integrity checks before writing to MASTER.
   * If validation fails, returns a report of issues WITHOUT writing.
   * If validation passes, writes data and invalidates cache.
   *
   * @param {string} sheetName Sheet name (validation only runs for MASTER).
   * @param {Array<Array>} data 2D array of data to write.
   * @param {number} [startRow=1] Start row (1-based).
   * @param {number} [startCol=1] Start column (1-based).
   * @param {Object} [options] Validation options.
   * @param {boolean} [options.skipValidation=false] Skip validation (for reconcile operations).
   * @returns {{success: boolean, issues: Array<string>, written: number}}
   */
  validatedBatchWrite(sheetName, data, startRow = 1, startCol = 1, options = {}) {
    if (!data || data.length === 0) {
      return { success: true, issues: [], written: 0 };
    }

    // Only validate writes to MASTER sheet
    const isMasterWrite = typeof BOM_CONFIG !== 'undefined' &&
                          sheetName === BOM_CONFIG.MASTER_SHEET_NAME;

    if (isMasterWrite && !options.skipValidation) {
      const issues = preWriteValidation_(this, data, startRow, startCol);
      if (issues.length > 0) {
        return { success: false, issues: issues, written: 0 };
      }
    }

    // Validation passed (or skipped) — execute write
    this.batchWrite(sheetName, data, startRow, startCol);
    return { success: true, issues: [], written: data.length };
  }

  /**
   * Find the first row matching a value in a specific column.
   * @param {string} sheetName Sheet name.
   * @param {number} colIndex Column index (0-based).
   * @param {*} value Value to find.
   * @returns {number} Row index (0-based) or -1 if not found.
   */
  findRow(sheetName, colIndex, value) {
    const data = this.getSheetData(sheetName);
    return data.findIndex(row => row[colIndex] === value);
  }
}


// ============================================================================
// PRE-WRITE VALIDATION GUARD
// ============================================================================

/**
 * Runs referential integrity checks on data before writing to MASTER.
 * Returns an array of issue descriptions. Empty array = all checks passed.
 *
 * Checks:
 *   1. Item Numbers exist in ITEMS sheet (no orphans)
 *   2. Level hierarchy is valid (no gaps > 1)
 *   3. Qty values are positive numbers (where present)
 *   4. Lifecycle states are recognized (if LIFECYCLE config exists)
 *
 * @param {SheetService} svc SheetService instance.
 * @param {Array<Array>} data Data rows to validate.
 * @param {number} startRow Start row (1-based, for error messages).
 * @param {number} startCol Start column (1-based).
 * @returns {Array<string>} Array of issue descriptions.
 */
function preWriteValidation_(svc, data, startRow, startCol) {
  const issues = [];

  // Bail out if BOM_CONFIG or COL not available
  if (typeof BOM_CONFIG === 'undefined' || typeof COL === 'undefined') {
    return issues;
  }

  // Build column index from first row if it looks like headers,
  // or from the existing MASTER headers
  let headers;
  try {
    const masterData = svc.getSheetData(BOM_CONFIG.MASTER_SHEET_NAME);
    headers = masterData[0];
  } catch (e) {
    return issues; // Can't read MASTER — skip validation
  }

  const colIdx = {};
  if (typeof getColumnIndexes === 'function') {
    const idx = getColumnIndexes(headers);
    Object.assign(colIdx, idx);
  } else {
    return issues; // Helper not available
  }

  // Build ITEMS lookup for FK validation
  let itemsSet = null;
  try {
    const itemsData = svc.getSheetData(BOM_CONFIG.ITEMS_SHEET_NAME);
    itemsSet = new Set();
    const pnIdx = itemsData[0].indexOf('Item Number');
    const col = pnIdx !== -1 ? pnIdx : 0;
    for (let i = 1; i < itemsData.length; i++) {
      const pn = itemsData[i][col];
      if (pn) itemsSet.add(pn.toString().trim());
    }
  } catch (e) {
    // ITEMS sheet not available — skip FK check
  }

  // Build lifecycle valid states set
  const validLifecycleStates = BOM_CONFIG.LIFECYCLE && BOM_CONFIG.LIFECYCLE.STATES
    ? new Set(BOM_CONFIG.LIFECYCLE.STATES)
    : null;

  // Determine column offsets relative to startCol
  const itemNumColData = colIdx[COL.ITEM_NUMBER] !== undefined ? colIdx[COL.ITEM_NUMBER] - (startCol - 1) : -1;
  const levelColData = colIdx[COL.LEVEL] !== undefined ? colIdx[COL.LEVEL] - (startCol - 1) : -1;
  const qtyColData = colIdx[COL.QTY] !== undefined ? colIdx[COL.QTY] - (startCol - 1) : -1;
  const lifecycleColData = colIdx[COL.LIFECYCLE] !== undefined ? colIdx[COL.LIFECYCLE] - (startCol - 1) : -1;

  let prevLevel = null;
  const orphanPNs = [];
  const levelGaps = [];
  const badQtys = [];
  const badLifecycles = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const sheetRow = startRow + i;

    // --- Check 1: Item Number exists in ITEMS ---
    if (itemNumColData >= 0 && itemNumColData < row.length && itemsSet) {
      const pn = row[itemNumColData];
      if (pn && pn.toString().trim() !== '') {
        const pnStr = pn.toString().trim();
        if (!itemsSet.has(pnStr)) {
          orphanPNs.push({ row: sheetRow, pn: pnStr });
        }
      }
    }

    // --- Check 2: Level hierarchy (no gaps > 1) ---
    if (levelColData >= 0 && levelColData < row.length) {
      const levelVal = row[levelColData];
      if (levelVal !== '' && levelVal !== null) {
        const level = parseInt(levelVal, 10);
        if (!isNaN(level) && prevLevel !== null && level > prevLevel + 1) {
          levelGaps.push({ row: sheetRow, from: prevLevel, to: level });
        }
        if (!isNaN(level)) prevLevel = level;
      }
    }

    // --- Check 3: Qty is positive ---
    if (qtyColData >= 0 && qtyColData < row.length) {
      const qtyVal = row[qtyColData];
      if (qtyVal !== '' && qtyVal !== null) {
        const num = Number(qtyVal);
        if (isNaN(num) || num <= 0) {
          badQtys.push({ row: sheetRow, val: qtyVal });
        }
      }
    }

    // --- Check 4: Lifecycle state is recognized ---
    if (lifecycleColData >= 0 && lifecycleColData < row.length && validLifecycleStates) {
      const lifeVal = row[lifecycleColData];
      if (lifeVal && lifeVal.toString().trim() !== '') {
        const normalized = typeof normalizeLifecycleState_ === 'function'
          ? normalizeLifecycleState_(lifeVal.toString().trim())
          : lifeVal.toString().trim().toUpperCase();
        if (!validLifecycleStates.has(normalized)) {
          badLifecycles.push({ row: sheetRow, val: lifeVal });
        }
      }
    }
  }

  // Compile issues
  if (orphanPNs.length > 0) {
    const sample = orphanPNs.slice(0, 5).map(o => `Row ${o.row}: ${o.pn}`).join(', ');
    const extra = orphanPNs.length > 5 ? ` (+${orphanPNs.length - 5} more)` : '';
    issues.push(`ORPHAN PARTS (${orphanPNs.length}): Item Numbers not found in ITEMS sheet. ${sample}${extra}`);
  }

  if (levelGaps.length > 0) {
    const sample = levelGaps.slice(0, 5).map(g => `Row ${g.row}: ${g.from}→${g.to}`).join(', ');
    issues.push(`LEVEL GAPS (${levelGaps.length}): Hierarchy jumps > 1. ${sample}`);
  }

  if (badQtys.length > 0) {
    const sample = badQtys.slice(0, 5).map(q => `Row ${q.row}: "${q.val}"`).join(', ');
    issues.push(`INVALID QTY (${badQtys.length}): Non-positive quantities. ${sample}`);
  }

  if (badLifecycles.length > 0) {
    const sample = badLifecycles.slice(0, 5).map(l => `Row ${l.row}: "${l.val}"`).join(', ');
    issues.push(`UNRECOGNIZED LIFECYCLE (${badLifecycles.length}): Invalid states. ${sample}`);
  }

  return issues;
}
