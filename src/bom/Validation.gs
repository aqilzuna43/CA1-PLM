// ============================================================================
// VALIDATION.gs — Real-Time BOM Validation & Data Integrity
// ============================================================================
// Layer 1: onEdit simple trigger — instant cell-level validation on MASTER
// Layer 2: onChange installable trigger — structural watchdog (row ins/del)
//
// Shared lookup builders (used by Validation, Reconcile, and Audit):
//   buildItemsLookup_(ss) → Map<PN, {desc, rev, lifecycle}>
//   buildAmlLookup_(ss)   → Map<PN, [{mfr, mpn}]>
// ============================================================================


// ========================
// SHARED LOOKUP BUILDERS
// ========================

/**
 * Builds a lookup map from the ITEMS sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Spreadsheet reference.
 * @returns {Map<string, {desc: string, rev: string, lifecycle: string}>}
 */
function buildItemsLookup_(ss) {
  const map = new Map();
  const sheet = ss.getSheetByName(BOM_CONFIG.ITEMS_SHEET_NAME);
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;

  // ITEMS layout: A=Item Number, B=Description, C=Revision, D=Lifecycle
  // Find columns by header to be resilient to reordering
  const headers = data[0];
  const pnIdx = headers.indexOf('Item Number');
  const descIdx = headers.indexOf('Part Description');
  const revIdx = headers.indexOf('Item Rev');
  const lifeIdx = headers.indexOf('Lifecycle');

  // Fallback to positional if headers don't match
  const pn = pnIdx !== -1 ? pnIdx : 0;
  const desc = descIdx !== -1 ? descIdx : 1;
  const rev = revIdx !== -1 ? revIdx : 2;
  const life = lifeIdx !== -1 ? lifeIdx : 3;

  for (let i = 1; i < data.length; i++) {
    const itemNumber = data[i][pn];
    if (!itemNumber || itemNumber.toString().trim() === '') continue;
    map.set(itemNumber.toString().trim(), {
      desc: data[i][desc] ? data[i][desc].toString() : '',
      rev: data[i][rev] ? data[i][rev].toString() : '',
      lifecycle: data[i][life] ? data[i][life].toString() : ''
    });
  }
  return map;
}

/**
 * Builds a lookup map from the AML sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Spreadsheet reference.
 * @returns {Map<string, Array<{mfr: string, mpn: string}>>}
 */
function buildAmlLookup_(ss) {
  const map = new Map();
  const sheet = ss.getSheetByName(BOM_CONFIG.AML_SHEET_NAME);
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;

  // AML layout: A=Item Number, B=Mfr. Name, C=Mfr. Part Number
  const headers = data[0];
  const pnIdx = headers.indexOf('Item Number');
  const mfrIdx = headers.indexOf('Mfr. Name');
  const mpnIdx = headers.indexOf('Mfr. Part Number');

  const pn = pnIdx !== -1 ? pnIdx : 0;
  const mfr = mfrIdx !== -1 ? mfrIdx : 1;
  const mpn = mpnIdx !== -1 ? mpnIdx : 2;

  for (let i = 1; i < data.length; i++) {
    const itemNumber = data[i][pn];
    if (!itemNumber || itemNumber.toString().trim() === '') continue;
    const key = itemNumber.toString().trim();
    if (!map.has(key)) map.set(key, []);
    map.get(key).push({
      mfr: data[i][mfr] ? data[i][mfr].toString() : '',
      mpn: data[i][mpn] ? data[i][mpn].toString() : ''
    });
  }
  return map;
}


// ========================
// LAYER 1: onEdit TRIGGER
// ========================

/**
 * Simple trigger — fires on every user edit in the spreadsheet.
 * Guards for MASTER sheet only, then branches by edited column.
 *
 * NOTE: This is a simple trigger so it can read any sheet in the same
 * spreadsheet (ITEMS, AML) without additional auth. It must complete
 * in <30 seconds. For paste operations, e.value is undefined — we read
 * directly from the sheet range.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Guard: skip header row
  const editedRow = e.range.getRow();
  if (editedRow <= 1) return;

  const ss = e.range.getSheet().getParent();

  // --- ITEMS sheet: Lifecycle state machine enforcement ---
  if (sheetName === BOM_CONFIG.ITEMS_SHEET_NAME) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const editedCol = e.range.getColumn();
    const editedHeader = headers[editedCol - 1];

    if (editedHeader === 'Lifecycle' && typeof handleItemsLifecycleEdit === 'function') {
      handleItemsLifecycleEdit(ss, sheet, editedRow, editedCol, e.oldValue, e.value);
    }
    return;
  }

  // --- MASTER sheet: existing validation logic ---
  if (sheetName !== BOM_CONFIG.MASTER_SHEET_NAME) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIdx = getColumnIndexes(headers);

  // Determine which column was edited (1-based → header name)
  const editedCol = e.range.getColumn();           // 1-based
  const editedHeader = headers[editedCol - 1];      // 0-based lookup

  // --- Multi-cell paste handling ---
  // For single-cell edits e.range is 1x1. For paste, it can span many cells.
  // We process only the first row of a paste to keep <30s budget.
  // Full paste reconciliation is handled by the batch Reconcile tool.
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();

  // Branch by edited column
  if (editedHeader === COL.ITEM_NUMBER) {
    // Item Number changed — auto-populate managed columns
    handleItemNumberChange_(ss, sheet, headers, colIdx, editedRow, numRows);
  } else if (isManagedColumn_(editedHeader)) {
    // User typed over a managed column — restore from source
    handleManagedColumnEdit_(ss, sheet, headers, colIdx, editedRow, editedCol, editedHeader);
  } else if (editedHeader === COL.LEVEL) {
    // Level changed — validate hierarchy gap
    validateLevelAtRow_(sheet, colIdx, editedRow);
  } else if (editedHeader === COL.QTY) {
    // Qty changed — validate positive number
    validateQtyAtRow_(sheet, colIdx, editedRow);
  }
}

/**
 * Checks if a header name is a script-managed column.
 */
function isManagedColumn_(headerName) {
  const allManaged = [
    ...BOM_CONFIG.MANAGED_COLUMNS.FROM_ITEMS,
    ...BOM_CONFIG.MANAGED_COLUMNS.FROM_AML
  ];
  return allManaged.includes(headerName);
}


// ========================
// HANDLER: Item Number Change
// ========================

/**
 * When an Item Number is entered or changed, auto-populate Description,
 * Rev, Lifecycle from ITEMS, and first AML entry as Mfr/MPN.
 * Writes plain values (not VLOOKUP formulas).
 */
function handleItemNumberChange_(ss, sheet, headers, colIdx, startRow, numRows) {
  const itemsMap = buildItemsLookup_(ss);
  const amlMap = buildAmlLookup_(ss);
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;

  const rowsToProcess = Math.min(numRows, 50); // Cap for performance

  for (let offset = 0; offset < rowsToProcess; offset++) {
    const row = startRow + offset;
    const itemNumber = sheet.getRange(row, colIdx[COL.ITEM_NUMBER] + 1).getValue();
    const pn = itemNumber ? itemNumber.toString().trim() : '';

    if (!pn) continue; // Empty cell — skip

    const itemData = itemsMap.get(pn);
    const amlEntries = amlMap.get(pn);

    // --- Populate from ITEMS ---
    if (itemData) {
      setManagedValue_(sheet, row, colIdx, COL.DESCRIPTION, itemData.desc, prefix);
      setManagedValue_(sheet, row, colIdx, COL.ITEM_REV, itemData.rev, prefix);
      setManagedValue_(sheet, row, colIdx, COL.LIFECYCLE, itemData.lifecycle, prefix);
    } else {
      // PN not found in ITEMS — flag as orphan
      const pnCol = colIdx[COL.ITEM_NUMBER] + 1;
      const pnCell = sheet.getRange(row, pnCol);
      pnCell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
      pnCell.setNote(prefix + 'Item Number not found in ITEMS sheet.');
    }

    // --- Populate from AML (first entry) ---
    if (amlEntries && amlEntries.length > 0) {
      setManagedValue_(sheet, row, colIdx, COL.MFR_NAME, amlEntries[0].mfr, prefix);
      setManagedValue_(sheet, row, colIdx, COL.MFR_PN, amlEntries[0].mpn, prefix);
    } else if (itemData) {
      // PN exists in ITEMS but has no AML — warning
      if (colIdx[COL.MFR_NAME] !== -1) {
        const mfrCell = sheet.getRange(row, colIdx[COL.MFR_NAME] + 1);
        mfrCell.setBackground(BOM_CONFIG.VALIDATION.COLORS.WARNING);
        mfrCell.setNote(prefix + 'No AML entry found for this Item Number.');
      }
    }

    // --- Validate Level (if present) ---
    if (colIdx[COL.LEVEL] !== -1) {
      validateLevelAtRow_(sheet, colIdx, row);
    }
  }
}


// ========================
// HANDLER: Managed Column Overwrite
// ========================

/**
 * When a user manually edits a managed column (Desc/Rev/Lifecycle/Mfr/MPN),
 * restore the correct value from the source sheet and provide visual feedback.
 */
function handleManagedColumnEdit_(ss, sheet, headers, colIdx, row, editedCol, editedHeader) {
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;

  // Get the Item Number for this row to look up correct value
  if (colIdx[COL.ITEM_NUMBER] === -1) return;
  const itemNumber = sheet.getRange(row, colIdx[COL.ITEM_NUMBER] + 1).getValue();
  const pn = itemNumber ? itemNumber.toString().trim() : '';
  if (!pn) return; // No PN on this row — nothing to restore

  let correctValue = null;
  let sourceName = '';

  if (BOM_CONFIG.MANAGED_COLUMNS.FROM_ITEMS.includes(editedHeader)) {
    const itemsMap = buildItemsLookup_(ss);
    const itemData = itemsMap.get(pn);
    if (itemData) {
      if (editedHeader === COL.DESCRIPTION) { correctValue = itemData.desc; }
      else if (editedHeader === COL.ITEM_REV) { correctValue = itemData.rev; }
      else if (editedHeader === COL.LIFECYCLE) { correctValue = itemData.lifecycle; }
      sourceName = BOM_CONFIG.ITEMS_SHEET_NAME;
    }
  } else if (BOM_CONFIG.MANAGED_COLUMNS.FROM_AML.includes(editedHeader)) {
    const amlMap = buildAmlLookup_(ss);
    const amlEntries = amlMap.get(pn);
    if (amlEntries && amlEntries.length > 0) {
      if (editedHeader === COL.MFR_NAME) { correctValue = amlEntries[0].mfr; }
      else if (editedHeader === COL.MFR_PN) { correctValue = amlEntries[0].mpn; }
      sourceName = BOM_CONFIG.AML_SHEET_NAME;
    }
  }

  if (correctValue !== null) {
    const cell = sheet.getRange(row, editedCol);
    const currentValue = cell.getValue();

    // Only restore if value actually differs from source
    if (currentValue !== null && currentValue.toString().trim() !== correctValue.trim()) {
      cell.setValue(correctValue);
      cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.RESTORED);
      cell.setNote(prefix + `Restored from ${sourceName}. This column is auto-managed.`);
    } else {
      // User set it to the correct value — clear any old warnings
      clearValidationNote_(cell, prefix);
      cell.setBackground(null);
    }
  }
}


// ========================
// HANDLER: Level Validation
// ========================

/**
 * Validates that the level at the given row doesn't create a hierarchy gap.
 * A gap occurs when the level jumps by more than 1 compared to the previous row.
 * E.g., going from Level 1 directly to Level 3 is invalid.
 */
function validateLevelAtRow_(sheet, colIdx, row) {
  if (colIdx[COL.LEVEL] === -1) return;
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;
  const levelCol = colIdx[COL.LEVEL] + 1;
  const cell = sheet.getRange(row, levelCol);
  const currentLevel = parseInt(cell.getValue(), 10);

  if (isNaN(currentLevel) || currentLevel < 0) {
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
    cell.setNote(prefix + 'Level must be a non-negative integer.');
    return;
  }

  // Check against previous row (skip if this is the first data row)
  if (row <= 2) {
    clearValidationNote_(cell, prefix);
    cell.setBackground(null);
    return;
  }

  const prevLevel = parseInt(sheet.getRange(row - 1, levelCol).getValue(), 10);
  if (!isNaN(prevLevel) && currentLevel > prevLevel + 1) {
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
    cell.setNote(prefix + `Level gap detected: jumped from ${prevLevel} to ${currentLevel}. Max allowed is ${prevLevel + 1}.`);
  } else {
    clearValidationNote_(cell, prefix);
    cell.setBackground(null);
  }
}


// ========================
// HANDLER: Qty Validation
// ========================

/**
 * Validates that quantity is a positive number.
 */
function validateQtyAtRow_(sheet, colIdx, row) {
  if (colIdx[COL.QTY] === -1) return;
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;
  const qtyCol = colIdx[COL.QTY] + 1;
  const cell = sheet.getRange(row, qtyCol);
  const val = cell.getValue();

  // Allow empty qty (some rows like top-level assemblies may not have qty)
  if (val === '' || val === null) {
    clearValidationNote_(cell, prefix);
    cell.setBackground(null);
    return;
  }

  const num = Number(val);
  if (isNaN(num) || num <= 0) {
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.WARNING);
    cell.setNote(prefix + 'Quantity should be a positive number.');
  } else {
    clearValidationNote_(cell, prefix);
    cell.setBackground(null);
  }
}


// ========================
// UTILITY: Managed Value Writer
// ========================

/**
 * Writes a value to a managed column cell with "restored" styling.
 * Only writes if the column exists in the sheet.
 */
function setManagedValue_(sheet, row, colIdx, colName, value, notePrefix) {
  if (colIdx[colName] === -1) return;
  const col = colIdx[colName] + 1;
  const cell = sheet.getRange(row, col);
  cell.setValue(value);
  cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.RESTORED);
  cell.setNote(notePrefix + 'Auto-populated from master data.');

  // Auto-fade the green background after setting (visual confirmation only)
  // Note: We leave the green so the user sees it was auto-set. It clears
  // on next reconcile or when the user acknowledges.
}


// ========================
// UTILITY: Note Cleanup
// ========================

/**
 * Removes only BOM Validation notes from a cell (preserves user notes).
 */
function clearValidationNote_(cell, prefix) {
  const existingNote = cell.getNote();
  if (existingNote && existingNote.startsWith(prefix)) {
    cell.setNote('');
  }
}


// ========================
// LAYER 2: onChange TRIGGER (Installable)
// ========================

/**
 * Installable trigger — responds to structural changes (row insert/delete).
 * Must be installed via menu: "Data Integrity > Install Change Watchdog"
 * or by running installChangeTrigger_() manually.
 */
function onChangeWatchdog(e) {
  if (!e) return;
  const changeType = e.changeType;

  if (changeType === 'INSERT_ROW') {
    handleRowInsert_();
  } else if (changeType === 'REMOVE_ROW') {
    handleRowDelete_();
  }
  // Other change types (EDIT, INSERT_COLUMN, REMOVE_COLUMN) are handled by onEdit or ignored
}

/**
 * After a row is inserted on MASTER, find rows that have an Item Number
 * but empty managed columns and populate them.
 */
function handleRowInsert_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BOM_CONFIG.MASTER_SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const colIdx = getColumnIndexes(headers);
  if (colIdx[COL.ITEM_NUMBER] === -1) return;

  const itemsMap = buildItemsLookup_(ss);
  const amlMap = buildAmlLookup_(ss);
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;
  const managedFromItems = BOM_CONFIG.MANAGED_COLUMNS.FROM_ITEMS;

  for (let i = 1; i < data.length; i++) {
    const pn = data[i][colIdx[COL.ITEM_NUMBER]];
    if (!pn || pn.toString().trim() === '') continue;

    const pnStr = pn.toString().trim();
    const row = i + 1;

    // Check if Description (first managed col) is empty — indicates new/unfilled row
    const descIdx = colIdx[COL.DESCRIPTION];
    if (descIdx !== -1) {
      const descVal = data[i][descIdx];
      if (!descVal || descVal.toString().trim() === '') {
        // Row has PN but empty managed cols — populate
        const itemData = itemsMap.get(pnStr);
        const amlEntries = amlMap.get(pnStr);

        if (itemData) {
          setManagedValue_(sheet, row, colIdx, COL.DESCRIPTION, itemData.desc, prefix);
          setManagedValue_(sheet, row, colIdx, COL.ITEM_REV, itemData.rev, prefix);
          setManagedValue_(sheet, row, colIdx, COL.LIFECYCLE, itemData.lifecycle, prefix);
        }

        if (amlEntries && amlEntries.length > 0) {
          setManagedValue_(sheet, row, colIdx, COL.MFR_NAME, amlEntries[0].mfr, prefix);
          setManagedValue_(sheet, row, colIdx, COL.MFR_PN, amlEntries[0].mpn, prefix);
        }
      }
    }
  }
}

/**
 * After rows are deleted from MASTER, scan for:
 *   1. Level hierarchy gaps
 *   2. AML row count mismatches (missing AML continuation rows)
 *
 * Performance: Uses batch reads (getBackgrounds/getNotes) and only writes
 * to cells that actually need changes. Avoids per-cell getRange() in loops.
 */
function handleRowDelete_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BOM_CONFIG.MASTER_SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const colIdx = getColumnIndexes(headers);
  if (colIdx[COL.LEVEL] === -1) return;

  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;
  const levelCol = colIdx[COL.LEVEL] + 1;

  // --- Pass 1: Level hierarchy gap detection ---
  // Batch-read backgrounds and notes for the Level column to avoid per-cell API calls
  const lastRow = data.length;
  const levelRange = sheet.getRange(2, levelCol, lastRow - 1, 1); // All data rows for Level col
  const levelBGs = levelRange.getBackgrounds();
  const levelNotes = levelRange.getNotes();

  for (let i = 2; i < data.length; i++) {
    const currentLevel = parseInt(data[i][colIdx[COL.LEVEL]], 10);
    const prevLevel = parseInt(data[i - 1][colIdx[COL.LEVEL]], 10);

    if (isNaN(currentLevel) || isNaN(prevLevel)) continue;

    const row = i + 1;
    const bgIdx = i - 1; // Index into levelBGs/levelNotes (0-based from row 2)

    if (currentLevel > prevLevel + 1) {
      // Flag the gap
      const cell = sheet.getRange(row, levelCol);
      cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
      cell.setNote(prefix + `Level gap: jumped from ${prevLevel} to ${currentLevel} (row deletion may have broken hierarchy).`);
    } else {
      // Only clear if this cell currently has a validation warning (avoid unnecessary writes)
      const existingNote = levelNotes[bgIdx][0];
      if (existingNote && existingNote.startsWith(prefix)) {
        const cell = sheet.getRange(row, levelCol);
        cell.setNote('');
        cell.setBackground(null);
      }
    }
  }

  // --- Pass 2: AML row count mismatch detection ---
  const hasMfrCol = colIdx[COL.MFR_NAME] !== -1 || colIdx[COL.MFR_PN] !== -1;
  if (!hasMfrCol || colIdx[COL.ITEM_NUMBER] === -1) return;

  const amlMap = buildAmlLookup_(ss);
  if (amlMap.size === 0) return;

  // First pass: collect all mismatches in memory (no API calls)
  const mismatches = []; // [{mainRow, pn, actual, expected}]

  let currentPN = '';
  let currentMainRow = -1;
  let actualAmlCount = 0;

  for (let i = 1; i < data.length; i++) {
    const itemNumber = data[i][colIdx[COL.ITEM_NUMBER]];
    const hasPN = itemNumber && itemNumber.toString().trim() !== '';

    if (hasPN) {
      // Flush previous part
      if (currentPN && currentMainRow > 0) {
        const expected = amlMap.get(currentPN);
        if (expected && expected.length > 1 && actualAmlCount < expected.length) {
          mismatches.push({ mainRow: currentMainRow, pn: currentPN, actual: actualAmlCount, expected: expected.length });
        }
      }

      currentPN = itemNumber.toString().trim();
      currentMainRow = i + 1;
      actualAmlCount = 0;

      const mfrName = colIdx[COL.MFR_NAME] !== -1 ? data[i][colIdx[COL.MFR_NAME]] : null;
      const mfrPN = colIdx[COL.MFR_PN] !== -1 ? data[i][colIdx[COL.MFR_PN]] : null;
      if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
        actualAmlCount = 1;
      }
    } else if (currentPN) {
      const mfrName = colIdx[COL.MFR_NAME] !== -1 ? data[i][colIdx[COL.MFR_NAME]] : null;
      const mfrPN = colIdx[COL.MFR_PN] !== -1 ? data[i][colIdx[COL.MFR_PN]] : null;
      if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
        actualAmlCount++;
      } else {
        if (currentPN && currentMainRow > 0) {
          const expected = amlMap.get(currentPN);
          if (expected && expected.length > 1 && actualAmlCount < expected.length) {
            mismatches.push({ mainRow: currentMainRow, pn: currentPN, actual: actualAmlCount, expected: expected.length });
          }
        }
        currentPN = '';
        currentMainRow = -1;
        actualAmlCount = 0;
      }
    }
  }

  // Flush last part
  if (currentPN && currentMainRow > 0) {
    const expected = amlMap.get(currentPN);
    if (expected && expected.length > 1 && actualAmlCount < expected.length) {
      mismatches.push({ mainRow: currentMainRow, pn: currentPN, actual: actualAmlCount, expected: expected.length });
    }
  }

  // Second pass: write only to mismatched cells (minimal API calls)
  const targetCol = colIdx[COL.MFR_NAME] !== -1 ? colIdx[COL.MFR_NAME] + 1 : colIdx[COL.MFR_PN] + 1;
  mismatches.forEach(m => {
    const cell = sheet.getRange(m.mainRow, targetCol);
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
    cell.setNote(prefix + `AML row mismatch: expected ${m.expected} AML entries but found ${m.actual} in BOM. A row may have been accidentally deleted.`);
  });
}


// ========================
// TRIGGER INSTALLER
// ========================

/**
 * Idempotent installer for the onChange trigger.
 * Safe to call multiple times — removes existing watchdog trigger first.
 */
function installChangeTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Remove any existing watchdog trigger
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onChangeWatchdog') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  // Install new trigger
  ScriptApp.newTrigger('onChangeWatchdog')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  ui.alert('Change Watchdog Installed',
    (removed > 0 ? `Replaced ${removed} existing trigger(s). ` : '') +
    'The onChange watchdog is now active.\n\n' +
    'It will auto-populate managed columns when rows are inserted ' +
    'and detect level hierarchy gaps when rows are deleted.',
    ui.ButtonSet.OK);
}
