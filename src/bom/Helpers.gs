// ============================================================================
// HELPERS.gs — Shared Utility Functions
// ============================================================================
// Low-level helpers used across multiple BOM modules: input validation,
// column indexing, hierarchy level calculation, parent-key tracking,
// and the AML row preparation tool.
// ============================================================================

/**
 * Prompts the user for input with validation. Returns trimmed text or null if cancelled.
 * @param {string} title Dialog title.
 * @param {string} message Prompt message.
 * @param {Object} [options] Validation options.
 * @param {number} [options.minLength=1] Minimum input length.
 * @param {number} [options.maxLength=500] Maximum input length.
 * @param {RegExp} [options.pattern] Regex pattern the input must match.
 * @param {string} [options.patternHint] Hint shown if pattern fails (e.g., "Use alphanumeric characters").
 * @returns {string|null} Validated input or null if cancelled.
 */
function promptWithValidation(title, message, options) {
  const ui = SpreadsheetApp.getUi();
  const opts = Object.assign({ minLength: 1, maxLength: 500 }, options || {});

  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const text = response.getResponseText().trim();
  if (text.length < opts.minLength) {
    ui.alert('Input Required', `Please enter at least ${opts.minLength} character(s).`, ui.ButtonSet.OK);
    return null;
  }
  if (text.length > opts.maxLength) {
    ui.alert('Input Too Long', `Input must be ${opts.maxLength} characters or less.`, ui.ButtonSet.OK);
    return null;
  }
  if (opts.pattern && !opts.pattern.test(text)) {
    ui.alert('Invalid Input', opts.patternHint || 'Input does not match the expected format.', ui.ButtonSet.OK);
    return null;
  }
  return text;
}

/**
 * Helper function to add all parent keys of a given locationKey to a Set.
 * e.g., "A/B/C" will add "A/B" and "A" to the set.
 * @param {string} locationKey The full location key of the changed item.
 * @param {Set<string>} affectedParentKeysSet The Set to add parent keys to.
 */
function addParentKeys(locationKey, affectedParentKeysSet) {
  if (!locationKey || !locationKey.includes('/')) {
    return; // This is a top-level item (or invalid), no parents to add.
  }

  const parts = locationKey.split('/');

  // Loop from the immediate parent up to the top level
  for (let i = parts.length - 1; i > 0; i--) {
    const parentKey = parts.slice(0, i).join('/');
    affectedParentKeysSet.add(parentKey);
  }
}

/**
 * Gets column indexes for all known BOM column names from headers.
 * Uses the explicit COLUMN_NAMES list instead of iterating all config properties.
 * @param {string[]} headers An array of header names from the sheet.
 * @returns {Object} An object mapping column name strings to their 0-based index (or -1 if not found).
 */
function getColumnIndexes(headers) {
  const indexes = {};

  if (!headers || headers.length === 0) {
    Logger.log("Warning: getColumnIndexes received empty or invalid headers.");
    for (const key in COL) {
      indexes[COL[key]] = -1;
    }
    return indexes;
  }

  // Map all known column names from the explicit COLUMN_NAMES object
  for (const key in COL) {
    indexes[COL[key]] = headers.indexOf(COL[key]);
  }

  // Also map any extra headers not in COLUMN_NAMES (e.g., dynamically added 'Change Impact')
  const knownValues = new Set(Object.values(COL));
  headers.forEach((header, index) => {
    if (header && !knownValues.has(header) && indexes[header] === undefined) {
      indexes[header] = index;
    }
  });

  return indexes;
}

/**
 * Helper: Calculates level depth.
 * Supports both Integer (1, 2, 3) and Dot-Notation (1.1, 1.1.1).
 */
function calculatePdmLevel(val) {
  if (val === "" || val === null) return 0;

  // Check if it is a simple integer (numeric type or string like "3")
  if (!isNaN(val) && !val.toString().includes('.')) {
    return parseInt(val, 10);
  }

  // Fallback: Count the dots for strings like "2.1.1"
  // "2.1" (1 dot) -> Level 2
  // "2.1.1" (2 dots) -> Level 3
  return (val.toString().match(/\./g) || []).length + 1;
}


// ============================================================================
// AML Row Preparation Tool
// ============================================================================

function runPrepareAMLRows() {
  const partInput = promptWithValidation('Prepare AML Rows', 'Enter Item Numbers, separated by commas:');
  if (!partInput) return;
  const partNumbers = partInput.split(',').map(pn => pn.trim()).filter(pn => pn);
  if (partNumbers.length > 0) {
    prepareAMLRows(partNumbers);
  } else {
    SpreadsheetApp.getUi().alert('No part numbers were entered.');
  }
}

function prepareAMLRows(partNumbers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const amlSheet = ss.getSheetByName(BOM_CONFIG.AML_SHEET_NAME);
  const activeSheet = ss.getActiveSheet();
  if (!amlSheet) {
    ui.alert(`Error: The master AML sheet named "${BOM_CONFIG.AML_SHEET_NAME}" could not be found.`);
    return;
  }
  if (!activeSheet || activeSheet.getName() === BOM_CONFIG.AML_SHEET_NAME || activeSheet.getName() === BOM_CONFIG.ITEMS_SHEET_NAME) {
    ui.alert('This function must be run on a valid BOM sheet, not on master data sheets.');
    return;
  }
  const amlData = amlSheet.getRange("A:A").getValues().flat().map(String);
  const bomData = activeSheet.getDataRange().getValues();
  const headers = bomData.length > 0 ?
    bomData[0] : [];
  const itemNumColIndex = headers.indexOf(COL.ITEM_NUMBER);
  if (itemNumColIndex === -1) {
    ui.alert(`Error: Could not find the column "${COL.ITEM_NUMBER}" in the active sheet.`);
    return;
  }
  let tasks = [];
  let summaryLog = [];
  const processedParts = new Set();
  partNumbers.forEach(partNumber => {
    const amlCount = amlData.filter(item => item.trim() === partNumber).length;
    if (amlCount <= 1) {
      if (!processedParts.has(partNumber)) {
        summaryLog.push(`- ${partNumber}: Skipped (has ${amlCount} AML entries).`);
        processedParts.add(partNumber);
      }
      return;
    }
    let instancesFound = 0;
    for (let i = 1; i < bomData.length; i++) {
      if (bomData[i][itemNumColIndex] && bomData[i][itemNumColIndex].toString().trim() === partNumber) {

        const foundRow = i + 1;
        tasks.push({
          partNumber: partNumber,
          row: foundRow,
          rowsToInsert: amlCount - 1
        });
        instancesFound++;
      }
    }
    if (instancesFound === 0) {
      if (!processedParts.has(partNumber)) {

        summaryLog.push(`- ${partNumber}: Not found in this sheet.`);
        processedParts.add(partNumber);
      }
    }
  });
  tasks.sort((a, b) => b.row - a.row);
  let successfulInsertions = {};
  tasks.forEach(task => {
    activeSheet.insertRowsAfter(task.row, task.rowsToInsert);
    if (!successfulInsertions[task.partNumber]) {
      successfulInsertions[task.partNumber] = {
        count: 0,
        rows: task.rowsToInsert
      };
    }
    successfulInsertions[task.partNumber].count++;
  });
  for (const partNum in successfulInsertions) {
    const info = successfulInsertions[partNum];
    summaryLog.push(`- ${partNum}: Success! Inserted ${info.rows} row(s) for ${info.count} instance(s).`);
  }
  if (summaryLog.length > 0) {
    ui.alert('Preparation Complete!', summaryLog.join('\n'), ui.ButtonSet.OK);
  } else {
    ui.alert('No actions were performed for the entered part numbers.');
  }
}
