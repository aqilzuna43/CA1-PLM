// ============================================================================
// BOMMAP.gs — Core BOM Map Building Logic
// ============================================================================
// Shared data-structure builders that convert raw sheet data into structured
// BOM Maps keyed by hierarchy path. Used by Comparison, Fabricator, and
// Audit modules.
// ============================================================================

/**
 * Core BOM Map building logic.
 * Creates a Map of BOM items, keyed by their structural location.
 * @param {string} sheetName The name of the sheet being processed (for error logging).
 * @param {Array<Array<string>>} data The 2D array of data from the sheet.
 * @param {Object} colIndexes An object mapping CONFIG keys (e.g., COL.LEVEL) to their column index.
 * @param {number} startRow The 0-based row index to start parsing from (e.g., 1 for header-less data).
 * @param {number} baseLevel The level to treat as "Level 0" (for subassembly extraction).
 * @returns {Map<string, object>} A Map of BOM item data.
 */
function buildBOMMap(sheetName, data, colIndexes, startRow, baseLevel) {
  const bomMap = new Map();
  if (data.length <= startRow) return bomMap; // No data to parse

  // Check for essential columns in the provided colIndexes
  const essentialCols = [
    COL.LEVEL,
    COL.ITEM_NUMBER,
    COL.DESCRIPTION,
    COL.ITEM_REV,
    COL.QTY,
    COL.MFR_NAME,
    COL.MFR_PN
  ];

  // Lifecycle is only essential if its index is not -1 (i.e., it's expected)
  if (colIndexes[COL.LIFECYCLE] !== -1) {
    essentialCols.push(COL.LIFECYCLE);
  }

  for (const colName of essentialCols) {
    if (colIndexes[colName] === -1) {
       // This check is now robust; it only fails if a *mapped* column is missing.
       throw new Error(`Required column for "${colName}" was not found in sheet "${sheetName}". Please check CONFIG or sheet headers.`);
    }
  }

  let pathStack = [];
  let currentItemData = null;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    const itemNumber = row[colIndexes[COL.ITEM_NUMBER]];
    const isMainPartRow = itemNumber && itemNumber.toString().trim() !== '';
    const levelVal = row[colIndexes[COL.LEVEL]];

    // Calculate normalized level using calculatePdmLevel() for consistency
    // Supports both integer levels (1, 2, 3) and dot-notation (1.1, 1.1.1)
    // Use a large negative number if level is blank to skip AML rows
    const level = (levelVal !== '' && levelVal !== null && !isNaN(parseFloat(levelVal)))
      ? (calculatePdmLevel(levelVal) - baseLevel)
      : -999;

    if (isMainPartRow) {
       // Stop processing if we've returned to a level at or above the base level
       if (level < 0) { // This handles "uncle" rows (e.g., level 1 when base was 2)
         if (currentItemData) {
           bomMap.set(currentItemData.locationKey, currentItemData);
         }
         break; // We have exited the subassembly
       }

       // This handles "sibling" rows
       if (level === 0 && i > startRow) { // i > startRow ensures we don't stop on the first item
         if (currentItemData) {
           bomMap.set(currentItemData.locationKey, currentItemData);
         }
         break; // We have hit a sibling assembly
       }

       if (currentItemData) {
           bomMap.set(currentItemData.locationKey, currentItemData);
       }

       pathStack.length = level;
       pathStack[level] = itemNumber;
       const locationKey = pathStack.join('/');
       const parent = level > 0 && pathStack[level - 1] ? pathStack[level - 1].toString().trim() : 'Top Level';

       currentItemData = {
           startRow: i + 1, // 1-based for Apps Script ranges
           endRow: i + 1,   // 1-based for Apps Script ranges
           parent: parent,
           locationKey: locationKey,
           colIndexes: colIndexes, // Pass normalized indexes
           mainRow: {
               [COL.ITEM_NUMBER]: itemNumber.toString().trim(),
               [COL.DESCRIPTION]: row[colIndexes[COL.DESCRIPTION]].toString().trim(),
               [COL.ITEM_REV]: row[colIndexes[COL.ITEM_REV]].toString().trim(),
               [COL.QTY]: row[colIndexes[COL.QTY]].toString().trim(),
               // Handle potentially missing lifecycle column
               [COL.LIFECYCLE]: colIndexes[COL.LIFECYCLE] !== -1 ? row[colIndexes[COL.LIFECYCLE]].toString().trim() : 'N/A',
           },
           aml: []
       };
    }

    if (currentItemData) {
       currentItemData.endRow = i + 1; // 1-based
       const mfrName = row[colIndexes[COL.MFR_NAME]];
       const mfrPN = row[colIndexes[COL.MFR_PN]];

       if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
           currentItemData.aml.push({
               [COL.MFR_NAME]: mfrName ? mfrName.toString().trim() : "",
               [COL.MFR_PN]: mfrPN ? mfrPN.toString().trim() : ""
           });
       }
    }
  }
   if (currentItemData) {
       bomMap.set(currentItemData.locationKey, currentItemData);
   }

  return bomMap;
}

/**
 * Wrapper for buildBOMMap for standard, full-sheet comparison.
 */
function createBOMMap(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Map();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers); // Gets standard CONFIG indexes
  // Start parsing at row 2 (index 1), treat level 0 as base
  return buildBOMMap(sheet.getName(), data, colIndexes, 1, 0);
}

/**
 * Creates a BOM Map for a specific subassembly from the Master BOM sheet.
 */
function createMasterSubassemblyMap(sheet, assemblyNumber) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Map();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers); // Gets standard CONFIG indexes

  const itemNumCol = colIndexes[COL.ITEM_NUMBER];
  const levelCol = colIndexes[COL.LEVEL];
  if (itemNumCol === -1 || levelCol === -1) {
    throw new Error(`Could not find Item Number or Level column in Master BOM.`);
  }

  let startRow = -1;
  let baseLevel = -1;

  // Find the assembly to start from
  for (let i = 1; i < data.length; i++) {
    if (data[i][itemNumCol] && data[i][itemNumCol].toString().trim() === assemblyNumber) {
      startRow = i; // 0-based index for buildBOMMap
      baseLevel = parseInt(data[i][levelCol], 10);
      if (isNaN(baseLevel)) {
         throw new Error(`Assembly "${assemblyNumber}" found, but its Level is not a valid number.`);
      }
      break;
    }
  }

  if (startRow === -1) {
    throw new Error(`Assembly "${assemblyNumber}" not found in sheet "${sheet.getName()}".`);
  }

  // Start parsing AT the found row (startRow), and subtract its level (baseLevel)
  return buildBOMMap(sheet.getName(), data, colIndexes, startRow, baseLevel);
}

/**
 * Creates a BOM Map from an external sheet using a provided header map.
 * Maps PDM header names to standard COL names so buildBOMMap can process them uniformly.
 * Assumes the external sheet starts at Level 0 for the assembly.
 */
function createExternalBOMMap(sheet, headerMap) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Map();
  const headers = data[0];

  // Build normalized column indexes: map each COL key to the PDM header's column index
  // headerMap keys match COL keys (LEVEL, ITEM_NUMBER, DESCRIPTION, etc.)
  const normalizedColIndexes = {};
  const missingHeaders = [];

  for (const colKey in COL) {
    const pdmHeaderName = headerMap[colKey];
    if (pdmHeaderName) {
      // This COL key has a PDM mapping — look it up in the actual headers
      const idx = headers.indexOf(pdmHeaderName);
      normalizedColIndexes[COL[colKey]] = idx;
      if (idx === -1) {
        missingHeaders.push(`"${pdmHeaderName}" (for ${COL[colKey]})`);
      }
    } else {
      // No PDM mapping for this column (e.g., LIFECYCLE, REFERENCE_NOTES) — mark as absent
      normalizedColIndexes[COL[colKey]] = -1;
    }
  }

  if (missingHeaders.length > 0) {
    throw new Error(`Missing headers in PDM sheet "${sheet.getName()}":\n${missingHeaders.join('\n')}`);
  }

  // Start parsing at row 2 (index 1), treat level 0 as base
  return buildBOMMap(sheet.getName(), data, normalizedColIndexes, 1, 0);
}
