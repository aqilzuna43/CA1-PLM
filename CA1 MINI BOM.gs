// --- CONFIGURATION ---
// !!!!! IMPORTANT: Adjust these names to exactly match your sheet headers and sheet names !!!!!
const CONFIG = {
  // Master Data Sheet Names
  ITEMS_SHEET_NAME: 'ITEMS',
  AML_SHEET_NAME: 'AML',

  // ECR/ECO Linking Sheet Name
  ECR_AFFECTED_ITEMS_SHEET: 'ECR_Affected_Items', // Name of your ECR details sheet

  // Column names used by various tools
  LEVEL_COL_NAME: 'Level',
  ITEM_NUM_COL_NAME: 'Item Number',
  DESC_COL_NAME: 'Part Description',
  ITEM_REV_COL_NAME: 'Item Rev',
  QTY_COL_NAME: 'Qty',
  LIFECYCLE_COL_NAME: 'Lifecycle',
  MFR_NAME_COL_NAME: 'Mfr. Name',
  MFR_PN_COL_NAME: 'Mfr. Part Number',
  REFERENCE_NOTES_COL_NAME: 'Reference Notes', // For MAKE/BUY/REF status

  // [NEW] Header mapping for the "Grafting/Import" feature (Matches your PDM Screenshot)
  PDM_GRAFT_SHEET_NAME: 'INPUT_PDM', // The tab name where you paste PDM data
  PDM_GRAFT_HEADERS: {
    HIERARCHY_COL: 'LEVEL',      // [UPDATED] Matches your screenshot (was ITEM NO.)
    PN_COL: 'PART NUMBER',
    REV_COL: 'REV',
    DESC_COL: 'DESCRIPTION',
    VENDOR_COL: 'VENDOR',
    MPN_COL: 'MPN',
    QTY_COL: 'QTY.'
  },

  // Header mapping for external PDM Comparison (existing feature)
  PDM_HEADER_MAP: {
    LEVEL_COL_NAME: 'Level',
    ITEM_NUM_COL_NAME: 'NR',
    DESC_COL_NAME: 'BENENNUNG',
    ITEM_REV_COL_NAME: 'Revision',
    QTY_COL_NAME: 'Qty',
    MFR_NAME_COL_NAME: 'Vendor',
    MFR_PN_COL_NAME: 'MPN'
  },

  // Column names for the 'Finalize and Release' process
  CHANGE_TRACKING_COLS_TO_DELETE: ['ECR #', 'Status', 'Change Impact'],

  // Sheet name for ECO Logging
  ECO_LOG_SHEET_NAME: 'ECO History'
};
// ---------------------


/**
 * Creates the custom "BOM Tools" menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BOM Tools')
    .addItem('1. Generate Detailed Comparison', 'runDetailedComparison')
    .addItem('2. Compare Master vs. PDM BOM', 'runCompareWithExternalBOM')
    .addItem('3. Import Children from PDM (Graft)', 'runImportPdmChildren') // [NEW]
    .addSeparator()
    .addItem('Generate Fabricator BOMs', 'runGenerateFabricatorBOMs')
    .addItem(`List 'BUY' Items with 'REF' Children`, 'runScreenBuyItems')
    .addItem('Audit BOM Lifecycle Status', 'runAuditBOMLifecycle')
    .addSeparator()
    .addItem('Prepare Rows for AML', 'runPrepareAMLRows')
    .addItem('Generate Master Lists from BOM', 'runGenerateMasterLists')
    .addItem('Where-Used Analysis', 'runWhereUsedAnalysis')
    .addSeparator()
    .addItem('Finalize and Release New BOM', 'runReleaseNewBOM')
    .addToUi();
}


//====================================================================================================
// === [NEW] IMPORT PDM CHILDREN (GRAFTING) ==========================================================
//====================================================================================================

/**
 * Imports children from the PDM sheet to the Master BOM, adjusting levels automatically.
 */
function runImportPdmChildren() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getActiveSheet(); // User should run this from the Master BOM
  const pdmSheet = ss.getSheetByName(CONFIG.PDM_GRAFT_SHEET_NAME);

  if (!pdmSheet) {
    ui.alert(`Error: Could not find the PDM input sheet named "${CONFIG.PDM_GRAFT_SHEET_NAME}". Please create it and paste your PDM export there.`);
    return;
  }

  // 1. Get Selected Parent in Master BOM
  const activeRange = masterSheet.getActiveRange();
  const activeRowIndex = activeRange.getRow(); // 1-based
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData[0];
  
  // Map Master Columns
  const mCols = getColumnIndexes(masterHeaders);

  // [FIX] Safety check for missing columns in Master BOM
  if (mCols[CONFIG.ITEM_NUM_COL_NAME] === -1 || mCols[CONFIG.LEVEL_COL_NAME] === -1) {
    ui.alert(`Error: Critical columns missing in Master Sheet.\n\nCould not find columns named:\n- "${CONFIG.ITEM_NUM_COL_NAME}"\n- "${CONFIG.LEVEL_COL_NAME}"\n\nPlease check the CONFIG section in the script and your sheet headers.`);
    return;
  }

  // Validate Master Selection
  if (activeRowIndex < 2 || activeRowIndex > masterData.length) {
    ui.alert("Please select a valid row containing a Parent Item.");
    return;
  }

  const parentRowData = masterData[activeRowIndex - 1]; // 0-based
  
  // [FIX] Safe string conversion to prevent "reading 'toString' of undefined"
  const masterParentPNVal = parentRowData[mCols[CONFIG.ITEM_NUM_COL_NAME]];
  const masterParentLevelVal = parentRowData[mCols[CONFIG.LEVEL_COL_NAME]];

  const masterParentPN = masterParentPNVal != null ? String(masterParentPNVal).trim() : "";
  const masterParentLevelStr = masterParentLevelVal != null ? String(masterParentLevelVal).trim() : "";
  const masterParentLevel = parseInt(masterParentLevelStr, 10);

  if (masterParentPN === "" || isNaN(masterParentLevel)) {
    ui.alert(`Error: The selected row must have a valid Part Number and numeric Level.\n\nFound PN: "${masterParentPN}"\nFound Level: "${masterParentLevelStr}"`);
    return;
  }

  // 2. Get PDM Data
  const pdmData = pdmSheet.getDataRange().getValues();
  if (pdmData.length < 2) { ui.alert("PDM Sheet is empty."); return; }
  const pdmHeaders = pdmData[0];

  // Map PDM Columns (Dynamic search based on CONFIG.PDM_GRAFT_HEADERS)
  const pCols = {};
  const missingPdmCols = [];
  for (const key in CONFIG.PDM_GRAFT_HEADERS) {
    pCols[key] = pdmHeaders.indexOf(CONFIG.PDM_GRAFT_HEADERS[key]);
    if (pCols[key] === -1) {
      missingPdmCols.push(CONFIG.PDM_GRAFT_HEADERS[key]);
    }
  }

  if (missingPdmCols.length > 0) {
    ui.alert(`Error: Missing columns in PDM Sheet "${CONFIG.PDM_GRAFT_SHEET_NAME}".\n\nCould not find: ${missingPdmCols.join(", ")}`);
    return;
  }

  // 3. Find Parent in PDM
  let pdmStartIndex = -1;
  let pdmParentLevel = -1;

  for (let i = 1; i < pdmData.length; i++) {
    const rowPN = String(pdmData[i][pCols.PN_COL]).trim();
    if (rowPN === masterParentPN) {
      pdmStartIndex = i;
      // Calculate PDM Level (Handles both "1.2.1" and simple integers "3")
      const hierarchyVal = pdmData[i][pCols.HIERARCHY_COL];
      pdmParentLevel = calculatePdmLevel(hierarchyVal); 
      break;
    }
  }

  if (pdmStartIndex === -1) {
    ui.alert(`Part Number "${masterParentPN}" not found in ${CONFIG.PDM_GRAFT_SHEET_NAME}.`);
    return;
  }

  // 4. Collect Children & Transform Levels
  const rowsToInsert = [];
  
  // Loop through PDM rows starting immediately after the parent
  for (let j = pdmStartIndex + 1; j < pdmData.length; j++) {
    const currentRow = pdmData[j];
    const hierarchyVal = currentRow[pCols.HIERARCHY_COL];
    
    // Safety check: Empty hierarchy string usually means end of data
    if (hierarchyVal === "" || hierarchyVal === null) break;

    const currentPdmLevel = calculatePdmLevel(hierarchyVal);

    // STOP CONDITION: If we find a level equal to or higher (smaller number) than the parent, 
    // we have left the sub-assembly.
    if (currentPdmLevel <= pdmParentLevel) break;

    // LEVEL TRANSFORMATION logic
    // Formula: NewLevel = MasterParentLevel + (PdmChildLevel - PdmParentLevel)
    const relativeDepth = currentPdmLevel - pdmParentLevel;
    const newMasterLevel = masterParentLevel + relativeDepth;

    // Construct the new row for Master BOM
    // We create an empty array matching the Master BOM width
    const newRow = new Array(masterHeaders.length).fill("");

    // Fill mapped columns
    if (mCols[CONFIG.LEVEL_COL_NAME] > -1) newRow[mCols[CONFIG.LEVEL_COL_NAME]] = newMasterLevel;
    if (mCols[CONFIG.ITEM_NUM_COL_NAME] > -1) newRow[mCols[CONFIG.ITEM_NUM_COL_NAME]] = currentRow[pCols.PN_COL];
    if (mCols[CONFIG.ITEM_REV_COL_NAME] > -1) newRow[mCols[CONFIG.ITEM_REV_COL_NAME]] = currentRow[pCols.REV_COL];
    if (mCols[CONFIG.DESC_COL_NAME] > -1) newRow[mCols[CONFIG.DESC_COL_NAME]] = currentRow[pCols.DESC_COL];
    if (mCols[CONFIG.QTY_COL_NAME] > -1) newRow[mCols[CONFIG.QTY_COL_NAME]] = currentRow[pCols.QTY_COL];
    if (mCols[CONFIG.MFR_NAME_COL_NAME] > -1) newRow[mCols[CONFIG.MFR_NAME_COL_NAME]] = currentRow[pCols.VENDOR_COL];
    if (mCols[CONFIG.MFR_PN_COL_NAME] > -1) newRow[mCols[CONFIG.MFR_PN_COL_NAME]] = currentRow[pCols.MPN_COL];
    
    // Set default status to "BUY" if it's a purchased part, or leave blank
    if (mCols[CONFIG.REFERENCE_NOTES_COL_NAME] > -1) newRow[mCols[CONFIG.REFERENCE_NOTES_COL_NAME]] = "Pending Review"; 

    rowsToInsert.push(newRow);
  }

  if (rowsToInsert.length === 0) {
    ui.alert("Found the parent in PDM, but it has no children listed below it.");
    return;
  }

  // 5. Insert into Master BOM
  masterSheet.insertRowsAfter(activeRowIndex, rowsToInsert.length);
  masterSheet.getRange(activeRowIndex + 1, 1, rowsToInsert.length, rowsToInsert[0].length)
             .setValues(rowsToInsert);

  // Optional: Highlight the new rows
  masterSheet.getRange(activeRowIndex + 1, 1, rowsToInsert.length, masterSheet.getLastColumn())
             .setBackground("#e6f7ff"); // Light blue highlight

  ui.alert(`Success! Grafted ${rowsToInsert.length} rows. \nNew levels adjusted relative to Level ${masterParentLevel}.`);
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


//====================================================================================================
// === CORE COMPARISON TOOL: DETAILED REPORT & HIGHLIGHTING (SIMPLIFIED ECR LINKING) ===============
//====================================================================================================
function runDetailedComparison() {
  const ui = SpreadsheetApp.getUi();
  const ecoNumberResponse = ui.prompt('Generate Comparison Report', 'Enter the base ECO Number for this report (e.g., ECO-12):', ui.ButtonSet.OK_CANCEL);
  if (ecoNumberResponse.getSelectedButton() !== ui.Button.OK || !ecoNumberResponse.getResponseText()) return;
  const ecoBase = ecoNumberResponse.getResponseText().trim();
  // --- Read ECR Affected Items Data ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ecrSheet = ss.getSheetByName(CONFIG.ECR_AFFECTED_ITEMS_SHEET);
  if (!ecrSheet) {
    ui.alert(`Error: The ECR Affected Items sheet named "${CONFIG.ECR_AFFECTED_ITEMS_SHEET}" was not found.`);
    return;
  }
  // Load and filter ECR data for the current ECO
  const ecrData = loadEcrData(ecrSheet, ecoBase);
  if (!ecrData) return; // Error handled in loadEcrData

  // Get unique ECRs associated with this ECO for the summary
  const allEcrsForEco = Array.from(new Set(ecrData.map(item => item.ecrNumber))).join(', ');
  const oldSheetName = ui.prompt('Detailed BOM Comparison', 'Enter the OLD BOM sheet name:', ui.ButtonSet.OK_CANCEL);
  if (oldSheetName.getSelectedButton() !== ui.Button.OK || !oldSheetName.getResponseText()) return;
  const newSheetName = ui.prompt('Detailed BOM Comparison', 'Enter the NEW BOM sheet name:', ui.ButtonSet.OK_CANCEL);
  if (newSheetName.getSelectedButton() !== ui.Button.OK || !newSheetName.getResponseText()) return;
  generateDetailedComparison(oldSheetName.getResponseText(), newSheetName.getResponseText(), ecoBase, allEcrsForEco, ecrData);
}

/**
 * Loads and filters data from the ECR Affected Items sheet based on ECO number.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} ecrSheet The ECR Affected Items sheet.
 * @param {string} ecoBase The ECO number to filter by.
 * @returns {Array<object>|null} Array of structured ECR data objects or null if error.
 */
function loadEcrData(ecrSheet, ecoBase) {
  const ui = SpreadsheetApp.getUi();
  const data = ecrSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert(`Warning: The ECR Affected Items sheet "${CONFIG.ECR_AFFECTED_ITEMS_SHEET}" is empty or has no data rows.`);
    return [];
  }
  const headers = data[0];
  // Required columns for Simplified matching
  const colIndexes = {
    ecr: headers.indexOf('ECR Number'),
    eco: headers.indexOf('ECO Number'),
    parent: headers.indexOf('Parent Assembly'),
    item: headers.indexOf('Item Number')
    // 'Change Type' is no longer strictly required for matching
  };
  // Check only essential columns for matching
  if (colIndexes.ecr === -1 || colIndexes.eco === -1 || colIndexes.parent === -1 || colIndexes.item === -1) {
    ui.alert(`Error: One or more required columns (ECR Number, ECO Number, Parent Assembly, Item Number) were not found in the "${CONFIG.ECR_AFFECTED_ITEMS_SHEET}" sheet.`);
    return null;
  }

  const filteredData = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colIndexes.eco] && row[colIndexes.eco].toString().trim() === ecoBase) {
      filteredData.push({
        ecrNumber: row[colIndexes.ecr] ? row[colIndexes.ecr].toString().trim() : '',
        parentAssembly: (row[colIndexes.parent] && row[colIndexes.parent].toString().trim() !== '') ? row[colIndexes.parent].toString().trim() : 'Top Level',
        itemNumber: row[colIndexes.item] ? row[colIndexes.item].toString().trim() : ''

        // 'changeType' is no longer stored as it's not used for matching
      });
    }
  }
  if (filteredData.length === 0) {
    ui.alert(`Warning: No ECR data found for ECO "${ecoBase}" in the "${CONFIG.ECR_AFFECTED_ITEMS_SHEET}" sheet.`);
  }
  return filteredData;
}


function generateDetailedComparison(oldSheetName, newSheetName, ecoBase, allEcrsForEco, ecrData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const oldSheet = ss.getSheetByName(oldSheetName);
  const newSheet = ss.getSheetByName(newSheetName);
  if (!oldSheet || !newSheet) {
    ui.alert('Error: One or both sheet names could not be found.');
    return;
  }
  newSheet.getDataRange().clearFormat();

  let oldBOMMap, newBOMMap;
  try {
    oldBOMMap = createBOMMap(oldSheet);
    newBOMMap = createBOMMap(newSheet);
  } catch (e) {
    ui.alert('Error creating BOM data maps:', e.message, ui.ButtonSet.OK);
    return;
  }

  const changes = [];
  const highlightColor = '#fff2cc'; // Light Yellow
  let changeCounter = 1;
  let addedCount = 0, removedCount = 0, modifiedItems = new Set();

  // --- [NEW] --- Sets to track direct and parent changes
  const affectedParentKeys = new Set();
  const directChangeKeys = new Set();
  // --- [END NEW] ---

  newBOMMap.forEach((newItem, locationKey) => {
    const oldItem = oldBOMMap.get(locationKey);
    const itemNumber = newItem.mainRow[CONFIG.ITEM_NUM_COL_NAME];
    const parentAssembly = newItem.parent;

    // --- Simplified ECR Lookup for any change affecting this Parent/Item ---
    const ecrNum = findMatchingEcrSimple(ecrData, parentAssembly, itemNumber);

    if (!oldItem) {
      const changeId = `${ecoBase}-${changeCounter++}`;
      changes.push([changeId, ecrNum, 'ADDED', itemNumber, parentAssembly, 'Component Added']);
      newSheet.getRange(newItem.startRow, 1, (newItem.endRow - newItem.startRow + 1), newSheet.getLastColumn()).setBackground('#d9ead3');
      addedCount++;
      // --- [NEW] --- Mark as direct change and flag parents
      directChangeKeys.add(locationKey);
      addParentKeys(locationKey, affectedParentKeys);
      // --- [END NEW] ---
    } else {
      let itemModified = false;
      const modifications = [];

      // Compare Main Row Attributes
      if (newItem.mainRow[CONFIG.DESC_COL_NAME] !== oldItem.mainRow[CONFIG.DESC_COL_NAME]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Description changed from "${oldItem.mainRow[CONFIG.DESC_COL_NAME]}" to "${newItem.mainRow[CONFIG.DESC_COL_NAME]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.DESC_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[CONFIG.ITEM_REV_COL_NAME] !== oldItem.mainRow[CONFIG.ITEM_REV_COL_NAME]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Rev changed from "${oldItem.mainRow[CONFIG.ITEM_REV_COL_NAME]}" to "${newItem.mainRow[CONFIG.ITEM_REV_COL_NAME]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.ITEM_REV_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[CONFIG.QTY_COL_NAME] !== oldItem.mainRow[CONFIG.QTY_COL_NAME]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Qty changed from "${oldItem.mainRow[CONFIG.QTY_COL_NAME]}" to "${newItem.mainRow[CONFIG.QTY_COL_NAME]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.QTY_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[CONFIG.LIFECYCLE_COL_NAME] !== oldItem.mainRow[CONFIG.LIFECYCLE_COL_NAME]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Lifecycle changed from "${oldItem.mainRow[CONFIG.LIFECYCLE_COL_NAME]}" to "${newItem.mainRow[CONFIG.LIFECYCLE_COL_NAME]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.LIFECYCLE_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }

      // Compare AML
      const oldAmlSet = new Set(oldItem.aml.map(a => `${a[CONFIG.MFR_NAME_COL_NAME]}|${a[CONFIG.MFR_PN_COL_NAME]}`));
      const newAmlSet = new Set(newItem.aml.map(a => `${a[CONFIG.MFR_NAME_COL_NAME]}|${a[CONFIG.MFR_PN_COL_NAME]}`));

      newItem.aml.forEach((aml, index) => {
        const amlString = `${aml[CONFIG.MFR_NAME_COL_NAME]}|${aml[CONFIG.MFR_PN_COL_NAME]}`;
        if (!oldAmlSet.has(amlString)) {
          const changeId = `${ecoBase}-${changeCounter++}`;
          modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `AML Added: ${aml[CONFIG.MFR_NAME_COL_NAME]} - ${aml[CONFIG.MFR_PN_COL_NAME]}`]);
          const amlRowNum = newItem.startRow + index;
          try {

            if (amlRowNum <= newSheet.getMaxRows()) {
              newSheet.getRange(amlRowNum, 1, 1, newSheet.getLastColumn()).setBackground(highlightColor);
            }
          } catch (e) { Logger.log(`Error highlighting AML row ${amlRowNum}: ${e}`); }
          itemModified = true;
        }
      });
      oldItem.aml.forEach(aml => {
        const amlString = `${aml[CONFIG.MFR_NAME_COL_NAME]}|${aml[CONFIG.MFR_PN_COL_NAME]}`;
        if (!newAmlSet.has(amlString)) {
          const changeId = `${ecoBase}-${changeCounter++}`;
          modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, newItem.parent, `AML Removed: ${aml[CONFIG.MFR_NAME_COL_NAME]} - ${aml[CONFIG.MFR_PN_COL_NAME]}`]);
          itemModified = true;
        }
      });
      if (modifications.length > 0) {
        changes.push(...modifications);
      }
      if (itemModified) {
        modifiedItems.add(locationKey);
        // --- [NEW] --- Mark as direct change and flag parents
        directChangeKeys.add(locationKey);
        addParentKeys(locationKey, affectedParentKeys);
        // --- [END NEW] ---
      }
      oldBOMMap.delete(locationKey);
    }
  });
  oldBOMMap.forEach((oldItem, locationKey) => {
    const changeId = `${ecoBase}-${changeCounter++}`;
    const itemNumber = oldItem.mainRow[CONFIG.ITEM_NUM_COL_NAME];
    const parentAssembly = oldItem.parent;
    const ecrNum = findMatchingEcrSimple(ecrData, parentAssembly, itemNumber);
    // --- [NEW] --- Flag parents of removed items
    addParentKeys(locationKey, affectedParentKeys);
    // --- [END NEW] ---
    changes.push([changeId, ecrNum, 'REMOVED', itemNumber, parentAssembly, 'Component Removed']);
    removedCount++;
  });

  // --- [NEW] Add "Change Impact" Column to New Sheet ---
  // This is done AFTER all comparison loops and highlighting,
  // but BEFORE the report is generated.
  try {
    const firstColHeader = newSheet.getRange(1, 1).getValue();
    let impactColIndex = 1;

    if (firstColHeader === 'Change Impact') {
      // Column already exists, just clear its content
      newSheet.getRange(2, impactColIndex, newSheet.getMaxRows() - 1, 1).clearContent();
    } else {
      // Insert new column at the beginning
      newSheet.insertColumnBefore(impactColIndex);
      newSheet.getRange(1, impactColIndex).setValue('Change Impact').setFontWeight('bold');
    }

    // Apply markers using the newBOMMap
    // The row numbers (item.startRow) are still correct.
    newBOMMap.forEach((item, locationKey) => {
      if (directChangeKeys.has(locationKey)) {
        newSheet.getRange(item.startRow, impactColIndex).setValue('●'); // Direct Change
      } else if (affectedParentKeys.has(locationKey)) {
        newSheet.getRange(item.startRow, impactColIndex).setValue('▼'); // Parent Impact
      }
    });

    newSheet.autoResizeColumn(impactColIndex);

  } catch (e) {
    Logger.log(`Error applying Change Impact markers: ${e}`);
    ui.alert(`Warning: Could not apply "Change Impact" markers to ${newSheetName}. Error: ${e.message}`);
  }
  // --- [END NEW] ---


  if (changes.length > 0) {
    changes.sort((a, b) => (a[4] < b[4] ? -1 : a[4] > b[4] ? 1 : a[3] < b[3] ? -1 : 1));
    // Sort by Parent(4), then Item(3)
    const reportSheetName = `Compare_Report_${new Date().toISOString().slice(0, 16).replace(/[:T]/g, '_')}`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName, 0);
    reportSheet.clear();

    // Add Summary Section
    reportSheet.insertRowsBefore(1, 8);
    // Adjusted to 8 rows for summary (was 9, removed blank row)
    reportSheet.getRange('A1').setValue('BOM Comparison Summary').setFontWeight('bold').setFontSize(12);
    reportSheet.getRange('A2').setValue(`ECO Base: ${ecoBase}`);
    reportSheet.getRange('A3').setValue(`Related ECR(s): ${allEcrsForEco || 'N/A'}`);
    reportSheet.getRange('A4').setValue(`Compared: "${oldSheetName}" vs "${newSheetName}"`);
    reportSheet.getRange('A5').setValue(`Date: ${new Date().toLocaleString()}`);
    // --- CORRECTED SUMMARY ROWS ---
    reportSheet.getRange('A6').setValue('Change Type').setFontWeight('bold');
    reportSheet.getRange('B6').setValue('Count').setFontWeight('bold');
    reportSheet.getRange('A7').setValue('Added Items:').setBackground('#d9ead3');
    reportSheet.getRange('B7').setValue(addedCount);
    // Count now in B7
    reportSheet.getRange('A8').setValue('Removed Items:').setBackground('#f4cccc');
    reportSheet.getRange('B8').setValue(removedCount);
    // Count now in B8
    reportSheet.getRange('A9').setValue('Modified Items:').setBackground('#fff2cc');
    reportSheet.getRange('B9').setValue(modifiedItems.size); // Count now in B9
    reportSheet.getRange('A6:B9').setHorizontalAlignment('left');
    // Adjusted range
    reportSheet.autoResizeColumns(1, 2);


    // Add Detailed Changes
    const headers = ['Change ID', 'ECR #', 'Change Type', 'Item Number', 'Parent Assembly', 'Details'];
    const headerRow = 11; // Start details below summary (Row 11 = 9 summary rows + 1 blank row)
    reportSheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (changes.length > 0) {
      reportSheet.getRange(headerRow + 1, 1, changes.length, headers.length).setValues(changes);
    }
    reportSheet.setFrozenRows(headerRow);
    reportSheet.autoResizeColumns(1, headers.length);
    applyReportFormatting(reportSheet, changes, headerRow);

    logECOComparison(ecoBase, allEcrsForEco, oldSheetName, newSheetName, reportSheetName, addedCount, removedCount, modifiedItems.size);
    ui.alert('Comparison Complete!', `A new formatted report named "${reportSheetName}" has been created, and "${newSheetName}" has been highlighted with change impact markers.`, ui.ButtonSet.OK);
    // [MODIFIED] Updated alert message.
  } else {
    ui.alert('Comparison Complete!', 'No differences were found.', ui.ButtonSet.OK);
  }
}

/**
 * Simplified ECR matching based only on Parent Assembly and Item Number.
 */
function findMatchingEcrSimple(ecrData, parentAssembly, itemNumber) {
  if (!ecrData || ecrData.length === 0) return '';
  const matchingEcrs = ecrData.filter(ecr =>
    (ecr.parentAssembly === parentAssembly || (parentAssembly === 'Top Level' && !ecr.parentAssembly)) &&
    ecr.itemNumber === itemNumber
  ).map(ecr => ecr.ecrNumber);
  return [...new Set(matchingEcrs)].join(', ');
}

/**
 * Applies formatting to the comparison report sheet.
 */
function applyReportFormatting(sheet, changes, headerRow) {
  const colors = ['#ffffff', '#e6e6e6'];
  const typeColors = { 'ADDED': '#d9ead3', 'MODIFIED': '#fff2cc', 'REMOVED': '#f4cccc' };
  let lastParent = null, colorIndex = 0;
  for (let i = 0; i < changes.length; i++) {
    const reportRow = i + headerRow + 1;
    const changeType = changes[i][2];
    const currentParent = changes[i][4];
    if (currentParent !== lastParent) {
      colorIndex = 1 - colorIndex;
      lastParent = currentParent;
    }
    sheet.getRange(reportRow, 1, 1, sheet.getLastColumn()).setBackground(colors[colorIndex]);
    sheet.getRange(reportRow, 3).setBackground(typeColors[changeType] || null);
    // Color 'Change Type' column (C)
  }
}

//====================================================================================================
// === [NEW] EXTERNAL PDM BOM COMPARISON =============================================================
//====================================================================================================

/**
 * [NEW] Runs the comparison between a subassembly of the Master BOM and an external PDM BOM.
 */
function runCompareWithExternalBOM() {
  const ui = SpreadsheetApp.getUi();
  const masterSheetName = ui.prompt('Compare Master vs. PDM', 'Enter the MASTER BOM sheet name (with standard headers):', ui.ButtonSet.OK_CANCEL);
  if (masterSheetName.getSelectedButton() !== ui.Button.OK || !masterSheetName.getResponseText()) return;

  const externalSheetName = ui.prompt('Compare Master vs. PDM', 'Enter the EXTERNAL PDM BOM sheet name (with PDM headers):', ui.ButtonSet.OK_CANCEL);
  if (externalSheetName.getSelectedButton() !== ui.Button.OK || !externalSheetName.getResponseText()) return;

  const assemblyNum = ui.prompt('Compare Master vs. PDM', 'Enter the Top-Level Assembly Number to compare (e.g., 13001430):', ui.ButtonSet.OK_CANCEL);
  if (assemblyNum.getSelectedButton() !== ui.Button.OK || !assemblyNum.getResponseText()) return;
  
  generateExternalComparison(masterSheetName.getResponseText().trim(), externalSheetName.getResponseText().trim(), assemblyNum.getResponseText().trim());
}

/**
 * [NEW] Core logic for comparing a Master Sub-BOM to an External PDM BOM.
 * This is a modified version of generateDetailedComparison, removing ECR logic for simplicity
 * and using custom BOM Map builders.
 */
function generateExternalComparison(masterSheetName, externalSheetName, assemblyNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getSheetByName(masterSheetName);
  const externalSheet = ss.getSheetByName(externalSheetName);

  if (!masterSheet || !externalSheet) {
    ui.alert('Error: One or both sheet names could not be found.');
    return;
  }
  externalSheet.getDataRange().clearFormat();

  let masterBOMMap, externalBOMMap;
  try {
      masterBOMMap = createMasterSubassemblyMap(masterSheet, assemblyNumber);
      externalBOMMap = createExternalBOMMap(externalSheet, CONFIG.PDM_HEADER_MAP);
  } catch (e) {
      ui.alert('Error creating BOM data maps:', e.message, ui.ButtonSet.OK);
      return;
  }

  const changes = [];
  const highlightColor = '#fff2cc'; // Light Yellow
  let changeCounter = 1;
  let addedCount = 0, removedCount = 0, modifiedItems = new Set();
  
  // Note: We are comparing the External BOM (as 'newItem') against the Master BOM (as 'oldItem')
  externalBOMMap.forEach((newItem, locationKey) => {
    const oldItem = masterBOMMap.get(locationKey);
    const itemNumber = newItem.mainRow[CONFIG.ITEM_NUM_COL_NAME];
    const parentAssembly = newItem.parent;
    const ecrNum = 'N/A'; // ECR linking is disabled for this comparison type

    if (!oldItem) {
      // Item exists in PDM BOM but not in Master BOM subassembly
      const changeId = `COMP-${changeCounter++}`;
      changes.push([changeId, ecrNum, 'ADDED', itemNumber, parentAssembly, 'Component Added (Exists in PDM, not Master)']);
      externalSheet.getRange(newItem.startRow, 1, (newItem.endRow - newItem.startRow + 1), externalSheet.getLastColumn()).setBackground('#d9ead3');
      addedCount++;
    } else 
    {
      let itemModified = false;
      const modifications = [];

      // Compare Main Row Attributes
      // Using CONFIG keys because createExternalBOMMap normalizes the keys
      if (newItem.mainRow[CONFIG.DESC_COL_NAME] !== oldItem.mainRow[CONFIG.DESC_COL_NAME]) {
        const changeId = `COMP-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Description changed from "${oldItem.mainRow[CONFIG.DESC_COL_NAME]}" to "${newItem.mainRow[CONFIG.DESC_COL_NAME]}"`]);
        externalSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.DESC_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[CONFIG.ITEM_REV_COL_NAME] !== oldItem.mainRow[CONFIG.ITEM_REV_COL_NAME]) {
        const changeId = `COMP-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Rev changed from "${oldItem.mainRow[CONFIG.ITEM_REV_COL_NAME]}" to "${newItem.mainRow[CONFIG.ITEM_REV_COL_NAME]}"`]);
        externalSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.ITEM_REV_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      // Note: Convert to String for comparison to handle numbers vs. text
      if (String(newItem.mainRow[CONFIG.QTY_COL_NAME]) !== String(oldItem.mainRow[CONFIG.QTY_COL_NAME])) {
        const changeId = `COMP-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Qty changed from "${oldItem.mainRow[CONFIG.QTY_COL_NAME]}" to "${newItem.mainRow[CONFIG.QTY_COL_NAME]}"`]);
        externalSheet.getRange(newItem.startRow, newItem.colIndexes[CONFIG.QTY_COL_NAME] + 1).setBackground(highlightColor);
        itemModified = true;
      }

      // Compare AML
      const oldAmlSet = new Set(oldItem.aml.map(a => `${a[CONFIG.MFR_NAME_COL_NAME]}|${a[CONFIG.MFR_PN_COL_NAME]}`));
      const newAmlSet = new Set(newItem.aml.map(a => `${a[CONFIG.MFR_NAME_COL_NAME]}|${a[CONFIG.MFR_PN_COL_NAME]}`));

      newItem.aml.forEach((aml, index) => {
        const amlString = `${aml[CONFIG.MFR_NAME_COL_NAME]}|${aml[CONFIG.MFR_PN_COL_NAME]}`;
        if (!oldAmlSet.has(amlString)) {
          const changeId = `COMP-${changeCounter++}`;
          modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `AML Added: ${aml[CONFIG.MFR_NAME_COL_NAME]} - ${aml[CONFIG.MFR_PN_COL_NAME]}`]);
          const amlRowNum = newItem.startRow + index;
          try {
             if(amlRowNum <= externalSheet.getMaxRows()) {
                externalSheet.getRange(amlRowNum, 1, 1, externalSheet.getLastColumn()).setBackground(highlightColor);
              }
          } catch(e) { Logger.log(`Error highlighting AML row ${amlRowNum}: ${e}`);}
          itemModified = true;
        }
      });
      oldItem.aml.forEach(aml => {
         const amlString = `${aml[CONFIG.MFR_NAME_COL_NAME]}|${aml[CONFIG.MFR_PN_COL_NAME]}`;
         if (!newAmlSet.has(amlString)) {
             const changeId = `COMP-${changeCounter++}`;
             modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, newItem.parent, `AML Removed: ${aml[CONFIG.MFR_NAME_COL_NAME]} - ${aml[CONFIG.MFR_PN_COL_NAME]}`]);
             itemModified = true;
         }
      });
      if (modifications.length > 0) {
        changes.push(...modifications);
      }
      if(itemModified) {
          modifiedItems.add(locationKey);
      }
      masterBOMMap.delete(locationKey); // Delete from master map to find REMOVED items
    }
  });
  
  // Items left in masterBOMMap exist in Master, but not in PDM BOM
  masterBOMMap.forEach((oldItem, locationKey) => {
    const changeId = `COMP-${changeCounter++}`;
    const itemNumber = oldItem.mainRow[CONFIG.ITEM_NUM_COL_NAME];
    const parentAssembly = oldItem.parent;
    changes.push([changeId, 'N/A', 'REMOVED', itemNumber, parentAssembly, 'Component Removed (Exists in Master, not PDM)']);
    removedCount++;
  });

  // --- Generate Report ---
  if (changes.length > 0) {
    changes.sort((a, b) => (a[4] < b[4] ? -1 : a[4] > b[4] ? 1 : a[3] < b[3] ? -1 : 1));
    const reportSheetName = `PDM_Compare_${assemblyNumber}_${new Date().toISOString().slice(0, 10)}`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName, 0);
    reportSheet.clear();

    // Add Summary Section
    reportSheet.insertRowsBefore(1, 8);
    reportSheet.getRange('A1').setValue('PDM vs. Master BOM Comparison').setFontWeight('bold').setFontSize(12);
    reportSheet.getRange('A2').setValue(`Assembly: ${assemblyNumber}`);
    reportSheet.getRange('A3').setValue(`Compared: "${masterSheetName}" (Master) vs "${externalSheetName}" (PDM)`);
    reportSheet.getRange('A4').setValue(`Date: ${new Date().toLocaleString()}`);
    reportSheet.getRange('A6').setValue('Change Type').setFontWeight('bold');
    reportSheet.getRange('B6').setValue('Count').setFontWeight('bold');
    reportSheet.getRange('A7').setValue('In PDM, Not in Master:').setBackground('#d9ead3');
    reportSheet.getRange('B7').setValue(addedCount);
    reportSheet.getRange('A8').setValue('In Master, Not in PDM:').setBackground('#f4cccc');
    reportSheet.getRange('B8').setValue(removedCount);
    reportSheet.getRange('A9').setValue('Modified Items:').setBackground('#fff2cc');
    reportSheet.getRange('B9').setValue(modifiedItems.size); 
    reportSheet.getRange('A6:B9').setHorizontalAlignment('left');
    reportSheet.autoResizeColumns(1, 2);

    // Add Detailed Changes
    const headers = ['Change ID', 'ECR #', 'Change Type', 'Item Number', 'Parent Assembly', 'Details'];
    const headerRow = 11; 
    reportSheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (changes.length > 0) {
        reportSheet.getRange(headerRow + 1, 1, changes.length, headers.length).setValues(changes);
    }
    reportSheet.setFrozenRows(headerRow);
    reportSheet.autoResizeColumns(1, headers.length);
    applyReportFormatting(reportSheet, changes, headerRow);

    ui.alert('Comparison Complete!', `A new formatted report named "${reportSheetName}" has been created, and "${externalSheetName}" has been highlighted.`, ui.ButtonSet.OK);
  } else {
    ui.alert('Comparison Complete!', 'No differences were found between the Master subassembly and the PDM BOM.', ui.ButtonSet.OK);
  }
}


//====================================================================================================
// === [REFACTORED] BOM MAP CREATION LOGIC ==========================================================
//====================================================================================================

/**
 * [NEW/REFACTORED] Core BOM Map building logic.
 * Creates a Map of BOM items, keyed by their structural location.
 * @param {string} sheetName The name of the sheet being processed (for error logging).
 * @param {Array<Array<string>>} data The 2D array of data from the sheet.
 * @param {Object} colIndexes An object mapping CONFIG keys (e.g., CONFIG.LEVEL_COL_NAME) to their column index.
 * @param {number} startRow The 0-based row index to start parsing from (e.g., 1 for header-less data).
 * @param {number} baseLevel The level to treat as "Level 0" (for subassembly extraction).
 * @returns {Map<string, object>} A Map of BOM item data.
 */
function buildBOMMap(sheetName, data, colIndexes, startRow, baseLevel) {
  const bomMap = new Map();
  if (data.length <= startRow) return bomMap; // No data to parse

  // Check for essential columns in the provided colIndexes
  const essentialCols = [
    CONFIG.LEVEL_COL_NAME, 
    CONFIG.ITEM_NUM_COL_NAME, 
    CONFIG.DESC_COL_NAME, 
    CONFIG.ITEM_REV_COL_NAME, 
    CONFIG.QTY_COL_NAME, 
    CONFIG.MFR_NAME_COL_NAME, 
    CONFIG.MFR_PN_COL_NAME
  ];
  
  // Lifecycle is only essential if its index is not -1 (i.e., it's expected)
  if (colIndexes[CONFIG.LIFECYCLE_COL_NAME] !== -1) {
    essentialCols.push(CONFIG.LIFECYCLE_COL_NAME);
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
    const itemNumber = row[colIndexes[CONFIG.ITEM_NUM_COL_NAME]];
    const isMainPartRow = itemNumber && itemNumber.toString().trim() !== '';
    const levelVal = row[colIndexes[CONFIG.LEVEL_COL_NAME]];
    
    // Calculate normalized level
    // Use a large negative number if level is blank to skip AML rows
    // Handle non-numeric levels like '3.1' by taking integer part
    const level = (levelVal !== '' && !isNaN(parseFloat(levelVal))) ? (parseInt(levelVal, 10) - baseLevel) : -999;
    
    if (isMainPartRow) {
       // Stop processing if we've returned to a level at or above the base level
       if (level < 0) { // This handles "uncle" rows (e.g., level 1 when base was 2)
         if (currentItemData) {
           bomMap.set(currentItemData.locationKey, currentItemData);
         }
         break; // We have exited the subassembly
       }
       
       // [FIX] This handles "sibling" rows
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
               [CONFIG.ITEM_NUM_COL_NAME]: itemNumber.toString().trim(),
               [CONFIG.DESC_COL_NAME]: row[colIndexes[CONFIG.DESC_COL_NAME]].toString().trim(),
               [CONFIG.ITEM_REV_COL_NAME]: row[colIndexes[CONFIG.ITEM_REV_COL_NAME]].toString().trim(),
               [CONFIG.QTY_COL_NAME]: row[colIndexes[CONFIG.QTY_COL_NAME]].toString().trim(),
               // Handle potentially missing lifecycle column
               [CONFIG.LIFECYCLE_COL_NAME]: colIndexes[CONFIG.LIFECYCLE_COL_NAME] !== -1 ? row[colIndexes[CONFIG.LIFECYCLE_COL_NAME]].toString().trim() : 'N/A',
           },
           aml: []
       };
    }

    if (currentItemData) {
       currentItemData.endRow = i + 1; // 1-based
       const mfrName = row[colIndexes[CONFIG.MFR_NAME_COL_NAME]];
       const mfrPN = row[colIndexes[CONFIG.MFR_PN_COL_NAME]];

       if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
           currentItemData.aml.push({
               [CONFIG.MFR_NAME_COL_NAME]: mfrName ? mfrName.toString().trim() : "",
               [CONFIG.MFR_PN_COL_NAME]: mfrPN ? mfrPN.toString().trim() : ""
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
 * [REFACTORED] Wrapper for buildBOMMap for standard, full-sheet comparison.
 * This is the original function, now pointing to the new core logic.
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
 * [NEW] Creates a BOM Map for a specific subassembly from the Master BOM sheet.
 */
function createMasterSubassemblyMap(sheet, assemblyNumber) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Map();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers); // Gets standard CONFIG indexes

  const itemNumCol = colIndexes[CONFIG.ITEM_NUM_COL_NAME];
  const levelCol = colIndexes[CONFIG.LEVEL_COL_NAME];
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
 * [NEW] Creates a BOM Map from an external sheet using a provided header map.
 * Assumes the external sheet starts at Level 0 for the assembly.
 */
function createExternalBOMMap(sheet, headerMap) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Map();
  const headers = data[0];

  // Build column indexes by mapping PDM headers to standard CONFIG keys
  const normalizedColIndexes = {
    [CONFIG.LEVEL_COL_NAME]: headers.indexOf(headerMap.LEVEL_COL_NAME),
    [CONFIG.ITEM_NUM_COL_NAME]: headers.indexOf(headerMap.ITEM_NUM_COL_NAME),
    [CONFIG.DESC_COL_NAME]: headers.indexOf(headerMap.DESC_COL_NAME),
    [CONFIG.ITEM_REV_COL_NAME]: headers.indexOf(headerMap.ITEM_REV_COL_NAME),
    [CONFIG.QTY_COL_NAME]: headers.indexOf(headerMap.QTY_COL_NAME),
    [CONFIG.LIFECYCLE_COL_NAME]: -1, // Not present in PDM BOM, set to -1
    [CONFIG.MFR_NAME_COL_NAME]: headers.indexOf(headerMap.MFR_NAME_COL_NAME),
    [CONFIG.MFR_PN_COL_NAME]: headers.indexOf(headerMap.MFR_PN_COL_NAME),
  };
  
  // Verify all mappings were successful
  for (const key in headerMap) {
    const configKey = Object.keys(CONFIG.PDM_HEADER_MAP).find(k => CONFIG.PDM_HEADER_MAP[k] === headerMap[key]);
    const standardConfigKey = Object.keys(CONFIG).find(k => k === key);
    if (normalizedColIndexes[standardConfigKey] === -1) {
      throw new Error(`Header "${headerMap[key]}" (mapped to ${standardConfigKey}) was not found in the PDM sheet.`);
    }
  }

  // Start parsing at row 2 (index 1), treat level 0 as base
  return buildBOMMap(sheet.getName(), data, normalizedColIndexes, 1, 0);
}


//====================================================================================================
// === GENERATE FABRICATOR BOM ======================================================================
//====================================================================================================

function runGenerateFabricatorBOMs() {
  const ui = SpreadsheetApp.getUi();
  const sourceResponse = ui.prompt('Generate Fabricator BOMs', 'Enter the name of the source Master BOM sheet:', ui.ButtonSet.OK_CANCEL);
  if (sourceResponse.getSelectedButton() !== ui.Button.OK || !sourceResponse.getResponseText()) return;

  const assemblyResponse = ui.prompt('Generate Fabricator BOMs', 'Enter one or more assembly part numbers, separated by commas:', ui.ButtonSet.OK_CANCEL);
  if (assemblyResponse.getSelectedButton() !== ui.Button.OK || !assemblyResponse.getResponseText()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceResponse.getResponseText());
  if (!sourceSheet) return ui.alert(`Sheet "${sourceResponse.getResponseText()}" not found.`);

  const assemblyNumbers = assemblyResponse.getResponseText().split(',').map(pn => pn.trim()).filter(Boolean);
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData[0];
  const colIndexes = getColumnIndexes(headers);

  if (colIndexes[CONFIG.LEVEL_COL_NAME] === -1 || colIndexes[CONFIG.ITEM_NUM_COL_NAME] === -1 || colIndexes[CONFIG.ITEM_REV_COL_NAME] === -1 || colIndexes[CONFIG.REFERENCE_NOTES_COL_NAME] === -1)
    return ui.alert('A required column (Level, Item Number, Item Rev, or Reference Notes) was not found.');
  let sheetsCreated = 0;
  assemblyNumbers.forEach(assemblyNum => {
    let parentRowIndex = -1;
    for (let i = 1; i < sourceData.length; i++) {
      if (sourceData[i][colIndexes[CONFIG.ITEM_NUM_COL_NAME]] && sourceData[i][colIndexes[CONFIG.ITEM_NUM_COL_NAME]].toString().trim() === assemblyNum) {
        parentRowIndex = i;
        break;
      }
    }

    if (parentRowIndex === -1) {
      ui.alert(`Assembly "${assemblyNum}" not found in the source BOM. Skipping.`);
      return;
    }


    const parentRow = sourceData[parentRowIndex];
    const parentLevel = parseInt(parentRow[colIndexes[CONFIG.LEVEL_COL_NAME]], 10);
    const parentRev = parentRow[colIndexes[CONFIG.ITEM_REV_COL_NAME]];
    const newSheetName = `Fabricator_${assemblyNum}_${parentRev}`;

    let newSheet = ss.getSheetByName(newSheetName);
    if (newSheet) newSheet.clear();
    else newSheet = ss.insertSheet(newSheetName);

    newSheet.appendRow(headers);
    let parentForSupplier = [...parentRow];
    parentForSupplier[colIndexes[CONFIG.LEVEL_COL_NAME]] = 0;
    newSheet.appendRow(parentForSupplier);

    for (let i = parentRowIndex + 1; i < sourceData.length; i++) {
      const childRow = sourceData[i];
      const childLevel = parseInt(childRow[colIndexes[CONFIG.LEVEL_COL_NAME]], 10);
      if (isNaN(childLevel) || childLevel <= parentLevel) break;

      const statusCell = childRow[colIndexes[CONFIG.REFERENCE_NOTES_COL_NAME]];
      const status = statusCell ? statusCell.toString().toUpperCase().trim() : '';
      if (status.includes('BUY') || status.includes('REF')) {
        let childForSupplier = [...childRow];
        childForSupplier[colIndexes[CONFIG.LEVEL_COL_NAME]] = childLevel - parentLevel;
        newSheet.appendRow(childForSupplier);
      }
    }
    sheetsCreated++;
  });
  if (sheetsCreated > 0) {
    ui.alert(`${sheetsCreated} Fabricator BOM sheet(s) have been created successfully.`);
  } else {
    ui.alert('No Fabricator BOM sheets were created (check assembly numbers?).');
  }
}

//====================================================================================================
// === AUDIT 'BUY' ITEMS ============================================================================
//====================================================================================================

function runScreenBuyItems() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('List \'BUY\' Items with \'REF\' Children', 'Enter the name of the BOM sheet to audit:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) return;

  const sourceSheetName = response.getResponseText();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Sheet "${sourceSheetName}" not found.`);

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers);
  if (colIndexes[CONFIG.LEVEL_COL_NAME] === -1 || colIndexes[CONFIG.ITEM_NUM_COL_NAME] === -1 || colIndexes[CONFIG.REFERENCE_NOTES_COL_NAME] === -1)
    return ui.alert('A required column was not found. Please check CONFIG settings.');
  const issues = new Set();

  for (let i = 1; i < data.length; i++) {
    const parentRow = data[i];
    const parentStatusCell = parentRow[colIndexes[CONFIG.REFERENCE_NOTES_COL_NAME]];
    const parentStatus = parentStatusCell ? parentStatusCell.toString().toUpperCase().trim() : '';
    if (parentStatus.includes('BUY')) {
      const parentLevel = parseInt(parentRow[colIndexes[CONFIG.LEVEL_COL_NAME]], 10);
      const parentItemNumber = parentRow[colIndexes[CONFIG.ITEM_NUM_COL_NAME]];
      if (isNaN(parentLevel) || !parentItemNumber) continue;

      for (let j = i + 1; j < data.length; j++) {
        const childRow = data[j];
        const childLevel = parseInt(childRow[colIndexes[CONFIG.LEVEL_COL_NAME]], 10);

        if (isNaN(childLevel) || childLevel <= parentLevel) break;

        const childStatusCell = childRow[colIndexes[CONFIG.REFERENCE_NOTES_COL_NAME]];
        const childStatus = childStatusCell ? childStatusCell.toString().toUpperCase().trim() : '';
        if (childStatus.includes('REF')) {
          issues.add(parentItemNumber);
          break;
        }
      }
    }
  }

  if (issues.size > 0) {
    const reportSheetName = `BUY_Item_Audit_List`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
    reportSheet.clear();
    const reportHeaders = ['\'BUY\' Items with \'REF\' Children'];
    const outputData = [reportHeaders, ...Array.from(issues).map(item => [item])];
    reportSheet.getRange(1, 1, 1, 1).setValues([reportHeaders]).setFontWeight('bold');
    if (outputData.length > 1) {
      reportSheet.getRange(2, 1, outputData.length - 1, 1).setValues(outputData.slice(1));
    }
    reportSheet.autoResizeColumn(1);
    ui.alert('Audit complete!', `Found ${issues.size} issue(s). See the "${reportSheetName}" sheet for the list.`, ui.ButtonSet.OK);
  } else {
    ui.alert('Audit complete!', 'No issues found. All \'BUY\' items have valid children.', ui.ButtonSet.OK);
  }
}

//====================================================================================================
// === AUDIT BOM LIFECYCLE =============================================================
//====================================================================================================
function runAuditBOMLifecycle() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Audit BOM Lifecycle Status', 'Enter the name of the BOM sheet to audit:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) return;
  const sourceSheetName = response.getResponseText();
  auditBOMLifecycle(sourceSheetName);
}

function auditBOMLifecycle(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Sheet "${sourceSheetName}" not found.`);

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers);
  if (colIndexes[CONFIG.ITEM_NUM_COL_NAME] === -1 || colIndexes[CONFIG.DESC_COL_NAME] === -1 || colIndexes[CONFIG.LIFECYCLE_COL_NAME] === -1)
    return ui.alert(`A required column ('${CONFIG.ITEM_NUM_COL_NAME}', '${CONFIG.DESC_COL_NAME}', or '${CONFIG.LIFECYCLE_COL_NAME}') was not found.`);
  const issues = [];
  const nonProductionStatuses = ['OBSOLETE', 'EOL', 'END OF LIFE', 'NRND', 'NOT RECOMMENDED'];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const lifecycleCell = row[colIndexes[CONFIG.LIFECYCLE_COL_NAME]];
    const lifecycleStatus = lifecycleCell ? lifecycleCell.toString().toUpperCase().trim() : '';
    const itemNumber = row[colIndexes[CONFIG.ITEM_NUM_COL_NAME]];
    if (itemNumber && lifecycleStatus && nonProductionStatuses.includes(lifecycleStatus)) {
      issues.push([
        itemNumber,
        row[colIndexes[CONFIG.DESC_COL_NAME]],
        row[colIndexes[CONFIG.LIFECYCLE_COL_NAME]],
        `Row ${i + 1}`
      ]);
    }
  }

  if (issues.length > 0) {
    const reportSheetName = `Lifecycle_Audit_Report_${sourceSheetName}`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
    reportSheet.clear();
    const reportHeaders = ['Item Number', 'Description', 'Lifecycle Status', 'Found at Sheet Row'];
    reportSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight('bold');
    reportSheet.getRange(2, 1, issues.length, reportHeaders.length).setValues(issues);
    reportSheet.autoResizeColumns(1, reportHeaders.length);
    ui.alert(`Lifecycle Audit complete! Found ${issues.length} components with non-production statuses. See "${reportSheetName}" for details.`);
  } else {
    ui.alert('Lifecycle Audit complete! No components with non-production lifecycle statuses found.');
  }
}


//====================================================================================================
// === UTILITIES ====================================================================================
//====================================================================================================

function runPrepareAMLRows() {
  const ui = SpreadsheetApp.getUi();
  const partNumberResponse = ui.prompt('Prepare AML Rows', 'Enter Item Numbers, separated by commas:', ui.ButtonSet.OK_CANCEL);
  if (partNumberResponse.getSelectedButton() != ui.Button.OK || !partNumberResponse.getResponseText()) return;
  const partNumbers = partNumberResponse.getResponseText().split(',').map(pn => pn.trim()).filter(pn => pn);
  if (partNumbers.length > 0) {
    prepareAMLRows(partNumbers);
  } else {
    ui.alert('No part numbers were entered.');
  }
}

function prepareAMLRows(partNumbers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const amlSheet = ss.getSheetByName(CONFIG.AML_SHEET_NAME);
  const activeSheet = ss.getActiveSheet();
  if (!amlSheet) {
    ui.alert(`Error: The master AML sheet named "${CONFIG.AML_SHEET_NAME}" could not be found.`);
    return;
  }
  if (!activeSheet || activeSheet.getName() === CONFIG.AML_SHEET_NAME || activeSheet.getName() === CONFIG.ITEMS_SHEET_NAME) {
    ui.alert('This function must be run on a valid BOM sheet, not on master data sheets.');
    return;
  }
  const amlData = amlSheet.getRange("A:A").getValues().flat().map(String);
  const bomData = activeSheet.getDataRange().getValues();
  const headers = bomData.length > 0 ?
    bomData[0] : [];
  const itemNumColIndex = headers.indexOf(CONFIG.ITEM_NUM_COL_NAME);
  if (itemNumColIndex === -1) {
    ui.alert(`Error: Could not find the column "${CONFIG.ITEM_NUM_COL_NAME}" in the active sheet.`);
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

function runGenerateMasterLists() {
  const ui = SpreadsheetApp.getUi();
  const bomSheetResponse = ui.prompt('Generate Master Lists', 'Enter the name of the source BOM sheet:', ui.ButtonSet.OK_CANCEL);
  if (bomSheetResponse.getSelectedButton() != ui.Button.OK || !bomSheetResponse.getResponseText()) return;
  const sourceSheetName = bomSheetResponse.getResponseText().trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Error: Source sheet "${sourceSheetName}" not found.`);
  generateItemList(sourceSheet);
  generateAmlList(sourceSheet);
  ui.alert(`Success! New ITEM and AML sheets have been generated from "${sourceSheetName}".`);
}

function generateItemList(sourceSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData.length > 0 ? sourceData[0] : [];
  const colIdx = getColumnIndexes(headers);
  const itemNumIdx = colIdx[CONFIG.ITEM_NUM_COL_NAME], descIdx = colIdx[CONFIG.DESC_COL_NAME], revIdx = colIdx[CONFIG.ITEM_REV_COL_NAME];
  if ([itemNumIdx, descIdx, revIdx].includes(-1)) return SpreadsheetApp.getUi().alert('Could not generate ITEM list. Required columns not found.');
  const uniqueItems = new Map();
  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i], itemNumber = row[itemNumIdx];
    const itemKey = itemNumber ? itemNumber.toString().trim() : null;
    if (itemKey && !uniqueItems.has(itemKey)) {
      uniqueItems.set(itemKey, { desc: row[descIdx], rev: row[revIdx] });
    }
  }
  const newSheetName = `ITEM_${sourceSheet.getName()}`;
  let newSheet = ss.getSheetByName(newSheetName) || ss.insertSheet(newSheetName);
  newSheet.clear();
  const newHeaders = [CONFIG.ITEM_NUM_COL_NAME, CONFIG.DESC_COL_NAME, CONFIG.ITEM_REV_COL_NAME];
  const outputData = [newHeaders, ...Array.from(uniqueItems, ([key, value]) => [key, value.desc, value.rev])];
  newSheet.getRange(1, 1, outputData.length, newHeaders.length).setValues(outputData);
  newSheet.autoResizeColumns(1, newHeaders.length);
}

function generateAmlList(sourceSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData.length > 0 ? sourceData[0] : [];
  const colIdx = getColumnIndexes(headers);
  const itemNumIdx = colIdx[CONFIG.ITEM_NUM_COL_NAME], mfrNameIdx = colIdx[CONFIG.MFR_NAME_COL_NAME], mfrPnIdx = colIdx[CONFIG.MFR_PN_COL_NAME];
  if ([itemNumIdx, mfrNameIdx, mfrPnIdx].includes(-1)) return SpreadsheetApp.getUi().alert('Could not generate AML list. Required columns not found.');
  const uniqueAml = new Set(), amlData = [];
  let currentItemNumber = '';
  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    if (row[itemNumIdx] && row[itemNumIdx].toString().trim() !== '') {
      currentItemNumber = row[itemNumIdx].toString().trim();
    }
    const mfrName = row[mfrNameIdx], mfrPn = row[mfrPnIdx];
    if (currentItemNumber && (mfrName || mfrPn)) {
      const uniqueKey = `${currentItemNumber}|${mfrName || ''}|${mfrPn || ''}`;
      if (!uniqueAml.has(uniqueKey)) {
        uniqueAml.add(uniqueKey);
        amlData.push([currentItemNumber, mfrName || '', mfrPn || '']);
      }
    }
  }
  const newSheetName = `AML_${sourceSheet.getName()}`;
  let newSheet = ss.getSheetByName(newSheetName) || ss.insertSheet(newSheetName);
  newSheet.clear();
  const newHeaders = [CONFIG.ITEM_NUM_COL_NAME, CONFIG.MFR_NAME_COL_NAME, CONFIG.MFR_PN_COL_NAME];
  const outputData = [newHeaders, ...amlData];
  newSheet.getRange(1, 1, outputData.length, newHeaders.length).setValues(outputData);
  newSheet.autoResizeColumns(1, newHeaders.length);
}

function runWhereUsedAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const partNumberResponse = ui.prompt('Where-Used Analysis', 'Enter the Part Number to search for:', ui.ButtonSet.OK_CANCEL);
  if (partNumberResponse.getSelectedButton() != ui.Button.OK || !partNumberResponse.getResponseText()) return;
  const sheetNameResponse = ui.prompt('Where-Used Analysis', `Enter the BOM sheet to search in:`, ui.ButtonSet.OK_CANCEL);
  if (sheetNameResponse.getSelectedButton() != ui.Button.OK || !sheetNameResponse.getResponseText()) return;
  performWhereUsed(partNumberResponse.getResponseText().trim(), sheetNameResponse.getResponseText().trim());
}

function performWhereUsed(partNumber, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) return ui.alert(`Error: Sheet "${sheetName}" not found.`);
  const data = sourceSheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : [];
  const colIdx = getColumnIndexes(headers);
  if (colIdx[CONFIG.LEVEL_COL_NAME] === -1 || colIdx[CONFIG.ITEM_NUM_COL_NAME] === -1) return ui.alert(`Error: Could not find required columns.`);
  const parentAssemblies = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colIdx[CONFIG.ITEM_NUM_COL_NAME]] && row[colIdx[CONFIG.ITEM_NUM_COL_NAME]].toString().trim() === partNumber) {
      const childLevel = parseInt(row[colIdx[CONFIG.LEVEL_COL_NAME]], 10);
      if (isNaN(childLevel)) continue;
      for (let j = i - 1; j >= 0; j--) {
        const parentRow = data[j];
        const parentLevel = parseInt(parentRow[colIdx[CONFIG.LEVEL_COL_NAME]], 10);
        if (isNaN(parentLevel)) continue;
        if (parentLevel < childLevel) {
          const parentPartNumber = parentRow[colIdx[CONFIG.ITEM_NUM_COL_NAME]].toString();
          const parentDesc = parentRow[colIdx[CONFIG.DESC_COL_NAME]] || 'N/A';
          parentAssemblies.set(parentPartNumber, parentDesc);
          break;
        }
      }
    }
  }
  let htmlOutput = `<b>Parent Assemblies for: ${partNumber}</b><br/><br/>`;
  if (parentAssemblies.size > 0) {
    htmlOutput += '<style>table,th,td{border:1px solid #ccc;border-collapse:collapse;padding:5px;font-family:Arial;}</style><table><tr><th>Parent Part Number</th><th>Description</th></tr>';
    parentAssemblies.forEach((desc, pn) => { htmlOutput += `<tr><td>${pn}</td><td>${desc}</td></tr>`; });
    htmlOutput += '</table>';
  } else {
    htmlOutput += 'Part number not found in any assemblies.';
  }
  ui.showModalDialog(HtmlService.createHtmlOutput(htmlOutput).setWidth(500).setHeight(300), `Where-Used Results`);
}


function runReleaseNewBOM() {
  const ui = SpreadsheetApp.getUi();

  // 1. Prompt for inputs
  const wipSheetName = ui.prompt('Release New BOM', 'Enter the name of the approved WIP sheet:', ui.ButtonSet.OK_CANCEL);
  if (wipSheetName.getSelectedButton() !== ui.Button.OK || !wipSheetName.getResponseText()) return;

  const newRevName = ui.prompt('Release New BOM', 'Enter the name for the new RELEASED sheet:', ui.ButtonSet.OK_CANCEL);
  if (newRevName.getSelectedButton() !== ui.Button.OK || !newRevName.getResponseText()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wipSheet = ss.getSheetByName(wipSheetName.getResponseText());

  // Validations
  if (!wipSheet) return ui.alert(`Error: Sheet "${wipSheetName.getResponseText()}" not found.`);
  if (ss.getSheetByName(newRevName.getResponseText())) return ui.alert(`A sheet named "${newRevName.getResponseText()}" already exists.`);

  // 2. Create the new sheet by copying
  const newSheet = wipSheet.copyTo(ss).setName(newRevName.getResponseText());
  ss.setActiveSheet(newSheet);

  // --- OPTIMIZATION SECTION START ---

  // 3. Batch Delete 'REF' Rows (Prevents Timeout)
  const data = newSheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : [];
  const statusColIndex = headers.indexOf(CONFIG.REFERENCE_NOTES_COL_NAME);

  if (statusColIndex === -1) {
    ui.alert('Warning', `Column "${CONFIG.REFERENCE_NOTES_COL_NAME}" not found. Could not filter for internal view.`, ui.ButtonSet.OK);
  } else {
    // We identify BLOCKS of rows to delete instead of one by one.
    const rangesToDelete = [];
    let currentBlockStart = -1;
    let currentBlockCount = 0;

    // Iterate backwards (Bottom to Top)
    // Row 0 is header, so we stop at i=1
    for (let i = data.length - 1; i >= 1; i--) {
      const cellValue = data[i][statusColIndex];
      const status = cellValue ? String(cellValue).toUpperCase() : "";

      if (status.includes('REF')) {
        if (currentBlockStart === -1) {
          // Found the bottom of a new block
          currentBlockStart = i;
          currentBlockCount = 1;
        } else {
          // Extend existing block upwards
          currentBlockStart = i;
          currentBlockCount++;
        }
      } else {
        // Found a gap (keep row). If we were tracking a block, save it now.
        if (currentBlockStart !== -1) {
          // currentBlockStart is 0-based array index. 
          // Sheet rows are 1-based. So row = index + 1.
          rangesToDelete.push({ row: currentBlockStart + 1, count: currentBlockCount });
          currentBlockStart = -1;
          currentBlockCount = 0;
        }
      }
    }
    // Capture final block if it ended at the top
    if (currentBlockStart !== -1) {
      rangesToDelete.push({ row: currentBlockStart + 1, count: currentBlockCount });
    }

    // Execute Batch Deletions
    // rangesToDelete contains blocks ordered from Bottom to Top, so it's safe to delete.
    rangesToDelete.forEach(range => {
      newSheet.deleteRows(range.row, range.count);
    });
  }

  // 4. Hard Limit: Delete all columns after Column R (Column 18)
  const maxCols = newSheet.getMaxColumns();
  const limitColIndex = 18; // Column R is the 18th column
  if (maxCols > limitColIndex) {
    newSheet.deleteColumns(limitColIndex + 1, maxCols - limitColIndex);
  }

  // 5. Delete specific "Change Tracking" columns if they exist within A-R
  // Refresh headers as we may have deleted rows/columns
  const finalHeaders = newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).getValues()[0];
  const colsToDelete = [];

  finalHeaders.forEach((header, index) => {
    if (CONFIG.CHANGE_TRACKING_COLS_TO_DELETE.includes(header)) {
      colsToDelete.push(index + 1); // Store 1-based index
    }
  });

  // Sort Descending (Highest Index First) to prevent index shifting when deleting
  colsToDelete.sort((a, b) => b - a);
  colsToDelete.forEach(colIndex => {
    newSheet.deleteColumn(colIndex);
  });

  // --- OPTIMIZATION SECTION END ---

  // 6. Protect the Sheet
  const protection = newSheet.protect().setDescription(`Released BOM ${newRevName.getResponseText()}`);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);

  ui.alert('Success!', `New BOM revision "${newRevName.getResponseText()}" has been created, filtered, cleaned (Column R limit), and protected.`, ui.ButtonSet.OK);
}

// [MODIFIED] The 'Change Impact' column will now be deleted upon release, as added to CONFIG.

//====================================================================================================
// === ECO LOGGING FUNCTION ========================================================================
//====================================================================================================
/**
 * Logs the details of a completed BOM comparison to the ECO History sheet.
 */
function logECOComparison(ecoBase, ecrString, oldSheetName, newSheetName, reportSheetName, addedCount, removedCount, modifiedCount) { // Added ecrString
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(CONFIG.ECO_LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.ECO_LOG_SHEET_NAME);
    const logHeaders = ['ECO Base', 'Related ECRs', 'Date', 'User', 'Old Sheet', 'New Sheet', 'Report Sheet', 'Added Items', 'Removed Items', 'Modified Items'];
    logSheet.appendRow(logHeaders);
    logSheet.setFrozenRows(1);
    logSheet.getRange(1, 1, 1, logHeaders.length).setFontWeight('bold');
  }

  const timestamp = new Date();
  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    userEmail = 'Unknown User';
  }

  logSheet.appendRow([
    ecoBase,
    ecrString || 'N/A', // Add ECR string
    timestamp,
    userEmail,
    oldSheetName,
    newSheetName,
    reportSheetName,
    addedCount,
    removedCount,
    modifiedCount
  ]);
}


//====================================================================================================
// === HELPER FUNCTIONS =============================================================================
//====================================================================================================

/**
 * [NEW] Helper function to add all parent keys of a given locationKey to a Set.
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
// --- [END NEW] ---


/**
 * Gets column indexes based on names from the CONFIG object. Robustly handles missing columns.
 * @param {string[]} headers An array of header names.
 * @returns {Object} An object mapping config names to column indexes (or -1 if not found).
 */
function getColumnIndexes(headers) {
  const indexes = {};
  if (!headers || headers.length === 0) {
    Logger.log("Warning: getColumnIndexes received empty or invalid headers.");
    // Initialize all expected keys to -1
    for (const key in CONFIG) {
      if (typeof CONFIG[key] === 'string' && CONFIG[key] !== CONFIG.ITEMS_SHEET_NAME && CONFIG[key] !== CONFIG.AML_SHEET_NAME && CONFIG[key] !== CONFIG.ECO_LOG_SHEET_NAME && CONFIG[key] !== CONFIG.ECR_AFFECTED_ITEMS_SHEET && key !== 'PDM_HEADER_MAP' && key !== 'PDM_GRAFT_HEADERS') { 
        indexes[CONFIG[key]] = -1;
      }
    }
    return indexes;
  }

  // [MODIFIED] Handle dynamically added columns like 'Change Impact'
  // First, map all known CONFIG columns
  for (const key in CONFIG) {
    if (typeof CONFIG[key] === 'string' && CONFIG[key] !== CONFIG.ITEMS_SHEET_NAME && CONFIG[key] !== CONFIG.AML_SHEET_NAME && CONFIG[key] !== CONFIG.ECO_LOG_SHEET_NAME && CONFIG[key] !== CONFIG.ECR_AFFECTED_ITEMS_SHEET && key !== 'PDM_HEADER_MAP' && key !== 'PDM_GRAFT_HEADERS') { 
      const index = headers.indexOf(CONFIG[key]);
      indexes[CONFIG[key]] = index; // Store index (-1 if not found)
    }
  }

  // [NEW] Also map any headers that aren't in CONFIG, just in case
  // This helps createBOMMap find columns even if they aren't in CONFIG
  headers.forEach((header, index) => {
    if (header && !Object.values(indexes).includes(index)) {
      // Check if this header name is NOT one of the config values
      const isConfigValue = Object.values(CONFIG).includes(header);
      if (!isConfigValue) {
        indexes[header] = index; // e.g., indexes['Change Impact'] = 0
      }
    }
  });

  return indexes;
}