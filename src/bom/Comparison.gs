// ============================================================================
// COMPARISON.gs — BOM Comparison Tools
// ============================================================================
// Detailed internal BOM comparison (ECO-linked) and external PDM-vs-Master
// comparison. Generates colour-coded reports and change-impact markers.
// ============================================================================

// ---------------------
// Internal BOM Comparison (ECO-linked)
// ---------------------

function runDetailedComparison() {
  const ecoBase = promptWithValidation('Generate Comparison Report', 'Enter the base ECO Number for this report (e.g., ECO-12):',
    { pattern: /^ECO-?\d+/i, patternHint: 'ECO number should start with "ECO" followed by a number (e.g., ECO-12).' });
  if (!ecoBase) return;

  // --- Read ECR Affected Items Data ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ecrSheet = ss.getSheetByName(BOM_CONFIG.ECR_AFFECTED_ITEMS_SHEET);
  if (!ecrSheet) {
    SpreadsheetApp.getUi().alert(`Error: The ECR Affected Items sheet named "${BOM_CONFIG.ECR_AFFECTED_ITEMS_SHEET}" was not found.`);
    return;
  }
  // Load and filter ECR data for the current ECO
  const ecrData = loadEcrData(ecrSheet, ecoBase);
  if (!ecrData) return; // Error handled in loadEcrData

  // Get unique ECRs associated with this ECO for the summary
  const allEcrsForEco = Array.from(new Set(ecrData.map(item => item.ecrNumber))).join(', ');
  const oldSheetName = promptWithValidation('Detailed BOM Comparison', 'Enter the OLD BOM sheet name:');
  if (!oldSheetName) return;
  const newSheetName = promptWithValidation('Detailed BOM Comparison', 'Enter the NEW BOM sheet name:');
  if (!newSheetName) return;
  generateDetailedComparison(oldSheetName, newSheetName, ecoBase, allEcrsForEco, ecrData);
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
    ui.alert(`Warning: The ECR Affected Items sheet "${BOM_CONFIG.ECR_AFFECTED_ITEMS_SHEET}" is empty or has no data rows.`);
    return [];
  }
  const headers = data[0];
  // Required columns for Simplified matching
  const colIndexes = {
    ecr: headers.indexOf('ECR Number'),
    eco: headers.indexOf('ECO Number'),
    parent: headers.indexOf('Parent Assembly'),
    item: headers.indexOf('Item Number')
  };
  // Check only essential columns for matching
  if (colIndexes.ecr === -1 || colIndexes.eco === -1 || colIndexes.parent === -1 || colIndexes.item === -1) {
    ui.alert(`Error: One or more required columns (ECR Number, ECO Number, Parent Assembly, Item Number) were not found in the "${BOM_CONFIG.ECR_AFFECTED_ITEMS_SHEET}" sheet.`);
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
      });
    }
  }
  if (filteredData.length === 0) {
    ui.alert(`Warning: No ECR data found for ECO "${ecoBase}" in the "${BOM_CONFIG.ECR_AFFECTED_ITEMS_SHEET}" sheet.`);
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

  // Sets to track direct and parent changes
  const affectedParentKeys = new Set();
  const directChangeKeys = new Set();

  newBOMMap.forEach((newItem, locationKey) => {
    const oldItem = oldBOMMap.get(locationKey);
    const itemNumber = newItem.mainRow[COL.ITEM_NUMBER];
    const parentAssembly = newItem.parent;

    // Simplified ECR Lookup for any change affecting this Parent/Item
    const ecrNum = findMatchingEcrSimple(ecrData, parentAssembly, itemNumber);

    if (!oldItem) {
      const changeId = `${ecoBase}-${changeCounter++}`;
      changes.push([changeId, ecrNum, 'ADDED', itemNumber, parentAssembly, 'Component Added']);
      newSheet.getRange(newItem.startRow, 1, (newItem.endRow - newItem.startRow + 1), newSheet.getLastColumn()).setBackground('#d9ead3');
      addedCount++;
      // Mark as direct change and flag parents
      directChangeKeys.add(locationKey);
      addParentKeys(locationKey, affectedParentKeys);
    } else {
      let itemModified = false;
      const modifications = [];

      // Compare Main Row Attributes
      if (newItem.mainRow[COL.DESCRIPTION] !== oldItem.mainRow[COL.DESCRIPTION]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Description changed from "${oldItem.mainRow[COL.DESCRIPTION]}" to "${newItem.mainRow[COL.DESCRIPTION]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[COL.DESCRIPTION] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[COL.ITEM_REV] !== oldItem.mainRow[COL.ITEM_REV]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Rev changed from "${oldItem.mainRow[COL.ITEM_REV]}" to "${newItem.mainRow[COL.ITEM_REV]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[COL.ITEM_REV] + 1).setBackground(highlightColor);
        itemModified = true;
        logRevisionChange(itemNumber, String(oldItem.mainRow[COL.ITEM_REV]), String(newItem.mainRow[COL.ITEM_REV]), ecoBase);
      }
      if (newItem.mainRow[COL.QTY] !== oldItem.mainRow[COL.QTY]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Qty changed from "${oldItem.mainRow[COL.QTY]}" to "${newItem.mainRow[COL.QTY]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[COL.QTY] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[COL.LIFECYCLE] !== oldItem.mainRow[COL.LIFECYCLE]) {
        const changeId = `${ecoBase}-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Lifecycle changed from "${oldItem.mainRow[COL.LIFECYCLE]}" to "${newItem.mainRow[COL.LIFECYCLE]}"`]);
        newSheet.getRange(newItem.startRow, newItem.colIndexes[COL.LIFECYCLE] + 1).setBackground(highlightColor);
        itemModified = true;
      }

      // Compare AML
      const oldAmlSet = new Set(oldItem.aml.map(a => `${a[COL.MFR_NAME]}|${a[COL.MFR_PN]}`));
      const newAmlSet = new Set(newItem.aml.map(a => `${a[COL.MFR_NAME]}|${a[COL.MFR_PN]}`));

      newItem.aml.forEach((aml, index) => {
        const amlString = `${aml[COL.MFR_NAME]}|${aml[COL.MFR_PN]}`;
        if (!oldAmlSet.has(amlString)) {
          const changeId = `${ecoBase}-${changeCounter++}`;
          modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `AML Added: ${aml[COL.MFR_NAME]} - ${aml[COL.MFR_PN]}`]);
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
        const amlString = `${aml[COL.MFR_NAME]}|${aml[COL.MFR_PN]}`;
        if (!newAmlSet.has(amlString)) {
          const changeId = `${ecoBase}-${changeCounter++}`;
          modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, newItem.parent, `AML Removed: ${aml[COL.MFR_NAME]} - ${aml[COL.MFR_PN]}`]);
          itemModified = true;
        }
      });
      if (modifications.length > 0) {
        changes.push(...modifications);
      }
      if (itemModified) {
        modifiedItems.add(locationKey);
        // Mark as direct change and flag parents
        directChangeKeys.add(locationKey);
        addParentKeys(locationKey, affectedParentKeys);
      }
      oldBOMMap.delete(locationKey);
    }
  });
  oldBOMMap.forEach((oldItem, locationKey) => {
    const changeId = `${ecoBase}-${changeCounter++}`;
    const itemNumber = oldItem.mainRow[COL.ITEM_NUMBER];
    const parentAssembly = oldItem.parent;
    const ecrNum = findMatchingEcrSimple(ecrData, parentAssembly, itemNumber);
    // Flag parents of removed items
    addParentKeys(locationKey, affectedParentKeys);
    changes.push([changeId, ecrNum, 'REMOVED', itemNumber, parentAssembly, 'Component Removed']);
    removedCount++;
  });

  // Add "Change Impact" Column to New Sheet
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

    // Apply markers using batch write instead of cell-by-cell setValue
    const maxRow = newSheet.getMaxRows();
    if (maxRow > 1) {
      const impactValues = newSheet.getRange(2, impactColIndex, maxRow - 1, 1).getValues();

      newBOMMap.forEach((item, locationKey) => {
        const rowIdx = item.startRow - 2; // Convert 1-based sheet row to 0-based array index (minus header)
        if (rowIdx >= 0 && rowIdx < impactValues.length) {
          if (directChangeKeys.has(locationKey)) {
            impactValues[rowIdx][0] = '●'; // Direct Change
          } else if (affectedParentKeys.has(locationKey)) {
            impactValues[rowIdx][0] = '▼'; // Parent Impact
          }
        }
      });

      newSheet.getRange(2, impactColIndex, impactValues.length, 1).setValues(impactValues);
    }
    newSheet.autoResizeColumn(impactColIndex);

  } catch (e) {
    Logger.log(`Error applying Change Impact markers: ${e}`);
    ui.alert(`Warning: Could not apply "Change Impact" markers to ${newSheetName}. Error: ${e.message}`);
  }


  if (changes.length > 0) {
    changes.sort((a, b) => (a[4] < b[4] ? -1 : a[4] > b[4] ? 1 : a[3] < b[3] ? -1 : 1));
    // Sort by Parent(4), then Item(3)
    const reportSheetName = `Compare_Report_${new Date().toISOString().slice(0, 16).replace(/[:T]/g, '_')}`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName, 0);
    reportSheet.clear();

    // Add Summary Section
    reportSheet.insertRowsBefore(1, 8);
    reportSheet.getRange('A1').setValue('BOM Comparison Summary').setFontWeight('bold').setFontSize(12);
    reportSheet.getRange('A2').setValue(`ECO Base: ${ecoBase}`);
    reportSheet.getRange('A3').setValue(`Related ECR(s): ${allEcrsForEco || 'N/A'}`);
    reportSheet.getRange('A4').setValue(`Compared: "${oldSheetName}" vs "${newSheetName}"`);
    reportSheet.getRange('A5').setValue(`Date: ${new Date().toLocaleString()}`);
    reportSheet.getRange('A6').setValue('Change Type').setFontWeight('bold');
    reportSheet.getRange('B6').setValue('Count').setFontWeight('bold');
    reportSheet.getRange('A7').setValue('Added Items:').setBackground('#d9ead3');
    reportSheet.getRange('B7').setValue(addedCount);
    reportSheet.getRange('A8').setValue('Removed Items:').setBackground('#f4cccc');
    reportSheet.getRange('B8').setValue(removedCount);
    reportSheet.getRange('A9').setValue('Modified Items:').setBackground('#fff2cc');
    reportSheet.getRange('B9').setValue(modifiedItems.size);
    reportSheet.getRange('A6:B9').setHorizontalAlignment('left');
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
  if (changes.length === 0) return;

  const colors = ['#ffffff', '#e6e6e6'];
  const typeColors = { 'ADDED': '#d9ead3', 'MODIFIED': '#fff2cc', 'REMOVED': '#f4cccc' };
  const numCols = sheet.getLastColumn();

  // Build background arrays in memory, then apply in a single batch
  const rowBackgrounds = [];
  let lastParent = null, colorIndex = 0;

  for (let i = 0; i < changes.length; i++) {
    const changeType = changes[i][2];
    const currentParent = changes[i][4];
    if (currentParent !== lastParent) {
      colorIndex = 1 - colorIndex;
      lastParent = currentParent;
    }
    // Create a row of background colors
    const rowBg = new Array(numCols).fill(colors[colorIndex]);
    // Override column C (index 2) with the change-type color
    rowBg[2] = typeColors[changeType] || colors[colorIndex];
    rowBackgrounds.push(rowBg);
  }

  // Single batch write for all backgrounds
  sheet.getRange(headerRow + 1, 1, changes.length, numCols).setBackgrounds(rowBackgrounds);
}


// ---------------------
// External PDM BOM Comparison
// ---------------------

/**
 * Runs the comparison between a subassembly of the Master BOM and an external PDM BOM.
 */
function runCompareWithExternalBOM() {
  const masterSheetName = promptWithValidation('Compare Master vs. PDM', 'Enter the MASTER BOM sheet name (with standard headers):');
  if (!masterSheetName) return;

  const externalSheetName = promptWithValidation('Compare Master vs. PDM', 'Enter the EXTERNAL PDM BOM sheet name (with PDM headers):');
  if (!externalSheetName) return;

  const assemblyNum = promptWithValidation('Compare Master vs. PDM', 'Enter the Top-Level Assembly Number to compare (e.g., 13001430):');
  if (!assemblyNum) return;

  generateExternalComparison(masterSheetName, externalSheetName, assemblyNum);
}

/**
 * Core logic for comparing a Master Sub-BOM to an External PDM BOM.
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
      externalBOMMap = createExternalBOMMap(externalSheet, BOM_CONFIG.PDM_HEADER_MAP);
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
    const itemNumber = newItem.mainRow[COL.ITEM_NUMBER];
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
      if (newItem.mainRow[COL.DESCRIPTION] !== oldItem.mainRow[COL.DESCRIPTION]) {
        const changeId = `COMP-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Description changed from "${oldItem.mainRow[COL.DESCRIPTION]}" to "${newItem.mainRow[COL.DESCRIPTION]}"`]);
        externalSheet.getRange(newItem.startRow, newItem.colIndexes[COL.DESCRIPTION] + 1).setBackground(highlightColor);
        itemModified = true;
      }
      if (newItem.mainRow[COL.ITEM_REV] !== oldItem.mainRow[COL.ITEM_REV]) {
        const changeId = `COMP-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Rev changed from "${oldItem.mainRow[COL.ITEM_REV]}" to "${newItem.mainRow[COL.ITEM_REV]}"`]);
        externalSheet.getRange(newItem.startRow, newItem.colIndexes[COL.ITEM_REV] + 1).setBackground(highlightColor);
        itemModified = true;
        logRevisionChange(itemNumber, String(oldItem.mainRow[COL.ITEM_REV]), String(newItem.mainRow[COL.ITEM_REV]), 'PDM Comparison');
      }
      // Note: Convert to String for comparison to handle numbers vs. text
      if (String(newItem.mainRow[COL.QTY]) !== String(oldItem.mainRow[COL.QTY])) {
        const changeId = `COMP-${changeCounter++}`;
        modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `Qty changed from "${oldItem.mainRow[COL.QTY]}" to "${newItem.mainRow[COL.QTY]}"`]);
        externalSheet.getRange(newItem.startRow, newItem.colIndexes[COL.QTY] + 1).setBackground(highlightColor);
        itemModified = true;
      }

      // Compare AML
      const oldAmlSet = new Set(oldItem.aml.map(a => `${a[COL.MFR_NAME]}|${a[COL.MFR_PN]}`));
      const newAmlSet = new Set(newItem.aml.map(a => `${a[COL.MFR_NAME]}|${a[COL.MFR_PN]}`));

      newItem.aml.forEach((aml, index) => {
        const amlString = `${aml[COL.MFR_NAME]}|${aml[COL.MFR_PN]}`;
        if (!oldAmlSet.has(amlString)) {
          const changeId = `COMP-${changeCounter++}`;
          modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, parentAssembly, `AML Added: ${aml[COL.MFR_NAME]} - ${aml[COL.MFR_PN]}`]);
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
         const amlString = `${aml[COL.MFR_NAME]}|${aml[COL.MFR_PN]}`;
         if (!newAmlSet.has(amlString)) {
             const changeId = `COMP-${changeCounter++}`;
             modifications.push([changeId, ecrNum, 'MODIFIED', itemNumber, newItem.parent, `AML Removed: ${aml[COL.MFR_NAME]} - ${aml[COL.MFR_PN]}`]);
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
    const itemNumber = oldItem.mainRow[COL.ITEM_NUMBER];
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
