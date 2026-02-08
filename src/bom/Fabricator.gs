// ============================================================================
// FABRICATOR.gs — Fabricator BOM Generation
// ============================================================================
// Generates flat, filtered BOM sheets for manufacturing. Extracts only BUY
// and REF items from specified assemblies, normalising levels and writing
// each assembly to its own sheet in a single batch operation.
// ============================================================================

function runGenerateFabricatorBOMs() {
  const ui = SpreadsheetApp.getUi();
  const sourceSheetName = promptWithValidation('Generate Fabricator BOMs', 'Enter the name of the source Master BOM sheet:');
  if (!sourceSheetName) return;

  const assemblyInput = promptWithValidation('Generate Fabricator BOMs', 'Enter one or more assembly part numbers, separated by commas:');
  if (!assemblyInput) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Sheet "${sourceSheetName}" not found.`);

  const assemblyNumbers = assemblyInput.split(',').map(pn => pn.trim()).filter(Boolean);
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData[0];
  const colIndexes = getColumnIndexes(headers);

  if (colIndexes[COL.LEVEL] === -1 || colIndexes[COL.ITEM_NUMBER] === -1 || colIndexes[COL.ITEM_REV] === -1 || colIndexes[COL.REFERENCE_NOTES] === -1)
    return ui.alert('A required column (Level, Item Number, Item Rev, or Reference Notes) was not found.');
  let sheetsCreated = 0;
  assemblyNumbers.forEach(assemblyNum => {
    let parentRowIndex = -1;
    for (let i = 1; i < sourceData.length; i++) {
      if (sourceData[i][colIndexes[COL.ITEM_NUMBER]] && sourceData[i][colIndexes[COL.ITEM_NUMBER]].toString().trim() === assemblyNum) {
        parentRowIndex = i;
        break;
      }
    }

    if (parentRowIndex === -1) {
      ui.alert(`Assembly "${assemblyNum}" not found in the source BOM. Skipping.`);
      return;
    }


    const parentRow = sourceData[parentRowIndex];
    const parentLevel = parseInt(parentRow[colIndexes[COL.LEVEL]], 10);
    const parentRev = parentRow[colIndexes[COL.ITEM_REV]];
    const newSheetName = `Fabricator_${assemblyNum}_${parentRev}`;

    let newSheet = ss.getSheetByName(newSheetName);
    if (newSheet) newSheet.clear();
    else newSheet = ss.insertSheet(newSheetName);

    // Collect all rows in memory first, then write in a single batch (4-15x faster)
    const outputRows = [];
    outputRows.push(headers);

    let parentForSupplier = [...parentRow];
    parentForSupplier[colIndexes[COL.LEVEL]] = 0;
    outputRows.push(parentForSupplier);

    for (let i = parentRowIndex + 1; i < sourceData.length; i++) {
      const childRow = sourceData[i];
      const childLevel = parseInt(childRow[colIndexes[COL.LEVEL]], 10);
      if (isNaN(childLevel) || childLevel <= parentLevel) break;

      const statusCell = childRow[colIndexes[COL.REFERENCE_NOTES]];
      const status = statusCell ? statusCell.toString().toUpperCase().trim() : '';
      if (status.includes('BUY') || status.includes('REF')) {
        let childForSupplier = [...childRow];
        childForSupplier[colIndexes[COL.LEVEL]] = childLevel - parentLevel;
        outputRows.push(childForSupplier);
      }
    }

    // Single batch write instead of row-by-row appendRow calls
    if (outputRows.length > 0) {
      newSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
    }
    sheetsCreated++;
  });
  if (sheetsCreated > 0) {
    ui.alert(`${sheetsCreated} Fabricator BOM sheet(s) have been created successfully.`);
  } else {
    ui.alert('No Fabricator BOM sheets were created (check assembly numbers?).');
  }
}
