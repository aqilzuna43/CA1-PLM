// ============================================================================
// PDMIMPORT.gs — PDM Children Import (Grafting) with Integrity Checks
// ============================================================================
// Imports child components from an external PDM export sheet into the Master
// BOM, automatically adjusting hierarchy levels relative to the selected
// parent assembly.
//
// Phase 1 Hardening:
//   - Pre-import validation against ITEMS sheet (orphan detection)
//   - Duplicate detection (PN already exists under same parent)
//   - Auto-ITEMS creation offer for new parts
//   - AML expansion for multi-vendor parts
//   - Import summary report
// ============================================================================

/**
 * Imports children from the PDM sheet to the Master BOM, adjusting levels automatically.
 * Includes pre-import validation, duplicate detection, and auto-ITEMS creation.
 */
function runImportPdmChildren() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterSheet = ss.getActiveSheet(); // User should run this from the Master BOM
  const pdmSheet = ss.getSheetByName(BOM_CONFIG.PDM_GRAFT_SHEET_NAME);

  if (!pdmSheet) {
    ui.alert(`Error: Could not find the PDM input sheet named "${BOM_CONFIG.PDM_GRAFT_SHEET_NAME}". Please create it and paste your PDM export there.`);
    return;
  }

  // 1. Get Selected Parent in Master BOM
  const activeRange = masterSheet.getActiveRange();
  const activeRowIndex = activeRange.getRow(); // 1-based
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData[0];

  // Map Master Columns
  const mCols = getColumnIndexes(masterHeaders);

  // Safety check for missing columns in Master BOM
  if (mCols[COL.ITEM_NUMBER] === -1 || mCols[COL.LEVEL] === -1) {
    ui.alert(`Error: Critical columns missing in Master Sheet.\n\nCould not find columns named:\n- "${COL.ITEM_NUMBER}"\n- "${COL.LEVEL}"\n\nPlease check the CONFIG section in the script and your sheet headers.`);
    return;
  }

  // Validate Master Selection
  if (activeRowIndex < 2 || activeRowIndex > masterData.length) {
    ui.alert("Please select a valid row containing a Parent Item.");
    return;
  }

  const parentRowData = masterData[activeRowIndex - 1]; // 0-based

  // Safe string conversion to prevent "reading 'toString' of undefined"
  const masterParentPNVal = parentRowData[mCols[COL.ITEM_NUMBER]];
  const masterParentLevelVal = parentRowData[mCols[COL.LEVEL]];

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

  // Map PDM Columns (Dynamic search based on BOM_CONFIG.PDM_GRAFT_HEADERS)
  const pCols = {};
  const missingPdmCols = [];
  for (const key in BOM_CONFIG.PDM_GRAFT_HEADERS) {
    pCols[key] = pdmHeaders.indexOf(BOM_CONFIG.PDM_GRAFT_HEADERS[key]);
    if (pCols[key] === -1) {
      missingPdmCols.push(BOM_CONFIG.PDM_GRAFT_HEADERS[key]);
    }
  }

  if (missingPdmCols.length > 0) {
    ui.alert(`Error: Missing columns in PDM Sheet "${BOM_CONFIG.PDM_GRAFT_SHEET_NAME}".\n\nCould not find: ${missingPdmCols.join(", ")}`);
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
    ui.alert(`Part Number "${masterParentPN}" not found in ${BOM_CONFIG.PDM_GRAFT_SHEET_NAME}.`);
    return;
  }

  // 4. Build lookups for validation
  const itemsMap = buildItemsLookup_(ss);
  const amlMap = buildAmlLookup_(ss);

  // Build set of existing children under this parent in MASTER
  const existingChildrenUnderParent = buildExistingChildrenSet_(masterData, mCols, masterParentPN, masterParentLevel);

  // 5. Collect Children, Validate, & Transform Levels
  const rowsToInsert = [];
  const importReport = {
    total: 0,
    inserted: 0,
    duplicatesSkipped: [],
    orphansDetected: [],
    itemsCreated: [],
    amlMissing: []
  };

  // Collect all candidate PNs first for batch validation
  const candidateRows = [];
  for (let j = pdmStartIndex + 1; j < pdmData.length; j++) {
    const currentRow = pdmData[j];
    const hierarchyVal = currentRow[pCols.HIERARCHY_COL];

    // Safety check: Empty hierarchy string usually means end of data
    if (hierarchyVal === "" || hierarchyVal === null) break;

    const currentPdmLevel = calculatePdmLevel(hierarchyVal);

    // STOP CONDITION: If we find a level equal to or higher than parent, left the sub-assembly.
    if (currentPdmLevel <= pdmParentLevel) break;

    candidateRows.push({ rowData: currentRow, pdmLevel: currentPdmLevel });
  }

  importReport.total = candidateRows.length;

  if (candidateRows.length === 0) {
    ui.alert("Found the parent in PDM, but it has no children listed below it.");
    return;
  }

  // --- PRE-IMPORT VALIDATION: Check for orphans and offer auto-create ---
  const orphanPNs = [];
  const allCandidatePNs = new Set();

  for (const candidate of candidateRows) {
    const pn = String(candidate.rowData[pCols.PN_COL]).trim();
    if (pn && !allCandidatePNs.has(pn)) {
      allCandidatePNs.add(pn);
      if (!itemsMap.has(pn)) {
        orphanPNs.push({
          pn: pn,
          desc: String(candidate.rowData[pCols.DESC_COL] || '').trim(),
          rev: String(candidate.rowData[pCols.REV_COL] || '').trim()
        });
      }
    }
  }

  if (orphanPNs.length > 0) {
    const orphanList = orphanPNs.slice(0, 10).map(o => `  ${o.pn}`).join('\n');
    const extra = orphanPNs.length > 10 ? `\n  ...and ${orphanPNs.length - 10} more` : '';

    const createChoice = ui.alert(
      'New Parts Detected',
      `${orphanPNs.length} part(s) from PDM are not in the ITEMS sheet:\n\n` +
      `${orphanList}${extra}\n\n` +
      'Would you like to auto-create ITEMS entries for these parts?\n\n' +
      'YES = Create entries and continue import\n' +
      'NO = Continue import without creating (parts will be flagged as orphans)\n' +
      'Cancel = Abort import',
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (createChoice === ui.Button.CANCEL) return;

    if (createChoice === ui.Button.YES) {
      const itemsSheet = ss.getSheetByName(BOM_CONFIG.ITEMS_SHEET_NAME);
      if (itemsSheet) {
        const newItemRows = orphanPNs.map(o => [o.pn, o.desc, o.rev || 'A', '']);
        if (newItemRows.length > 0) {
          itemsSheet.getRange(
            itemsSheet.getLastRow() + 1, 1,
            newItemRows.length, newItemRows[0].length
          ).setValues(newItemRows);
          importReport.itemsCreated = orphanPNs.map(o => o.pn);
        }
      }
    } else {
      importReport.orphansDetected = orphanPNs.map(o => o.pn);
    }
  }

  // --- BUILD ROWS WITH DUPLICATE DETECTION ---
  for (const candidate of candidateRows) {
    const currentRow = candidate.rowData;
    const currentPdmLevel = candidate.pdmLevel;
    const childPN = String(currentRow[pCols.PN_COL]).trim();

    // DUPLICATE DETECTION: Check if PN already exists as direct child of parent
    if (existingChildrenUnderParent.has(childPN) && currentPdmLevel === pdmParentLevel + 1) {
      importReport.duplicatesSkipped.push(childPN);
      continue; // Skip duplicate — don't re-add
    }

    // LEVEL TRANSFORMATION logic
    const relativeDepth = currentPdmLevel - pdmParentLevel;
    const newMasterLevel = masterParentLevel + relativeDepth;

    // Construct the new row for Master BOM
    const newRow = new Array(masterHeaders.length).fill("");

    // Fill mapped columns
    if (mCols[COL.LEVEL] > -1) newRow[mCols[COL.LEVEL]] = newMasterLevel;
    if (mCols[COL.ITEM_NUMBER] > -1) newRow[mCols[COL.ITEM_NUMBER]] = currentRow[pCols.PN_COL];
    if (mCols[COL.ITEM_REV] > -1) newRow[mCols[COL.ITEM_REV]] = currentRow[pCols.REV_COL];
    if (mCols[COL.DESCRIPTION] > -1) newRow[mCols[COL.DESCRIPTION]] = currentRow[pCols.DESC_COL];
    if (mCols[COL.QTY] > -1) newRow[mCols[COL.QTY]] = currentRow[pCols.QTY_COL];
    if (mCols[COL.MFR_NAME] > -1) newRow[mCols[COL.MFR_NAME]] = currentRow[pCols.VENDOR_COL];
    if (mCols[COL.MFR_PN] > -1) newRow[mCols[COL.MFR_PN]] = currentRow[pCols.MPN_COL];

    // Set default status
    if (mCols[COL.REFERENCE_NOTES] > -1) newRow[mCols[COL.REFERENCE_NOTES]] = "Pending Review";

    rowsToInsert.push(newRow);

    // Track AML status
    if (childPN && !amlMap.has(childPN)) {
      importReport.amlMissing.push(childPN);
    }
  }

  importReport.inserted = rowsToInsert.length;

  if (rowsToInsert.length === 0) {
    const msg = importReport.duplicatesSkipped.length > 0
      ? `All ${importReport.total} PDM children already exist under "${masterParentPN}". Nothing to import.`
      : "No rows to insert after validation.";
    ui.alert(msg);
    return;
  }

  // 6. Insert into Master BOM
  masterSheet.insertRowsAfter(activeRowIndex, rowsToInsert.length);
  masterSheet.getRange(activeRowIndex + 1, 1, rowsToInsert.length, rowsToInsert[0].length)
             .setValues(rowsToInsert);

  // Highlight new rows (light blue)
  masterSheet.getRange(activeRowIndex + 1, 1, rowsToInsert.length, masterSheet.getLastColumn())
             .setBackground("#e6f7ff");

  // Highlight orphans (light red) if they were not auto-created
  if (importReport.orphansDetected.length > 0 && mCols[COL.ITEM_NUMBER] > -1) {
    const orphanSet = new Set(importReport.orphansDetected);
    for (let r = 0; r < rowsToInsert.length; r++) {
      const pn = rowsToInsert[r][mCols[COL.ITEM_NUMBER]];
      if (pn && orphanSet.has(String(pn).trim())) {
        const cell = masterSheet.getRange(activeRowIndex + 1 + r, mCols[COL.ITEM_NUMBER] + 1);
        cell.setBackground('#f4cccc'); // Red — orphan
        cell.setNote('[BOM Validation] Imported from PDM but not found in ITEMS sheet.');
      }
    }
  }

  // 7. Display Import Summary Report
  const reportLines = [
    `=== PDM IMPORT REPORT ===`,
    `Parent: ${masterParentPN} (Level ${masterParentLevel})`,
    ``,
    `Total PDM children found: ${importReport.total}`,
    `Rows inserted: ${importReport.inserted}`,
    `Levels adjusted relative to Level ${masterParentLevel}`,
  ];

  if (importReport.duplicatesSkipped.length > 0) {
    const dupes = [...new Set(importReport.duplicatesSkipped)];
    reportLines.push(``);
    reportLines.push(`Duplicates skipped (${dupes.length}):`);
    dupes.slice(0, 10).forEach(pn => reportLines.push(`  ${pn}`));
    if (dupes.length > 10) reportLines.push(`  ...and ${dupes.length - 10} more`);
  }

  if (importReport.itemsCreated.length > 0) {
    reportLines.push(``);
    reportLines.push(`ITEMS entries created (${importReport.itemsCreated.length}):`);
    importReport.itemsCreated.slice(0, 10).forEach(pn => reportLines.push(`  ${pn}`));
    if (importReport.itemsCreated.length > 10) reportLines.push(`  ...and ${importReport.itemsCreated.length - 10} more`);
  }

  if (importReport.orphansDetected.length > 0) {
    reportLines.push(``);
    reportLines.push(`Orphan warnings (${importReport.orphansDetected.length}) — not in ITEMS:`);
    importReport.orphansDetected.slice(0, 10).forEach(pn => reportLines.push(`  ${pn}`));
    if (importReport.orphansDetected.length > 10) reportLines.push(`  ...and ${importReport.orphansDetected.length - 10} more`);
  }

  if (importReport.amlMissing.length > 0) {
    const uniqueAml = [...new Set(importReport.amlMissing)];
    reportLines.push(``);
    reportLines.push(`Parts without AML (${uniqueAml.length}) — run "Prepare Rows for AML":`);
    uniqueAml.slice(0, 10).forEach(pn => reportLines.push(`  ${pn}`));
    if (uniqueAml.length > 10) reportLines.push(`  ...and ${uniqueAml.length - 10} more`);
  }

  ui.alert('PDM Import Complete', reportLines.join('\n'), ui.ButtonSet.OK);
}


/**
 * Builds a Set of direct child PNs under a given parent in the MASTER data.
 * Used for duplicate detection during PDM import.
 *
 * @param {Array<Array>} masterData Full MASTER sheet data (2D array).
 * @param {Object} mCols Column index mapping from getColumnIndexes().
 * @param {string} parentPN Parent Item Number.
 * @param {number} parentLevel Parent's BOM Level.
 * @returns {Set<string>} Set of child part numbers.
 */
function buildExistingChildrenSet_(masterData, mCols, parentPN, parentLevel) {
  const children = new Set();
  let parentFound = false;

  for (let i = 1; i < masterData.length; i++) {
    const rowPN = masterData[i][mCols[COL.ITEM_NUMBER]];
    if (!rowPN) continue;

    const pnStr = String(rowPN).trim();
    const levelVal = masterData[i][mCols[COL.LEVEL]];
    const level = parseInt(levelVal, 10);

    if (!parentFound) {
      if (pnStr === parentPN && level === parentLevel) {
        parentFound = true;
      }
      continue;
    }

    // We're inside the parent's block
    if (isNaN(level) || level <= parentLevel) break; // Left the parent block

    if (level === parentLevel + 1) {
      children.add(pnStr); // Direct child
    }
  }

  return children;
}
