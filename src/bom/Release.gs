// ============================================================================
// RELEASE.gs — BOM Release & Effectivity
// ============================================================================
// Finalize-and-Release process (copy, filter REF rows, protect, clean up
// change-tracking columns) and BOM Effectivity Date stamping.
// ============================================================================

// ---------------------
// 1. Finalize and Release New BOM
// ---------------------

function runReleaseNewBOM() {
  const ui = SpreadsheetApp.getUi();

  // 1. Prompt for inputs (with validation)
  const wipSheetInput = promptWithValidation('Release New BOM', 'Enter the name of the approved WIP sheet:');
  if (!wipSheetInput) return;

  const newRevInput = promptWithValidation('Release New BOM', 'Enter the name for the new RELEASED sheet:',
    { pattern: /^[A-Za-z0-9_\-\.\s]+$/, patternHint: 'Sheet name can only contain letters, numbers, spaces, hyphens, underscores, and dots.' });
  if (!newRevInput) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wipSheet = ss.getSheetByName(wipSheetInput);

  // Validations
  if (!wipSheet) return ui.alert(`Error: Sheet "${wipSheetInput}" not found.`);
  if (ss.getSheetByName(newRevInput)) return ui.alert(`A sheet named "${newRevInput}" already exists.`);

  // --- Pre-release lifecycle audit ---
  // Check for OBSOLETE/EOL components before allowing release
  const wipData = wipSheet.getDataRange().getValues();
  const wipHeaders = wipData.length > 0 ? wipData[0] : [];
  const wipColIdx = getColumnIndexes(wipHeaders);

  if (wipColIdx[COL.LIFECYCLE] !== -1 && wipColIdx[COL.ITEM_NUMBER] !== -1) {
    const nonProductionStatuses = ['OBSOLETE', 'EOL', 'END OF LIFE', 'NRND', 'NOT RECOMMENDED'];
    const blockers = [];

    for (let i = 1; i < wipData.length; i++) {
      const lifecycleCell = wipData[i][wipColIdx[COL.LIFECYCLE]];
      const lifecycleStatus = lifecycleCell ? lifecycleCell.toString().toUpperCase().trim() : '';
      const itemNumber = wipData[i][wipColIdx[COL.ITEM_NUMBER]];
      if (itemNumber && lifecycleStatus && nonProductionStatuses.includes(lifecycleStatus)) {
        blockers.push(`  • ${itemNumber} — ${lifecycleCell}`);
      }
    }

    if (blockers.length > 0) {
      const proceed = ui.alert('⚠ Lifecycle Warning',
        `The following ${blockers.length} component(s) have non-production lifecycle status:\n\n${blockers.slice(0, 15).join('\n')}` +
        (blockers.length > 15 ? `\n  ... and ${blockers.length - 15} more` : '') +
        '\n\nDo you still want to proceed with the release?',
        ui.ButtonSet.YES_NO);
      if (proceed !== ui.Button.YES) return;
    }
  }

  // 2. Create the new sheet by copying
  const newSheet = wipSheet.copyTo(ss).setName(newRevInput);
  ss.setActiveSheet(newSheet);

  // --- OPTIMIZATION SECTION START ---

  // 3. Batch Delete 'REF' Rows (Prevents Timeout)
  const data = newSheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : [];
  const statusColIndex = headers.indexOf(COL.REFERENCE_NOTES);

  if (statusColIndex === -1) {
    ui.alert('Warning', `Column "${COL.REFERENCE_NOTES}" not found. Could not filter for internal view.`, ui.ButtonSet.OK);
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
    if (BOM_CONFIG.CHANGE_TRACKING_COLS_TO_DELETE.includes(header)) {
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
  const protection = newSheet.protect().setDescription(`Released BOM ${newRevInput}`);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);

  ui.alert('Success!', `New BOM revision "${newRevInput}" has been created, filtered, cleaned (Column R limit), and protected.`, ui.ButtonSet.OK);
}


// ---------------------
// 2. BOM Effectivity Dates
// ---------------------

/**
 * Adds "Effective From" and "Effective Until" columns to a BOM sheet
 * if they don't already exist, then populates "Effective From" with today's date
 * for all rows that have an Item Number but no existing effectivity date.
 *
 * Usage: Run this after releasing a new BOM revision to stamp effectivity dates.
 * Parts being phased out should have their "Effective Until" set manually or via ECR.
 */
function runSetEffectivityDates() {
  const sheetName = promptWithValidation('Set BOM Effectivity Dates',
    'Enter the BOM sheet name to add/update effectivity dates:');
  if (!sheetName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return ui.alert(`Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ui.alert('Sheet has no data rows.');

  const headers = data[0];
  const colIdx = getColumnIndexes(headers);
  const effFrom = BOM_CONFIG.EFFECTIVITY.EFFECTIVE_FROM;
  const effUntil = BOM_CONFIG.EFFECTIVITY.EFFECTIVE_UNTIL;

  // Add columns if they don't exist
  let effFromIdx = headers.indexOf(effFrom);
  let effUntilIdx = headers.indexOf(effUntil);
  let colsAdded = 0;

  if (effFromIdx === -1) {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(effFrom).setFontWeight('bold');
    effFromIdx = newCol - 1;
    colsAdded++;
  }

  if (effUntilIdx === -1) {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(effUntil).setFontWeight('bold');
    effUntilIdx = newCol - 1;
    colsAdded++;
  }

  // Re-read data if columns were added
  const freshData = colsAdded > 0 ? sheet.getDataRange().getValues() : data;
  const today = new Date();
  let stamped = 0;

  // Read existing "Effective From" values, stamp where missing
  if (colIdx[COL.ITEM_NUMBER] !== -1) {
    const effFromValues = sheet.getRange(2, effFromIdx + 1, freshData.length - 1, 1).getValues();

    for (let i = 1; i < freshData.length; i++) {
      const itemNumber = freshData[i][colIdx[COL.ITEM_NUMBER]];
      if (itemNumber && itemNumber.toString().trim() !== '') {
        const existingDate = effFromValues[i - 1][0];
        if (!existingDate || existingDate === '') {
          effFromValues[i - 1][0] = today;
          stamped++;
        }
      }
    }

    sheet.getRange(2, effFromIdx + 1, effFromValues.length, 1).setValues(effFromValues);

    // Format date columns
    sheet.getRange(2, effFromIdx + 1, freshData.length - 1, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, effUntilIdx + 1, freshData.length - 1, 1).setNumberFormat('yyyy-mm-dd');
  }

  ui.alert('Effectivity Dates Updated',
    `${colsAdded > 0 ? `Added ${colsAdded} new column(s). ` : ''}` +
    `Stamped "Effective From" date (${today.toLocaleDateString()}) on ${stamped} row(s).\n\n` +
    'Set "Effective Until" manually for parts being phased out.',
    ui.ButtonSet.OK);
}
