// ============================================================================
// RECONCILE.gs — Batch Data Sync & Sheet Protection
// ============================================================================
// Layer 3 batch tools for periodic full-sync and master data protection:
//   1. Reconcile Master Data — overwrite all managed columns from ITEMS/AML
//   2. Protect Master Sheets — warning-level protection on ITEMS/AML +
//      data-validation warnings on MASTER managed columns
// ============================================================================


// ---------------------
// 1. Reconcile Master Data
// ---------------------

/**
 * Two-phase batch sync of ALL managed columns on the MASTER sheet.
 *
 * Phase 1 — AML Row Repair:
 *   Scans MASTER to find parts with fewer AML rows than the AML tab expects.
 *   Inserts missing continuation rows and populates them with AML data.
 *
 * Phase 2 — Data Sync (multi-AML aware):
 *   Re-reads MASTER (since Phase 1 may have inserted rows), then overwrites
 *   Description/Rev/Lifecycle from ITEMS and Mfr/MPN from AML.
 *   Assigns correct AML entry to each row (main row → [0], continuation → [1], etc.)
 *   Also populates Reference Notes on continuation rows from the main part row.
 *
 * NOTE: setValues() does NOT fire onEdit triggers, so this won't
 * cause infinite loops with the real-time validation.
 */
function runReconcileMasterData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BOM_CONFIG.MASTER_SHEET_NAME);

  if (!sheet) {
    return ui.alert('Error', `Sheet "${BOM_CONFIG.MASTER_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
  }

  let data = sheet.getDataRange().getValues();
  if (data.length < 2) return ui.alert('The MASTER sheet has no data rows.');

  const headers = data[0];
  const colIdx = getColumnIndexes(headers);

  if (colIdx[COL.ITEM_NUMBER] === -1) {
    return ui.alert('Error', `Column "${COL.ITEM_NUMBER}" not found on MASTER sheet.`, ui.ButtonSet.OK);
  }

  // Build lookup maps
  const itemsMap = buildItemsLookup_(ss);
  const amlMap = buildAmlLookup_(ss);
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;

  if (itemsMap.size === 0) {
    return ui.alert('Warning', `ITEMS sheet is empty or not found. Cannot reconcile.`, ui.ButtonSet.OK);
  }

  const mfrCol = colIdx[COL.MFR_NAME];
  const mpnCol = colIdx[COL.MFR_PN];
  const refNotesCol = colIdx[COL.REFERENCE_NOTES];

  // ===================================================================
  // PHASE 1: AML Row Repair — insert missing continuation rows
  // ===================================================================
  let amlRowsInserted = 0;

  if (mfrCol !== -1 || mpnCol !== -1) {
    // Scan MASTER to find parts with missing AML rows
    // Collect: [{sheetRow, pn, actual, expected, refNotes}]
    const insertTasks = [];
    let currentPN = '';
    let currentMainSheetRow = -1;
    let currentRefNotes = '';
    let actualAmlCount = 0;

    for (let i = 1; i < data.length; i++) {
      const itemNumber = data[i][colIdx[COL.ITEM_NUMBER]];
      const hasPN = itemNumber && itemNumber.toString().trim() !== '';

      if (hasPN) {
        // Flush previous part
        flushAmlInsertTask_(insertTasks, amlMap, currentPN, currentMainSheetRow, actualAmlCount, currentRefNotes);

        // Start tracking new part
        currentPN = itemNumber.toString().trim();
        currentMainSheetRow = i + 1; // 1-based
        currentRefNotes = refNotesCol !== -1 && data[i][refNotesCol]
          ? data[i][refNotesCol].toString()
          : '';
        actualAmlCount = 0;

        // Count main row's own AML entry
        const mfrName = mfrCol !== -1 ? data[i][mfrCol] : null;
        const mfrPN = mpnCol !== -1 ? data[i][mpnCol] : null;
        if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
          actualAmlCount = 1;
        }
      } else if (currentPN) {
        // Blank Item Number — check if it's an AML continuation row
        const mfrName = mfrCol !== -1 ? data[i][mfrCol] : null;
        const mfrPN = mpnCol !== -1 ? data[i][mpnCol] : null;
        if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
          actualAmlCount++;
        } else {
          // Non-AML blank row — flush
          flushAmlInsertTask_(insertTasks, amlMap, currentPN, currentMainSheetRow, actualAmlCount, currentRefNotes);
          currentPN = '';
          currentMainSheetRow = -1;
          actualAmlCount = 0;
        }
      }
    }
    // Flush last part
    flushAmlInsertTask_(insertTasks, amlMap, currentPN, currentMainSheetRow, actualAmlCount, currentRefNotes);

    // Insert missing rows bottom-up (highest row first to prevent row-shift cascading)
    if (insertTasks.length > 0) {
      insertTasks.sort((a, b) => b.afterRow - a.afterRow);

      insertTasks.forEach(task => {
        // afterRow is the last existing row for this part (main row + existing continuations)
        sheet.insertRowsAfter(task.afterRow, task.rowsToInsert);
        amlRowsInserted += task.rowsToInsert;

        // Populate the newly inserted continuation rows
        const numCols = headers.length;
        for (let r = 0; r < task.rowsToInsert; r++) {
          const newRow = task.afterRow + 1 + r; // 1-based sheet row
          const blankRow = new Array(numCols).fill('');

          // Populate Mfr. Name, Mfr. Part Number, and Reference Notes
          const amlIdx = task.existingCount + r; // Which AML entry this row gets
          if (amlIdx < task.amlEntries.length) {
            if (mfrCol !== -1) blankRow[mfrCol] = task.amlEntries[amlIdx].mfr;
            if (mpnCol !== -1) blankRow[mpnCol] = task.amlEntries[amlIdx].mpn;
          }
          if (refNotesCol !== -1) blankRow[refNotesCol] = task.refNotes;

          sheet.getRange(newRow, 1, 1, numCols).setValues([blankRow]);
        }
      });

      // Re-read MASTER data since rows were inserted
      data = sheet.getDataRange().getValues();
    }
  }

  // ===================================================================
  // PHASE 2: Data Sync — multi-AML aware managed column overwrite
  // ===================================================================

  // Stats
  let rowsUpdated = 0;
  let staleValues = 0;
  let orphanPNs = [];
  let missingAML = [];
  let formulasReplaced = 0;

  const numRows = data.length - 1;
  const descCol = colIdx[COL.DESCRIPTION];
  const revCol = colIdx[COL.ITEM_REV];
  const lifeCol = colIdx[COL.LIFECYCLE];

  // Read formulas to detect VLOOKUP replacements
  const formulaRange = sheet.getRange(2, 1, numRows, headers.length);
  const formulas = formulaRange.getFormulas();

  // Build updated data arrays
  const updatedData = data.slice(1).map(row => [...row]);

  // Track current part for multi-AML assignment
  let trackPN = '';
  let trackAmlEntries = null;
  let trackAmlIdx = 0;        // Which AML entry the current row corresponds to
  let trackRefNotes = '';

  for (let i = 0; i < numRows; i++) {
    const pn = updatedData[i][colIdx[COL.ITEM_NUMBER]];
    const hasPn = pn && pn.toString().trim() !== '';

    if (hasPn) {
      // Main part row — new Item Number
      const pnStr = pn.toString().trim();
      const itemData = itemsMap.get(pnStr);
      const amlEntries = amlMap.get(pnStr);
      let rowChanged = false;

      trackPN = pnStr;
      trackAmlEntries = amlEntries || null;
      trackAmlIdx = 0;
      trackRefNotes = refNotesCol !== -1 && updatedData[i][refNotesCol]
        ? updatedData[i][refNotesCol].toString()
        : '';

      // --- ITEMS columns ---
      if (itemData) {
        const mappings = [
          { col: descCol, newVal: itemData.desc },
          { col: revCol, newVal: itemData.rev },
          { col: lifeCol, newVal: itemData.lifecycle }
        ];

        mappings.forEach(({ col, newVal }) => {
          if (col === -1) return;
          const currentVal = updatedData[i][col] ? updatedData[i][col].toString().trim() : '';
          const formula = formulas[i][col];

          if (formula && formula !== '') {
            updatedData[i][col] = newVal;
            formulasReplaced++;
            rowChanged = true;
          } else if (currentVal !== newVal.trim()) {
            updatedData[i][col] = newVal;
            staleValues++;
            rowChanged = true;
          }
        });
      } else {
        if (pnStr) orphanPNs.push(pnStr);
      }

      // --- AML columns (entry [0] for main row) ---
      if (trackAmlEntries && trackAmlEntries.length > 0) {
        const amlMappings = [
          { col: mfrCol, newVal: trackAmlEntries[0].mfr },
          { col: mpnCol, newVal: trackAmlEntries[0].mpn }
        ];

        amlMappings.forEach(({ col, newVal }) => {
          if (col === -1) return;
          const currentVal = updatedData[i][col] ? updatedData[i][col].toString().trim() : '';
          const formula = formulas[i][col];

          if (formula && formula !== '') {
            updatedData[i][col] = newVal;
            formulasReplaced++;
            rowChanged = true;
          } else if (currentVal !== newVal.trim()) {
            updatedData[i][col] = newVal;
            staleValues++;
            rowChanged = true;
          }
        });
        trackAmlIdx = 1; // Next continuation row gets entry [1]
      } else if (itemData) {
        missingAML.push(pnStr);
      }

      if (rowChanged) rowsUpdated++;

    } else if (trackPN && trackAmlEntries && trackAmlIdx < trackAmlEntries.length) {
      // AML continuation row — blank Item Number, assign next AML entry
      const mfrName = mfrCol !== -1 ? updatedData[i][mfrCol] : null;
      const mfrPN = mpnCol !== -1 ? updatedData[i][mpnCol] : null;
      const isContinuationRow = (mfrName && mfrName.toString().trim() !== '')
        || (mfrPN && mfrPN.toString().trim() !== '');

      if (isContinuationRow || trackAmlIdx < trackAmlEntries.length) {
        let rowChanged = false;
        const entry = trackAmlEntries[trackAmlIdx];

        // Overwrite Mfr/MPN with correct AML entry
        if (mfrCol !== -1) {
          const currentMfr = updatedData[i][mfrCol] ? updatedData[i][mfrCol].toString().trim() : '';
          if (currentMfr !== entry.mfr.trim()) {
            updatedData[i][mfrCol] = entry.mfr;
            staleValues++;
            rowChanged = true;
          }
        }
        if (mpnCol !== -1) {
          const currentMpn = updatedData[i][mpnCol] ? updatedData[i][mpnCol].toString().trim() : '';
          if (currentMpn !== entry.mpn.trim()) {
            updatedData[i][mpnCol] = entry.mpn;
            staleValues++;
            rowChanged = true;
          }
        }

        // Populate Reference Notes from main part row
        if (refNotesCol !== -1 && trackRefNotes) {
          const currentRef = updatedData[i][refNotesCol] ? updatedData[i][refNotesCol].toString().trim() : '';
          if (currentRef !== trackRefNotes.trim()) {
            updatedData[i][refNotesCol] = trackRefNotes;
            rowChanged = true;
          }
        }

        if (rowChanged) rowsUpdated++;
        trackAmlIdx++;
      }
    } else {
      // Blank row with no AML context — reset tracking
      trackPN = '';
      trackAmlEntries = null;
      trackAmlIdx = 0;
    }
  }

  // --- Batch write all data at once ---
  formulaRange.setValues(updatedData);

  // --- Clear any AML mismatch error flags (Phase 1 fixed them) ---
  if (amlRowsInserted > 0 && mfrCol !== -1) {
    const mfrRange = sheet.getRange(2, mfrCol + 1, numRows, 1);
    const mfrNotes = mfrRange.getNotes();
    for (let i = 0; i < numRows; i++) {
      if (mfrNotes[i][0] && mfrNotes[i][0].includes('AML row mismatch')) {
        const cell = sheet.getRange(i + 2, mfrCol + 1);
        cell.setNote('');
        cell.setBackground(null);
      }
    }
  }

  // --- Apply background colors for orphans and missing AML ---
  const pnCol = colIdx[COL.ITEM_NUMBER] + 1;
  const errorColor = BOM_CONFIG.VALIDATION.COLORS.ERROR;
  const warningColor = BOM_CONFIG.VALIDATION.COLORS.WARNING;

  if (orphanPNs.length > 0) {
    const orphanSet = new Set(orphanPNs);
    for (let i = 0; i < numRows; i++) {
      const pn = updatedData[i][colIdx[COL.ITEM_NUMBER]];
      if (pn && orphanSet.has(pn.toString().trim())) {
        const cell = sheet.getRange(i + 2, pnCol);
        cell.setBackground(errorColor);
        cell.setNote(prefix + 'Orphaned: Item Number not found in ITEMS sheet.');
      }
    }
  }

  if (missingAML.length > 0 && mfrCol !== -1) {
    const missingSet = new Set(missingAML);
    for (let i = 0; i < numRows; i++) {
      const pn = updatedData[i][colIdx[COL.ITEM_NUMBER]];
      if (pn && missingSet.has(pn.toString().trim())) {
        const cell = sheet.getRange(i + 2, mfrCol + 1);
        cell.setBackground(warningColor);
        cell.setNote(prefix + 'No AML entry found for this Item Number.');
      }
    }
  }

  // --- Summary ---
  const uniqueOrphans = [...new Set(orphanPNs)];
  const uniqueMissingAML = [...new Set(missingAML)];

  let summary = `Reconciliation Complete!\n\n`;
  summary += `• ${rowsUpdated} row(s) updated\n`;
  summary += `• ${amlRowsInserted} AML continuation row(s) inserted\n`;
  summary += `• ${formulasReplaced} VLOOKUP formula(s) replaced with values\n`;
  summary += `• ${staleValues} stale value(s) corrected\n`;
  summary += `• ${uniqueOrphans.length} orphaned PN(s) flagged (red)\n`;
  summary += `• ${uniqueMissingAML.length} PN(s) missing AML entries (yellow)\n`;

  if (uniqueOrphans.length > 0) {
    summary += `\nOrphaned PNs:\n${uniqueOrphans.slice(0, 10).join(', ')}`;
    if (uniqueOrphans.length > 10) summary += ` ... and ${uniqueOrphans.length - 10} more`;
  }

  ui.alert('Reconcile Master Data', summary, ui.ButtonSet.OK);
}


/**
 * Helper: checks if a part has missing AML rows and adds an insert task.
 */
function flushAmlInsertTask_(insertTasks, amlMap, pn, mainSheetRow, actualCount, refNotes) {
  if (!pn || mainSheetRow < 1) return;
  const expectedEntries = amlMap.get(pn);
  if (!expectedEntries || expectedEntries.length <= 1) return;
  if (actualCount >= expectedEntries.length) return;

  const rowsToInsert = expectedEntries.length - actualCount;
  // afterRow = main row + existing continuation rows (insert after the last existing row for this part)
  const afterRow = mainSheetRow + (actualCount - 1);

  insertTasks.push({
    afterRow: afterRow > 0 ? afterRow : mainSheetRow,
    rowsToInsert: rowsToInsert,
    pn: pn,
    existingCount: actualCount,
    amlEntries: expectedEntries,
    refNotes: refNotes
  });
}




// ---------------------
// 2. Protect Master Sheets
// ---------------------

/**
 * Applies warning-level protection on ITEMS and AML sheets, and
 * adds data-validation warnings on managed columns of the MASTER sheet.
 *
 * Warning-level protection shows a confirmation dialog when users
 * try to edit but does NOT block them (unlike full protection).
 * Data validation warnings show a small triangle icon with help text.
 *
 * IMPORTANT: Removes ALL existing protections (including manually-applied
 * hard protections) before applying warning-only. Hard protections block
 * scripts like ECR commitToMaster() from writing to ITEMS/AML sheets.
 */
function runProtectMasterSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let actionsPerformed = [];

  // --- Protect ITEMS sheet ---
  const itemsSheet = ss.getSheetByName(BOM_CONFIG.ITEMS_SHEET_NAME);
  if (itemsSheet) {
    // Remove ALL existing protections: sheet-level AND range-level
    let removedCount = 0;
    itemsSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
      try { p.remove(); removedCount++; } catch (e) { /* skip if not owner */ }
    });
    itemsSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
      try { p.remove(); removedCount++; } catch (e) { /* skip if not owner */ }
    });

    const protection = itemsSheet.protect()
      .setDescription('[BOM Tools] ITEMS sheet — edit with caution');
    protection.setWarningOnly(true);
    actionsPerformed.push(`ITEMS sheet: Warning-level protection applied` +
      (removedCount > 0 ? ` (replaced ${removedCount} existing protection(s))` : ''));
  } else {
    actionsPerformed.push(`ITEMS sheet: Not found (skipped)`);
  }

  // --- Protect AML sheet ---
  const amlSheet = ss.getSheetByName(BOM_CONFIG.AML_SHEET_NAME);
  if (amlSheet) {
    let removedCount = 0;
    amlSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
      try { p.remove(); removedCount++; } catch (e) { /* skip if not owner */ }
    });
    amlSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
      try { p.remove(); removedCount++; } catch (e) { /* skip if not owner */ }
    });

    const protection = amlSheet.protect()
      .setDescription('[BOM Tools] AML sheet — edit with caution');
    protection.setWarningOnly(true);
    actionsPerformed.push(`AML sheet: Warning-level protection applied` +
      (removedCount > 0 ? ` (replaced ${removedCount} existing protection(s))` : ''));
  } else {
    actionsPerformed.push(`AML sheet: Not found (skipped)`);
  }

  // --- Clean up MASTER sheet protections + apply data validation warnings ---
  const masterSheet = ss.getSheetByName(BOM_CONFIG.MASTER_SHEET_NAME);
  if (masterSheet) {
    // Remove ALL existing protections on MASTER (sheet-level AND range-level)
    // These block cross-project scripts like ECR commitToMaster() from writing/deleting rows
    let masterRemovedCount = 0;
    const masterSheetProtections = masterSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    masterSheetProtections.forEach(p => {
      try { p.remove(); masterRemovedCount++; } catch (e) { /* skip if not owner */ }
    });
    const masterRangeProtections = masterSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    masterRangeProtections.forEach(p => {
      try { p.remove(); masterRemovedCount++; } catch (e) { /* skip if not owner */ }
    });

    if (masterRemovedCount > 0) {
      actionsPerformed.push(`MASTER sheet: Removed ${masterRemovedCount} existing protection(s)`);
    }

    // Apply data validation warnings on managed columns (cosmetic only, does not block)
    const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    const allManaged = [...BOM_CONFIG.MANAGED_COLUMNS.FROM_ITEMS, ...BOM_CONFIG.MANAGED_COLUMNS.FROM_AML];
    const lastRow = masterSheet.getLastRow();
    let dvCount = 0;

    if (lastRow > 1) {
      allManaged.forEach(colName => {
        const colIndex = headers.indexOf(colName);
        if (colIndex === -1) return;

        const range = masterSheet.getRange(2, colIndex + 1, lastRow - 1, 1);

        // Create a validation rule that always shows a warning triangle
        // using requireFormulaSatisfied with =TRUE (always passes, but shows help text)
        const rule = SpreadsheetApp.newDataValidation()
          .requireFormulaSatisfied('=TRUE')
          .setAllowInvalid(true) // Allow any input (warning only, no blocking)
          .setHelpText(`Auto-managed by BOM Tools from ${
            BOM_CONFIG.MANAGED_COLUMNS.FROM_ITEMS.includes(colName) ? 'ITEMS' : 'AML'
          } sheet. Manual edits will be restored on next sync.`)
          .build();

        range.setDataValidation(rule);
        dvCount++;
      });
    }

    actionsPerformed.push(`MASTER sheet: Data validation warnings on ${dvCount} managed column(s)`);
  } else {
    actionsPerformed.push(`MASTER sheet: Not found (skipped)`);
  }

  ui.alert('Master Sheets Protected',
    actionsPerformed.join('\n') +
    '\n\nWarning-level protection shows a confirmation dialog but does NOT block editing.\n' +
    'Data validation triangles indicate auto-managed columns.',
    ui.ButtonSet.OK);
}
