// ============================================================================
// AUDIT.gs — BOM Audit & Quality Tools
// ============================================================================
// Four audit tools for data-quality checks:
//   1. Lifecycle Status — flags OBSOLETE / EOL / NRND components
//   2. Structural Integrity — detects assemblies with inconsistent children
//   3. MAKE Item Screen — finds MAKE items that contain REF children
//   4. Comprehensive BOM Validation — 10-check scan with categorized report
// ============================================================================

// ---------------------
// 1. MAKE Items with REF Children
// ---------------------

function runScreenMakeItemsWithRef() {
  const sourceSheetName = promptWithValidation("List 'MAKE' Items with 'REF' Children", 'Enter the name of the BOM sheet to audit:');
  if (!sourceSheetName) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Sheet "${sourceSheetName}" not found.`);

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers);
  if (colIndexes[COL.LEVEL] === -1 || colIndexes[COL.ITEM_NUMBER] === -1 || colIndexes[COL.REFERENCE_NOTES] === -1)
    return ui.alert('A required column was not found. Please check CONFIG settings.');
  const issues = new Set();

  for (let i = 1; i < data.length; i++) {
    const parentRow = data[i];
    const parentStatusCell = parentRow[colIndexes[COL.REFERENCE_NOTES]];
    const parentStatus = parentStatusCell ? parentStatusCell.toString().toUpperCase().trim() : '';
    if (parentStatus.includes('MAKE')) {
      const parentLevel = parseInt(parentRow[colIndexes[COL.LEVEL]], 10);
      const parentItemNumber = parentRow[colIndexes[COL.ITEM_NUMBER]];
      if (isNaN(parentLevel) || !parentItemNumber) continue;

      for (let j = i + 1; j < data.length; j++) {
        const childRow = data[j];
        const childLevel = parseInt(childRow[colIndexes[COL.LEVEL]], 10);

        if (isNaN(childLevel) || childLevel <= parentLevel) break;

        const childStatusCell = childRow[colIndexes[COL.REFERENCE_NOTES]];
        const childStatus = childStatusCell ? childStatusCell.toString().toUpperCase().trim() : '';
        if (childStatus.includes('REF')) {
          issues.add(parentItemNumber);
          break;
        }
      }
    }
  }

  if (issues.size > 0) {
    const reportSheetName = `MAKE_REF_Audit_List`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
    reportSheet.clear();
    const reportHeaders = ["'MAKE' Items with 'REF' Children"];
    const outputData = [reportHeaders, ...Array.from(issues).map(item => [item])];
    reportSheet.getRange(1, 1, 1, 1).setValues([reportHeaders]).setFontWeight('bold');
    if (outputData.length > 1) {
      reportSheet.getRange(2, 1, outputData.length - 1, 1).setValues(outputData.slice(1));
    }
    reportSheet.autoResizeColumn(1);
    ui.alert('Audit complete!', `Found ${issues.size} issue(s). See the "${reportSheetName}" sheet for the list.`, ui.ButtonSet.OK);
  } else {
    ui.alert('Audit complete!', "No issues found. No 'MAKE' items have 'REF' children.", ui.ButtonSet.OK);
  }
}


// ---------------------
// 2. Lifecycle Status Audit
// ---------------------

function runAuditBOMLifecycle() {
  const sourceSheetName = promptWithValidation('Audit BOM Lifecycle Status', 'Enter the name of the BOM sheet to audit:');
  if (!sourceSheetName) return;
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
  if (colIndexes[COL.ITEM_NUMBER] === -1 || colIndexes[COL.DESCRIPTION] === -1 || colIndexes[COL.LIFECYCLE] === -1)
    return ui.alert(`A required column ('${COL.ITEM_NUMBER}', '${COL.DESCRIPTION}', or '${COL.LIFECYCLE}') was not found.`);
  const issues = [];
  const nonProductionStatuses = ['OBSOLETE', 'EOL', 'END OF LIFE', 'NRND', 'NOT RECOMMENDED'];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const lifecycleCell = row[colIndexes[COL.LIFECYCLE]];
    const lifecycleStatus = lifecycleCell ? lifecycleCell.toString().toUpperCase().trim() : '';
    const itemNumber = row[colIndexes[COL.ITEM_NUMBER]];
    if (itemNumber && lifecycleStatus && nonProductionStatuses.includes(lifecycleStatus)) {
      issues.push([
        itemNumber,
        row[colIndexes[COL.DESCRIPTION]],
        row[colIndexes[COL.LIFECYCLE]],
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


// ---------------------
// 3. Structural Integrity Audit
// ---------------------

/**
 * BOM Integrity Check: Finds assemblies (MAKE items) that appear multiple times in the BOM
 * but have DIFFERENT child structures at each occurrence.
 *
 * WHY: A part number used in multiple subassemblies is perfectly normal (e.g., a common
 * connector used in 3 different boards). However, if that part is itself an ASSEMBLY
 * (i.e., it has children), then its children MUST be identical everywhere it appears.
 * If PN 1233 has children [A, B, C] under parent 12345 but [A, B, D] under parent 12354,
 * that's a structural inconsistency — the BOM is broken.
 *
 * WHAT IT CHECKS:
 *   1. Finds every PN that appears as a parent (has children below it)
 *   2. For each occurrence, collects its direct child list (PN + Qty)
 *   3. Compares child lists across all occurrences
 *   4. Reports mismatches as integrity violations
 */
function runAuditDuplicatePartNumbers() {
  const sourceSheetName = promptWithValidation('BOM Structural Integrity Audit', 'Enter the BOM sheet name to audit:');
  if (!sourceSheetName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Sheet "${sourceSheetName}" not found.`);

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const colIndexes = getColumnIndexes(headers);

  if (colIndexes[COL.LEVEL] === -1 || colIndexes[COL.ITEM_NUMBER] === -1)
    return ui.alert('Required columns (Level, Item Number) not found.');

  const hasQty = colIndexes[COL.QTY] !== -1;

  // PASS 1: Parse BOM into structured rows
  const bomRows = []; // [{row, level, pn}]
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemNumber = row[colIndexes[COL.ITEM_NUMBER]];
    if (!itemNumber || itemNumber.toString().trim() === '') continue;

    const levelVal = row[colIndexes[COL.LEVEL]];
    const level = parseInt(levelVal, 10);
    if (isNaN(level)) continue;

    const qty = hasQty ? String(row[colIndexes[COL.QTY]]).trim() : '';
    bomRows.push({ dataIndex: i, sheetRow: i + 1, level: level, pn: itemNumber.toString().trim(), qty: qty });
  }

  // PASS 2: For each row that IS a parent (has children at level+1),
  // collect its direct children as a "child signature"
  const assemblyChildMap = new Map();

  for (let idx = 0; idx < bomRows.length; idx++) {
    const current = bomRows[idx];
    const directChildren = [];

    // Look ahead for direct children (level === current.level + 1)
    for (let j = idx + 1; j < bomRows.length; j++) {
      const next = bomRows[j];
      if (next.level <= current.level) break; // Exited this parent's block
      if (next.level === current.level + 1) {
        // Direct child
        directChildren.push(hasQty ? `${next.pn}:${next.qty}` : next.pn);
      }
    }

    // Only track this PN if it HAS children (it's an assembly/MAKE item)
    if (directChildren.length > 0) {
      // Sort to make signature order-independent (BOM order may vary)
      const signature = directChildren.slice().sort().join(' | ');

      if (!assemblyChildMap.has(current.pn)) assemblyChildMap.set(current.pn, []);
      assemblyChildMap.get(current.pn).push({
        sheetRow: current.sheetRow,
        level: current.level,
        signature: signature,
        children: directChildren
      });
    }
  }

  // PASS 3: Find mismatches — assemblies that appear multiple times with DIFFERENT child structures
  const mismatches = [];
  const mismatchRows = new Set();

  assemblyChildMap.forEach((occurrences, pn) => {
    if (occurrences.length < 2) return; // Only one occurrence, nothing to compare

    // Group by signature
    const signatureGroups = new Map();
    occurrences.forEach(occ => {
      if (!signatureGroups.has(occ.signature)) signatureGroups.set(occ.signature, []);
      signatureGroups.get(occ.signature).push(occ);
    });

    if (signatureGroups.size > 1) {
      // MISMATCH FOUND — same PN has different child structures
      let groupIndex = 1;
      signatureGroups.forEach((group, signature) => {
        const rows = group.map(g => g.sheetRow).join(', ');
        const childList = group[0].children.join(', ');
        mismatches.push([
          pn,
          `Variant ${groupIndex}`,
          group.length,
          rows,
          childList
        ]);
        group.forEach(g => mismatchRows.add(g.sheetRow));
        groupIndex++;
      });
    }
  });

  // Generate report
  if (mismatches.length > 0) {
    const reportSheetName = `Integrity_Audit_${sourceSheetName}`;
    let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
    reportSheet.clear();

    const reportHeaders = ['Assembly PN', 'Structure Variant', 'Occurrences', 'Found at Rows', 'Direct Children (PN:Qty)'];
    const outputData = [reportHeaders, ...mismatches];
    reportSheet.getRange(1, 1, outputData.length, reportHeaders.length).setValues(outputData);
    reportSheet.getRange(1, 1, 1, reportHeaders.length).setFontWeight('bold');
    reportSheet.autoResizeColumns(1, reportHeaders.length);

    // Highlight mismatched rows on source sheet
    const highlightColor = '#fff3e0'; // Light orange for structural mismatch
    mismatchRows.forEach(rowNum => {
      try {
        sourceSheet.getRange(rowNum, 1, 1, sourceSheet.getLastColumn()).setBackground(highlightColor);
      } catch (e) { /* skip */ }
    });

    // Count unique PNs with issues
    const uniqueMismatchPNs = new Set(mismatches.map(m => m[0]));
    ui.alert('Integrity Audit Complete!',
      `Found ${uniqueMismatchPNs.size} assembly(ies) with inconsistent child structures.\n\n` +
      `See "${reportSheetName}" for details.\nAffected rows highlighted in orange on source sheet.`,
      ui.ButtonSet.OK);
  } else {
    // Also report summary stats
    const totalAssemblies = assemblyChildMap.size;
    const multiUse = Array.from(assemblyChildMap.values()).filter(v => v.length > 1).length;
    ui.alert('Integrity Audit Complete!',
      `All clear! BOM structure is consistent.\n\n` +
      `• ${totalAssemblies} unique assemblies found\n` +
      `• ${multiUse} assembly(ies) used in multiple locations — all have matching child structures.`,
      ui.ButtonSet.OK);
  }
}


// ---------------------
// 4. Comprehensive BOM Validation (9-Check Scan)
// ---------------------

/**
 * Runs a comprehensive 10-check validation scan on a BOM sheet.
 * Cross-references ITEMS and AML sheets for data integrity.
 * Generates a categorized report sheet with summary + color-coded rows.
 *
 * Checks:
 *   1. Orphaned parts (MASTER PN not in ITEMS)
 *   2. Missing AML entries
 *   3. Level hierarchy gaps
 *   4. Duplicate assemblies with different children (structural)
 *   5. Stale managed values (MASTER ≠ ITEMS/AML)
 *   6. MAKE items with REF children
 *   7. Lifecycle warnings (OBSOLETE/EOL/NRND)
 *   8. Circular dependency detection
 *   9. Blank Item Number on rows with Level
 *  10. AML row count mismatch (missing AML continuation rows)
 */
function runValidateBOM() {
  const sheetName = promptWithValidation('Validate BOM (Full Audit)',
    'Enter the BOM sheet name to validate:');
  if (!sheetName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return ui.alert(`Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ui.alert('Sheet has no data rows.');

  const headers = data[0];
  const colIdx = getColumnIndexes(headers);

  if (colIdx[COL.ITEM_NUMBER] === -1 || colIdx[COL.LEVEL] === -1) {
    return ui.alert('Error', `Required columns ("${COL.ITEM_NUMBER}", "${COL.LEVEL}") not found.`, ui.ButtonSet.OK);
  }

  // Build lookup maps
  const itemsMap = buildItemsLookup_(ss);
  const amlMap = buildAmlLookup_(ss);

  // Collect all findings: {check, severity, row, pn, message}
  const findings = [];
  const SEV = { ERROR: 'ERROR', WARNING: 'WARNING' };

  // Parse BOM rows once
  const bomRows = [];
  for (let i = 1; i < data.length; i++) {
    const pn = data[i][colIdx[COL.ITEM_NUMBER]];
    const level = parseInt(data[i][colIdx[COL.LEVEL]], 10);
    bomRows.push({
      dataIdx: i,
      sheetRow: i + 1,
      pn: pn ? pn.toString().trim() : '',
      level: isNaN(level) ? null : level,
      rawData: data[i]
    });
  }

  // === CHECK 1: Orphaned Parts ===
  bomRows.forEach(row => {
    if (row.pn && !itemsMap.has(row.pn)) {
      findings.push({
        check: '1. Orphaned Part',
        severity: SEV.ERROR,
        row: row.sheetRow,
        pn: row.pn,
        message: `Item Number "${row.pn}" not found in ${BOM_CONFIG.ITEMS_SHEET_NAME} sheet.`
      });
    }
  });

  // === CHECK 2: Missing AML ===
  bomRows.forEach(row => {
    if (row.pn && itemsMap.has(row.pn) && !amlMap.has(row.pn)) {
      findings.push({
        check: '2. Missing AML',
        severity: SEV.WARNING,
        row: row.sheetRow,
        pn: row.pn,
        message: `No AML entry in ${BOM_CONFIG.AML_SHEET_NAME} sheet.`
      });
    }
  });

  // === CHECK 3: Level Hierarchy Gaps ===
  for (let i = 1; i < bomRows.length; i++) {
    const curr = bomRows[i];
    const prev = bomRows[i - 1];
    if (curr.level !== null && prev.level !== null && curr.level > prev.level + 1) {
      findings.push({
        check: '3. Level Gap',
        severity: SEV.ERROR,
        row: curr.sheetRow,
        pn: curr.pn || '(empty)',
        message: `Level jumps from ${prev.level} (row ${prev.sheetRow}) to ${curr.level}. Max allowed: ${prev.level + 1}.`
      });
    }
  }

  // === CHECK 4: Duplicate Assemblies with Different Children ===
  // Reuse the structural integrity pattern from runAuditDuplicatePartNumbers
  const assemblyChildMap = new Map();
  const hasQty = colIdx[COL.QTY] !== -1;

  for (let idx = 0; idx < bomRows.length; idx++) {
    const current = bomRows[idx];
    if (current.level === null || !current.pn) continue;

    const directChildren = [];
    for (let j = idx + 1; j < bomRows.length; j++) {
      const next = bomRows[j];
      if (next.level === null || next.level <= current.level) break;
      if (next.level === current.level + 1) {
        directChildren.push(hasQty
          ? `${next.pn}:${next.rawData[colIdx[COL.QTY]]}`
          : next.pn);
      }
    }

    if (directChildren.length > 0) {
      const signature = directChildren.slice().sort().join(' | ');
      if (!assemblyChildMap.has(current.pn)) assemblyChildMap.set(current.pn, []);
      assemblyChildMap.get(current.pn).push({
        sheetRow: current.sheetRow,
        signature: signature
      });
    }
  }

  assemblyChildMap.forEach((occurrences, pn) => {
    if (occurrences.length < 2) return;
    const signatures = new Set(occurrences.map(o => o.signature));
    if (signatures.size > 1) {
      const rows = occurrences.map(o => o.sheetRow).join(', ');
      findings.push({
        check: '4. Structural Mismatch',
        severity: SEV.ERROR,
        row: occurrences[0].sheetRow,
        pn: pn,
        message: `Assembly appears ${occurrences.length} times with ${signatures.size} different child structures (rows: ${rows}).`
      });
    }
  });

  // === CHECK 5: Stale Managed Values ===
  bomRows.forEach(row => {
    if (!row.pn) return;
    const itemData = itemsMap.get(row.pn);
    const amlEntries = amlMap.get(row.pn);

    if (itemData) {
      const checks = [
        { col: COL.DESCRIPTION, expected: itemData.desc, label: 'Description' },
        { col: COL.ITEM_REV, expected: itemData.rev, label: 'Rev' },
        { col: COL.LIFECYCLE, expected: itemData.lifecycle, label: 'Lifecycle' }
      ];

      checks.forEach(({ col, expected, label }) => {
        if (colIdx[col] === -1) return;
        const actual = row.rawData[colIdx[col]] ? row.rawData[colIdx[col]].toString().trim() : '';
        if (expected && actual !== expected.trim()) {
          findings.push({
            check: '5. Stale Value',
            severity: SEV.WARNING,
            row: row.sheetRow,
            pn: row.pn,
            message: `${label}: MASTER has "${actual}", ITEMS has "${expected}".`
          });
        }
      });
    }

    if (amlEntries && amlEntries.length > 0) {
      const amlChecks = [
        { col: COL.MFR_NAME, expected: amlEntries[0].mfr, label: 'Mfr. Name' },
        { col: COL.MFR_PN, expected: amlEntries[0].mpn, label: 'Mfr. Part Number' }
      ];

      amlChecks.forEach(({ col, expected, label }) => {
        if (colIdx[col] === -1) return;
        const actual = row.rawData[colIdx[col]] ? row.rawData[colIdx[col]].toString().trim() : '';
        if (expected && actual !== expected.trim()) {
          findings.push({
            check: '5. Stale Value',
            severity: SEV.WARNING,
            row: row.sheetRow,
            pn: row.pn,
            message: `${label}: MASTER has "${actual}", AML has "${expected}".`
          });
        }
      });
    }
  });

  // === CHECK 6: MAKE Items with REF Children ===
  // REF means supplier-managed subassembly — valid under BUY, invalid under MAKE
  if (colIdx[COL.REFERENCE_NOTES] !== -1) {
    for (let i = 0; i < bomRows.length; i++) {
      const row = bomRows[i];
      if (row.level === null || !row.pn) continue;

      const status = row.rawData[colIdx[COL.REFERENCE_NOTES]];
      const statusStr = status ? status.toString().toUpperCase().trim() : '';

      if (statusStr.includes('MAKE')) {
        for (let j = i + 1; j < bomRows.length; j++) {
          const child = bomRows[j];
          if (child.level === null || child.level <= row.level) break;

          const childStatus = child.rawData[colIdx[COL.REFERENCE_NOTES]];
          const childStr = childStatus ? childStatus.toString().toUpperCase().trim() : '';
          if (childStr.includes('REF')) {
            findings.push({
              check: '6. MAKE with REF Child',
              severity: SEV.ERROR,
              row: row.sheetRow,
              pn: row.pn,
              message: `MAKE item has REF child "${child.pn}" at row ${child.sheetRow}.`
            });
            break;
          }
        }
      }
    }
  }

  // === CHECK 7: Lifecycle Warnings ===
  if (colIdx[COL.LIFECYCLE] !== -1) {
    const nonProduction = ['OBSOLETE', 'EOL', 'END OF LIFE', 'NRND', 'NOT RECOMMENDED'];
    bomRows.forEach(row => {
      if (!row.pn) return;
      const lifecycle = row.rawData[colIdx[COL.LIFECYCLE]];
      const lifeStr = lifecycle ? lifecycle.toString().toUpperCase().trim() : '';
      if (lifeStr && nonProduction.includes(lifeStr)) {
        findings.push({
          check: '7. Lifecycle Risk',
          severity: SEV.WARNING,
          row: row.sheetRow,
          pn: row.pn,
          message: `Component has non-production lifecycle status: "${lifecycle}".`
        });
      }
    });
  }

  // === CHECK 8: Circular Dependency Detection ===
  // Build parent-child map, then check if any PN appears as its own ancestor
  const parentChildMap = new Map(); // PN → Set of child PNs (direct)

  for (let i = 0; i < bomRows.length; i++) {
    const current = bomRows[i];
    if (current.level === null || !current.pn) continue;

    for (let j = i + 1; j < bomRows.length; j++) {
      const child = bomRows[j];
      if (child.level === null || child.level <= current.level) break;
      if (child.level === current.level + 1 && child.pn) {
        if (!parentChildMap.has(current.pn)) parentChildMap.set(current.pn, new Set());
        parentChildMap.get(current.pn).add(child.pn);
      }
    }
  }

  // DFS for cycles
  const visited = new Set();
  const inStack = new Set();
  const circularPNs = new Set();

  function detectCycle_(pn) {
    if (inStack.has(pn)) {
      circularPNs.add(pn);
      return true;
    }
    if (visited.has(pn)) return false;

    visited.add(pn);
    inStack.add(pn);

    const children = parentChildMap.get(pn);
    if (children) {
      for (const child of children) {
        if (detectCycle_(child)) {
          circularPNs.add(pn);
        }
      }
    }

    inStack.delete(pn);
    return false;
  }

  parentChildMap.forEach((_, pn) => {
    if (!visited.has(pn)) detectCycle_(pn);
  });

  if (circularPNs.size > 0) {
    circularPNs.forEach(pn => {
      const matchRow = bomRows.find(r => r.pn === pn);
      findings.push({
        check: '8. Circular Dependency',
        severity: SEV.ERROR,
        row: matchRow ? matchRow.sheetRow : 0,
        pn: pn,
        message: `Part number appears in its own descendant tree (circular reference).`
      });
    });
  }

  // === CHECK 9: Blank Item Number on Rows with Level ===
  bomRows.forEach(row => {
    if (row.level !== null && row.level > 0 && !row.pn) {
      findings.push({
        check: '9. Blank Item Number',
        severity: SEV.WARNING,
        row: row.sheetRow,
        pn: '(empty)',
        message: `Row has Level ${row.level} but no Item Number.`
      });
    }
  });

  // === CHECK 10: AML Row Count Mismatch ===
  // Parts with multiple AML entries occupy multiple consecutive rows in MASTER.
  // If an AML continuation row is accidentally deleted, the BOM has fewer rows
  // than the AML tab expects for that part.
  if (amlMap.size > 0 && (colIdx[COL.MFR_NAME] !== -1 || colIdx[COL.MFR_PN] !== -1)) {
    let currentPN = '';
    let currentMainRow = -1;
    let actualAmlCount = 0;

    for (let i = 0; i < data.length - 1; i++) {
      const dataRow = data[i + 1]; // skip header (i=0 → data row 1)
      const itemNumber = dataRow[colIdx[COL.ITEM_NUMBER]];
      const hasPN = itemNumber && itemNumber.toString().trim() !== '';

      if (hasPN) {
        // Flush previous part
        if (currentPN && currentMainRow > 0) {
          const expected = amlMap.get(currentPN);
          if (expected && expected.length > 1 && actualAmlCount < expected.length) {
            findings.push({
              check: '10. AML Row Mismatch',
              severity: SEV.ERROR,
              row: currentMainRow,
              pn: currentPN,
              message: `Expected ${expected.length} AML rows but found ${actualAmlCount}. An AML continuation row may have been deleted.`
            });
          }
        }

        currentPN = itemNumber.toString().trim();
        currentMainRow = i + 2; // 1-based sheet row
        actualAmlCount = 0;

        // Count main row's own AML entry
        const mfrName = colIdx[COL.MFR_NAME] !== -1 ? dataRow[colIdx[COL.MFR_NAME]] : null;
        const mfrPN = colIdx[COL.MFR_PN] !== -1 ? dataRow[colIdx[COL.MFR_PN]] : null;
        if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
          actualAmlCount = 1;
        }
      } else if (currentPN) {
        // Blank Item Number — check if it's an AML continuation row
        const mfrName = colIdx[COL.MFR_NAME] !== -1 ? dataRow[colIdx[COL.MFR_NAME]] : null;
        const mfrPN = colIdx[COL.MFR_PN] !== -1 ? dataRow[colIdx[COL.MFR_PN]] : null;
        if ((mfrName && mfrName.toString().trim() !== '') || (mfrPN && mfrPN.toString().trim() !== '')) {
          actualAmlCount++;
        } else {
          // Non-AML blank row — flush
          if (currentPN && currentMainRow > 0) {
            const expected = amlMap.get(currentPN);
            if (expected && expected.length > 1 && actualAmlCount < expected.length) {
              findings.push({
                check: '10. AML Row Mismatch',
                severity: SEV.ERROR,
                row: currentMainRow,
                pn: currentPN,
                message: `Expected ${expected.length} AML rows but found ${actualAmlCount}. An AML continuation row may have been deleted.`
              });
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
        findings.push({
          check: '10. AML Row Mismatch',
          severity: SEV.ERROR,
          row: currentMainRow,
          pn: currentPN,
          message: `Expected ${expected.length} AML rows but found ${actualAmlCount}. An AML continuation row may have been deleted.`
        });
      }
    }
  }

  // === Generate Report ===
  generateValidationReport_(ss, ui, sheetName, findings);
}


/**
 * Generates a categorized validation report sheet.
 */
function generateValidationReport_(ss, ui, sourceSheetName, findings) {
  const reportSheetName = `BOM_Validation_${sourceSheetName}`;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }

  const errorColor = BOM_CONFIG.VALIDATION.COLORS.ERROR;
  const warningColor = BOM_CONFIG.VALIDATION.COLORS.WARNING;

  // --- Summary Section ---
  const errors = findings.filter(f => f.severity === 'ERROR');
  const warnings = findings.filter(f => f.severity === 'WARNING');

  const summaryRows = [
    ['BOM Validation Report'],
    [`Source Sheet: ${sourceSheetName}`, `Date: ${new Date().toLocaleString()}`],
    [''],
    ['Summary'],
    [`Total Findings: ${findings.length}`, `Errors: ${errors.length}`, `Warnings: ${warnings.length}`],
    ['']
  ];

  // Check category counts
  const checkCounts = {};
  findings.forEach(f => {
    checkCounts[f.check] = (checkCounts[f.check] || 0) + 1;
  });

  Object.entries(checkCounts).forEach(([check, count]) => {
    const sev = findings.find(f => f.check === check).severity;
    summaryRows.push([`  ${check}`, `${count} finding(s)`, sev]);
  });

  summaryRows.push(['']);
  summaryRows.push(['─── Detailed Findings ───']);

  // Pad all summary rows to 5 columns
  const maxCols = 5;
  summaryRows.forEach(row => {
    while (row.length < maxCols) row.push('');
  });

  // --- Detail Section ---
  const detailHeader = ['Check', 'Severity', 'Row', 'Item Number', 'Details'];
  summaryRows.push(detailHeader);

  const startDetailRow = summaryRows.length;

  // Sort findings: errors first, then by check number, then by row
  findings.sort((a, b) => {
    if (a.severity !== b.severity) return a.severity === 'ERROR' ? -1 : 1;
    if (a.check !== b.check) return a.check.localeCompare(b.check);
    return a.row - b.row;
  });

  findings.forEach(f => {
    summaryRows.push([f.check, f.severity, f.row, f.pn, f.message]);
  });

  // Write all data at once
  if (summaryRows.length > 0) {
    reportSheet.getRange(1, 1, summaryRows.length, maxCols).setValues(summaryRows);
  }

  // --- Formatting ---
  // Title
  reportSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  // Summary header
  reportSheet.getRange(4, 1).setFontWeight('bold');
  // Detail header row
  reportSheet.getRange(startDetailRow, 1, 1, maxCols).setFontWeight('bold').setBackground('#d9d9d9');

  // Color-code detail rows by severity
  for (let i = 0; i < findings.length; i++) {
    const rowNum = startDetailRow + 1 + i;
    const color = findings[i].severity === 'ERROR' ? errorColor : warningColor;
    reportSheet.getRange(rowNum, 2).setBackground(color); // Color the Severity cell
  }

  // Auto-resize columns
  for (let c = 1; c <= maxCols; c++) {
    reportSheet.autoResizeColumn(c);
  }

  // Activate report sheet
  ss.setActiveSheet(reportSheet);

  // Show summary alert
  if (findings.length === 0) {
    ui.alert('BOM Validation Complete!',
      `All 10 checks passed. No issues found in "${sourceSheetName}".\n\n` +
      `Report saved to "${reportSheetName}".`,
      ui.ButtonSet.OK);
  } else {
    ui.alert('BOM Validation Complete!',
      `Found ${errors.length} error(s) and ${warnings.length} warning(s) in "${sourceSheetName}".\n\n` +
      `See "${reportSheetName}" for the full categorized report.`,
      ui.ButtonSet.OK);
  }
}
