// ============================================================================
// ANALYSIS.gs — Analysis & Reporting Tools
// ============================================================================
// Where-Used analysis (full ancestor chain), BOM Dashboard with key metrics,
// and Master List generation (unique ITEMS & AML extraction).
// ============================================================================

// ---------------------
// 1. Where-Used Analysis
// ---------------------

function runWhereUsedAnalysis() {
  const partNumber = promptWithValidation('Where-Used Analysis', 'Enter the Part Number to search for:');
  if (!partNumber) return;
  const sheetName = promptWithValidation('Where-Used Analysis', 'Enter the BOM sheet to search in:');
  if (!sheetName) return;
  performWhereUsedFullChain(partNumber, sheetName);
}

/**
 * Legacy single-level where-used (kept for compatibility).
 */
function performWhereUsed(partNumber, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) return ui.alert(`Error: Sheet "${sheetName}" not found.`);
  const data = sourceSheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : [];
  const colIdx = getColumnIndexes(headers);
  if (colIdx[COL.LEVEL] === -1 || colIdx[COL.ITEM_NUMBER] === -1) return ui.alert(`Error: Could not find required columns.`);
  const parentAssemblies = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colIdx[COL.ITEM_NUMBER]] && row[colIdx[COL.ITEM_NUMBER]].toString().trim() === partNumber) {
      const childLevel = parseInt(row[colIdx[COL.LEVEL]], 10);
      if (isNaN(childLevel)) continue;
      for (let j = i - 1; j >= 0; j--) {
        const parentRow = data[j];
        const parentLevel = parseInt(parentRow[colIdx[COL.LEVEL]], 10);
        if (isNaN(parentLevel)) continue;
        if (parentLevel < childLevel) {
          const parentPartNumber = parentRow[colIdx[COL.ITEM_NUMBER]].toString();
          const parentDesc = parentRow[colIdx[COL.DESCRIPTION]] || 'N/A';
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

/**
 * Full ancestor chain where-used analysis.
 * For each occurrence, shows: Part → Parent → Grandparent → ... → Top-Level Assembly
 */
function performWhereUsedFullChain(partNumber, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) return ui.alert(`Error: Sheet "${sheetName}" not found.`);

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : [];
  const colIdx = getColumnIndexes(headers);
  if (colIdx[COL.LEVEL] === -1 || colIdx[COL.ITEM_NUMBER] === -1)
    return ui.alert(`Error: Could not find required columns.`);

  // Build path stack to track full ancestry at each row
  const rowPaths = []; // rowPaths[i] = array of ancestors from top to current
  let pathStack = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemNumber = row[colIdx[COL.ITEM_NUMBER]];
    const levelVal = row[colIdx[COL.LEVEL]];
    const level = parseInt(levelVal, 10);

    if (itemNumber && itemNumber.toString().trim() !== '' && !isNaN(level)) {
      pathStack.length = level;
      pathStack[level] = itemNumber.toString().trim();
      rowPaths[i] = pathStack.slice(0, level + 1); // Copy current path
    } else {
      rowPaths[i] = null; // AML or blank row
    }
  }

  // Find all occurrences and their chains
  const chains = [];
  for (let i = 1; i < data.length; i++) {
    if (!rowPaths[i]) continue;
    const pn = rowPaths[i][rowPaths[i].length - 1];
    if (pn === partNumber) {
      chains.push(rowPaths[i].slice().reverse()); // Reverse: part → parent → grandparent → top
    }
  }

  // Build HTML output
  let htmlOutput = `<style>
    body { font-family: Arial, sans-serif; font-size: 13px; }
    table { border-collapse: collapse; width: 100%; margin-top: 10px; }
    th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
    th { background: #4285f4; color: #fff; }
    .chain { color: #666; font-size: 12px; }
    .arrow { color: #999; }
  </style>`;
  htmlOutput += `<b>Where-Used: ${partNumber}</b> (${chains.length} occurrence(s))<br/>`;

  if (chains.length > 0) {
    htmlOutput += '<table><tr><th>#</th><th>Full Assembly Chain (Part → Top Level)</th></tr>';
    chains.forEach((chain, idx) => {
      const chainStr = chain.map((pn, i) =>
        i === 0 ? `<b>${pn}</b>` : pn
      ).join(' <span class="arrow">→</span> ');
      htmlOutput += `<tr><td>${idx + 1}</td><td class="chain">${chainStr}</td></tr>`;
    });
    htmlOutput += '</table>';
  } else {
    htmlOutput += '<br/>Part number not found in any assemblies.';
  }

  ui.showModalDialog(HtmlService.createHtmlOutput(htmlOutput).setWidth(700).setHeight(400), 'Where-Used Results (Full Chain)');
}


// ---------------------
// 2. BOM Dashboard
// ---------------------

/**
 * Generates a summary dashboard sheet with key BOM statistics.
 */
function runGenerateDashboard() {
  const sourceSheetName = promptWithValidation('Generate BOM Dashboard', 'Enter the BOM sheet name:');
  if (!sourceSheetName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return ui.alert(`Sheet "${sourceSheetName}" not found.`);

  const data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) return ui.alert('Sheet has no data rows.');

  const headers = data[0];
  const colIdx = getColumnIndexes(headers);

  // Gather statistics
  const uniqueParts = new Set();
  const uniqueAssemblies = new Set(); // Parts that have children (i.e., are parents)
  const lifecycleCounts = {};
  const statusCounts = {}; // MAKE/BUY/REF
  let maxDepth = 0;
  let totalAmlEntries = 0;
  let partsWithoutAml = new Set();
  let pathStack = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemNumber = row[colIdx[COL.ITEM_NUMBER]];
    const levelVal = row[colIdx[COL.LEVEL]];
    const level = parseInt(levelVal, 10);

    if (itemNumber && itemNumber.toString().trim() !== '' && !isNaN(level)) {
      const pn = itemNumber.toString().trim();
      uniqueParts.add(pn);
      if (level > maxDepth) maxDepth = level;

      // Track parent assemblies
      pathStack.length = level;
      pathStack[level] = pn;
      if (level > 0 && pathStack[level - 1]) {
        uniqueAssemblies.add(pathStack[level - 1]);
      }

      // Lifecycle distribution
      if (colIdx[COL.LIFECYCLE] !== -1) {
        const lifecycle = row[colIdx[COL.LIFECYCLE]];
        const lcStr = lifecycle ? lifecycle.toString().trim() : '(blank)';
        lifecycleCounts[lcStr] = (lifecycleCounts[lcStr] || 0) + 1;
      }

      // Status distribution (MAKE/BUY/REF)
      if (colIdx[COL.REFERENCE_NOTES] !== -1) {
        const statusCell = row[colIdx[COL.REFERENCE_NOTES]];
        const status = statusCell ? statusCell.toString().trim().toUpperCase() : '(blank)';
        statusCounts[status] = (statusCounts[status] || 0) + 1;
      }

      // Check AML presence on the main part row
      const mfrName = colIdx[COL.MFR_NAME] !== -1 ? row[colIdx[COL.MFR_NAME]] : null;
      const mfrPN = colIdx[COL.MFR_PN] !== -1 ? row[colIdx[COL.MFR_PN]] : null;
      if (mfrName && mfrName.toString().trim() !== '') {
        totalAmlEntries++;
      } else {
        partsWithoutAml.add(pn);
      }
    } else if (itemNumber === '' || itemNumber === null) {
      // AML continuation row
      const mfrName = colIdx[COL.MFR_NAME] !== -1 ? row[colIdx[COL.MFR_NAME]] : null;
      if (mfrName && mfrName.toString().trim() !== '') {
        totalAmlEntries++;
        // Remove from "no AML" list since it has at least one vendor
        if (pathStack.length > 0) {
          const lastPN = pathStack[pathStack.length - 1];
          if (lastPN) partsWithoutAml.delete(lastPN);
        }
      }
    }
  }

  // Build dashboard sheet
  const dashSheetName = `Dashboard_${sourceSheetName}`;
  let dashSheet = ss.getSheetByName(dashSheetName) || ss.insertSheet(dashSheetName);
  dashSheet.clear();

  const rows = [];
  rows.push(['BOM Dashboard — ' + sourceSheetName, '', '']);
  rows.push(['Generated', new Date().toLocaleString(), '']);
  rows.push(['', '', '']);

  // Key Metrics
  rows.push(['KEY METRICS', '', '']);
  rows.push(['Total Unique Parts', uniqueParts.size, '']);
  rows.push(['Assembly (MAKE) Parts', uniqueAssemblies.size, '']);
  rows.push(['Leaf (BUY/REF) Parts', uniqueParts.size - uniqueAssemblies.size, '']);
  rows.push(['Max BOM Depth', maxDepth, '']);
  rows.push(['Total AML Entries', totalAmlEntries, '']);
  rows.push(['Parts Without AML', partsWithoutAml.size, partsWithoutAml.size > 0 ? '⚠' : '✓']);
  rows.push(['', '', '']);

  // Lifecycle Distribution
  rows.push(['LIFECYCLE DISTRIBUTION', 'Count', '']);
  const sortedLifecycle = Object.entries(lifecycleCounts).sort((a, b) => b[1] - a[1]);
  sortedLifecycle.forEach(([status, count]) => {
    rows.push([status, count, '']);
  });
  rows.push(['', '', '']);

  // Status Distribution
  rows.push(['STATUS DISTRIBUTION (MAKE/BUY/REF)', 'Count', '']);
  const sortedStatus = Object.entries(statusCounts).sort((a, b) => b[1] - a[1]);
  sortedStatus.forEach(([status, count]) => {
    rows.push([status, count, '']);
  });
  rows.push(['', '', '']);

  // Parts without AML (list first 20)
  if (partsWithoutAml.size > 0) {
    rows.push(['PARTS WITHOUT AML (Top 20)', '', '']);
    const noAmlArr = Array.from(partsWithoutAml).slice(0, 20);
    noAmlArr.forEach(pn => rows.push([pn, '', '']));
    if (partsWithoutAml.size > 20) {
      rows.push([`... and ${partsWithoutAml.size - 20} more`, '', '']);
    }
  }

  // Write to sheet
  dashSheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // Formatting
  dashSheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  dashSheet.getRange(4, 1).setFontWeight('bold').setFontSize(11);
  const lifecycleHeaderRow = 4 + 7 + 1; // after key metrics + blank
  dashSheet.getRange(lifecycleHeaderRow, 1).setFontWeight('bold').setFontSize(11);
  dashSheet.autoResizeColumns(1, 3);

  ui.alert('Dashboard Generated!', `See "${dashSheetName}" for BOM statistics.`, ui.ButtonSet.OK);
}


// ---------------------
// 3. Master List Generation
// ---------------------

function runGenerateMasterLists() {
  const sourceSheetName = promptWithValidation('Generate Master Lists', 'Enter the name of the source BOM sheet:');
  if (!sourceSheetName) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
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
  const itemNumIdx = colIdx[COL.ITEM_NUMBER], descIdx = colIdx[COL.DESCRIPTION], revIdx = colIdx[COL.ITEM_REV];
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
  const newHeaders = [COL.ITEM_NUMBER, COL.DESCRIPTION, COL.ITEM_REV];
  const outputData = [newHeaders, ...Array.from(uniqueItems, ([key, value]) => [key, value.desc, value.rev])];
  newSheet.getRange(1, 1, outputData.length, newHeaders.length).setValues(outputData);
  newSheet.autoResizeColumns(1, newHeaders.length);
}

function generateAmlList(sourceSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData.length > 0 ? sourceData[0] : [];
  const colIdx = getColumnIndexes(headers);
  const itemNumIdx = colIdx[COL.ITEM_NUMBER], mfrNameIdx = colIdx[COL.MFR_NAME], mfrPnIdx = colIdx[COL.MFR_PN];
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
  const newHeaders = [COL.ITEM_NUMBER, COL.MFR_NAME, COL.MFR_PN];
  const outputData = [newHeaders, ...amlData];
  newSheet.getRange(1, 1, outputData.length, newHeaders.length).setValues(outputData);
  newSheet.autoResizeColumns(1, newHeaders.length);
}
