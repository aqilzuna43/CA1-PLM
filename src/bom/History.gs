// ============================================================================
// HISTORY.gs — Change History & Logging
// ============================================================================
// Revision history tracking (per-item rev changes) and ECO comparison
// logging. Both functions auto-create their target sheets if missing.
// ============================================================================

// ---------------------
// 1. Revision Change Logging
// ---------------------

/**
 * Logs a revision change to the Rev_History sheet.
 * Called automatically during BOM comparison when rev changes are detected.
 * Can also be called manually for ad-hoc tracking.
 *
 * @param {string} itemNumber The part number whose revision changed.
 * @param {string} oldRev The previous revision.
 * @param {string} newRev The new revision.
 * @param {string} source What triggered the change (e.g., 'ECO-12', 'Manual', 'PDM Import').
 */
function logRevisionChange(itemNumber, oldRev, newRev, source) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let revSheet = ss.getSheetByName(REV_HISTORY_SHEET_NAME);

  if (!revSheet) {
    revSheet = ss.insertSheet(REV_HISTORY_SHEET_NAME);
    const revHeaders = ['Item Number', 'Old Rev', 'New Rev', 'Changed By', 'Date', 'Source'];
    revSheet.appendRow(revHeaders);
    revSheet.setFrozenRows(1);
    revSheet.getRange(1, 1, 1, revHeaders.length).setFontWeight('bold');
  }

  let userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail(); } catch (e) { userEmail = 'Unknown'; }

  revSheet.appendRow([itemNumber, oldRev, newRev, userEmail, new Date(), source || 'Manual']);
}


// ---------------------
// 2. ECO Comparison Logging
// ---------------------

/**
 * Logs the details of a completed BOM comparison to the ECO History sheet.
 */
function logECOComparison(ecoBase, ecrString, oldSheetName, newSheetName, reportSheetName, addedCount, removedCount, modifiedCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(BOM_CONFIG.ECO_LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(BOM_CONFIG.ECO_LOG_SHEET_NAME);
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
    ecrString || 'N/A',
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
