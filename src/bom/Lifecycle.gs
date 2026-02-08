// ============================================================================
// LIFECYCLE.gs — Lifecycle State Machine & Transition Governance
// ============================================================================
// Enforces governed lifecycle transitions on the ITEMS sheet:
//   - Forward transitions follow LIFECYCLE.TRANSITIONS rules
//   - Backward transitions (deviations) require an ECR reference
//   - Every transition is logged to Lifecycle_History sheet
//
// Integrates with:
//   - Validation.gs (onEdit hook for ITEMS sheet lifecycle column)
//   - Reconcile.gs  (batch lifecycle audit)
//   - Release.gs    (pre-release lifecycle check)
// ============================================================================


// ========================
// STATE MACHINE CORE
// ========================

/**
 * Validates whether a lifecycle transition is allowed.
 *
 * @param {string} currentState The current lifecycle state.
 * @param {string} newState The proposed new lifecycle state.
 * @returns {{valid: boolean, isDeviation: boolean, message: string}}
 */
function validateLifecycleTransition(currentState, newState) {
  const config = BOM_CONFIG.LIFECYCLE;
  const current = normalizeLifecycleState_(currentState);
  const next = normalizeLifecycleState_(newState);

  // Same state — no transition needed
  if (current === next) {
    return { valid: true, isDeviation: false, message: 'No change.' };
  }

  // Blank → any state is always valid (initial assignment)
  if (current === '') {
    if (config.STATES.includes(next)) {
      return { valid: true, isDeviation: false, message: 'Initial lifecycle assignment.' };
    }
    return {
      valid: false,
      isDeviation: false,
      message: `"${newState}" is not a recognized lifecycle state. Valid states: ${config.STATES.join(', ')}`
    };
  }

  // Validate both states are recognized
  if (!config.STATES.includes(current)) {
    return {
      valid: false,
      isDeviation: false,
      message: `Current state "${currentState}" is not recognized. Valid states: ${config.STATES.join(', ')}`
    };
  }
  if (!config.STATES.includes(next)) {
    return {
      valid: false,
      isDeviation: false,
      message: `"${newState}" is not a recognized lifecycle state. Valid states: ${config.STATES.join(', ')}`
    };
  }

  // Check forward transitions
  const allowedForward = config.TRANSITIONS[current] || [];
  if (allowedForward.includes(next)) {
    return { valid: true, isDeviation: false, message: `Forward transition: ${current} → ${next}` };
  }

  // Not a valid forward transition — must be a deviation (backward)
  return {
    valid: false,
    isDeviation: true,
    message: `${current} → ${next} is a backward transition. ${config.DEVIATION_REQUIRED_MSG}`
  };
}

/**
 * Normalizes lifecycle state strings for comparison.
 * Handles common aliases (e.g., "END OF LIFE" → "EOL").
 * @param {string} state Raw state string.
 * @returns {string} Normalized uppercase state.
 */
function normalizeLifecycleState_(state) {
  if (!state) return '';
  const trimmed = state.toString().trim().toUpperCase();

  // Common aliases
  const ALIASES = {
    'END OF LIFE': 'EOL',
    'NOT RECOMMENDED': 'NRND',
    'NOT RECOMMENDED FOR NEW DESIGNS': 'NRND',
    'PRODUCTION': 'ACTIVE',
    'RELEASED': 'ACTIVE',
    'IN PRODUCTION': 'ACTIVE',
    'PROTOTYPE': 'PROTOTYPE',
    'PROTO': 'PROTOTYPE'
  };

  return ALIASES[trimmed] || trimmed;
}


// ========================
// LIFECYCLE HISTORY LOGGING
// ========================

/**
 * Logs a lifecycle transition to the Lifecycle_History sheet.
 * Auto-creates the sheet if it doesn't exist.
 *
 * @param {string} itemNumber The part number.
 * @param {string} oldState Previous lifecycle state.
 * @param {string} newState New lifecycle state.
 * @param {string} source What triggered the change (e.g., 'onEdit', 'ECR-2025-042', 'Reconcile').
 * @param {string} [ecrRef] ECR reference for deviation transitions.
 */
function logLifecycleTransition(itemNumber, oldState, newState, source, ecrRef) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = BOM_CONFIG.LIFECYCLE.HISTORY_SHEET_NAME;
  let historySheet = ss.getSheetByName(sheetName);

  if (!historySheet) {
    historySheet = ss.insertSheet(sheetName);
    const headers = [
      'Item Number', 'Old State', 'New State', 'Transition Type',
      'Changed By', 'Date', 'Source', 'ECR Reference'
    ];
    historySheet.appendRow(headers);
    historySheet.setFrozenRows(1);
    historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  const validation = validateLifecycleTransition(oldState, newState);
  const transitionType = validation.isDeviation ? 'DEVIATION' : 'FORWARD';

  let userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail(); } catch (e) { userEmail = 'Unknown'; }

  historySheet.appendRow([
    itemNumber,
    normalizeLifecycleState_(oldState) || '(blank)',
    normalizeLifecycleState_(newState),
    transitionType,
    userEmail,
    new Date(),
    source || 'Manual',
    ecrRef || ''
  ]);
}


// ========================
// ITEMS SHEET LIFECYCLE VALIDATION (onEdit hook)
// ========================

/**
 * Handles lifecycle column edits on the ITEMS sheet.
 * Called from the onEdit trigger when the edited sheet is ITEMS
 * and the edited column is Lifecycle.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Spreadsheet reference.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet ITEMS sheet.
 * @param {number} row Edited row (1-based).
 * @param {number} col Edited column (1-based).
 * @param {*} oldValue Previous cell value.
 * @param {*} newValue New cell value.
 */
function handleItemsLifecycleEdit(ss, sheet, row, col, oldValue, newValue) {
  const prefix = BOM_CONFIG.VALIDATION.NOTE_PREFIX;
  const cell = sheet.getRange(row, col);
  const oldState = oldValue ? oldValue.toString().trim() : '';
  const newState = newValue ? newValue.toString().trim() : '';

  if (oldState === newState) return; // No actual change

  const validation = validateLifecycleTransition(oldState, newState);

  if (validation.valid) {
    // Valid forward transition or initial assignment
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.RESTORED);
    cell.setNote(prefix + validation.message);

    // Get item number for logging
    const itemNumber = getItemNumberForRow_(sheet, row);
    if (itemNumber) {
      logLifecycleTransition(itemNumber, oldState, newState, 'onEdit (ITEMS)');
    }
  } else if (validation.isDeviation) {
    // Backward transition — prompt for ECR reference
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.WARNING);
    cell.setNote(prefix + validation.message);

    // Attempt to get ECR reference via prompt (only works in non-trigger context)
    // In onEdit context, we can't use ui.prompt(), so we flag and require manual resolution
    const itemNumber = getItemNumberForRow_(sheet, row);
    if (itemNumber) {
      // Revert the value — user must use the menu tool for deviations
      cell.setValue(oldValue || '');
      cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
      cell.setNote(
        prefix + `BLOCKED: ${normalizeLifecycleState_(oldState)} → ${normalizeLifecycleState_(newState)} ` +
        'is a backward transition.\nUse BOM Tools > Data Integrity > Lifecycle Deviation to change this state with an ECR reference.'
      );
    }
  } else {
    // Invalid state entirely
    cell.setValue(oldValue || '');
    cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.ERROR);
    cell.setNote(prefix + validation.message);
  }
}

/**
 * Gets the Item Number for a given row on the ITEMS sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The ITEMS sheet.
 * @param {number} row Row number (1-based).
 * @returns {string} Item number or empty string.
 */
function getItemNumberForRow_(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const pnIdx = headers.indexOf('Item Number');
  if (pnIdx === -1) return '';
  const val = sheet.getRange(row, pnIdx + 1).getValue();
  return val ? val.toString().trim() : '';
}


// ========================
// LIFECYCLE DEVIATION TOOL (Menu-driven)
// ========================

/**
 * Menu-driven tool for performing backward lifecycle transitions.
 * Requires an ECR reference number as justification.
 * Accessible via: BOM Tools > Data Integrity > Lifecycle Deviation
 */
function runLifecycleDeviation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Item Number
  const itemInput = promptWithValidation(
    'Lifecycle Deviation',
    'Enter the Item Number to change lifecycle state:',
    { minLength: 1, maxLength: 100 }
  );
  if (!itemInput) return;

  // 2. Look up current state in ITEMS
  const itemsSheet = ss.getSheetByName(BOM_CONFIG.ITEMS_SHEET_NAME);
  if (!itemsSheet) {
    ui.alert('Error', `ITEMS sheet "${BOM_CONFIG.ITEMS_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }

  const itemsData = itemsSheet.getDataRange().getValues();
  const headers = itemsData[0];
  const pnIdx = headers.indexOf('Item Number');
  const lifeIdx = headers.indexOf('Lifecycle');

  if (pnIdx === -1 || lifeIdx === -1) {
    ui.alert('Error', 'ITEMS sheet missing required columns (Item Number, Lifecycle).', ui.ButtonSet.OK);
    return;
  }

  let targetRow = -1;
  let currentState = '';
  for (let i = 1; i < itemsData.length; i++) {
    const pn = itemsData[i][pnIdx] ? itemsData[i][pnIdx].toString().trim() : '';
    if (pn === itemInput) {
      targetRow = i + 1; // 1-based sheet row
      currentState = itemsData[i][lifeIdx] ? itemsData[i][lifeIdx].toString().trim() : '';
      break;
    }
  }

  if (targetRow === -1) {
    ui.alert('Not Found', `Item "${itemInput}" not found in ITEMS sheet.`, ui.ButtonSet.OK);
    return;
  }

  // 3. Get new state
  const validStates = BOM_CONFIG.LIFECYCLE.STATES.join(', ');
  const newStateInput = promptWithValidation(
    'Lifecycle Deviation',
    `Current state: ${currentState || '(blank)'}\n\nEnter the NEW lifecycle state:\nValid states: ${validStates}`,
    { minLength: 1, maxLength: 50 }
  );
  if (!newStateInput) return;

  const normalized = normalizeLifecycleState_(newStateInput);
  if (!BOM_CONFIG.LIFECYCLE.STATES.includes(normalized)) {
    ui.alert('Invalid State', `"${newStateInput}" is not a valid lifecycle state.\nValid: ${validStates}`, ui.ButtonSet.OK);
    return;
  }

  // 4. Require ECR reference
  const ecrRef = promptWithValidation(
    'ECR Reference Required',
    'This is a deviation (backward transition).\nEnter the authorizing ECR number:',
    {
      minLength: 1,
      maxLength: 50,
      pattern: /^ECR[-\s]?\d+/i,
      patternHint: 'ECR reference must start with "ECR-" followed by a number (e.g., ECR-2025-042).'
    }
  );
  if (!ecrRef) return;

  // 5. Confirm
  const confirm = ui.alert(
    'Confirm Deviation',
    `Item: ${itemInput}\n` +
    `Transition: ${currentState || '(blank)'} → ${normalized}\n` +
    `ECR: ${ecrRef}\n\nProceed?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // 6. Execute
  const cell = itemsSheet.getRange(targetRow, lifeIdx + 1);
  cell.setValue(normalized);
  cell.setBackground(BOM_CONFIG.VALIDATION.COLORS.WARNING);
  cell.setNote(
    BOM_CONFIG.VALIDATION.NOTE_PREFIX +
    `Deviation: ${currentState || '(blank)'} → ${normalized}. Authorized by ${ecrRef}.`
  );

  // 7. Log
  logLifecycleTransition(itemInput, currentState, normalized, 'Lifecycle Deviation (Menu)', ecrRef);

  ui.alert('Deviation Applied',
    `${itemInput}: ${currentState || '(blank)'} → ${normalized}\nLogged with ECR ref: ${ecrRef}`,
    ui.ButtonSet.OK
  );
}


// ========================
// BATCH LIFECYCLE AUDIT
// ========================

/**
 * Audits all items in the ITEMS sheet for lifecycle state validity.
 * Reports unrecognized states and items missing lifecycle values.
 * Accessible via: BOM Tools > Audit & Quality > Audit Lifecycle States
 */
function runAuditLifecycleStates() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemsSheet = ss.getSheetByName(BOM_CONFIG.ITEMS_SHEET_NAME);
  if (!itemsSheet) {
    ui.alert('Error', `ITEMS sheet "${BOM_CONFIG.ITEMS_SHEET_NAME}" not found.`, ui.ButtonSet.OK);
    return;
  }

  const data = itemsSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert('Empty', 'ITEMS sheet has no data rows.', ui.ButtonSet.OK);
    return;
  }

  const headers = data[0];
  const pnIdx = headers.indexOf('Item Number');
  const lifeIdx = headers.indexOf('Lifecycle');

  if (pnIdx === -1 || lifeIdx === -1) {
    ui.alert('Error', 'ITEMS sheet missing required columns.', ui.ButtonSet.OK);
    return;
  }

  const validStates = new Set(BOM_CONFIG.LIFECYCLE.STATES);
  const issues = { unrecognized: [], blank: [], nonProduction: [] };
  const stateCounts = {};

  for (let i = 1; i < data.length; i++) {
    const pn = data[i][pnIdx] ? data[i][pnIdx].toString().trim() : '';
    if (!pn) continue;

    const rawState = data[i][lifeIdx] ? data[i][lifeIdx].toString().trim() : '';
    const normalized = normalizeLifecycleState_(rawState);

    // Count states
    const displayState = normalized || '(blank)';
    stateCounts[displayState] = (stateCounts[displayState] || 0) + 1;

    if (!rawState) {
      issues.blank.push(pn);
    } else if (!validStates.has(normalized)) {
      issues.unrecognized.push({ pn: pn, state: rawState });
    } else if (BOM_CONFIG.LIFECYCLE.NON_PRODUCTION.includes(normalized)) {
      issues.nonProduction.push({ pn: pn, state: normalized });
    }
  }

  // Build report
  const lines = ['=== LIFECYCLE STATE AUDIT ===', ''];

  // State distribution
  lines.push('State Distribution:');
  const sorted = Object.entries(stateCounts).sort((a, b) => b[1] - a[1]);
  sorted.forEach(([state, count]) => lines.push(`  ${state}: ${count}`));
  lines.push('');

  if (issues.unrecognized.length > 0) {
    lines.push(`⚠ Unrecognized States (${issues.unrecognized.length}):`);
    issues.unrecognized.slice(0, 20).forEach(i => lines.push(`  ${i.pn}: "${i.state}"`));
    if (issues.unrecognized.length > 20) lines.push(`  ...and ${issues.unrecognized.length - 20} more`);
    lines.push('');
  }

  if (issues.blank.length > 0) {
    lines.push(`⚠ Missing Lifecycle (${issues.blank.length}):`);
    issues.blank.slice(0, 20).forEach(pn => lines.push(`  ${pn}`));
    if (issues.blank.length > 20) lines.push(`  ...and ${issues.blank.length - 20} more`);
    lines.push('');
  }

  if (issues.nonProduction.length > 0) {
    lines.push(`⚠ Non-Production Parts (${issues.nonProduction.length}):`);
    issues.nonProduction.slice(0, 20).forEach(i => lines.push(`  ${i.pn}: ${i.state}`));
    if (issues.nonProduction.length > 20) lines.push(`  ...and ${issues.nonProduction.length - 20} more`);
    lines.push('');
  }

  const totalIssues = issues.unrecognized.length + issues.blank.length;
  if (totalIssues === 0) {
    lines.push('✓ All lifecycle states are valid.');
  } else {
    lines.push(`Valid states: ${BOM_CONFIG.LIFECYCLE.STATES.join(', ')}`);
  }

  ui.alert('Lifecycle Audit', lines.join('\n'), ui.ButtonSet.OK);
}
