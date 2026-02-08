// ============================================================================
// CONFIG.gs — BOM Configuration & Data Access
// ============================================================================
// Central configuration for all BOM tools. Defines sheet names, column
// mappings, and data-access helpers used across every module.
// ============================================================================

// --- CONFIGURATION ---
// !!!!! IMPORTANT: Adjust these names to exactly match your sheet headers and sheet names !!!!!
const BOM_CONFIG = {
  // Master Data Sheet Names
  ITEMS_SHEET_NAME: 'ITEMS',
  AML_SHEET_NAME: 'AML',

  // ECR/ECO Linking Sheet Name
  ECR_AFFECTED_ITEMS_SHEET: 'ECR_Affected_Items',

  // Column names used by various tools (must match sheet headers exactly)
  COLUMN_NAMES: {
    LEVEL: 'Level',
    ITEM_NUMBER: 'Item Number',
    DESCRIPTION: 'Part Description',
    ITEM_REV: 'Item Rev',
    QTY: 'Qty',
    LIFECYCLE: 'Lifecycle',
    MFR_NAME: 'Mfr. Name',
    MFR_PN: 'Mfr. Part Number',
    REFERENCE_NOTES: 'Reference Notes'  // For MAKE/BUY/REF status
  },

  // Header mapping for the "Grafting/Import" feature (Matches PDM export)
  PDM_GRAFT_SHEET_NAME: 'INPUT_PDM',
  PDM_GRAFT_HEADERS: {
    HIERARCHY_COL: 'LEVEL',
    PN_COL: 'PART NUMBER',
    REV_COL: 'REV',
    DESC_COL: 'DESCRIPTION',
    VENDOR_COL: 'VENDOR',
    MPN_COL: 'MPN',
    QTY_COL: 'QTY.'
  },

  // Header mapping for external PDM Comparison (existing feature)
  // Keys must match COLUMN_NAMES keys above. Columns not present in PDM
  // (e.g., LIFECYCLE, REFERENCE_NOTES) should be omitted — createExternalBOMMap
  // will treat missing keys as -1 (absent).
  PDM_HEADER_MAP: {
    LEVEL: 'Level',
    ITEM_NUMBER: 'NR',
    DESCRIPTION: 'BENENNUNG',
    ITEM_REV: 'Revision',
    QTY: 'Qty',
    MFR_NAME: 'Vendor',
    MFR_PN: 'MPN'
  },

  // Column names for the 'Finalize and Release' process
  CHANGE_TRACKING_COLS_TO_DELETE: ['ECR #', 'Status', 'Change Impact'],

  // Sheet name for ECO Logging
  ECO_LOG_SHEET_NAME: 'ECO History',

  // BOM Effectivity date column names
  EFFECTIVITY: {
    EFFECTIVE_FROM: 'Effective From',
    EFFECTIVE_UNTIL: 'Effective Until'
  },

  // --- Lifecycle State Machine ---
  // Governed states with allowed forward transitions.
  // Backward transitions require an ECR reference (deviation).
  LIFECYCLE: {
    STATES: ['DRAFT', 'PROTOTYPE', 'ACTIVE', 'NRND', 'EOL', 'OBSOLETE'],

    // Allowed forward transitions: state → [valid next states]
    TRANSITIONS: {
      'DRAFT':     ['PROTOTYPE', 'ACTIVE'],
      'PROTOTYPE': ['ACTIVE', 'OBSOLETE'],
      'ACTIVE':    ['NRND', 'EOL', 'OBSOLETE'],
      'NRND':      ['EOL', 'OBSOLETE'],
      'EOL':       ['OBSOLETE'],
      'OBSOLETE':  []   // Terminal state — no forward transitions
    },

    // States that block inclusion in a released BOM
    NON_PRODUCTION: ['OBSOLETE', 'EOL', 'NRND'],

    // Backward transitions are any transition not in TRANSITIONS above.
    // They are allowed ONLY with an ECR reference number.
    DEVIATION_REQUIRED_MSG: 'Backward lifecycle transition requires an ECR reference (deviation).',

    // History sheet for lifecycle audit trail
    HISTORY_SHEET_NAME: 'Lifecycle_History'
  },

  // --- Data Integrity & Validation ---

  // Name of the primary MASTER BOM sheet (for onEdit trigger guard)
  MASTER_SHEET_NAME: 'MASTER',

  // Columns whose values are managed by script (auto-populated from ITEMS/AML).
  // Users should not manually edit these — onEdit will restore them.
  MANAGED_COLUMNS: {
    FROM_ITEMS: ['Part Description', 'Item Rev', 'Lifecycle'],
    FROM_AML: ['Mfr. Name', 'Mfr. Part Number']
  },

  // Visual feedback config for real-time validation (no popups)
  VALIDATION: {
    COLORS: {
      ERROR: '#f4cccc',      // Red — critical errors (orphan, circular, gap)
      WARNING: '#fff2cc',    // Yellow — warnings (stale, missing AML)
      STALE: '#fce5cd',      // Orange — stale/overwritten managed value
      RESTORED: '#d9ead3',   // Green — value auto-restored from source
      OK: null               // Clear background (no issue)
    },
    NOTE_PREFIX: '[BOM Validation] '
  }
};

// Convenience aliases for column names (used throughout the codebase)
const COL = BOM_CONFIG.COLUMN_NAMES;

// Config for revision history sheet name
const REV_HISTORY_SHEET_NAME = 'Rev_History';

// ---------------------
// Data Access Helpers
// ---------------------

/**
 * Lazy-initialized SheetService for the active spreadsheet.
 * Provides cached sheet data reads and batch writes via SheetService utility.
 * Falls back to direct API if SheetService/CacheManager are not loaded.
 */
function getActiveSheetService() {
  if (typeof SheetService === 'undefined' || typeof cacheManager === 'undefined') {
    return null; // Utility modules not loaded — fall back to direct API
  }
  if (!getActiveSheetService._instance) {
    getActiveSheetService._instance = new SheetService(
      SpreadsheetApp.getActiveSpreadsheet().getId()
    );
  }
  return getActiveSheetService._instance;
}

/**
 * Reads sheet data using SheetService (with caching) if available,
 * otherwise falls back to direct Sheets API.
 * @param {string} sheetName Name of the sheet tab.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] Optional spreadsheet (defaults to active).
 * @returns {Array<Array>} 2D array of cell values.
 */
function readSheetData(sheetName, ss) {
  const svc = getActiveSheetService();
  if (svc && !ss) {
    try {
      return svc.getSheetData(sheetName);
    } catch (e) {
      Logger.log(`SheetService read failed for "${sheetName}", falling back: ${e.message}`);
    }
  }
  // Fallback: direct API
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  return sheet.getDataRange().getValues();
}
