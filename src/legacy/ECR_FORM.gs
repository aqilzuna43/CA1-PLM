/**
 * ==================================================
 * ECR AUTOMATION SUITE v2.0 (REDLINE PATTERN)
 *
 * Architecture: ENOVIA/ARAS-style "Current vs. New" columns
 * - Auto-fill ONLY writes to "Curr *" columns (gray, read-only)
 * - User-editable "New *" columns are NEVER overwritten
 * - Background colors: Gray = auto-filled, White = user-editable, Yellow = user-modified
 *
 * Sheet Layouts:
 *   Affected_Items: Parent | Child PN | Curr Desc | Type | Curr Rev | New Rev | Curr Qty | New Qty | New Desc | Disp | Reason | Status
 *   AVL_Changes:    Item PN | Curr Desc | Type | Curr Mfr | Curr MPN | New Mfr | New MPN | Disp | Reason | Status
 *
 *   // Master Sheet IDs
 *   //1CNvUEd4BQrh35LBQQa-hEZa9glhCOxhXytilTnTXnxg - EVAL
 *   //1rytdEh6xi8FrHKVN_ZOOTGXJFIZqrNWEYrPD0-ogk-Q - MAIN
 * ==================================================
 */

const ECR_CONFIG = {
  MASTER_ID: '1rytdEh6xi8FrHKVN_ZOOTGXJFIZqrNWEYrPD0-ogk-Q',

  // Tab Names
  FORM_TAB_NAME: 'ECR_Form',
  BOM_LIST_TAB: 'Affected_Items',
  AVL_LIST_TAB: 'AVL_Changes',

  // Master Source Tabs
  ITEMS_TAB_NAME: 'ITEMS',
  BOM_TAB_NAME: 'MASTER',
  AML_TAB_NAME: 'AML',
  ECR_LOG_TAB_NAME: 'ECR_Affected_Items',

  FORM_CELLS: { ECR_NUM: 'B4', ECO_NUM: 'E4', ECR_STATUS: 'B6', REVIEWER: 'E6', REVIEW_DATE: 'B8', REJECTION_REASON: 'E8' },
  TABLE_START_ROW: 2,

  // --- ECR Workflow State Machine ---
  WORKFLOW: {
    STATES: ['DRAFT', 'SUBMITTED', 'UNDER_REVIEW', 'APPROVED', 'REJECTED', 'IMPLEMENTING', 'CLOSED'],

    // Allowed transitions: state → [valid next states]
    TRANSITIONS: {
      'DRAFT':          ['SUBMITTED'],
      'SUBMITTED':      ['UNDER_REVIEW', 'REJECTED'],
      'UNDER_REVIEW':   ['APPROVED', 'REJECTED'],
      'APPROVED':       ['IMPLEMENTING'],
      'REJECTED':       ['DRAFT'],       // Rejection loops back for revision
      'IMPLEMENTING':   ['CLOSED'],
      'CLOSED':         []               // Terminal state
    },

    // States that allow data editing
    EDITABLE_STATES: ['DRAFT', 'REJECTED'],

    // State that allows commit to master
    COMMIT_REQUIRED_STATE: 'APPROVED',

    // ECR History sheet on Master spreadsheet
    HISTORY_SHEET_NAME: 'ECR_Workflow_History'
  },

  // Master BOM Indexes (Zero-based for array logic)
  M_IDX: { LEVEL: 1, ITEM: 2, QTY: 6 },
  A_IDX: { ITEM: 0, MFR: 1, MPN: 2 },
  I_IDX: { ITEM: 0, DESC: 1, REV: 2 },

  // =====================================================================
  // Affected_Items column indexes (0-based) — "Redline" layout
  // A=0       B=1        C=2         D=3    E=4       F=5      G=6       H=7      I=8       J=9   K=10    L=11
  // Parent | Child PN | Curr Desc | Type | Curr Rev | New Rev | Curr Qty | New Qty | New Desc | Disp | Reason | Status
  // =====================================================================
  BOM_COL: {
    PARENT:    0,  // A — Auto-fill or manual (smart parent lookup)
    CHILD_PN:  1,  // B — USER KEY (input trigger)
    CURR_DESC: 2,  // C — AUTO (gray, from ITEMS)
    TYPE:      3,  // D — USER (dropdown: ADDED, REMOVED, MODIFIED, QTY CHANGE, REV ROLL, DESC CHANGE)
    CURR_REV:  4,  // E — AUTO (gray, from ITEMS)
    NEW_REV:   5,  // F — USER (editable, never overwritten)
    CURR_QTY:  6,  // G — AUTO (gray, from MASTER BOM)
    NEW_QTY:   7,  // H — USER (editable, never overwritten)
    NEW_DESC:  8,  // I — USER (editable, never overwritten)
    DISP:      9,  // J — USER
    REASON:   10,  // K — USER
    STATUS:   11   // L — AUTO (validation status)
  },
  BOM_TOTAL_COLS: 12,

  // =====================================================================
  // AVL_Changes column indexes (0-based) — "Redline" layout
  // A=0       B=1         C=2    D=3       E=4       F=5      G=6      H=7   I=8     J=9
  // Item PN | Curr Desc | Type | Curr Mfr | Curr MPN | New Mfr | New MPN | Disp | Reason | Status
  // =====================================================================
  AVL_COL: {
    ITEM_PN:   0,  // A — USER KEY (carried from Affected_Items)
    CURR_DESC: 1,  // B — AUTO (gray, from ITEMS)
    TYPE:      2,  // C — USER (dropdown: AVL_ADD, AVL_REMOVE, AVL_REPLACE)
    CURR_MFR:  3,  // D — AUTO (gray, from AML — split from old "Mfr: MPN" format)
    CURR_MPN:  4,  // E — AUTO (gray, from AML)
    NEW_MFR:   5,  // F — USER (editable, never overwritten)
    NEW_MPN:   6,  // G — USER (editable, never overwritten)
    DISP:      7,  // H — USER
    REASON:    8,  // I — USER
    STATUS:    9   // J — AUTO (validation status)
  },
  AVL_TOTAL_COLS: 10,

  // Background colors (ARAS convention)
  COLORS: {
    AUTO_FILL: '#f3f3f3',  // Gray — auto-populated, read-only
    EDITABLE:  '#ffffff',  // White — user-editable
    MODIFIED:  '#fff9c4',  // Light yellow — user has entered data
    ERROR:     '#ffcccc',  // Light red — validation error
    SUCCESS:   '#e8f5e9'   // Light green — validated OK
  }
};


// ====================================================================================================
// === MENU ==========================================================================================
// ====================================================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ECR Actions')
    .addItem('▶ Auto-Fill Data (BOM + AVL)', 'populateCurrentData')
    .addItem('▶ Submit to Master Log', 'submitComprehensiveECR')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Workflow')
        .addItem('Submit ECR for Review', 'ecrTransition_Submit')
        .addItem('Begin Review', 'ecrTransition_BeginReview')
        .addItem('Approve ECR', 'ecrTransition_Approve')
        .addItem('Reject ECR (with reason)', 'ecrTransition_Reject')
        .addItem('Reopen as Draft (after rejection)', 'ecrTransition_Reopen')
        .addItem('View Workflow Status', 'viewEcrWorkflowStatus')
    )
    .addSeparator()
    .addItem('⚠ ADMIN: Commit to Master', 'commitToMaster')
    .addSeparator()
    .addItem('Setup Sheet Headers', 'setupSheetHeaders')
    .addItem('Set Admin Password', 'setAdminPassword')
    .addItem('Troubleshoot Link', 'DEBUG_Check_Parent_Child_Link')
    .addToUi();
}


// ====================================================================================================
// === SHEET HEADER SETUP (Run once to initialize templates) =========================================
// ====================================================================================================

/**
 * Sets up the Affected_Items and AVL_Changes sheet headers with the new redline column layout.
 * Safe to run multiple times — only writes headers in row 1, never touches data rows.
 */
function setupSheetHeaders() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const BC = ECR_CONFIG.BOM_COL;
  const AC = ECR_CONFIG.AVL_COL;
  const C = ECR_CONFIG.COLORS;

  // --- Affected_Items ---
  let bomSheet = ss.getSheetByName(ECR_CONFIG.BOM_LIST_TAB);
  if (!bomSheet) bomSheet = ss.insertSheet(ECR_CONFIG.BOM_LIST_TAB);

  const bomHeaders = ['Parent', 'Child PN', 'Curr Desc', 'Change Type', 'Curr Rev', 'New Rev', 'Curr Qty', 'New Qty', 'New Desc', 'Disp', 'Reason', 'Status'];
  bomSheet.getRange(1, 1, 1, bomHeaders.length).setValues([bomHeaders]).setFontWeight('bold');

  // Color-code headers: gray for auto-fill columns, white for user columns
  const bomHeaderBg = new Array(bomHeaders.length).fill(C.EDITABLE);
  [BC.CURR_DESC, BC.CURR_REV, BC.CURR_QTY, BC.STATUS].forEach(idx => { bomHeaderBg[idx] = C.AUTO_FILL; });
  bomSheet.getRange(1, 1, 1, bomHeaders.length).setBackgrounds([bomHeaderBg]);

  // --- AVL_Changes ---
  let avlSheet = ss.getSheetByName(ECR_CONFIG.AVL_LIST_TAB);
  if (!avlSheet) avlSheet = ss.insertSheet(ECR_CONFIG.AVL_LIST_TAB);

  const avlHeaders = ['Item PN', 'Curr Desc', 'Change Type', 'Curr Mfr', 'Curr MPN', 'New Mfr', 'New MPN', 'Disp', 'Reason', 'Status'];
  avlSheet.getRange(1, 1, 1, avlHeaders.length).setValues([avlHeaders]).setFontWeight('bold');

  const avlHeaderBg = new Array(avlHeaders.length).fill(C.EDITABLE);
  [AC.CURR_DESC, AC.CURR_MFR, AC.CURR_MPN, AC.STATUS].forEach(idx => { avlHeaderBg[idx] = C.AUTO_FILL; });
  avlSheet.getRange(1, 1, 1, avlHeaders.length).setBackgrounds([avlHeaderBg]);

  ui.alert('Headers Updated!',
    'Both sheet headers have been set up with the new Current/New column layout.\n\n' +
    'Gray columns = auto-filled (read-only)\nWhite columns = user-editable',
    ui.ButtonSet.OK);
}


// ====================================================================================================
// === FEATURE 1: AUTO-FILL (REDLINE PATTERN) ========================================================
// ====================================================================================================

/**
 * Auto-fills "Current" columns from the Master BOM while preserving all user-entered "New" values.
 *
 * RULES:
 * 1. Gray "Curr *" columns are ALWAYS overwritten with fresh master data
 * 2. White "New *" columns are NEVER touched if user has entered data
 * 3. Parent column: auto-filled via smart lookup, but preserved if user already selected
 * 4. Status column: set to validation result (OK, ERROR, or NEW)
 */
function populateCurrentData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const BC = ECR_CONFIG.BOM_COL;
  const AC = ECR_CONFIG.AVL_COL;
  const C = ECR_CONFIG.COLORS;

  // --- Connect to Master ---
  let masterSS, itemsSheet, bomSheet, amlSheet;
  try {
    masterSS = SpreadsheetApp.openById(ECR_CONFIG.MASTER_ID);
    itemsSheet = masterSS.getSheetByName(ECR_CONFIG.ITEMS_TAB_NAME);
    bomSheet = masterSS.getSheetByName(ECR_CONFIG.BOM_TAB_NAME);
    amlSheet = masterSS.getSheetByName(ECR_CONFIG.AML_TAB_NAME);
    if (!itemsSheet || !bomSheet || !amlSheet) throw new Error("Master tabs missing.");
  } catch(e) { return ui.alert('Connection Failed: ' + e.message); }

  // --- CACHE: ITEMS (Desc/Rev) ---
  const itemsRaw = itemsSheet.getDataRange().getValues();
  const itemMap = new Map();
  for (let i = 1; i < itemsRaw.length; i++) {
    const r = itemsRaw[i];
    if (r.length > ECR_CONFIG.I_IDX.REV) {
      itemMap.set(String(r[ECR_CONFIG.I_IDX.ITEM]).trim(), {
        desc: String(r[ECR_CONFIG.I_IDX.DESC]).trim(),
        rev: String(r[ECR_CONFIG.I_IDX.REV]).trim()
      });
    }
  }

  // --- CACHE: AML (Mfr/MPN per item) ---
  const amlMap = new Map(); // item → [{mfr, mpn}]
  const amlRaw = amlSheet.getDataRange().getValues();
  for (let i = 1; i < amlRaw.length; i++) {
    const item = String(amlRaw[i][0]).trim();
    const mfrName = String(amlRaw[i][1]).trim();
    const mfrPN = String(amlRaw[i][2]).trim();
    if (item) {
      if (!amlMap.has(item)) amlMap.set(item, []);
      const existing = amlMap.get(item);
      // Avoid duplicates
      if (!existing.some(e => e.mfr === mfrName && e.mpn === mfrPN)) {
        existing.push({ mfr: mfrName, mpn: mfrPN });
      }
    }
  }

  // --- CACHE: BOM DATA (Qty & Parent Lookup) ---
  const bomData = bomSheet.getDataRange().getValues();


  // ===================================================================
  // PROCESS: AFFECTED ITEMS (Redline BOM Changes)
  // ===================================================================
  const bomListSheet = ss.getSheetByName(ECR_CONFIG.BOM_LIST_TAB);
  if (bomListSheet) {
    const lastRow = bomListSheet.getLastRow();
    if (lastRow >= ECR_CONFIG.TABLE_START_ROW) {
      const numRows = lastRow - ECR_CONFIG.TABLE_START_ROW + 1;
      const range = bomListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, numRows, ECR_CONFIG.BOM_TOTAL_COLS);
      const values = range.getValues();
      const backgrounds = [];
      const validationsToSet = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        let parentItem = String(row[BC.PARENT]).trim();
        const childItem = String(row[BC.CHILD_PN]).trim();
        const type = String(row[BC.TYPE]).toUpperCase();

        // Initialize row background: gray for auto-fill cols, white for user cols
        const rowBg = new Array(ECR_CONFIG.BOM_TOTAL_COLS).fill(C.EDITABLE);
        [BC.CURR_DESC, BC.CURR_REV, BC.CURR_QTY, BC.STATUS].forEach(idx => { rowBg[idx] = C.AUTO_FILL; });

        if (!childItem) {
          backgrounds.push(rowBg);
          continue;
        }

        const isParentRequired = !["REV ROLL", "DESC CHANGE", "REV ONLY"].includes(type);
        let lookupSuccess = false;

        // --- SMART PARENT LOOKUP ---
        if (isParentRequired && (!parentItem || parentItem === "\u26a0 SELECT PARENT")) {
          const foundParents = findParentsInBom(childItem, bomData);
          if (foundParents.length > 1) {
            values[i][BC.PARENT] = "\u26a0 SELECT PARENT";
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(foundParents, true)
              .setAllowInvalid(false)
              .build();
            validationsToSet.push({ row: i + ECR_CONFIG.TABLE_START_ROW, col: BC.PARENT + 1, rule: rule });
            parentItem = "";
          } else if (foundParents.length === 1) {
            parentItem = foundParents[0];
            values[i][BC.PARENT] = parentItem;
          }
        }

        // --- AUTO-FILL: Curr Desc & Curr Rev (ALWAYS overwrite — these are read-only mirror) ---
        if (itemMap.has(childItem)) {
          const info = itemMap.get(childItem);
          values[i][BC.CURR_DESC] = info.desc;
          values[i][BC.CURR_REV] = info.rev;
        } else if (type === "ADDED") {
          values[i][BC.CURR_DESC] = '(NEW ITEM)';
          values[i][BC.CURR_REV] = '-';
        } else {
          values[i][BC.CURR_DESC] = '(NOT FOUND)';
          values[i][BC.CURR_REV] = '?';
        }

        // --- AUTO-FILL: Curr Qty (ALWAYS overwrite) ---
        if (isParentRequired && parentItem && parentItem !== "\u26a0 SELECT PARENT") {
          let foundQty = "";
          let parentFound = false;
          let parentLevel = -1;
          for (let r = 0; r < bomData.length; r++) {
            const bRow = bomData[r];
            if (bRow.length <= ECR_CONFIG.M_IDX.QTY) continue;
            const rowItem = String(bRow[ECR_CONFIG.M_IDX.ITEM]).trim();
            if (!rowItem) continue;

            const rowLevelRaw = bRow[ECR_CONFIG.M_IDX.LEVEL];
            let rowLevel = isNaN(parseFloat(rowLevelRaw)) ? 0 : parseFloat(rowLevelRaw);

            if (!parentFound && rowItem === parentItem) {
              parentFound = true; parentLevel = rowLevel; continue;
            }
            if (parentFound) {
              if (rowLevel <= parentLevel) break;
              if (rowLevel === parentLevel + 1 && rowItem === childItem) {
                foundQty = bRow[ECR_CONFIG.M_IDX.QTY];
                break;
              }
            }
          }

          if (foundQty !== "") {
            values[i][BC.CURR_QTY] = foundQty;
            lookupSuccess = true;
          }
        } else if (type === "ADDED") {
          values[i][BC.CURR_QTY] = '-';
          lookupSuccess = true; // ADDED items don't need qty lookup
        }

        // --- HIGHLIGHT user "New" columns if they have data (yellow = modified) ---
        if (String(row[BC.NEW_REV]).trim() !== '')  rowBg[BC.NEW_REV]  = C.MODIFIED;
        if (String(row[BC.NEW_QTY]).trim() !== '')  rowBg[BC.NEW_QTY]  = C.MODIFIED;
        if (String(row[BC.NEW_DESC]).trim() !== '') rowBg[BC.NEW_DESC] = C.MODIFIED;

        // --- STATUS & VALIDATION ---
        let status = 'OK';
        if (values[i][BC.PARENT] === "\u26a0 SELECT PARENT") {
          status = 'SELECT PARENT';
          for (let k = 0; k < ECR_CONFIG.BOM_TOTAL_COLS; k++) rowBg[k] = C.ERROR;
          rowBg[BC.STATUS] = C.ERROR;
        } else if (isParentRequired && parentItem && !lookupSuccess && type !== "ADDED") {
          status = 'LOOKUP FAILED';
          for (let k = 0; k < ECR_CONFIG.BOM_TOTAL_COLS; k++) rowBg[k] = C.ERROR;
          rowBg[BC.STATUS] = C.ERROR;
        } else if (isParentRequired && !parentItem && type !== "ADDED") {
          status = 'MISSING PARENT';
          for (let k = 0; k < ECR_CONFIG.BOM_TOTAL_COLS; k++) rowBg[k] = C.ERROR;
          rowBg[BC.STATUS] = C.ERROR;
        } else {
          rowBg[BC.STATUS] = C.SUCCESS;
        }
        values[i][BC.STATUS] = status;

        backgrounds.push(rowBg);
      }

      // WRITE BACK — single batch operation
      range.setValues(values);
      range.setBackgrounds(backgrounds);

      validationsToSet.forEach(v => {
        bomListSheet.getRange(v.row, v.col).setDataValidation(v.rule);
      });
    }
  }


  // ===================================================================
  // PROCESS: AVL CHANGES (Row Expansion + Preserve Manual Data)
  // ===================================================================
  const avlListSheet = ss.getSheetByName(ECR_CONFIG.AVL_LIST_TAB);
  if (avlListSheet) {
    const lastRow = avlListSheet.getLastRow();

    // 1. PRESERVE existing manual data using composite key
    const existingDataMap = new Map();
    const uniqueItems = new Set();

    if (lastRow >= ECR_CONFIG.TABLE_START_ROW) {
      const dataRange = avlListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, lastRow - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.AVL_TOTAL_COLS);
      const currentValues = dataRange.getValues();

      currentValues.forEach(row => {
        const item = String(row[AC.ITEM_PN]).trim();
        if (!item) return;

        uniqueItems.add(item);
        // Composite key: item + current mfr + current mpn (unique per AML row)
        const mfr = String(row[AC.CURR_MFR]).trim();
        const mpn = String(row[AC.CURR_MPN]).trim();
        const compositeKey = `${item}::${mfr}::${mpn}`;

        existingDataMap.set(compositeKey, {
          changeType: row[AC.TYPE],
          newMfr:     row[AC.NEW_MFR],
          newMpn:     row[AC.NEW_MPN],
          disp:       row[AC.DISP],
          reason:     row[AC.REASON]
        });
      });
    }

    // 2. BUILD new table — expand AML rows for each unique item
    if (uniqueItems.size > 0) {
      const newRows = [];
      const newBgs = [];

      uniqueItems.forEach(item => {
        const desc = itemMap.has(item) ? itemMap.get(item).desc : "";

        if (amlMap.has(item)) {
          const vendors = amlMap.get(item);
          vendors.forEach(vendor => {
            const compositeKey = `${item}::${vendor.mfr}::${vendor.mpn}`;
            const preserved = existingDataMap.get(compositeKey) || {};

            newRows.push([
              item,
              desc,
              preserved.changeType || "",
              vendor.mfr,
              vendor.mpn,
              preserved.newMfr || "",
              preserved.newMpn || "",
              preserved.disp || "",
              preserved.reason || "",
              ""  // Status — set below
            ]);

            // Build row background
            const rowBg = new Array(ECR_CONFIG.AVL_TOTAL_COLS).fill(ECR_CONFIG.COLORS.EDITABLE);
            [AC.CURR_DESC, AC.CURR_MFR, AC.CURR_MPN, AC.STATUS].forEach(idx => { rowBg[idx] = ECR_CONFIG.COLORS.AUTO_FILL; });
            if (String(preserved.newMfr || "").trim() !== '') rowBg[AC.NEW_MFR] = ECR_CONFIG.COLORS.MODIFIED;
            if (String(preserved.newMpn || "").trim() !== '') rowBg[AC.NEW_MPN] = ECR_CONFIG.COLORS.MODIFIED;
            newBgs.push(rowBg);
          });
        } else {
          // No existing AVL — show placeholder
          const compositeKey = `${item}::::`; // empty mfr/mpn
          const preserved = existingDataMap.get(compositeKey) || {};

          newRows.push([
            item, desc, preserved.changeType || "",
            "(No AML)", "", preserved.newMfr || "", preserved.newMpn || "",
            preserved.disp || "", preserved.reason || "", ""
          ]);

          const rowBg = new Array(ECR_CONFIG.AVL_TOTAL_COLS).fill(ECR_CONFIG.COLORS.EDITABLE);
          [AC.CURR_DESC, AC.CURR_MFR, AC.CURR_MPN, AC.STATUS].forEach(idx => { rowBg[idx] = ECR_CONFIG.COLORS.AUTO_FILL; });
          newBgs.push(rowBg);
        }
      });

      // 3. WRITE TO SHEET — clear and rewrite
      const numRowsToClear = avlListSheet.getMaxRows() - ECR_CONFIG.TABLE_START_ROW + 1;
      if (numRowsToClear > 0) {
        avlListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, numRowsToClear, ECR_CONFIG.AVL_TOTAL_COLS).clearContent().setBackground(null);
      }

      if (newRows.length > 0) {
        avlListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, newRows.length, ECR_CONFIG.AVL_TOTAL_COLS).setValues(newRows);
        avlListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, newRows.length, ECR_CONFIG.AVL_TOTAL_COLS).setBackgrounds(newBgs);
      }
    }
  }

  ui.alert("Auto-Fill Complete.\n\nGray columns = current master data (auto-filled)\nYellow cells = your manual entries (preserved)");
}


// ====================================================================================================
// === FEATURE 2: SUBMIT TO MASTER LOG ================================================================
// ====================================================================================================

/**
 * Submits ECR data to the Master's ECR_Affected_Items log sheet.
 * Reads from the new redline column layout.
 */
function submitComprehensiveECR() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const BC = ECR_CONFIG.BOM_COL;
  const AC = ECR_CONFIG.AVL_COL;

  const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);
  if (!formSheet) return ui.alert("Form tab missing.");

  const ecrNum = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_NUM).getValue();
  const ecoNum = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECO_NUM).getValue();
  if (!ecrNum || !ecoNum) return ui.alert("Missing ECR/ECO Number.");

  const rowsToPush = [];

  // --- BOM Sheet (Affected_Items) ---
  const bomSheet = ss.getSheetByName(ECR_CONFIG.BOM_LIST_TAB);
  if (bomSheet && bomSheet.getLastRow() >= ECR_CONFIG.TABLE_START_ROW) {
    const data = bomSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, bomSheet.getLastRow() - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.BOM_TOTAL_COLS).getValues();
    const validTypes = ["ADDED","REMOVED","MODIFIED","QTY CHANGE","REV ROLL","DESC CHANGE"];

    data.forEach(row => {
      const item = row[BC.CHILD_PN];
      const type = String(row[BC.TYPE]).toUpperCase();

      if (item && validTypes.includes(type)) {
        rowsToPush.push([
          ecrNum, ecoNum, item, row[BC.PARENT], type,
          row[BC.CURR_REV], row[BC.NEW_REV], row[BC.CURR_QTY], row[BC.NEW_QTY],
          "", "",
          row[BC.NEW_DESC], row[BC.DISP], row[BC.REASON]
        ]);
      }
    });
  }

  // --- AVL Sheet (AVL_Changes) ---
  const avlSheet = ss.getSheetByName(ECR_CONFIG.AVL_LIST_TAB);
  if (avlSheet && avlSheet.getLastRow() >= ECR_CONFIG.TABLE_START_ROW) {
    const data = avlSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, avlSheet.getLastRow() - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.AVL_TOTAL_COLS).getValues();

    data.forEach(row => {
      const item = row[AC.ITEM_PN];
      const type = String(row[AC.TYPE]).toUpperCase();

      if (item && ["AVL_ADD", "AVL_REMOVE", "AVL_REPLACE"].includes(type)) {
        // For REMOVE: log the current mfr/mpn being removed
        // For ADD/REPLACE: log the new mfr/mpn
        let logMfr = row[AC.NEW_MFR];
        let logMpn = row[AC.NEW_MPN];
        if (type === "AVL_REMOVE") {
          logMfr = row[AC.CURR_MFR];
          logMpn = row[AC.CURR_MPN];
        }

        rowsToPush.push([
          ecrNum, ecoNum, item, "N/A", type,
          "", "", "", "",
          logMfr, logMpn,
          "", row[AC.DISP], row[AC.REASON]
        ]);
      }
    });
  }

  if (rowsToPush.length === 0) return ui.alert("No valid rows to submit.");

  try {
    const masterSS = SpreadsheetApp.openById(ECR_CONFIG.MASTER_ID);
    const logSheet = masterSS.getSheetByName(ECR_CONFIG.ECR_LOG_TAB_NAME);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, rowsToPush.length, rowsToPush[0].length).setValues(rowsToPush);
    ui.alert(`Success! Submitted ${rowsToPush.length} items.`);
  } catch(e) {
    ui.alert("Error: " + e.message);
  }
}


// ====================================================================================================
// === FEATURE 3: COMMIT TO MASTER (ADMIN) ============================================================
// ====================================================================================================

/**
 * Commits ECR changes to the Master BOM/AML/ITEMS sheets.
 * Reads from the new redline column layout — uses "New *" columns for the target values.
 */
function commitToMaster() {
  const ui = SpreadsheetApp.getUi();
  const BC = ECR_CONFIG.BOM_COL;
  const AC = ECR_CONFIG.AVL_COL;

  // --- PASSWORD PROTECTION ---
  const storedPwd = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!storedPwd) {
    return ui.alert('\u26a0 Setup Required',
      'Admin password has not been set.\nUse the menu: ECR Actions \u2192 \ud83d\udd11 Set Admin Password',
      ui.ButtonSet.OK);
  }

  const pwdResponse = ui.prompt('\ud83d\udd12 ADMIN AUTHENTICATION', 'Enter Admin Password to proceed:', ui.ButtonSet.OK_CANCEL);
  if (pwdResponse.getSelectedButton() !== ui.Button.OK) return;
  if (pwdResponse.getResponseText() !== storedPwd) {
    return ui.alert('\u274c ACCESS DENIED', 'Incorrect Password.', ui.ButtonSet.OK);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Connect to Master ---
  let masterSS, masterSheet, amlSheet, itemsSheet;
  let masterData, amlData, itemsData;
  let masterAmlItems = new Set();

  try {
    masterSS = SpreadsheetApp.openById(ECR_CONFIG.MASTER_ID);
    masterSheet = masterSS.getSheetByName(ECR_CONFIG.BOM_TAB_NAME);
    amlSheet = masterSS.getSheetByName(ECR_CONFIG.AML_TAB_NAME);
    itemsSheet = masterSS.getSheetByName(ECR_CONFIG.ITEMS_TAB_NAME);

    amlData = amlSheet.getDataRange().getValues();
    amlData.forEach(r => masterAmlItems.add(String(r[ECR_CONFIG.A_IDX.ITEM]).trim()));
    masterData = masterSheet.getDataRange().getValues();
    itemsData = itemsSheet.getDataRange().getValues();
  } catch(e) { return ui.alert('Connection Failed: ' + e.message); }

  // --- PRE-FLIGHT: Check for missing AVL on new items ---
  const localBomSheet = ss.getSheetByName(ECR_CONFIG.BOM_LIST_TAB);
  const localAvlSheet = ss.getSheetByName(ECR_CONFIG.AVL_LIST_TAB);

  if (localBomSheet && localAvlSheet &&
      localBomSheet.getLastRow() >= ECR_CONFIG.TABLE_START_ROW &&
      localAvlSheet.getLastRow() >= ECR_CONFIG.TABLE_START_ROW) {
    const bRows = localBomSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, localBomSheet.getLastRow() - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.BOM_TOTAL_COLS).getValues();
    const aRows = localAvlSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, localAvlSheet.getLastRow() - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.AVL_TOTAL_COLS).getValues();

    const newBomItems = new Set(
      bRows.filter(r => String(r[BC.TYPE]).toUpperCase() === 'ADDED').map(r => String(r[BC.CHILD_PN]).trim())
    );
    const newAvlItems = new Set(
      aRows.filter(r => String(r[AC.TYPE]).toUpperCase() === 'AVL_ADD').map(r => String(r[AC.ITEM_PN]).trim())
    );

    const missingAvl = [...newBomItems].filter(x =>
      x !== "" && !newAvlItems.has(x) && !masterAmlItems.has(x)
    );

    if (missingAvl.length > 0) {
      const confirm = ui.alert('\u26a0\ufe0f POTENTIAL DATA GAP',
        `Items being ADDED with no AVL:\n\n${missingAvl.join(', ')}\n\nProceed anyway?`,
        ui.ButtonSet.YES_NO);
      if (confirm !== ui.Button.YES) return;
    }
  }

  // --- WORKFLOW GATE: Only APPROVED ECRs can be committed ---
  const ecrStatus = getEcrWorkflowState_();
  if (ecrStatus !== ECR_CONFIG.WORKFLOW.COMMIT_REQUIRED_STATE) {
    const stateMsg = ecrStatus ? ecrStatus : '(no status set)';
    return ui.alert('\u26d4 WORKFLOW BLOCKED',
      `ECR must be in "${ECR_CONFIG.WORKFLOW.COMMIT_REQUIRED_STATE}" state to commit.\n\n` +
      `Current state: ${stateMsg}\n\n` +
      'Use ECR Actions > Workflow to advance the ECR through the approval process.',
      ui.ButtonSet.OK);
  }

  const response = ui.alert('\u26a0 ADMIN ACTION: COMMIT TO MASTER',
    'This will permanently modify the MASTER BOM and AML sheets.\n\nAre you sure?',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // BUILD ITEMS MAP
  const itemsRowMap = new Map();
  for (let r = 1; r < itemsData.length; r++) {
    const iName = String(itemsData[r][ECR_CONFIG.I_IDX.ITEM]).trim();
    if (iName) itemsRowMap.set(iName, r + 1);
  }

  // --- 1. PROCESS AVL CHANGES ---
  const avlListSheet = ss.getSheetByName(ECR_CONFIG.AVL_LIST_TAB);
  if (avlListSheet && avlListSheet.getLastRow() >= ECR_CONFIG.TABLE_START_ROW) {
    const changes = avlListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, avlListSheet.getLastRow() - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.AVL_TOTAL_COLS).getValues();

    changes.forEach(row => {
      const item = String(row[AC.ITEM_PN]).trim();
      const type = String(row[AC.TYPE]).toUpperCase();
      const currMfr = String(row[AC.CURR_MFR]).trim();
      const newMfr = String(row[AC.NEW_MFR]).trim();
      const newMpn = String(row[AC.NEW_MPN]).trim();

      if (!item) return;

      const findAmlIndex = (searchMfr) => {
        for (let i = 0; i < amlData.length; i++) {
          if (String(amlData[i][ECR_CONFIG.A_IDX.ITEM]).trim() === item &&
              String(amlData[i][ECR_CONFIG.A_IDX.MFR]).trim() === searchMfr) {
            return i;
          }
        }
        return -1;
      };

      if (type === "AVL_ADD") {
        let exists = false;
        for (let i = 0; i < amlData.length; i++) {
          if (String(amlData[i][ECR_CONFIG.A_IDX.ITEM]).trim() === item &&
              String(amlData[i][ECR_CONFIG.A_IDX.MFR]).trim() === newMfr &&
              String(amlData[i][ECR_CONFIG.A_IDX.MPN]).trim() === newMpn) {
            exists = true; break;
          }
        }
        if (!exists && newMfr && newMpn) {
          amlSheet.appendRow([item, newMfr, newMpn]);
          amlData.push([item, newMfr, newMpn]);
        }
      }
      else if (type === "AVL_REMOVE" || type === "AVL_REPLACE") {
        const arrayIndex = findAmlIndex(currMfr);
        if (arrayIndex > -1) {
          const sheetRowIndex = arrayIndex + 1;
          if (type === "AVL_REMOVE") {
            amlSheet.deleteRow(sheetRowIndex);
            amlData.splice(arrayIndex, 1);
          } else if (type === "AVL_REPLACE") {
            amlSheet.getRange(sheetRowIndex, ECR_CONFIG.A_IDX.MFR + 1).setValue(newMfr);
            amlSheet.getRange(sheetRowIndex, ECR_CONFIG.A_IDX.MPN + 1).setValue(newMpn);
            amlData[arrayIndex][ECR_CONFIG.A_IDX.MFR] = newMfr;
            amlData[arrayIndex][ECR_CONFIG.A_IDX.MPN] = newMpn;
          }
        }
      }
    });
  }

  // --- 2. PROCESS BOM CHANGES ---
  const bomListSheet = ss.getSheetByName(ECR_CONFIG.BOM_LIST_TAB);
  if (bomListSheet && bomListSheet.getLastRow() >= ECR_CONFIG.TABLE_START_ROW) {
    const changes = bomListSheet.getRange(ECR_CONFIG.TABLE_START_ROW, 1, bomListSheet.getLastRow() - ECR_CONFIG.TABLE_START_ROW + 1, ECR_CONFIG.BOM_TOTAL_COLS).getValues();

    changes.forEach(row => {
      const parent = String(row[BC.PARENT]).trim();
      const child = String(row[BC.CHILD_PN]).trim();
      const type = String(row[BC.TYPE]).toUpperCase();
      const newRev = String(row[BC.NEW_REV]).trim();
      const newQty = row[BC.NEW_QTY];
      const newDesc = String(row[BC.NEW_DESC]).trim();

      if (!child) return;

      // SEARCH for parent→child link in masterData
      let parentFound = false;
      let parentLevel = -1;
      let targetArrayIndex = -1;
      let insertAfterArrayIndex = -1;

      const isBomUpdateNeeded = ["ADDED", "REMOVED", "MODIFIED", "QTY CHANGE"].includes(type);

      if (isBomUpdateNeeded && parent) {
        for (let i = 0; i < masterData.length; i++) {
          const rItem = String(masterData[i][ECR_CONFIG.M_IDX.ITEM]).trim();
          if (!rItem) continue;

          const rLevelRaw = masterData[i][ECR_CONFIG.M_IDX.LEVEL];
          const rLevel = isNaN(parseFloat(rLevelRaw)) ? 0 : parseFloat(rLevelRaw);

          if (!parentFound && rItem === parent) {
            parentFound = true;
            parentLevel = rLevel;
            insertAfterArrayIndex = i;
            continue;
          }

          if (parentFound) {
            if (rLevel <= parentLevel) break;
            if (rItem === child && (rLevel === parentLevel + 1)) {
              targetArrayIndex = i;
              break;
            }
          }
        }
      }

      // --- EXECUTION BRANCHES ---

      // CASE A: REV ROLL or DESC CHANGE
      if (type === "REV ROLL" || type === "DESC CHANGE") {
        if (itemsRowMap.has(child)) {
          const itemRow = itemsRowMap.get(child);
          if (newRev !== "") itemsSheet.getRange(itemRow, ECR_CONFIG.I_IDX.REV + 1).setValue(newRev);
          if (newDesc !== "") itemsSheet.getRange(itemRow, ECR_CONFIG.I_IDX.DESC + 1).setValue(newDesc);
        }
      }

      // CASE B: QTY CHANGE
      else if (type === "QTY CHANGE") {
        if (targetArrayIndex > -1 && newQty !== "" && newQty !== null) {
          masterSheet.getRange(targetArrayIndex + 1, ECR_CONFIG.M_IDX.QTY + 1).setValue(newQty);
          masterData[targetArrayIndex][ECR_CONFIG.M_IDX.QTY] = newQty;
        }
      }

      // CASE C: MODIFIED
      else if (type === "MODIFIED") {
        if (targetArrayIndex > -1 && newQty !== "" && newQty !== null) {
          masterSheet.getRange(targetArrayIndex + 1, ECR_CONFIG.M_IDX.QTY + 1).setValue(newQty);
          masterData[targetArrayIndex][ECR_CONFIG.M_IDX.QTY] = newQty;
        }
        if (itemsRowMap.has(child)) {
          const itemRow = itemsRowMap.get(child);
          if (newRev !== "") itemsSheet.getRange(itemRow, ECR_CONFIG.I_IDX.REV + 1).setValue(newRev);
          if (newDesc !== "") itemsSheet.getRange(itemRow, ECR_CONFIG.I_IDX.DESC + 1).setValue(newDesc);
        }
      }

      // CASE D: REMOVED (CASCADE)
      else if (type === "REMOVED") {
        if (targetArrayIndex > -1) {
          const targetLevel = parseFloat(masterData[targetArrayIndex][ECR_CONFIG.M_IDX.LEVEL]);
          let rowsToDelete = 1;
          for (let scan = targetArrayIndex + 1; scan < masterData.length; scan++) {
            const scanItem = String(masterData[scan][ECR_CONFIG.M_IDX.ITEM]).trim();
            if (!scanItem) { rowsToDelete++; continue; }
            const scanLevel = parseFloat(masterData[scan][ECR_CONFIG.M_IDX.LEVEL]);
            if (isNaN(scanLevel) || scanLevel <= targetLevel) break;
            rowsToDelete++;
          }
          for (let d = rowsToDelete - 1; d >= 0; d--) {
            masterSheet.deleteRow(targetArrayIndex + 1 + d);
          }
          masterData.splice(targetArrayIndex, rowsToDelete);
        }
      }

      // CASE E: ADDED
      else if (type === "ADDED" && parentFound) {
        let desc = "";
        if (itemsRowMap.has(child)) {
          const rIndex = itemsRowMap.get(child) - 1;
          if (rIndex < itemsData.length) {
            desc = itemsData[rIndex][ECR_CONFIG.I_IDX.DESC];
          } else {
            desc = newDesc !== "" ? newDesc : "";
          }
        } else {
          desc = newDesc !== "" ? newDesc : "";
          const rev = newRev !== "" ? newRev : "A";
          itemsSheet.appendRow([child, desc, rev]);
          itemsRowMap.set(child, itemsSheet.getLastRow());
        }

        if (targetArrayIndex > -1) {
          if (newQty !== "" && newQty !== null) {
            masterSheet.getRange(targetArrayIndex + 1, ECR_CONFIG.M_IDX.QTY + 1).setValue(newQty);
            masterData[targetArrayIndex][ECR_CONFIG.M_IDX.QTY] = newQty;
          }
        } else {
          const sheetInsertRow = insertAfterArrayIndex + 2;
          const itemRef = `C${sheetInsertRow}`;
          const f_Desc = `=IF(ISBLANK(${itemRef}), "", VLOOKUP(${itemRef}, ITEMS!A:C, 2, FALSE))`;
          const f_Rev = `=IF(ISBLANK(${itemRef}), "", VLOOKUP(${itemRef}, ITEMS!A:C, 3, FALSE))`;
          const f_Life = `=IF(ISBLANK(${itemRef}), "", VLOOKUP(${itemRef}, ITEMS!A:D, 4, FALSE))`;
          const f_Mfr = `=IFNA(IF(ISBLANK(${itemRef}), "", FILTER(AML!B:C, AML!A:A = ${itemRef})), "No AML Found")`;

          const newRowData = [];
          for (let c = 0; c < ECR_CONFIG.M_IDX.LEVEL; c++) newRowData.push("");
          newRowData.push(parentLevel + 1);
          newRowData.push(child);
          newRowData.push(f_Desc);
          newRowData.push(f_Rev);
          newRowData.push(f_Life);
          newRowData.push(newQty !== "" && newQty !== null ? newQty : 1);
          for (let k = 0; k < 5; k++) newRowData.push("");
          newRowData.push(f_Mfr);
          newRowData.push("");

          masterSheet.insertRowAfter(insertAfterArrayIndex + 1);
          masterSheet.getRange(sheetInsertRow, 1, 1, newRowData.length).setValues([newRowData]);
          masterData.splice(insertAfterArrayIndex + 1, 0, newRowData);
        }
      }
    });
  }

  // --- AUTO-TRANSITION: APPROVED → IMPLEMENTING → CLOSED ---
  transitionEcrState_('IMPLEMENTING', 'commitToMaster (auto)');
  transitionEcrState_('CLOSED', 'commitToMaster (auto)');

  ui.alert("Commit Complete. Master Data has been updated.\n\nECR status has been set to CLOSED.");
}


// ====================================================================================================
// === HELPER FUNCTIONS ===============================================================================
// ====================================================================================================

/**
 * HELPER: Find parent assemblies for a child item in the BOM data.
 */
function findParentsInBom(childItem, bomData) {
  const parents = new Set();
  const childItemStr = String(childItem).trim();

  for (let i = 0; i < bomData.length; i++) {
    const row = bomData[i];
    if (row.length <= ECR_CONFIG.M_IDX.ITEM) continue;

    const rowItem = String(row[ECR_CONFIG.M_IDX.ITEM]).trim();

    if (rowItem === childItemStr) {
      const rowLevelRaw = row[ECR_CONFIG.M_IDX.LEVEL];
      const childLevel = isNaN(parseFloat(rowLevelRaw)) ? 0 : parseFloat(rowLevelRaw);

      for (let j = i - 1; j >= 0; j--) {
        const parentRow = bomData[j];
        const pItem = String(parentRow[ECR_CONFIG.M_IDX.ITEM]).trim();
        if (!pItem) continue;

        const parentLevelRaw = parentRow[ECR_CONFIG.M_IDX.LEVEL];
        const parentLevel = isNaN(parseFloat(parentLevelRaw)) ? 0 : parseFloat(parentLevelRaw);

        if (parentLevel < childLevel) {
          parents.add(pItem);
          break;
        }
      }
    }
  }
  return Array.from(parents);
}


// ====================================================================================================
// === ADMIN PASSWORD MANAGEMENT ======================================================================
// ====================================================================================================

function setAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const currentPwd = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');

  if (currentPwd) {
    const verifyResponse = ui.prompt('\ud83d\udd12 Verify Current Password',
      'Enter your CURRENT admin password:', ui.ButtonSet.OK_CANCEL);
    if (verifyResponse.getSelectedButton() !== ui.Button.OK) return;
    if (verifyResponse.getResponseText() !== currentPwd) {
      return ui.alert('\u274c ACCESS DENIED', 'Incorrect current password.', ui.ButtonSet.OK);
    }
  }

  const newPwdResponse = ui.prompt('\ud83d\udd11 Set Admin Password',
    'Enter the NEW admin password (minimum 6 characters):', ui.ButtonSet.OK_CANCEL);
  if (newPwdResponse.getSelectedButton() !== ui.Button.OK) return;

  const newPwd = newPwdResponse.getResponseText().trim();
  if (newPwd.length < 6) {
    return ui.alert('\u274c Too Short', 'Password must be at least 6 characters.', ui.ButtonSet.OK);
  }

  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', newPwd);
  ui.alert('\u2705 Password Updated', 'Admin password set successfully.', ui.ButtonSet.OK);
}


// ====================================================================================================
// === DEBUGGER TOOL ==================================================================================
// ====================================================================================================

function DEBUG_Check_Parent_Child_Link() {
  const ui = SpreadsheetApp.getUi();
  const promptParent = ui.prompt('DEBUGGER', 'Enter Exact PARENT Name:', ui.ButtonSet.OK_CANCEL);
  if (promptParent.getSelectedButton() !== ui.Button.OK) return;
  const pName = promptParent.getResponseText().trim();

  const promptChild = ui.prompt('DEBUGGER', 'Enter Exact CHILD Name:', ui.ButtonSet.OK_CANCEL);
  if (promptChild.getSelectedButton() !== ui.Button.OK) return;
  const cName = promptChild.getResponseText().trim();

  Logger.log(`STARTING DEBUG for Parent: [${pName}] and Child: [${cName}]`);

  try {
    const masterSS = SpreadsheetApp.openById(ECR_CONFIG.MASTER_ID);
    const bomSheet = masterSS.getSheetByName(ECR_CONFIG.BOM_TAB_NAME);
    const bomData = bomSheet.getDataRange().getValues();

    let parentFound = false;
    let parentLevel = -1;
    let parentRowIndex = -1;

    for (let i = 0; i < bomData.length; i++) {
      const row = bomData[i];
      const rItem = String(row[ECR_CONFIG.M_IDX.ITEM]).trim();
      if (!rItem) continue;

      const rLevel = isNaN(parseFloat(row[ECR_CONFIG.M_IDX.LEVEL])) ? 0 : parseFloat(row[ECR_CONFIG.M_IDX.LEVEL]);

      if (!parentFound && rItem === pName) {
        parentFound = true;
        parentLevel = rLevel;
        parentRowIndex = i + 1;
        Logger.log(`\u2705 PARENT FOUND at Row ${parentRowIndex}. Level: ${parentLevel}`);
        continue;
      }

      if (parentFound) {
        if (rLevel <= parentLevel) {
          Logger.log(`\u26d4 End of Parent Block at Row ${i + 1}. Item: ${rItem} (Level ${rLevel})`);
          break;
        }

        if (rLevel === parentLevel + 1) {
          const match = (rItem === cName) ? "MATCH!" : "No Match";
          Logger.log(`   -> Child Candidate Row ${i+1}: [${rItem}] vs [${cName}] ... ${match}`);
        }
      }
    }

    if (!parentFound) Logger.log("\u274c PARENT NOT FOUND in Master BOM.");

  } catch(e) {
    Logger.log("ERROR: " + e.message);
  }

  ui.alert("Debug Complete. Check 'View > Executions' or 'View > Logs' to see the report.");
}


// ====================================================================================================
// === ECR WORKFLOW STATE MACHINE ======================================================================
// ====================================================================================================

/**
 * Gets the current ECR workflow state from the form sheet.
 * @returns {string} Current state or '' if not set.
 */
function getEcrWorkflowState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);
  if (!formSheet) return '';
  const val = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_STATUS).getValue();
  return val ? val.toString().trim().toUpperCase() : '';
}

/**
 * Sets the ECR workflow state on the form sheet.
 * @param {string} state New state value.
 */
function setEcrWorkflowState_(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);
  if (!formSheet) return;
  formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_STATUS).setValue(state);
}

/**
 * Validates and executes a workflow transition.
 * @param {string} targetState The desired next state.
 * @param {string} source What triggered the transition.
 * @param {string} [rejectionReason] Reason for rejection (only for REJECTED transitions).
 * @returns {boolean} True if transition succeeded.
 */
function transitionEcrState_(targetState, source, rejectionReason) {
  const currentState = getEcrWorkflowState_() || 'DRAFT';
  const target = targetState.toUpperCase();

  // Validate transition
  const allowed = ECR_CONFIG.WORKFLOW.TRANSITIONS[currentState] || [];
  if (!allowed.includes(target)) {
    return false;
  }

  // Execute transition
  setEcrWorkflowState_(target);

  // Set reviewer info for review/approval/rejection
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);
  if (formSheet) {
    let userEmail = '';
    try { userEmail = Session.getActiveUser().getEmail(); } catch (e) { userEmail = 'Unknown'; }

    if (['UNDER_REVIEW', 'APPROVED', 'REJECTED'].includes(target)) {
      formSheet.getRange(ECR_CONFIG.FORM_CELLS.REVIEWER).setValue(userEmail);
      formSheet.getRange(ECR_CONFIG.FORM_CELLS.REVIEW_DATE).setValue(new Date());
    }
    if (target === 'REJECTED' && rejectionReason) {
      formSheet.getRange(ECR_CONFIG.FORM_CELLS.REJECTION_REASON).setValue(rejectionReason);
    }
    if (target === 'DRAFT') {
      // Clear rejection reason when reopened
      formSheet.getRange(ECR_CONFIG.FORM_CELLS.REJECTION_REASON).setValue('');
    }
  }

  // Log to workflow history on Master spreadsheet
  logEcrWorkflowTransition_(currentState, target, source, rejectionReason);

  return true;
}

/**
 * Logs a workflow transition to the ECR_Workflow_History sheet on the Master spreadsheet.
 */
function logEcrWorkflowTransition_(fromState, toState, source, rejectionReason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);
    const ecrNum = formSheet ? formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_NUM).getValue() : '';
    const ecoNum = formSheet ? formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECO_NUM).getValue() : '';

    const masterSS = SpreadsheetApp.openById(ECR_CONFIG.MASTER_ID);
    const sheetName = ECR_CONFIG.WORKFLOW.HISTORY_SHEET_NAME;
    let historySheet = masterSS.getSheetByName(sheetName);

    if (!historySheet) {
      historySheet = masterSS.insertSheet(sheetName);
      const headers = ['ECR #', 'ECO #', 'From State', 'To State', 'Transitioned By', 'Date', 'Source', 'Rejection Reason'];
      historySheet.appendRow(headers);
      historySheet.setFrozenRows(1);
      historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    let userEmail = '';
    try { userEmail = Session.getActiveUser().getEmail(); } catch (e) { userEmail = 'Unknown'; }

    historySheet.appendRow([
      ecrNum, ecoNum, fromState, toState, userEmail, new Date(), source || 'Manual', rejectionReason || ''
    ]);
  } catch (e) {
    Logger.log('Failed to log workflow transition: ' + e.message);
  }
}


// --- Workflow Transition Menu Actions ---

/** Submit ECR for review: DRAFT → SUBMITTED */
function ecrTransition_Submit() {
  const ui = SpreadsheetApp.getUi();
  const currentState = getEcrWorkflowState_() || 'DRAFT';

  if (currentState !== 'DRAFT') {
    return ui.alert('Cannot Submit', `ECR must be in DRAFT state to submit.\nCurrent state: ${currentState}`, ui.ButtonSet.OK);
  }

  // Validate ECR/ECO numbers are filled
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);
  if (formSheet) {
    const ecrNum = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_NUM).getValue();
    const ecoNum = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECO_NUM).getValue();
    if (!ecrNum || !ecoNum) {
      return ui.alert('Missing Data', 'ECR and ECO numbers must be filled before submitting.', ui.ButtonSet.OK);
    }
  }

  if (transitionEcrState_('SUBMITTED', 'Menu: Submit')) {
    ui.alert('ECR Submitted', 'ECR has been submitted for review.\nStatus: SUBMITTED', ui.ButtonSet.OK);
  }
}

/** Begin review: SUBMITTED → UNDER_REVIEW */
function ecrTransition_BeginReview() {
  const ui = SpreadsheetApp.getUi();
  const currentState = getEcrWorkflowState_();

  if (currentState !== 'SUBMITTED') {
    return ui.alert('Cannot Review', `ECR must be in SUBMITTED state.\nCurrent state: ${currentState}`, ui.ButtonSet.OK);
  }

  if (transitionEcrState_('UNDER_REVIEW', 'Menu: Begin Review')) {
    ui.alert('Review Started', 'ECR is now under review.\nStatus: UNDER_REVIEW', ui.ButtonSet.OK);
  }
}

/** Approve ECR: UNDER_REVIEW → APPROVED */
function ecrTransition_Approve() {
  const ui = SpreadsheetApp.getUi();
  const currentState = getEcrWorkflowState_();

  if (currentState !== 'UNDER_REVIEW') {
    return ui.alert('Cannot Approve', `ECR must be in UNDER_REVIEW state.\nCurrent state: ${currentState}`, ui.ButtonSet.OK);
  }

  const confirm = ui.alert('Confirm Approval',
    'Are you sure you want to approve this ECR?\nThis will allow the Admin to commit changes to the Master BOM.',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  if (transitionEcrState_('APPROVED', 'Menu: Approve')) {
    ui.alert('ECR Approved', 'ECR has been approved.\nStatus: APPROVED\n\nThe Admin can now use "Commit to Master" to apply changes.', ui.ButtonSet.OK);
  }
}

/** Reject ECR: SUBMITTED/UNDER_REVIEW → REJECTED */
function ecrTransition_Reject() {
  const ui = SpreadsheetApp.getUi();
  const currentState = getEcrWorkflowState_();

  if (currentState !== 'SUBMITTED' && currentState !== 'UNDER_REVIEW') {
    return ui.alert('Cannot Reject', `ECR must be in SUBMITTED or UNDER_REVIEW state.\nCurrent state: ${currentState}`, ui.ButtonSet.OK);
  }

  const reasonResponse = ui.prompt('Rejection Reason',
    'Please provide a reason for rejecting this ECR:', ui.ButtonSet.OK_CANCEL);
  if (reasonResponse.getSelectedButton() !== ui.Button.OK) return;

  const reason = reasonResponse.getResponseText().trim();
  if (!reason) {
    return ui.alert('Reason Required', 'A rejection reason is required.', ui.ButtonSet.OK);
  }

  if (transitionEcrState_('REJECTED', 'Menu: Reject', reason)) {
    ui.alert('ECR Rejected', `ECR has been rejected.\nReason: ${reason}\n\nThe submitter can reopen as DRAFT to revise.`, ui.ButtonSet.OK);
  }
}

/** Reopen rejected ECR: REJECTED → DRAFT */
function ecrTransition_Reopen() {
  const ui = SpreadsheetApp.getUi();
  const currentState = getEcrWorkflowState_();

  if (currentState !== 'REJECTED') {
    return ui.alert('Cannot Reopen', `Only REJECTED ECRs can be reopened.\nCurrent state: ${currentState}`, ui.ButtonSet.OK);
  }

  if (transitionEcrState_('DRAFT', 'Menu: Reopen')) {
    ui.alert('ECR Reopened', 'ECR has been reopened as DRAFT for revision.\nPrevious rejection reason has been cleared.', ui.ButtonSet.OK);
  }
}

/** View current workflow status */
function viewEcrWorkflowStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(ECR_CONFIG.FORM_TAB_NAME);

  if (!formSheet) {
    return ui.alert('Error', 'ECR_Form sheet not found.', ui.ButtonSet.OK);
  }

  const ecrNum = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_NUM).getValue() || '(not set)';
  const ecoNum = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECO_NUM).getValue() || '(not set)';
  const status = formSheet.getRange(ECR_CONFIG.FORM_CELLS.ECR_STATUS).getValue() || 'DRAFT';
  const reviewer = formSheet.getRange(ECR_CONFIG.FORM_CELLS.REVIEWER).getValue() || '(none)';
  const reviewDate = formSheet.getRange(ECR_CONFIG.FORM_CELLS.REVIEW_DATE).getValue() || '(none)';
  const rejection = formSheet.getRange(ECR_CONFIG.FORM_CELLS.REJECTION_REASON).getValue() || '';

  const currentState = status.toString().toUpperCase() || 'DRAFT';
  const nextStates = ECR_CONFIG.WORKFLOW.TRANSITIONS[currentState] || [];

  let statusLines = [
    `ECR #: ${ecrNum}`,
    `ECO #: ${ecoNum}`,
    ``,
    `Current Status: ${currentState}`,
    `Reviewer: ${reviewer}`,
    `Review Date: ${reviewDate}`,
  ];

  if (rejection) {
    statusLines.push(`Rejection Reason: ${rejection}`);
  }

  statusLines.push('');
  statusLines.push(`Allowed Next States: ${nextStates.length > 0 ? nextStates.join(', ') : '(none — terminal state)'}`);
  statusLines.push('');
  statusLines.push('Workflow: DRAFT → SUBMITTED → UNDER_REVIEW → APPROVED → IMPLEMENTING → CLOSED');

  ui.alert('ECR Workflow Status', statusLines.join('\n'), ui.ButtonSet.OK);
}
