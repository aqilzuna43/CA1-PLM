/**
 * ==================================================
 * ECR AUTOMATION SUITE (ROW-EXPANSION FOR AVL)
 * 1. Rev from ITEMS (Master Definition)
 * 2. Qty from BOM (Usage Definition)
 * 3. AVL Expanded Rows (One Vendor per Row)
 *   // UPDATED: Master Sheet ID
  //1CNvUEd4BQrh35LBQQa-hEZa9glhCOxhXytilTnTXnxg - EVAL
  //1rytdEh6xi8FrHKVN_ZOOTGXJFIZqrNWEYrPD0-ogk-Q - MAIN
 * ==================================================
 */

const CONFIG = {
  // UPDATED: Master Sheet ID
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

  FORM_CELLS: { ECR_NUM: 'B4', ECO_NUM: 'E4' },
  TABLE_START_ROW: 2, 

  // Master BOM Indexes (Zero-based for array logic, Add 1 for Sheet logic if needed)
  M_IDX: { LEVEL: 1, ITEM: 2, QTY: 6 },
  // AML Indexes
  A_IDX: { ITEM: 0, MFR: 1, MPN: 2 },
  // ITEMS Sheet Indexes
  I_IDX: { ITEM: 0, DESC: 1, REV: 2 }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ECR Actions')
    .addItem('▶ Auto-Fill Data (BOM + AVL)', 'populateCurrentData')
    .addItem('▶ Submit to Master Log', 'submitComprehensiveECR')
    .addSeparator()
    .addItem('⚠ ADMIN: Commit to Master', 'commitToMaster')
    .addSeparator()
    .addItem('❓ Troubleshoot Link', 'DEBUG_Check_Parent_Child_Link')
    .addToUi();
}

/**
 * FEATURE 1: AUTO-FILL (EXPANDS AVL ROWS)
 * UPDATED: Includes Smart Dropdowns and Validation Coloring
 * UPDATED: Skips Parent Validation for "REV ROLL" / "DESC CHANGE"
 * UPDATED: Ignores Blank Rows in Master BOM
 */
function populateCurrentData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Connect
  let masterSS, itemsSheet, bomSheet, amlSheet;
  try {
    masterSS = SpreadsheetApp.openById(CONFIG.MASTER_ID);
    itemsSheet = masterSS.getSheetByName(CONFIG.ITEMS_TAB_NAME);
    bomSheet = masterSS.getSheetByName(CONFIG.BOM_TAB_NAME);
    amlSheet = masterSS.getSheetByName(CONFIG.AML_TAB_NAME);
    if (!itemsSheet || !bomSheet || !amlSheet) throw new Error("Master tabs missing.");
  } catch(e) { return ui.alert('Connection Failed: ' + e.message); }

  // --- CACHE 1: ITEMS (Desc/Rev) ---
  const itemsRaw = itemsSheet.getDataRange().getValues();
  const itemMap = new Map(); 
  for (let i = 1; i < itemsRaw.length; i++) {
    const r = itemsRaw[i];
    if (r.length > CONFIG.I_IDX.REV) {
      itemMap.set(String(r[CONFIG.I_IDX.ITEM]).trim(), {
        desc: String(r[CONFIG.I_IDX.DESC]).trim(),
        rev: String(r[CONFIG.I_IDX.REV]).trim()
      });
    }
  }

  // --- CACHE 2: AVL DATA (Mfr Info) ---
  const amlMap = new Map();
  const amlRaw = amlSheet.getDataRange().getValues();
  for (let i = 1; i < amlRaw.length; i++) {
    const item = String(amlRaw[i][0]).trim();
    const mfrName = String(amlRaw[i][1]).trim();
    const mfrPN = String(amlRaw[i][2]).trim();
    const entry = `${mfrName}: ${mfrPN}`; 
    
    if (item) {
      if (!amlMap.has(item)) amlMap.set(item, []);
      if (!amlMap.get(item).includes(entry)) amlMap.get(item).push(entry);
    }
  }

  // --- CACHE 3: BOM DATA (Qty & Parent Lookup) ---
  const bomData = bomSheet.getDataRange().getValues();


  // === PROCESS SHEET 2: AFFECTED ITEMS (Standard Logic) ===
  const bomListSheet = ss.getSheetByName(CONFIG.BOM_LIST_TAB);
  if (bomListSheet) {
    const lastRow = bomListSheet.getLastRow();
    if (lastRow >= CONFIG.TABLE_START_ROW) {
      const numRows = lastRow - CONFIG.TABLE_START_ROW + 1;
      const range = bomListSheet.getRange(CONFIG.TABLE_START_ROW, 1, numRows, 7); // A:G
      const values = range.getValues();
      const backgrounds = range.getBackgrounds(); // Get existing colors
      const validationsToSet = []; // Store validations to apply later

      for (let i = 0; i < values.length; i++) {
        let parentItem = String(values[i][0]).trim();
        const childItem = String(values[i][1]).trim();
        const type = String(values[i][3]).toUpperCase(); // Col D (Change Type)
        let lookupSuccess = false; 

        if (!childItem) continue;

        // CHECK IF PARENT IS REQUIRED
        const isParentRequired = !["REV ROLL", "DESC CHANGE", "REV ONLY"].includes(type);

        // --- REC A: SMART PARENT POPULATION ---
        if (isParentRequired && (!parentItem || parentItem === "⚠ SELECT PARENT")) {
          const foundParents = findParentsInBom(childItem, bomData);
          
          if (foundParents.length > 1) {
            // MULTIPLE PARENTS FOUND: Create Dropdown
            values[i][0] = "⚠ SELECT PARENT"; 
            
            // Build Validation Rule
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(foundParents, true)
              .setAllowInvalid(false)
              .build();
            
            validationsToSet.push({
              row: i + CONFIG.TABLE_START_ROW,
              col: 1, // Col A
              rule: rule
            });
            parentItem = ""; // Clear for Qty lookup this round

          } else if (foundParents.length === 1) {
            // SINGLE PARENT FOUND: Fill directly
            parentItem = foundParents[0];
            values[i][0] = parentItem;
          }
        }

        // Desc & Rev
        if (itemMap.has(childItem)) {
          const info = itemMap.get(childItem);
          values[i][2] = info.desc; 
          values[i][4] = info.rev;  
        }

        // Qty Lookup (Only if Parent exists AND is required)
        if (isParentRequired && parentItem && parentItem !== "⚠ SELECT PARENT") {
           let foundQty = "";
           let parentFound = false;
           let parentLevel = -1;
           for (let r = 0; r < bomData.length; r++) {
             const row = bomData[r];
             if (row.length <= CONFIG.M_IDX.QTY) continue;
             const rowItem = String(row[CONFIG.M_IDX.ITEM]).trim();
             
             // BUG FIX: SKIP EMPTY ROWS
             if (!rowItem) continue;

             const rowLevelRaw = row[CONFIG.M_IDX.LEVEL];
             let rowLevel = isNaN(parseFloat(rowLevelRaw)) ? 0 : parseFloat(rowLevelRaw);
             
             if (!parentFound && rowItem === parentItem) {
               parentFound = true; parentLevel = rowLevel; continue; 
             }
             if (parentFound) {
               if (rowLevel <= parentLevel) break;
               if (rowLevel === parentLevel + 1 && rowItem === childItem) {
                 foundQty = row[CONFIG.M_IDX.QTY]; 
                 break; 
               }
             }
           }
           
           if (foundQty !== "") {
             values[i][6] = foundQty;
             lookupSuccess = true;
           }
        }

        // --- REC B: STATUS VALIDATION CHECK ---
        // 1. If "SELECT PARENT" is active -> RED
        if (values[i][0] === "⚠ SELECT PARENT") {
           for (let k = 0; k < 7; k++) backgrounds[i][k] = "#FFCCCC"; 
        } 
        // 2. If Parent is required, but lookup failed -> RED
        else if (isParentRequired && parentItem && !lookupSuccess && type !== "ADDED") {
           // (We ignore ADDED here because new items might not have Qty yet)
           for (let k = 0; k < 7; k++) backgrounds[i][k] = "#FFCCCC"; 
        } 
        // 3. If Parent is MISSING but required -> RED
        else if (isParentRequired && !parentItem && type !== "ADDED") {
           for (let k = 0; k < 7; k++) backgrounds[i][k] = "#FFCCCC"; 
        }
        // 4. Success -> White
        else {
           for (let k = 0; k < 7; k++) backgrounds[i][k] = "#FFFFFF"; 
        }
      }

      // WRITE BACK TO SHEET
      range.setValues(values);
      range.setBackgrounds(backgrounds);

      // APPLY VALIDATIONS (One by one as they differ per row)
      validationsToSet.forEach(v => {
        bomListSheet.getRange(v.row, v.col).setDataValidation(v.rule);
      });
    }
  }

  // === PROCESS SHEET 3: AVL CHANGES (ROW EXPANSION LOGIC + PRESERVE MANUAL DATA) ===
  const avlListSheet = ss.getSheetByName(CONFIG.AVL_LIST_TAB);
  if (avlListSheet) {
    const lastRow = avlListSheet.getLastRow();
    
    // 1. PRESERVE EXISTING MANUAL DATA
    const existingDataMap = new Map();
    const uniqueItems = new Set();

    if (lastRow >= CONFIG.TABLE_START_ROW) {
      const dataRange = avlListSheet.getRange(CONFIG.TABLE_START_ROW, 1, lastRow - CONFIG.TABLE_START_ROW + 1, 8);
      const currentValues = dataRange.getValues();

      currentValues.forEach(row => {
        const item = String(row[0]).trim();
        if (!item) return;

        uniqueItems.add(item);
        const vendorKey = String(row[3]).trim(); 
        const compositeKey = `${item}::${vendorKey}`;

        existingDataMap.set(compositeKey, {
          changeType: row[2],
          targetMfr: row[4],
          targetPn: row[5],
          disp: row[6],
          reason: row[7]
        });
      });
    }

    // 2. BUILD NEW TABLE
    if (uniqueItems.size > 0) {
      const newRows = [];
      
      uniqueItems.forEach(item => {
        const desc = itemMap.has(item) ? itemMap.get(item).desc : "";
        
        if (amlMap.has(item)) {
          const vendors = amlMap.get(item);
          vendors.forEach(vendorString => {
            const compositeKey = `${item}::${vendorString}`;
            const preserved = existingDataMap.get(compositeKey) || {};

            newRows.push([
              item, desc, preserved.changeType || "", vendorString, 
              preserved.targetMfr || "", preserved.targetPn || "", 
              preserved.disp || "", preserved.reason || ""
            ]);
          });
        } else {
          const vendorString = "No Existing AVL";
          const compositeKey = `${item}::${vendorString}`;
          const preserved = existingDataMap.get(compositeKey) || {};

          newRows.push([
            item, desc, preserved.changeType || "", vendorString, 
            preserved.targetMfr || "", preserved.targetPn || "", 
            preserved.disp || "", preserved.reason || ""
          ]);
        }
      });

      // 3. WRITE TO SHEET
      const numRowsToClear = avlListSheet.getMaxRows() - CONFIG.TABLE_START_ROW + 1;
      if (numRowsToClear > 0) {
        avlListSheet.getRange(CONFIG.TABLE_START_ROW, 1, numRowsToClear, 8).clearContent();
      }
      
      if (newRows.length > 0) {
        avlListSheet.getRange(CONFIG.TABLE_START_ROW, 1, newRows.length, 8).setValues(newRows);
        avlListSheet.getRange(CONFIG.TABLE_START_ROW, 2, newRows.length, 1).setBackground("#f3f3f3"); // Desc
        avlListSheet.getRange(CONFIG.TABLE_START_ROW, 4, newRows.length, 1).setBackground("#f3f3f3"); // Current AVL
      }
    }
  }

  ui.alert("Auto-Fill Complete. AVL Rows Expanded & Data Preserved.");
}

/**
 * FEATURE 2: SUBMIT LOG
 */
function submitComprehensiveECR() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(CONFIG.FORM_TAB_NAME);
  if (!formSheet) return ui.alert("Form tab missing.");

  const ecrNum = formSheet.getRange(CONFIG.FORM_CELLS.ECR_NUM).getValue();
  const ecoNum = formSheet.getRange(CONFIG.FORM_CELLS.ECO_NUM).getValue();
  if (!ecrNum || !ecoNum) return ui.alert("Missing ECR/ECO Number.");

  const rowsToPush = [];

  // BOM Sheet (Affected_Items)
  const bomSheet = ss.getSheetByName(CONFIG.BOM_LIST_TAB);
  if (bomSheet && bomSheet.getLastRow() >= CONFIG.TABLE_START_ROW) {
    const data = bomSheet.getRange(CONFIG.TABLE_START_ROW, 1, bomSheet.getLastRow() - CONFIG.TABLE_START_ROW + 1, 11).getValues();
    data.forEach(row => {
      const item = row[1]; // Col B
      // Accept more types now
      const validTypes = ["ADDED","REMOVED","MODIFIED","QTY CHANGE","REV ROLL","DESC CHANGE"];
      const type = String(row[3]).toUpperCase(); // Col D
      
      if (item && validTypes.includes(type)) {
        rowsToPush.push([
          ecrNum, ecoNum, item, row[0], type, // Common
          row[4], row[5], row[6], row[7],     // BOM Specific
          "", "",                             // AVL Specific
          row[8], row[9], row[10]             // NEW: DescChange, Disp, Reason
        ]);
      }
    });
  }

  // AVL Sheet (AVL_Changes)
  const avlSheet = ss.getSheetByName(CONFIG.AVL_LIST_TAB);
  if (avlSheet && avlSheet.getLastRow() >= CONFIG.TABLE_START_ROW) {
    const data = avlSheet.getRange(CONFIG.TABLE_START_ROW, 1, avlSheet.getLastRow() - CONFIG.TABLE_START_ROW + 1, 8).getValues();
    data.forEach(row => {
      const item = row[0]; 
      const type = String(row[2]).toUpperCase(); 
      const currentAvl = String(row[3]); 
      
      let logMfrName = row[4]; 
      let logMfrPN = row[5];   
      const disp = row[6];     
      const reason = row[7];   

      if (item && ["AVL_ADD", "AVL_REMOVE", "AVL_REPLACE"].includes(type)) {
        if (type === "AVL_REMOVE") {
          const parts = currentAvl.split(":");
          if (parts.length >= 1) logMfrName = parts[0].trim();
          if (parts.length >= 2) logMfrPN = parts[1].trim();
        } 
        
        rowsToPush.push([
          ecrNum, ecoNum, item, "N/A", type, 
          "", "", "", "",                    
          logMfrName, logMfrPN,              
          "", disp, reason                   
        ]);
      }
    });
  }

  if (rowsToPush.length === 0) return ui.alert("No valid rows to submit.");

  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_ID);
    const logSheet = masterSS.getSheetByName(CONFIG.ECR_LOG_TAB_NAME);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, rowsToPush.length, rowsToPush[0].length).setValues(rowsToPush);
    ui.alert(`Success! Submitted ${rowsToPush.length} items.`);
  } catch(e) {
    ui.alert("Error: " + e.message);
  }
}

/**
 * FEATURE 3: COMMIT TO MASTER (ADMIN ONLY)
 * Directly modifies the MASTER and AML sheets based on the ECR.
 * OPTIMIZED: Uses in-memory caching to reduce API calls (faster execution).
 * INCLUDES: Validation to ensure new items have AVL.
 * UPDATED: Handles REV ROLL, DESC CHANGE, QTY CHANGE separately.
 * UPDATED: Ignores Blank Rows in Master BOM
 */
function commitToMaster() {
  const ui = SpreadsheetApp.getUi();

  // --- PASSWORD PROTECTION ---
  const pwdResponse = ui.prompt('🔒 ADMIN AUTHENTICATION', 'Enter Admin Password to proceed:', ui.ButtonSet.OK_CANCEL);
  
  if (pwdResponse.getSelectedButton() !== ui.Button.OK) {
    return; // User clicked Cancel
  }
  
  const enteredPwd = pwdResponse.getResponseText();
  if (enteredPwd !== 'akfmsl21') {
    return ui.alert('❌ ACCESS DENIED', 'Incorrect Password.', ui.ButtonSet.OK);
  }
  // ---------------------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 0. PRE-FLIGHT VALIDATION: Check for "Ghost" Items ---
  let masterSS, masterSheet, amlSheet, itemsSheet;
  let masterData, amlData, itemsData;
  let masterAmlItems = new Set();

  try {
    masterSS = SpreadsheetApp.openById(CONFIG.MASTER_ID);
    masterSheet = masterSS.getSheetByName(CONFIG.BOM_TAB_NAME);
    amlSheet = masterSS.getSheetByName(CONFIG.AML_TAB_NAME);
    itemsSheet = masterSS.getSheetByName(CONFIG.ITEMS_TAB_NAME); 
    
    // FETCH DATA
    amlData = amlSheet.getDataRange().getValues();
    amlData.forEach(r => masterAmlItems.add(String(r[CONFIG.A_IDX.ITEM]).trim()));
    masterData = masterSheet.getDataRange().getValues();
    itemsData = itemsSheet.getDataRange().getValues();

  } catch(e) { return ui.alert('Connection Failed during Pre-Flight: ' + e.message); }

  const localBomSheet = ss.getSheetByName(CONFIG.BOM_LIST_TAB);
  const localAvlSheet = ss.getSheetByName(CONFIG.AVL_LIST_TAB);
  
  if (localBomSheet && localAvlSheet && localBomSheet.getLastRow() >= CONFIG.TABLE_START_ROW && localAvlSheet.getLastRow() >= CONFIG.TABLE_START_ROW) {
    const bRows = localBomSheet.getRange(CONFIG.TABLE_START_ROW, 1, localBomSheet.getLastRow() - CONFIG.TABLE_START_ROW + 1, 11).getValues();
    const aRows = localAvlSheet.getRange(CONFIG.TABLE_START_ROW, 1, localAvlSheet.getLastRow() - CONFIG.TABLE_START_ROW + 1, 8).getValues();

    const newBomItems = new Set(
      bRows.filter(r => String(r[3]).toUpperCase() === 'ADDED').map(r => String(r[1]).trim())
    );
    const newAvlItems = new Set(
      aRows.filter(r => String(r[2]).toUpperCase() === 'AVL_ADD').map(r => String(r[0]).trim())
    );

    const missingAvl = [...newBomItems].filter(x => 
      x !== "" && 
      !newAvlItems.has(x) && 
      !masterAmlItems.has(x)
    );

    if (missingAvl.length > 0) {
       const confirm = ui.alert('⚠️ POTENTIAL DATA GAP DETECTED', 
         `You are adding the following items to the BOM, but they have NO defined AVL in either the current changes or the Master AML:\n\n${missingAvl.join(', ')}\n\nThese items will show "No AML Found".\n\nAre you sure you want to proceed?`,
         ui.ButtonSet.YES_NO);
       if (confirm !== ui.Button.YES) return;
    }
  }
  // -------------------------------------------------------------

  const response = ui.alert('⚠ ADMIN ACTION: COMMIT TO MASTER', 
    'This will permanently modify the MASTER BOM and AML sheets based on the current ECR data.\n\nAre you sure you want to proceed?', 
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  // BUILD ITEMS MAP (for fast lookup)
  const itemsRowMap = new Map();
  for(let r=1; r<itemsData.length; r++) {
    const iName = String(itemsData[r][CONFIG.I_IDX.ITEM]).trim();
    if(iName) itemsRowMap.set(iName, r + 1); // r+1 is Sheet Row Index
  }

  // --- 1. PROCESS AVL CHANGES ---
  const avlListSheet = ss.getSheetByName(CONFIG.AVL_LIST_TAB);
  if (avlListSheet) {
    const changes = avlListSheet.getRange(CONFIG.TABLE_START_ROW, 1, avlListSheet.getLastRow() - CONFIG.TABLE_START_ROW + 1, 8).getValues();
    
    changes.forEach(row => {
      const item = String(row[0]).trim();
      const type = String(row[2]).toUpperCase();
      const currentAvl = String(row[3]); 
      const targetMfr = String(row[4]).trim();
      const targetMpn = String(row[5]).trim();
      
      if (!item) return;

      const findAmlIndex = (searchMfr) => {
        for (let i = 0; i < amlData.length; i++) {
          if (String(amlData[i][CONFIG.A_IDX.ITEM]).trim() === item && 
              String(amlData[i][CONFIG.A_IDX.MFR]).trim() === searchMfr) {
            return i; 
          }
        }
        return -1;
      };

      if (type === "AVL_ADD") {
         let exists = false;
         for (let i = 0; i < amlData.length; i++) {
            if (String(amlData[i][CONFIG.A_IDX.ITEM]).trim() === item && 
                String(amlData[i][CONFIG.A_IDX.MFR]).trim() === targetMfr &&
                String(amlData[i][CONFIG.A_IDX.MPN]).trim() === targetMpn) {
              exists = true; break;
            }
         }
         
         if (!exists && targetMfr && targetMpn) {
           amlSheet.appendRow([item, targetMfr, targetMpn]);
           amlData.push([item, targetMfr, targetMpn]);
         }
      } 
      else if (type === "AVL_REMOVE" || type === "AVL_REPLACE") {
        const parts = currentAvl.split(":");
        if (parts.length < 2) return; 
        const currentMfr = parts[0].trim();
        const arrayIndex = findAmlIndex(currentMfr); 

        if (arrayIndex > -1) {
          const sheetRowIndex = arrayIndex + 1; 

          if (type === "AVL_REMOVE") {
            amlSheet.deleteRow(sheetRowIndex);
            amlData.splice(arrayIndex, 1);
          } else if (type === "AVL_REPLACE") {
             amlSheet.getRange(sheetRowIndex, CONFIG.A_IDX.MFR + 1).setValue(targetMfr);
             amlSheet.getRange(sheetRowIndex, CONFIG.A_IDX.MPN + 1).setValue(targetMpn);
             amlData[arrayIndex][CONFIG.A_IDX.MFR] = targetMfr;
             amlData[arrayIndex][CONFIG.A_IDX.MPN] = targetMpn;
          }
        }
      }
    });
  }

  // --- 2. PROCESS BOM CHANGES ---
  const bomListSheet = ss.getSheetByName(CONFIG.BOM_LIST_TAB);
  if (bomListSheet) {
    const changes = bomListSheet.getRange(CONFIG.TABLE_START_ROW, 1, bomListSheet.getLastRow() - CONFIG.TABLE_START_ROW + 1, 11).getValues();
    
    changes.forEach(row => {
      const parent = String(row[0]).trim();
      const child = String(row[1]).trim();
      const type = String(row[3]).toUpperCase();
      const newRev = String(row[5]).trim();
      const newQty = row[7];
      const newDesc = String(row[8]).trim();
      
      if (!child) return;

      // SEARCH in IN-MEMORY masterData if Parent exists
      let parentFound = false;
      let parentLevel = -1;
      let targetArrayIndex = -1; 
      let insertAfterArrayIndex = -1;

      // Only search BOM if type is relevant (i.e., not just a pure Rev Roll)
      // OR if user provided a parent for a Modified/Qty Change
      const isBomUpdateNeeded = ["ADDED", "REMOVED", "MODIFIED", "QTY CHANGE"].includes(type);

      if (isBomUpdateNeeded && parent) {
        for (let i = 0; i < masterData.length; i++) {
          const rItem = String(masterData[i][CONFIG.M_IDX.ITEM]).trim();
          
          // BUG FIX: SKIP EMPTY ROWS
          if (!rItem) continue;

          const rLevelRaw = masterData[i][CONFIG.M_IDX.LEVEL];
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
      
      // CASE A: REV ROLL or DESC CHANGE (Update ITEMS only, Skip BOM)
      if (type === "REV ROLL" || type === "DESC CHANGE") {
        if (itemsRowMap.has(child)) {
          const itemRow = itemsRowMap.get(child);
          if (newRev !== "") itemsSheet.getRange(itemRow, CONFIG.I_IDX.REV + 1).setValue(newRev);
          if (newDesc !== "") itemsSheet.getRange(itemRow, CONFIG.I_IDX.DESC + 1).setValue(newDesc);
        }
      }

      // CASE B: QTY CHANGE (Update BOM Qty only)
      else if (type === "QTY CHANGE") {
        if (targetArrayIndex > -1) {
          masterSheet.getRange(targetArrayIndex + 1, CONFIG.M_IDX.QTY + 1).setValue(newQty);
          masterData[targetArrayIndex][CONFIG.M_IDX.QTY] = newQty;
        }
      }

      // CASE C: MODIFIED (Update EVERYTHING)
      else if (type === "MODIFIED") {
        // 1. Update BOM Qty if valid parent link found
        if (targetArrayIndex > -1) {
          masterSheet.getRange(targetArrayIndex + 1, CONFIG.M_IDX.QTY + 1).setValue(newQty);
          masterData[targetArrayIndex][CONFIG.M_IDX.QTY] = newQty;
        }
        // 2. Update ITEMS Definition (Global)
        if (itemsRowMap.has(child)) {
          const itemRow = itemsRowMap.get(child);
          if (newRev !== "") itemsSheet.getRange(itemRow, CONFIG.I_IDX.REV + 1).setValue(newRev);
          if (newDesc !== "") itemsSheet.getRange(itemRow, CONFIG.I_IDX.DESC + 1).setValue(newDesc);
        }
      }

      // CASE D: REMOVED (BOM Only)
      else if (type === "REMOVED") {
        if (targetArrayIndex > -1) {
          masterSheet.deleteRow(targetArrayIndex + 1);
          masterData.splice(targetArrayIndex, 1);
        }
      }

      // CASE E: ADDED (BOM Only, plus Item Creation if needed)
      else if (type === "ADDED" && parentFound) {
        // Ensure Item Definition Exists
        let desc = "";
        if (itemsRowMap.has(child)) {
           const rIndex = itemsRowMap.get(child) - 1; 
           if (rIndex < itemsData.length) {
              desc = itemsData[rIndex][CONFIG.I_IDX.DESC];
           } else {
              desc = newDesc !== "" ? newDesc : String(row[2]).trim();
           }
        } else {
           desc = newDesc !== "" ? newDesc : String(row[2]).trim();
           const rev = newRev !== "" ? newRev : (String(row[4]).trim() || "A");
           itemsSheet.appendRow([child, desc, rev]);
           itemsRowMap.set(child, itemsSheet.getLastRow()); 
        }

        if (targetArrayIndex > -1) {
          masterSheet.getRange(targetArrayIndex + 1, CONFIG.M_IDX.QTY + 1).setValue(newQty);
          masterData[targetArrayIndex][CONFIG.M_IDX.QTY] = newQty;
        } else {
          const sheetInsertRow = insertAfterArrayIndex + 2; 
          const itemRef = `C${sheetInsertRow}`;
          const f_Desc = `=IF(ISBLANK(${itemRef}), "", VLOOKUP(${itemRef}, ITEMS!A:C, 2, FALSE))`;
          const f_Rev = `=IF(ISBLANK(${itemRef}), "", VLOOKUP(${itemRef}, ITEMS!A:C, 3, FALSE))`;
          const f_Life = `=IF(ISBLANK(${itemRef}), "", VLOOKUP(${itemRef}, ITEMS!A:D, 4, FALSE))`;
          const f_Mfr = `=IFNA(IF(ISBLANK(${itemRef}), "", FILTER(AML!B:C, AML!A:A = ${itemRef})), "No AML Found")`;

          const newRowData = [];
          for(let c=0; c<CONFIG.M_IDX.LEVEL; c++) newRowData.push(""); 
          
          newRowData.push(parentLevel + 1); 
          newRowData.push(child);           
          newRowData.push(f_Desc);          
          newRowData.push(f_Rev);           
          newRowData.push(f_Life);          
          newRowData.push(newQty);          

          for(let k=0; k<5; k++) newRowData.push(""); 

          newRowData.push(f_Mfr);           
          newRowData.push("");              

          masterSheet.insertRowAfter(insertAfterArrayIndex + 1);
          masterSheet.getRange(sheetInsertRow, 1, 1, newRowData.length).setValues([newRowData]);
          masterData.splice(insertAfterArrayIndex + 1, 0, newRowData);
        }
      }
    });
  }

  ui.alert("Commit Complete. Master Data has been updated.");
}

/**
 * HELPER: FIND PARENT ASSEMBLIES IN BOM
 * Updated to accept passed data array
 * UPDATED: Ignores Blank Rows in Master BOM
 */
function findParentsInBom(childItem, bomData) {
  const parents = new Set();
  const childItemStr = String(childItem).trim();
  
  for (let i = 0; i < bomData.length; i++) {
    const row = bomData[i];
    if (row.length <= CONFIG.M_IDX.ITEM) continue;
    
    const rowItem = String(row[CONFIG.M_IDX.ITEM]).trim();
    
    if (rowItem === childItemStr) {
      const rowLevelRaw = row[CONFIG.M_IDX.LEVEL];
      const childLevel = isNaN(parseFloat(rowLevelRaw)) ? 0 : parseFloat(rowLevelRaw);
      
      for (let j = i - 1; j >= 0; j--) {
        const parentRow = bomData[j];
        
        // BUG FIX: IGNORE BLANK ROWS
        const pItem = String(parentRow[CONFIG.M_IDX.ITEM]).trim();
        if (!pItem) continue;

        const parentLevelRaw = parentRow[CONFIG.M_IDX.LEVEL];
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

/**
 * ----------------------------------------------------
 * DEBUGGER TOOL
 * ----------------------------------------------------
 * Use this to verify why a link isn't found.
 * UPDATED: Ignores Blank Rows
 */
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
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_ID);
    const bomSheet = masterSS.getSheetByName(CONFIG.BOM_TAB_NAME);
    const bomData = bomSheet.getDataRange().getValues();
    
    let parentFound = false;
    let parentLevel = -1;
    let parentRowIndex = -1;
    
    for (let i = 0; i < bomData.length; i++) {
      const row = bomData[i];
      const rItem = String(row[CONFIG.M_IDX.ITEM]).trim();
      
      // BUG FIX: SKIP EMPTY ROWS
      if (!rItem) continue;

      const rLevel = isNaN(parseFloat(row[CONFIG.M_IDX.LEVEL])) ? 0 : parseFloat(row[CONFIG.M_IDX.LEVEL]);
      
      // LOOK FOR PARENT
      if (!parentFound && rItem === pName) {
        parentFound = true;
        parentLevel = rLevel;
        parentRowIndex = i + 1;
        Logger.log(`✅ PARENT FOUND at Row ${parentRowIndex}. Level: ${parentLevel}`);
        continue;
      }
      
      // LOOK FOR CHILD UNDER PARENT
      if (parentFound) {
        if (rLevel <= parentLevel) {
          Logger.log(`⛔ End of Parent Block at Row ${i + 1}. Item: ${rItem} (Level ${rLevel})`);
          break;
        }
        
        // Log every child to see what script sees
        if (rLevel === parentLevel + 1) {
           const match = (rItem === cName) ? "MATCH!" : "No Match";
           Logger.log(`   -> Child Candidate Row ${i+1}: [${rItem}] vs [${cName}] ... ${match}`);
        }
      }
    }
    
    if (!parentFound) Logger.log("❌ PARENT NOT FOUND in Master BOM.");
    
  } catch(e) {
    Logger.log("ERROR: " + e.message);
  }
  
  ui.alert("Debug Complete. Check 'View > Executions' or 'View > Logs' to see the report.");
}