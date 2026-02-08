// ============================================================================
// MENU.gs — Google Sheets Custom Menu
// ============================================================================
// Builds the "BOM Tools" menu with categorised submenus for better
// discoverability in the Google Sheets UI.
// ============================================================================

/**
 * Creates the custom "BOM Tools" menu when the spreadsheet is opened.
 * Organises tools into logical submenus for easier navigation.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('BOM Tools')

    // --- Comparison ---
    .addSubMenu(
      ui.createMenu('Comparison')
        .addItem('Generate Detailed Comparison (ECO)', 'runDetailedComparison')
        .addItem('Compare Master vs. PDM BOM', 'runCompareWithExternalBOM')
    )

    // --- PDM Integration ---
    .addSubMenu(
      ui.createMenu('PDM Integration')
        .addItem('Import Children from PDM (Graft)', 'runImportPdmChildren')
    )

    // --- Fabrication ---
    .addSubMenu(
      ui.createMenu('Fabrication')
        .addItem('Generate Fabricator BOMs', 'runGenerateFabricatorBOMs')
    )

    // --- Data Integrity ---
    .addSubMenu(
      ui.createMenu('Data Integrity')
        .addItem('Reconcile Master Data (Full Sync)', 'runReconcileMasterData')
        .addItem('Validate BOM (9-Check Audit)', 'runValidateBOM')
        .addItem('Lifecycle Deviation (ECR Required)', 'runLifecycleDeviation')
        .addItem('Protect Master Sheets', 'runProtectMasterSheets')
        .addItem('Install Change Watchdog', 'installChangeTrigger_')
    )

    // --- Audit & Quality ---
    .addSubMenu(
      ui.createMenu('Audit & Quality')
        .addItem('Audit BOM Lifecycle Status', 'runAuditBOMLifecycle')
        .addItem('Audit Lifecycle States (State Machine)', 'runAuditLifecycleStates')
        .addItem('Audit BOM Structural Integrity', 'runAuditDuplicatePartNumbers')
        .addItem("List 'MAKE' Items with 'REF' Children", 'runScreenMakeItemsWithRef')
    )

    // --- Analysis & Reports ---
    .addSubMenu(
      ui.createMenu('Analysis & Reports')
        .addItem('Where-Used Analysis (Full Chain)', 'runWhereUsedAnalysis')
        .addItem('Generate BOM Dashboard', 'runGenerateDashboard')
        .addItem('Generate Master Lists from BOM', 'runGenerateMasterLists')
    )

    .addSeparator()

    // --- Utilities (top-level for quick access) ---
    .addItem('Prepare Rows for AML', 'runPrepareAMLRows')
    .addItem('Set BOM Effectivity Dates', 'runSetEffectivityDates')
    .addItem('Finalize and Release New BOM', 'runReleaseNewBOM')

    .addToUi();
}
