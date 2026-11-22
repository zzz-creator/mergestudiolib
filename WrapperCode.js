/**
 * ============================================================================
 * MERGE STUDIO — BOUND SPREADSHEET WRAPPER (WrapperCode.gs)
 * ============================================================================
 * Lightweight container-bound script for any Google Sheet.
 * Connects the custom "Merge Studio" menu and SidebarUI.html to MergeStudioLib.
 *
 * ----------------------------------------------------------------------------
 * SETUP & LINKING INSTRUCTIONS (BOUND SPREADSHEET)
 * ----------------------------------------------------------------------------
 * 1. OPEN BOUND SCRIPT EDITOR:
 *    - Open your Google Sheet.
 *    - Click "Extensions" > "Apps Script".
 *
 * 2. ADD FILES:
 *    - Replace Code.gs with this file (WrapperCode.gs).
 *    - Click "+" next to "Files" > "HTML", name it exactly: SidebarUI
 *      (Apps Script automatically stores it as SidebarUI.html), and paste the
 *      contents of SidebarUI.html.
 *
 * 3. LINK THE MERGE STUDIO LIBRARY:
 *    - In the left sidebar next to "Libraries", click "+".
 *    - Paste the **Script ID** of your MergeStudioLib project and click "Look up".
 *    - Select the latest numbered **Version** (do NOT use HEAD if sharing with
 *      teammates who only have Viewer access on the library).
 *    - Set **Identifier** to exactly: MergeStudioLib
 *    - Click "Add".
 *
 * 4. FIRST-RUN AUTHORIZATION:
 *    - Reload the Google Sheet. A "Merge Studio" menu will appear after ~2 seconds.
 *    - Click "Merge Studio" > "Open Merge Studio".
 *    - On first run, Google prompts for authorization. For unverified internal
 *      or personal scripts, click "Advanced" > "Go to <Project Name> (unsafe)"
 *      > "Allow". No public Workspace Marketplace verification is required.
 * ============================================================================
 */

/**
 * Adds the custom "Merge Studio" menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Merge Studio')
    .addItem('Open Merge Studio', 'openMergeSidebar')
    .addItem('Run Pending Rows', 'runPendingRowsFromMenu')
    .addToUi();
}

/**
 * Opens the AutoCrat-style Merge Studio sidebar.
 */
function openMergeSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SidebarUI')
    .setTitle('Merge Studio')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Menu shortcut to run pending rows using saved configuration.
 */
function runPendingRowsFromMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = MergeStudioLib.runMerge(SpreadsheetApp.getActiveSpreadsheet());
    ui.alert('Merge Studio', result.message, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Merge Studio Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

/* ============================================================================
 * THIN WRAPPERS FOR google.script.run (called by SidebarUI.html)
 * ============================================================================ */

function apiGetSpreadsheetContext() {
  return MergeStudioLib.getSpreadsheetContext(SpreadsheetApp.getActiveSpreadsheet());
}

function apiGetSheetHeaders(sheetName) {
  return MergeStudioLib.getSheetHeaders(SpreadsheetApp.getActiveSpreadsheet(), sheetName);
}

function apiGetTemplateTags(docIdOrUrl) {
  return MergeStudioLib.getTemplateTags(docIdOrUrl);
}

function apiSaveConfig(config) {
  return MergeStudioLib.saveConfig(SpreadsheetApp.getActiveSpreadsheet(), config);
}

function apiRunMerge(options) {
  return MergeStudioLib.runMerge(SpreadsheetApp.getActiveSpreadsheet(), options);
}

function apiConfigureTrigger(mode) {
  return MergeStudioLib.configureTrigger(SpreadsheetApp.getActiveSpreadsheet(), mode);
}

function apiClearStatuses() {
  return MergeStudioLib.clearStatuses(SpreadsheetApp.getActiveSpreadsheet());
}

/* ============================================================================
 * TRIGGER ENTRY POINTS (Invoked by ScriptApp triggers)
 * ============================================================================ */

function runMergeFromTrigger() {
  MergeStudioLib.runMerge(SpreadsheetApp.getActiveSpreadsheet());
}

function onFormSubmitMerge(e) {
  MergeStudioLib.runMerge(SpreadsheetApp.getActiveSpreadsheet());
}