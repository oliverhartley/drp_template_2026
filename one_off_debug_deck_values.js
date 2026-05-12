/**
 * ****************************************
 * Google Apps Script - One-Off Deck Values Debugger
 * File: one_off_debug_deck_values.js
 * Description: Inspects specific cell values and formulas in the Atos deck.
 * ****************************************
 */

function debugDeckValues() {
  const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID);
  const files = folder.getFilesByName("atos.net - Partner Dashboard 2026");
  if (!files.hasNext()) {
    Logger.log("Error: Atos deck not found.");
    return;
  }
  
  const ss = SpreadsheetApp.open(files.next());
  
  // 1. Check formula in Tier Dashboard
  const dashSheet = ss.getSheetByName("Tier Dashboard");
  if (dashSheet) {
    const cell = dashSheet.getRange("C22");
    Logger.log(`Tier Dashboard C22 Value: "${cell.getValue()}"`);
    Logger.log(`Tier Dashboard C22 Formula: "${cell.getFormula()}"`);
  } else {
    Logger.log("Error: Tier Dashboard sheet not found.");
  }
  
  // 2. Check values in Profile Deep Dive
  const diveSheet = ss.getSheetByName("Profile Deep Dive");
  if (diveSheet) {
    const range = diveSheet.getRange("AE7:AE30");
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
      Logger.log(`Row ${i + 7}: "${values[i][0]}"`);
    }
  } else {
    Logger.log("Error: Profile Deep Dive sheet not found.");
  }
}
