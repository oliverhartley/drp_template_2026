/**
 * ****************************************
 * Google Apps Script - One-Off Deck Inspector
 * File: one_off_inspect_deck_2026.js
 * Description: Opens the generated Atos deck and inspects formulas.
 * ****************************************
 */

function inspectAtosDeck() {
  const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID);
  const files = folder.getFilesByName("atos.net - Partner Dashboard 2026");
  if (!files.hasNext()) {
    Logger.log("Error: Atos deck not found.");
    return;
  }
  
  const ss = SpreadsheetApp.open(files.next());
  const sheet = ss.getSheetByName("Tier Dashboard");
  if (!sheet) {
    Logger.log("Error: Tier Dashboard sheet not found.");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 2, lastRow - 1, 2); // Read Columns B (Product) and C (Tier 1)
  const values = range.getValues();
  const formulas = range.getFormulas();
  
  for (let i = 0; i < values.length; i++) {
    Logger.log(`Row ${i + 2}: Product="${values[i][0]}", Formula="${formulas[i][1]}"`);
  }
}
