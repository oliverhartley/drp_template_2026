/**
 * ****************************************
 * Google Apps Script - One-Off Sheet Cleanup
 * File: one_off_cleanup_sheets_2026.js
 * Description: Lists all sheets and allows deleting non-_w2 sheets.
 * ****************************************
 */

function listAllSheets() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheets = ss.getSheets();
  
  Logger.log("Current sheets in spreadsheet:");
  sheets.forEach(sheet => {
    const name = sheet.getName();
    const isNew = name.endsWith("_v3");
    Logger.log(`- "${name}" [Keep suggested: ${isNew ? "YES" : "NO"}]`);
  });
}

function deleteNonW2Sheets() {
  const ss = SpreadsheetApp.openById(DESTINATION_SS_ID);
  const sheets = ss.getSheets();
  
  Logger.log("Starting cleanup of non-_w2 sheets...");
  
  // We need to keep at least one sheet, so we should find a _w2 sheet first
  let keptSheet = null;
  for (const sheet of sheets) {
    if (sheet.getName().endsWith("_w2")) {
      keptSheet = sheet;
      break;
    }
  }
  
  if (!keptSheet) {
    Logger.log("Error: No sheet ending with _w2 found. Aborting cleanup to avoid deleting everything!");
    return;
  }
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.endsWith("_w2")) {
      try {
        Logger.log(`Deleting sheet: ${name}`);
        ss.deleteSheet(sheet);
      } catch (e) {
        Logger.log(`  Failed to delete ${name}: ${e.toString()}`);
      }
    }
  });
  
  Logger.log("Cleanup complete.");
}
