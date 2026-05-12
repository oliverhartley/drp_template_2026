/**
 * ****************************************
 * Google Apps Script - One-Off Force Security Values
 * File: one_off_force_atos_security_values.js
 * Description: Counts values in Column AE of Deep Dive and forces summary into Tier Dashboard.
 * ****************************************
 */

function forceAtosSecurityValues() {
  const folder = DriveApp.getFolderById(PARTNER_FOLDER_ID);
  const files = folder.getFilesByName("atos.net - Partner Dashboard 2026");
  if (!files.hasNext()) {
    Logger.log("Error: Atos deck not found.");
    return;
  }
  
  const ss = SpreadsheetApp.open(files.next());
  const diveSheet = ss.getSheetByName("Profile Deep Dive");
  
  if (diveSheet) {
    const lastRow = diveSheet.getLastRow();
    const range = diveSheet.getRange("AE7:AE" + lastRow);
    const values = range.getValues();
    
    let count1 = 0, count2 = 0, count3 = 0, count4 = 0, countDash = 0;
    
    for (let i = 0; i < values.length; i++) {
      const val = String(values[i][0]).trim();
      if (val === "Tier 1") count1++;
      else if (val === "Tier 2") count2++;
      else if (val === "Tier 3") count3++;
      else if (val === "Tier 4") count4++;
      else if (val === "-") countDash++;
    }
    
    Logger.log(`Script Counts for Column AE:`);
    Logger.log(`Tier 1: ${count1}`);
    Logger.log(`Tier 2: ${count2}`);
    Logger.log(`Tier 3: ${count3}`);
    Logger.log(`Tier 4: ${count4}`);
    Logger.log(`-: ${countDash}`);
  } else {
    Logger.log("Error: Profile Deep Dive sheet not found.");
  }
  
  const dashSheet = ss.getSheetByName("Tier Dashboard");
  if (dashSheet) {
    Logger.log("Forcing values into Tier Dashboard row 22 (Cloud Security)...");
    
    // User provided values
    // Tier 1: 1 (C22)
    // Tier 2: 2 (D22)
    // Tier 3: 10 (E22)
    // Tier 4: 26 (F22)
    
    dashSheet.getRange("C22").setValue(1);
    dashSheet.getRange("D22").setValue(2);
    dashSheet.getRange("E22").setValue(10);
    dashSheet.getRange("F22").setValue(26);
    
    Logger.log("Values forced successfully.");
  } else {
    Logger.log("Error: Tier Dashboard sheet not found.");
  }
}
