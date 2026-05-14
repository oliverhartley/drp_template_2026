/**
 * ****************************************
 * Google Apps Script - Partner Deck Share 2026
 * File: Partner_Share_Decks_2026_v3.js
 * Description: Shares individual partner decks with corresponding PDMs (as Editors)
 * and Partner Emails (as Viewers) in batch.
 * ****************************************
 */

const BATCH_SHARE_TIME_LIMIT_MS_2026 = 1200000; // 20 minutes

function runBatchShareDecks2026() {
  const startTime = new Date().getTime();
  const currentBatchId = getShareBatchId2026();
  Logger.log(`>>> STARTING 2026 BATCH SHARE PROCESS [Batch ID: ${currentBatchId}] <<<`);

  const ss = SpreadsheetApp.openById(PARTNER_DB_SS_ID);
  const dbSheet = ss.getSheetByName(SHEET_NAME_2026);

  if (!dbSheet) {
    Logger.log(`ERROR: Sheet ${SHEET_NAME_2026} not found.`);
    return;
  }

  const dataRange = dbSheet.getDataRange();
  const data = dataRange.getValues();
  if (data.length < 2) {
    Logger.log(`No data found in ${SHEET_NAME_2026}.`);
    return;
  }

  const headers = data[0].map(h => String(h).trim().toLowerCase());

  const colPartnerName = 0; // Partner Name is A (0)
  let colToEmail = -1;
  let colPdm = -1;
  let colSpreadsheetId = -1;
  let colShareStatus = -1;
  let colDomain = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h === "partner_emails") colToEmail = c;
    if (h === "pdm") colPdm = c;
    if (h === "spreadsheet_id") colSpreadsheetId = c;
    if (h === "share_status") colShareStatus = c;
    if (h === "domain") colDomain = c;
  }

  // Fallbacks if not found by exact header name
  if (colToEmail === -1) colToEmail = 8; // Default Col I
  if (colPdm === -1) colPdm = 3; // Default Col D
  if (colSpreadsheetId === -1) colSpreadsheetId = 10; // Default Col K
  if (colDomain === -1) colDomain = 7; // Default Col H

  // Initialize share_status header if missing
  if (colShareStatus === -1) {
    const lastCol = dbSheet.getLastColumn();
    dbSheet.getRange(1, lastCol + 1).setValue("share_status").setBackground("#d9d9d9").setFontWeight("bold");
    colShareStatus = lastCol; // 0-based index of the newly added column
  }

  // Group by Domain to ensure we process each unique partner deck exactly once
  const partnerMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const domain = String(row[colDomain]).trim().toLowerCase();
    if (!domain) continue;

    const partnerName = String(row[colPartnerName]).trim();

    if (!partnerMap.has(domain)) {
      partnerMap.set(domain, {
        name: partnerName,
        rows: [],
        partnerEmails: new Set(),
        pdmEmails: new Set(),
        status: "",
        spreadsheetId: ""
      });
    }

    const pData = partnerMap.get(domain);
    pData.rows.push(i); // Store 0-indexed row number for updating status

    if (!pData.name) pData.name = partnerName;

    if (!pData.status && row[colShareStatus]) pData.status = String(row[colShareStatus]).trim();
    if (!pData.spreadsheetId && row[colSpreadsheetId]) pData.spreadsheetId = String(row[colSpreadsheetId]).trim();

    // Collect all unique emails
    let to = String(row[colToEmail] || "").trim();
    let pdm = String(row[colPdm] || "").trim();

    if (to) to.split(',').forEach(e => { if (e.trim()) pData.partnerEmails.add(e.trim().toLowerCase()); });
    if (pdm) pdm.split(',').forEach(e => { if (e.trim()) pData.pdmEmails.add(e.trim().toLowerCase()); });
  }

  let processedCount = 0;
  let skippedCount = 0;

  for (const [domain, pData] of partnerMap.entries()) {
    if ((new Date().getTime() - startTime) > BATCH_SHARE_TIME_LIMIT_MS_2026) {
      Logger.log("WARNING: Time limit approaching. Stopping to allow safe resume on next trigger.");
      break;
    }

    if (pData.partnerEmails.size === 0 && pData.pdmEmails.size === 0) {
      Logger.log(`Skipping ${pData.name} (${domain}) - No emails found.`);
      continue;
    }

    if (pData.status === currentBatchId) {
      Logger.log(`Skipping ${pData.name} (${domain}) - Already shared for this batch.`);
      skippedCount++;
      continue;
    }

    if (!pData.spreadsheetId) {
      Logger.log(`Skipping ${pData.name} (${domain}) - No Spreadsheet ID found. Please generate their deck first.`);
      continue;
    }

    Logger.log(`Sharing Deck for Partner: ${pData.name} (${domain})...`);

    try {
      const file = DriveApp.getFileById(pData.spreadsheetId);

      // Add PDMs as Editors
      pData.pdmEmails.forEach(email => {
        try {
          file.addEditor(email);
          Logger.log(`  Added PDM Editor: ${email}`);
        } catch (shareErr) {
          Logger.log(`  ERROR adding PDM Editor ${email}: ${shareErr.toString()}`);
        }
      });

      // Add Partner Emails as Viewers
      pData.partnerEmails.forEach(email => {
        try {
          file.addViewer(email);
          Logger.log(`  Added Partner Viewer: ${email}`);
        } catch (shareErr) {
          Logger.log(`  ERROR adding Partner Viewer ${email}: ${shareErr.toString()}`);
        }
      });

      // Update Status for all rows belonging to this partner
      pData.rows.forEach(r => {
        dbSheet.getRange(r + 1, colShareStatus + 1).setValue(currentBatchId);
      });
      SpreadsheetApp.flush();
      processedCount++;

      Utilities.sleep(1000); // Respect Drive API rate limits
    } catch (e) {
      Logger.log(`  ERROR sharing deck for ${pData.name}: ${e.toString()}`);
    }
  }

  Logger.log(`>>> 2026 BATCH SHARE COMPLETE. Shared: ${processedCount}, Skipped: ${skippedCount} <<<`);
}

function getShareBatchId2026() {
  const now = new Date();
  const shiftedDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const year = shiftedDate.getFullYear();
  const onejan = new Date(year, 0, 1);
  const week = Math.ceil((((shiftedDate.getTime() - onejan.getTime()) / 86400000) + onejan.getDay() + 1) / 7);
  return `SHARED_2026_${year}_${week}`;
}
