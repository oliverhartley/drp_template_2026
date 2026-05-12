/**
 * ****************************************
 * Google Apps Script - Partner DB Generator 2026
 * File: Partner_DB_Generator_2026.js
 * Description: Reads from the new partner source spreadsheet,
 * matches against BQ using domain, and populates the DB sheet.
 * ****************************************
 */

function getNewSourceSpreadsheetData2026() {
  const ss = SpreadsheetApp.openById(NEW_SOURCE_SS_ID);
  const sheet = ss.getSheetByName(NEW_SHEET_NAME_PARTNER);
  if (!sheet) throw new Error(`Sheet "${NEW_SHEET_NAME_PARTNER}" not found in Source Spreadsheet.`);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ""; // Assuming row 1 is headers
  
  const range = sheet.getRange(2, 1, lastRow - 1, 5); // Reading Columns A to E
  const values = range.getValues();
  
  let structList = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    let domain = String(row[0] || "").trim();
    let partnerEmails = String(row[1] || "").trim();
    let pdm = String(row[2] || "").trim();
    let ccTechLead = String(row[3] || "").trim();
    let partnerType = String(row[4] || "").trim();

    if (domain) {
      const internalId = `ROW_${i + 2}`; // Generate Internal ID
      structList.push(`STRUCT('${internalId}' AS internal_id, '${domain.replace(/'/g, "\\'")}' AS domain, '${partnerEmails.replace(/'/g, "\\'")}' AS partner_emails, '${pdm.replace(/'/g, "\\'")}' AS pdm, '${ccTechLead.replace(/'/g, "\\'")}' AS cc_tech_lead, '${partnerType.replace(/'/g, "\\'")}' AS partner_type)`);
    }
  }
  return structList.join(',\n');
}

function runPartnerDbGeneration2026() {
  try {
    Logger.log("Generating virtual table from new source spreadsheet...");
    const VIRTUAL_TABLE_DATA = getNewSourceSpreadsheetData2026();
    if (!VIRTUAL_TABLE_DATA) { Logger.log("Error: No data found in Source Sheet."); return; }

    const SQL_QUERY = `
      WITH Spreadsheet_Data AS ( SELECT * FROM UNNEST([ ${VIRTUAL_TABLE_DATA} ]) ),
      
      -- Flatten BigQuery Data to get domains and corresponding partner details
      BQ_Flattened AS (
        SELECT DISTINCT
           t1.partner_id,
           t1.partner_name,
           t1.profile_details.residing_country,
           REPLACE(LOWER(bq_domain), '@', '') as bq_domain_clean
        FROM \`concord-prod.service_partnercoe.drp_partner_master\` AS t1
        LEFT JOIN UNNEST(t1.partner_details.email_domain) AS bq_domain
        WHERE 'LATAM' IN UNNEST(t1.profile_details.delivery_region)
        AND t1.profile_details.residing_country IN ('Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela')
      ),
      
      -- Match Spreadsheet Data with BQ Data on Domain
      MatchedData AS (
          SELECT
              sheet.internal_id,
              sheet.domain,
              sheet.partner_emails,
              sheet.pdm,
              sheet.cc_tech_lead,
              sheet.partner_type,
              bq.partner_id,
              bq.partner_name as bq_partner_name,
              bq.residing_country,
              CASE
                WHEN bq.residing_country = 'Brazil' THEN 'Brazil'
                WHEN bq.residing_country = 'Mexico' THEN 'Mexico'
                WHEN bq.residing_country IN ('Argentina', 'Bolivia', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominican Republic', 'Ecuador', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Paraguay', 'Peru', 'Uruguay', 'Venezuela') THEN 'MCO'
                ELSE 'Other'
              END AS derived_sub_region
          FROM Spreadsheet_Data AS sheet
          LEFT JOIN BQ_Flattened AS bq
            ON sheet.domain = bq.bq_domain_clean
      ),
      
      -- Aggregate to Partner and Country Level
      PartnerAggregation AS (
          SELECT
              md.internal_id,
              MAX(md.domain) as Domain,
              MAX(md.partner_emails) as Partner_Emails,
              MAX(md.pdm) as PDM,
              MAX(md.cc_tech_lead) as CC_Tech_Lead,
              MAX(md.partner_type) as Type_of_Partner,
              MAX(md.partner_id) as Partner_ID,
              MAX(md.bq_partner_name) as Partner_Name,
              md.residing_country as Country,
              MAX(md.derived_sub_region) as Sub_Region
          FROM MatchedData md
          GROUP BY md.internal_id, md.residing_country
      )
      SELECT 
          pa.Partner_Name,
          pa.Country,
          pa.Sub_Region,
          pa.PDM,
          pa.Type_of_Partner,
          pa.internal_id AS Internal_ID,
          pa.Partner_ID,
          pa.Domain,
          pa.Partner_Emails,
          pa.CC_Tech_Lead
      FROM PartnerAggregation AS pa
      ORDER BY pa.Partner_Name, pa.Country;
    `;

    const ss = SpreadsheetApp.openById(PARTNER_DB_SS_ID);
    let sheet = ss.getSheetByName(SHEET_NAME_2026);
    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME_2026); }
    
    // Read existing data to preserve manual edits and deck links
    const existingData = sheet.getDataRange().getValues();
    const existingMap = new Map();
    let existingHeaders = [];
    let colEmails = -1;
    let colPDM = -1;
    if (existingData.length > 1) {
      existingHeaders = existingData[0].map(h => String(h).trim().toLowerCase());
      const colDomain = existingHeaders.indexOf("domain");
      const colCountry = existingHeaders.indexOf("country");
      colEmails = existingHeaders.indexOf("partner_emails");
      colPDM = existingHeaders.indexOf("pdm");
      
      for (let i = 1; i < existingData.length; i++) {
        const domain = String(existingData[i][colDomain]).trim().toLowerCase();
        const country = String(existingData[i][colCountry]).trim().toLowerCase();
        const key = `${domain}|${country}`;
        existingMap.set(key, existingData[i]);
      }
    }

    Logger.log("Executing Partner DB Generation query...");
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    if (!queryResults.rows || queryResults.rows.length === 0) { 
      sheet.getRange('A2').setValue("No results returned."); 
      return; 
    }
    
    const bqHeaders = queryResults.schema.fields.map(field => field.name);
    const bqRows = [];
    queryResults.rows.forEach(row => { 
      const rowData = row.f.map(field => field.v === null ? "" : field.v); 
      bqRows.push(rowData); 
    });
    
    // Merge data
    const mergedData = [];
    const combinedHeaders = [...bqHeaders];
    
    // Force headers for deck generator columns if missing or empty
    const expectedExtraHeaders = ["Spreadsheet_ID", "Spreadsheet_URL", "Deck_Status"];
    
    for (let i = 0; i < expectedExtraHeaders.length; i++) {
      const colIdx = bqHeaders.length + i;
      let header = "";
      if (existingHeaders.length > colIdx) {
        header = String(existingData[0][colIdx]).trim();
      }
      if (!header || header === "") {
        header = expectedExtraHeaders[i];
      }
      combinedHeaders.push(header);
    }
    
    // If there are even more columns in existing data, keep them
    if (existingHeaders.length > combinedHeaders.length) {
      for (let i = combinedHeaders.length; i < existingHeaders.length; i++) {
        combinedHeaders.push(existingData[0][i]);
      }
    }
    
    mergedData.push(combinedHeaders);

    bqRows.forEach(newRow => {
      const domain = String(newRow[7]).trim().toLowerCase(); // Domain is index 7 in new data
      const country = String(newRow[1]).trim().toLowerCase(); // Country is index 1 in new data
      const key = `${domain}|${country}`;
      
      if (existingMap.has(key)) {
        const existingRow = existingMap.get(key);
        const mergedRow = [...existingRow];
        
        // Pad mergedRow if it's shorter than combinedHeaders
        while (mergedRow.length < combinedHeaders.length) {
          mergedRow.push("");
        }
        
        // Merge PDM if PDM column was found in existing sheet
        const existingPDM = colPDM >= 0 ? String(existingRow[colPDM]).trim() : "";
        const newPDM = String(newRow[3]).trim(); // Index 3 is PDM in SQL
        
        const pdmSet = new Set();
        if (existingPDM) existingPDM.split(',').forEach(p => pdmSet.add(p.trim()));
        if (newPDM) newPDM.split(',').forEach(p => pdmSet.add(p.trim()));
        
        const mergedPDM = Array.from(pdmSet).filter(p => p !== "").join(', ');
        if (colPDM >= 0) {
          mergedRow[colPDM] = mergedPDM;
        }
        
        // Merge emails if email column was found in existing sheet
        const existingEmails = colEmails >= 0 ? String(existingRow[colEmails]).trim() : "";
        const newEmails = String(newRow[8]).trim(); // Index 8 is Partner_Emails in SQL SELECT
        
        const emailsSet = new Set();
        if (existingEmails) existingEmails.split(',').forEach(e => emailsSet.add(e.trim()));
        if (newEmails) newEmails.split(',').forEach(e => emailsSet.add(e.trim()));
        
        const mergedEmails = Array.from(emailsSet).filter(e => e !== "").join(', ');
        if (colEmails >= 0) {
          mergedRow[colEmails] = mergedEmails;
        }
        mergedData.push(mergedRow);
        existingMap.delete(key); // Mark as processed
      } else {
        // Pad new row with empty strings
        const paddedRow = [...newRow];
        while (paddedRow.length < combinedHeaders.length) {
          paddedRow.push("");
        }
        mergedData.push(paddedRow);
      }
    });
    
    // Append remaining existing rows that were not in BQ results
    existingMap.forEach((existingRow) => {
      mergedData.push(existingRow);
    });
    
    // Clear and write back
    sheet.clear();
    sheet.getRange(1, 1, mergedData.length, combinedHeaders.length).setValues(mergedData);
    
    // Formatting
    sheet.getRange(1, 1, 1, combinedHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
    sheet.autoResizeColumns(1, combinedHeaders.length);
    
    Logger.log("Partner DB load complete!");
  } catch (e) { 
    Logger.log("ERROR: " + e.toString()); 
  }
}
