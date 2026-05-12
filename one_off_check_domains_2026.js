/**
 * ****************************************
 * Google Apps Script - One-Off Domain Check
 * File: one_off_check_domains_2026.js
 * Description: Queries BigQuery to check the format of email domains.
 * ****************************************
 */

function checkDomainsFormat() {
  const SQL_QUERY = `
    SELECT partner_name, partner_details.email_domain
    FROM \`concord-prod.service_partnercoe.drp_partner_master\`
    WHERE partner_details.email_domain IS NOT NULL AND ARRAY_LENGTH(partner_details.email_domain) > 0
    LIMIT 5
  `;

  Logger.log("Executing Domain Check Query...");
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    if (!queryResults.rows || queryResults.rows.length === 0) { 
      Logger.log("No query results returned from BigQuery.");
      return; 
    }
    
    queryResults.rows.forEach(row => {
      Logger.log("Row: " + JSON.stringify(row.f.map(cell => cell.v)));
    });

  } catch (e) {
    Logger.log("ERROR in Domain Check: " + e.toString());
  }
}
