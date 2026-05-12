/**
 * ****************************************
 * Google Apps Script - One-Off Schema Check
 * File: one_off_schema_check_2026.js
 * Description: Queries BigQuery to inspect the schema of drp_partner_master.
 * ****************************************
 */

function checkSchemaOfPartnerMaster() {
  const SQL_QUERY = `
    SELECT * FROM \`concord-prod.service_partnercoe.drp_partner_master\` LIMIT 1
  `;

  Logger.log("Executing Schema Check Query...");
  try {
    const request = { query: SQL_QUERY, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
    
    if (!queryResults.rows || queryResults.rows.length === 0) { 
      Logger.log("No query results returned from BigQuery.");
      return; 
    }
    
    const headers = queryResults.schema.fields.map(field => field.name);
    Logger.log("Top-Level Schema Fields: " + JSON.stringify(headers));
    
    // Also log the structure of record fields
    queryResults.schema.fields.forEach(field => {
      if (field.type === 'RECORD') {
        const subFields = field.fields.map(f => f.name);
        Logger.log(`Record Field [${field.name}] sub-fields: ` + JSON.stringify(subFields));
        
        // Check for nested records (e.g., vector_details inside partner_details)
        field.fields.forEach(subField => {
          if (subField.type === 'RECORD') {
             Logger.log(`  Nested Record Field [${field.name}.${subField.name}] sub-fields: ` + JSON.stringify(subField.fields.map(f => f.name)));
          }
        });
      }
    });

    Logger.log("Check complete. Please provide the log output.");

  } catch (e) {
    Logger.log("ERROR in Schema Check: " + e.toString());
  }
}
