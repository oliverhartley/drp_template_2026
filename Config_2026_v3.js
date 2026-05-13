/**
 * ****************************************
 * Google Apps Script - Global Configuration
 * File: Config.gs
 * Description: Centralized variables used across all scripts.
 * ****************************************
 */

const PROJECT_ID = "concord-prod"; 

// The Spreadsheet where the scripts run (Destination)
const DESTINATION_SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); 

// New Source Spreadsheet for 2026 Partners (now the Active Spreadsheet)
const NEW_SOURCE_SS_ID = DESTINATION_SS_ID;
const NEW_SHEET_NAME_PARTNER = "Partners";

// Dedicated Spreadsheet for Partner DB
const PARTNER_DB_SS_ID = "1Cvr1uL9Bj2uPCC7kgTdV3WTBDSiAFHEbajdgpU6MaRM";

// The Folder where individual decks are saved
const PARTNER_FOLDER_ID = "14BEnsiPHs9ldiCE7aGT06E_QXutMq8c-"; 

// Sheet Names
const SHEET_NAME_2026 = "LATAM_Partner_DB_2026_v3";
const SHEET_NAME_SCORE_2026 = "LATAM_Partner_Score_2026_v3";
const SHEET_NAME_DEEPDIVE_2026 = "LATAM_DeepDive_2026_v3";
const SHEET_NAME_DASHBOARD_2026 = "DRP Dashboard by Partner / Region";
const SHEET_NAME_CONSOLIDATED_DASHBOARD = "DRP Consolidated Dashboard";
const SHEET_NAME_CACHE_2026 = "CACHE_Dashboard_2026_v3";
const DECK_SHEET_NAME = "Tier Dashboard";

const PRODUCT_SCHEMA = [
  { solution: 'Infrastructure Modernization', color: '#fce5cd', products: ['Google Compute Engine', 'Google Cloud Networking', 'SAP on Google Cloud', 'Google Cloud VMware Engine', 'Google Distributed Cloud'] },
  { solution: 'Application Modernization', color: '#d9d2e9', products: ['Google Kubernetes Engine', 'Apigee API Management'] },
  { solution: 'Databases', color: '#fce5cd', products: ['Cloud SQL', 'AlloyDB for PostgreSQL', 'Spanner', 'Cloud Run', 'Oracle'] },
  { solution: 'Data & Analytics', color: '#d9ead3', products: ['BigQuery', 'Looker', 'Dataflow', 'Dataproc'] },
  { solution: 'Artificial Intelligence', color: '#c9daf8', products: ['Vertex AI Platform', 'AI Applications', 'Gemini Enterprise', 'Gemini Enterprise for Customer Experience'] },
  { solution: 'Security', color: '#f4cccc', products: ['Cloud Security', 'Security Command Center', 'Security Operations', 'Google Threat Intelligence'] },
  { solution: 'Workspace', color: '#fff2cc', products: ['Workspace'] }
];