/**
 * ****************************************
 * Google Apps Script - Step-by-Step Update
 * File: update_step_by_step.js
 * Description: Functions to run each update step individually or all at once.
 * WARNING: Running all steps in one go might hit the execution time limit (6 mins).
 * ****************************************
 */

function updateEverything2026() {
  Logger.log(">>> STARTING ALL-IN-ONE UPDATE PROCESS <<<");

  step_1_generateDb();
  step_2_updateScoring();
  step_3_updateDeepDive();
  step_4_recreateDashboard();
  step_5_generateDecks();
  step_6_sendEmails();

  Logger.log(">>> ALL-IN-ONE UPDATE COMPLETE <<<");
}

function step_1_generateDb() {
  Logger.log("Step 1: Generating Partner Database...");
  runPartnerDbGeneration2026();
}

function step_2_updateScoring() {
  Logger.log("Step 2: Updating Scoring Pivot...");
  runPartnerScorePivot2026();
}

function step_3_updateDeepDive() {
  Logger.log("Step 3: Updating Profile Deep Dive...");
  runDeepDive2026();
}

function step_4_recreateDashboard() {
  Logger.log("Step 4: Recreating Dashboard...");
  setupDashboard2026();
}

function step_5_generateDecks() {
  Logger.log("Step 5: Updating Partner Decks (Full Batch)...");
  Logger.log("WARNING: This step might take a long time and could cause timeout.");
  runFullBatchDecks2026();
}

function step_6_sendEmails() {
  Logger.log("Step 6: Sending Emails (Full Batch)...");
  Logger.log("WARNING: This step might also take a long time.");
  runBatchEmailSender2026();
}
