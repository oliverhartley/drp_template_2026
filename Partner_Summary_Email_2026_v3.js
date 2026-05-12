/**
 * ****************************************
 * Google Apps Script - Partner Summary Email 2026
 * File: Partner_Summary_Email_2026.js
 * Description: Generates an executive summary using Gemini and sends it via email.
 * Iterates over LATAM_Partner_DB_2026 and aggregates by Partner Name.
 * ****************************************
 */

const BATCH_EMAIL_TIME_LIMIT_MS_2026 = 1200000; // 20 minutes

function runBatchEmailSender2026() {
  const startTime = new Date().getTime();
  const currentBatchId = getBatchId2026(); 
  Logger.log(`>>> STARTING 2026 BATCH EMAIL PROCESS [Batch ID: ${currentBatchId}] <<<`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = SpreadsheetApp.openById(PARTNER_DB_SS_ID).getSheetByName(SHEET_NAME_2026);
  
  if (!dbSheet) {
    Logger.log(`ERROR: Sheet ${SHEET_NAME_2026} not found.`);
    return;
  }

  const dataRange = dbSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());

  // Find columns based on new DB schema
  const colPartnerName = 0; // Partner Name is A (0)
  
  let colToEmail = -1;
  let colCcEmail = -1;
  let colPdm = -1;
  let colStatus = -1;
  let colSpreadsheetId = -1;
  let colDomain = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h === "partner_emails") colToEmail = c;
    if (h === "cc_tech_lead") colCcEmail = c;
    if (h === "pdm") colPdm = c;
    if (h === "spreadsheet_id") colSpreadsheetId = c;
    if (h === "email_status") colStatus = c;
    if (h === "domain") colDomain = c;
  }

  // Fallback if not found by exact name
  if (colToEmail === -1) colToEmail = 8; // Default to Col I
  if (colCcEmail === -1) colCcEmail = 9; // Default to Col J
  if (colPdm === -1) colPdm = 3; // Default to Col D
  if (colSpreadsheetId === -1) colSpreadsheetId = 10; // Default to Col K
  if (colDomain === -1) colDomain = 7; // Default to Col H (index 7)
  
  // Set up header for Email Status in Col N (13) if missing
  if (colStatus === -1) {
    dbSheet.getRange("N1").setValue("email_status").setBackground("#d9d9d9").setFontWeight("bold");
    colStatus = 13;
  }

  if (colToEmail === -1 || colCcEmail === -1 || colSpreadsheetId === -1 || colStatus === -1 || colPdm === -1) {
      Logger.log("ERROR: Could not find one or more required columns. Please ensure they exist.");
      return;
  }

  // Group by Domain to avoid variations in Partner Name
  const partnerMap = new Map();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const domain = String(row[colDomain]).trim().toLowerCase();
    if (!domain) continue;

    const partnerName = String(row[colPartnerName]).trim();

    if (!partnerMap.has(domain)) {
      partnerMap.set(domain, {
        name: partnerName, // Representative name (first one encountered)
        rows: [],
        toEmails: new Set(),
        ccEmails: new Set(),
        status: "",
        spreadsheetId: ""
      });
    }

    const pData = partnerMap.get(domain);
    pData.rows.push(i); // Store 0-indexed row number for updating status
    
    if (!pData.name) pData.name = partnerName;
    
    // Default to the first found status and ID for the group
    if (!pData.status && row[colStatus]) pData.status = String(row[colStatus]).trim();
    if (!pData.spreadsheetId && row[colSpreadsheetId]) pData.spreadsheetId = String(row[colSpreadsheetId]).trim();

    // Collect all unique emails for this partner
    let to = String(row[colToEmail] || "").trim();
    let cc = String(row[colCcEmail] || "").trim();
    let pdm = String(row[colPdm] || "").trim();
    
    // Fallback logic: If no partner email, send to PDM
    if (!to || to === "") {
      to = pdm;
    } else {
      // If partner email exists, add PDM to CC
      if (pdm) pdm.split(',').forEach(e => pData.ccEmails.add(e.trim()));
    }
    
    if (to) to.split(',').forEach(e => pData.toEmails.add(e.trim()));
    if (cc) cc.split(',').forEach(e => pData.ccEmails.add(e.trim()));
  }

  let processedCount = 0;
  let skippedCount = 0;

  for (const [domain, pData] of partnerMap.entries()) {
    if (isTimeLimitApproaching2026(startTime)) {
      Logger.log("WARNING: Time limit approaching. Stopping to allow safe resume on next trigger.");
      break;
    }

    if (pData.ccEmails.size === 0 && pData.toEmails.size === 0) {
      Logger.log(`Skipping ${pData.name} (${domain}) - No emails found.`);
      continue;
    }

    if (pData.status === currentBatchId) {
      Logger.log(`Skipping ${pData.name} (${domain}) - Already processed for this batch.`);
      skippedCount++;
      continue;
    }

    if (!pData.spreadsheetId) {
      Logger.log(`Skipping ${pData.name} (${domain}) - No Spreadsheet ID found. Please generate their deck first.`);
      continue;
    }

    Logger.log(`Processing Partner: ${pData.name} (${domain})...`);
    
    const toEmailStr = Array.from(pData.toEmails).join(",");
    const ccEmailStr = Array.from(pData.ccEmails).join(",");

    try {
      generateAndSendPartnerSummary2026(pData.name, pData.spreadsheetId, toEmailStr, ccEmailStr);

      // Update Status for all rows belonging to this partner
      pData.rows.forEach(r => {
        dbSheet.getRange(r + 1, colStatus + 1).setValue(currentBatchId);
      });
      SpreadsheetApp.flush();
      processedCount++;

      Utilities.sleep(5000); // Respect Gemini quotas
    } catch (e) {
      Logger.log(`  ERROR processing ${pData.name}: ${e.toString()}`);
      // Clear status on failure to ensure retry later
      try {
        pData.rows.forEach(r => {
          dbSheet.getRange(r + 1, colStatus + 1).setValue("");
        });
        SpreadsheetApp.flush();
      } catch (statusError) {
        Logger.log(`    Failed to clear status for failed partner: ${statusError.toString()}`);
      }
    }
  }

  Logger.log(`>>> 2026 BATCH RUN COMPLETE. Sent: ${processedCount}, Skipped: ${skippedCount} <<<`);
}

function getBatchId2026() {
  const now = new Date();
  const shiftedDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const year = shiftedDate.getFullYear();
  const onejan = new Date(year, 0, 1);
  const week = Math.ceil((((shiftedDate.getTime() - onejan.getTime()) / 86400000) + onejan.getDay() + 1) / 7);
  return `SENT_2026_${year}_${week}`;
}

function isTimeLimitApproaching2026(startTime) {
  return (new Date().getTime() - startTime) > BATCH_EMAIL_TIME_LIMIT_MS_2026;
}

function generateAndSendPartnerSummary2026(partnerName, ssId, toEmails, ccEmails) {
  Logger.log(`  Generating 2026 summary for ${partnerName}...`);

  const sheetData = getPartnerSheetData2026(ssId);
  if (!sheetData) {
    Logger.log("  ERROR: Failed to retrieve sheet data.");
    return;
  }
  
  const fullPrompt = `
    You are an expert Data Analyst and Executive Assistant.
    Please analyze the following 2026 data for partner: "${partnerName}".
    
    Data from "Tier Dashboard":
    ${sheetData.tierDashboard}
    
    Data from "Profile Deep Dive":
    ${sheetData.profileDeepDive}
    
    Task:
    Create a comprehensive Email Report containing TWO SECTIONS:
    
    SECTION 1: VISUAL EXECUTIVE DASHBOARD (The "Infographic")
    - Start with this EXACT greeting: "Hola ${partnerName},<br><br>Aquí su informe semanal del DRP Status para su análisis. Cualquier duda puedes contactar al equipo de Partner (copiado en este correo)."
    - This must be a graphical representation using ONLY HTML/CSS, fully responsive and optimized for email clients (inline CSS is mandatory).
    - Do NOT use images or external charts. Use HTML/CSS to create visual "Scorecards", clean tables, and visual "Bar Charts".
    - **CRITICAL STRATEGY & TARGET FOCUS**:
        - Focus all analysis (visual and narrative) exclusively on **Tier 1 (Delivery Ready / Expert)** and **Tier 2 (Advanced / Intermediate)** profiles.
        - Omit detailed analysis of Tiers 3 and 4; treat them simply as the pool of 'not yet ready' profiles.
    - **CRITICAL LAYOUT REQUIREMENT - STRICT SINGLE-COLUMN VERTICAL STACK**:
        - Enforce a strictly **single-column, vertical stack layout** for the entire email (max width 700px).
        - **Do NOT use multi-column layouts** where Solutions or different sections are placed side-by-side. Everything must flow vertically, one section/row below the other, to prevent messy horizontal wrapping in email clients.
    - Layout & Design Guidelines:
        - **Overall Style**: Premium, modern look. Use a clean sans-serif font (e.g., Arial, Helvetica). Centered with a subtle light border or shadow wrapper.
        - **Header**: Partner Name & "Readiness Snapshot" styled as a prominent title.
        - **KPI Row**: 3 Scorecard Cards aligned horizontally.
            - Show: Card 1: **Total Profiles**, Card 2: **Tier 1 (Delivery Ready)** count, Card 3: **Tier 2 (Advanced)** count.
            - Do NOT show "Profiles with No Tiers" or "Beginner" counts here.
            - Each card should have: subtle light grey background ('#f8f9fa'), rounded corners (8px), border ('1px solid #e0e0e0'), and a top accent border (4px thick) in a Google Brand color (Blue, Green, or Yellow).
            - Metrics (numbers) should be large, bold, and centered.
        - **Solution Readiness Snapshot Table**:
            - Display a **clean, compact table** listing the high-level Solutions.
            - Do NOT list individual products. Only show the Solutions (e.g. Infrastructure Modernization, Application Modernization, etc.).
            - Table Structure:
                - Column 1: **Solution** (bold, aligned left).
                - Column 2: **Tier 1 (Delivery Ready)** count.
                - Column 3: **Tier 2 (Advanced)** count.
            - Table styling:
                - Header row: Prominent background color (e.g., Google Blue '#4285f4' or dark grey '#3c4043') with white text, bold, centered headers.
                - Zebra striping (alternating '#f9f9f9' and white rows).
                - Thin grey borders ('1px solid #eeeeee').
                - Explicit padding (8px to 10px) for all cells.
                - Center align the numeric counts in Columns 2 and 3.
        - **Upskilling Gaps (Focus on Tier 1/2 Target)**:
            - A clean table showing "Solution", "Current Ready (Tier 1/2)", "Target (Tier 1/2)", and "Gap to Target".
            - Table styling: Alternating row colors (zebra striping with '#f9f9f9'), thin grey borders ('1px solid #eeeeee'), explicit column padding (10px), bold header row with light grey background ('#eaecf0').
        - **Top Talent (Tier 1/2)**:
            - A clean table showing the top 3-5 individuals who have achieved **Tier 1 or Tier 2** status.
            - Columns: "Profile ID", "Country", "Job Title", "Best Tier Achieved".
            - Do NOT list individuals who are only Tier 3 or Tier 4.
    - Style: Use Google Brand colors (Blue #4285f4, Red #ea4335, Yellow #fbbc04, Green #34a853). Use Grey #f1f3f4 for backgrounds.
    
    SECTION 2: DETAILED EXECUTIVE SUMMARY
    - Written narrative explaining the data, focusing on Tier 1 and Tier 2 talent availability, distribution, and upskilling strategies to move talent to Tier 1 & 2.
    - Tiers Definitions for context (focus on Tiers 1 & 2 as target):
      - Tier 1: Delivery Ready (Expert).
      - Tier 2: Advanced / Intermediate.
      - Tier 3: Beginner-Intermediate (Not Ready).
      - Tier 4: Beginner (Not Ready).
    - Sections: "Key Strengths (Tier 1 & 2 Availability)", "Critical Gaps (Tier 1 & 2 Sufficiency)", "Recommendations (Upskilling to Tier 1 & 2)".
    - Output Format:
    Return ONE block of clean, professional HTML.
    - Use Inline CSS for everything (Gmail compatible).
    - Make it look premium (padding, border-radius, shadows).
  `;

  const finalHtml = callGeminiWithFallback2026(fullPrompt);
  if (!finalHtml) {
    Logger.log("  ERROR: Failed to generate summary from Gemini.");
    throw new Error("Failed to generate summary from Gemini.");
  }

  const subject = `[GCP DRP Readiness 2026] Partner Executive Summary: ${partnerName}`;
  const fileUrl = `https://docs.google.com/spreadsheets/d/${ssId}/edit`;
  
  let cleanHtml = finalHtml.replace(/```html/g, "").replace(/```/g, "").trim();

  const emailBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: 0 auto;">
      ${cleanHtml}
      <br><br>
      <hr>
      <p style="text-align: center;">
        <a href="${fileUrl}" style="background-color: #4285f4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">
          Open 2026 Partner Dashboard
        </a>
      </p>
      <p style="font-size: 12px; color: #666; text-align: center;">
        Link to file: <a href="${fileUrl}">${fileUrl}</a>
      </p>
      
      <!-- Footer -->
      <br>
      <div style="text-align: center; color: #999; font-size: 11px; margin-top: 20px;">
        <p>&copy; 2026 Google Cloud Partner Team. Confidential.</p>
        <p style="font-style: italic;">
          This summary was generated by Gemini. Any imprecision, please let the team know.
        </p>
      </div>
    </div>
  `;

  sendEmail2026(subject, emailBody, toEmails, ccEmails);
}

function getPartnerSheetData2026(ssId) {
  try {
    const ss = SpreadsheetApp.openById(ssId);

    const tierSheet = ss.getSheetByName("Tier Dashboard");
    const deepDiveSheet = ss.getSheetByName("Profile Deep Dive");

    if (!tierSheet || !deepDiveSheet) {
      Logger.log("ERROR: Missing required sheets in deck.");
      return null;
    }

    const tierData = tierSheet.getDataRange().getValues().map(row => row.join(", ")).join("\\n");
    const deepDiveData = deepDiveSheet.getRange(1, 1, Math.min(deepDiveSheet.getLastRow(), 200), deepDiveSheet.getLastColumn()).getValues().map(row => row.join(", ")).join("\\n");

    return {
      tierDashboard: tierData,
      profileDeepDive: deepDiveData
    };
  } catch (e) {
    Logger.log(`Error reading sheets for deck ${ssId}: ${e.toString()}`);
    return null;
  }
}

function callGeminiWithFallback2026(prompt) {
  const userModels = [
    { name: 'gemini-2.5-flash-lite', version: 'v1' },
    { name: 'gemini-2.5-flash', version: 'v1' }
  ];

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties.");
    return null;
  }

  for (const model of userModels) {
    try {
      const url = `https://generativelanguage.googleapis.com/${model.version}/models/${model.name}:generateContent?key=${apiKey}`;

      const payload = {
        contents: [{
          parts: [{ text: prompt }]
        }]
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        const json = JSON.parse(responseText);
        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts) {
          return json.candidates[0].content.parts[0].text;
        }
      } else {
        Logger.log(`FAILED: Model ${model.name} returned code ${responseCode}. Response: ${responseText}`);
      }
    } catch (e) {
      Logger.log(`EXCEPTION: Model ${model.name} failed with error: ${e.toString()}`);
    }
  }
  return null;
}

function sendEmail2026(subject, htmlBody, to, cc) {
  try {
    const emailOptions = {
      to: to,
      subject: subject,
      htmlBody: htmlBody
    };

    if (cc && String(cc).trim() !== "") {
      emailOptions.cc = cc;
    }

    if (!to || String(to).trim() === "") {
      Logger.log("  WARNING: 'TO' email is empty. Attempting to send using CC only if possible, or aborting.");
      if (emailOptions.cc) {
        emailOptions.to = emailOptions.cc;
      } else {
        Logger.log("  ERROR: No recipients defined. Skipping email.");
        return;
      }
    }

    MailApp.sendEmail(emailOptions);
    Logger.log(`  Email sent to: ${emailOptions.to} (CC: ${cc || 'None'})`);
  } catch (e) {
    Logger.log(`  Error sending email: ${e.toString()}`);
  }
}

function createPartnerEmailTriggers2026() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runBatchEmailSender2026') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const tuesdayHours = [13, 14, 15, 16, 17, 18, 19, 20, 21, 22];
  
  tuesdayHours.forEach(hour => {
    ScriptApp.newTrigger('runBatchEmailSender2026')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .atHour(hour)
      .create();
  });
  
  const mondayHours = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
  
  mondayHours.forEach(hour => {
    ScriptApp.newTrigger('runBatchEmailSender2026')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(hour)
      .create();
  });
  
  Logger.log(`Successfully created triggers for 'runBatchEmailSender2026':`);
  Logger.log(`- Tuesdays at hours: ${tuesdayHours.join(', ')}`);
  Logger.log(`- Mondays at hours: ${mondayHours.join(', ')}`);
}
