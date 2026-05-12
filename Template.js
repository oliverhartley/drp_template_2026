/*
# Template Guide - DRP 2026

This file serves as a guide to replicate this project for a new region or a new set of partners.

## Step 1: Copy the Files

You need to create copies of the two main spreadsheets used in this workflow:

1.  **Source Spreadsheet**: This is the file that contains the list of partners and their emails.
    *   Action: Make a copy of the source file provided for your region.
    *   Note the new Spreadsheet ID.

2.  **Master Spreadsheet (Destination)**: This is the file where the scripts run, and it contains the Dashboard, Database, and Score sheets.
    *   Action: Make a copy of the master file. This copy will also copy the attached Apps Script project!
    *   Note the new Spreadsheet ID.

## Step 2: Create a New Drive Folder

Individual decks for each partner will be generated and saved in a specific Google Drive folder.
1.  Go to Google Drive.
2.  Create a new folder for your partner decks.
3.  Copy the ID of this folder (you can find it in the URL when you open the folder).

## Step 3: Update the Configuration

Open the `Config_2026_w2.js` file in the Apps Script editor of your *copied* Master Spreadsheet and update the following variables with your new IDs:

1.  `DESTINATION_SS_ID`: Set this to the ID of your newly copied **Master Spreadsheet**.
2.  `NEW_SOURCE_SS_ID`: Set this to the ID of your newly copied **Source Spreadsheet**.
3.  `PARTNER_FOLDER_ID`: Set this to the ID of the new **Drive Folder** you created in Step 2.

## Step 4: Run the Setup

Once the configuration is updated, open the `update_step_by_step.js` file and run the functions in order to generate the database, update data, and generate decks for your new set of partners!
*/
