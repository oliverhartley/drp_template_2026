/*
# Skill: Google Apps Script & GitHub Sync Setup

This document describes the workflow for setting up a new Google Apps Script project with local `clasp` development and GitHub synchronization. Future instances of the AI can read this file to execute the workflow.

## Instructions for the AI Assistant

When the user asks to "Start a new Apps Script project using the sync skill", follow these steps:

### Step 1: Gather Information
Ask the user for:
1.  **Google Apps Script ID** (Script ID).
2.  **GitHub Repository Name** (to create).

### Step 2: Create GitHub Repository
Use the GitHub CLI (`gh`) to create a new repository under the user's account:
```bash
gh repo create <RepositoryName> --public --confirm
```
*(Note: Defaulting to public, ask if private is needed).*

### Step 3: Configure Clasp Locally
1. Create a new directory for the project named `<RepositoryName>`.
2. Run `npx @google/clasp clone <ScriptID>` in that directory.

### Step 4: Initialize Git and Sync
1. In the project directory, run `git init` (if not already part of a parent repo).
2. Add the remote: `git remote add origin https://github.com/oliverhartley/<RepositoryName>.git`
3. Stage files: `git add .`
4. Commit: `git commit -m "Initial commit of Apps Script files"`
5. Push: `git push -u origin main`

### Step 5: Maintain Sync (Ongoing)
Whenever changes are made to the files:
1. Run `npx @google/clasp push` to update Google Apps Script.
2. Run `git add .`
3. Run `git commit -m "<Meaningful message explaining the changes>"`
4. Run `git push` to update GitHub.

---
Created on 2026-04-06.
*/
