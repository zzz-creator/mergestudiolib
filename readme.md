# Merge Studio

A free, self-hosted AutoCrat alternative built on Google Apps Script. Designed as a standalone Library project so you can maintain a single codebase and attach it to any number of Google Sheets without requiring Google Workspace Marketplace verification.

**Only one job per spreadsheet is supported**
---

## Features


- **Smart Tag Scanning:** Automatically detects `<<Tag>>` placeholders across your Google Doc body, headers, and footers.
- **Branching Logic:** Supports row-level merge conditions alongside value-based branch rules to dynamically override the Template Doc, Destination Folder, and Email Templates.
- **Two-Stage Application Review & Multi-State Decision Workflow:**
  - **Stage 1:** Automatically notifies a designated reviewer when a document is merged.
  - **Stage 2:** Monitors a dedicated review column and sends state-specific emails (e.g., *Accepted*, *Waitlisted*, *Interview*, *Rejected*) as statuses change.
- **Flexible Output & Automation:** Export as **PDF**, **Google Doc**, or **Both**, and automate execution via *On Form Submit* or *Time-Driven* triggers.

---

## Setup & Deployment Instructions

### 1. Create the Library Project
1. Go to [script.google.com](https://script.google.com) and click **New project**.
2. Name the project **Merge Studio Library** (or `MergeStudioLib`).
3. Replace the default `Code.gs` contents with the entire `LibraryCode.gs` file (you can rename `Code.gs` to `LibraryCode.gs` inside the editor).
4. Click the **Save** icon (`Ctrl+S` / `Cmd+S`).

### 2. Deploy a Versioned Library
1. Click **Deploy** > **New deployment** in the top-right corner.
2. Click the gear icon next to **Select type** and choose **Library**.
3. Enter a version description (e.g., `v1.2.0 - Multi-State Review & Branching`) and click **Deploy**.
4. Open **Project Settings** (gear icon on the left sidebar) and copy the **Script ID**.

### 3. Share with Collaborators / Teammates
1. Click **Share** in the top-right corner of the Apps Script editor (or share the script file directly from Google Drive).
2. Grant teammates **Viewer** access (read-only access is sufficient for them to run the library from their own spreadsheets without modifying your source code).

> **Note:** Users with Viewer access must select a numbered version when linking the library. They cannot run `HEAD (Development mode)`.

### 4. Link to a Google Sheet & First-Run Authorization
1. Open your target Google Sheet and go to **Extensions** > **Apps Script**.
2. Paste `WrapperCode.js` into `Code.gs` and add `SidebarUI.html` as an HTML file named `SidebarUI`.
3. In the left sidebar, click **+** next to **Libraries**, paste your **Script ID**, select the latest version, and set the Identifier to `MergeStudioLib`.
4. Reload your Google Sheet and open **Merge Studio** from the top menu bar.
5. On the first run, Google will display an **Authorization required** prompt (and *"This app isn't verified"* on standard `@gmail.com` accounts). Click **Advanced** > **Go to \<Project Name\> (unsafe)** > **Allow** to grant Drive, Docs, Sheets, and Gmail permissions to your private script.
