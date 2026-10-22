# Party Invite & RSVP Platform

A lightweight, zero-cost, self-hosted event invitation and RSVP platform built with static HTML, Tailwind CSS, and Google Apps Script.

It automatically routes RSVPs from different event pages into corresponding tabs within a single Google Sheet based on your repository's directory structure.

------------------------------------------------------------------------

## 🏗️ Architecture & How It Works

```         
┌───────────────────────────────────────────────────────────┐
│                      GitHub Pages                         │
│                                                           │
│  .../party-invite/friendsgiving-2025/  (index.html)       │
│  .../party-invite/friendsgiving-2026/  (index.html)       │
└──────────────────────────────┬────────────────────────────┘
                               │
                               │ POST Request (JSON payload)
                               │ includes tabName extracted from URL path
                               ▼
┌───────────────────────────────────────────────────────────┐
│                   Google Apps Script                      │
│                                                           │
│  1. Receives payload (Name, Attendance, Board Idea, etc.) │
│  2. Checks Google Sheet for matching tabName              │
│  3. Creates new tab + headers if it doesn't exist yet     │
│  4. Appends guest submission as a new row                 │
└──────────────────────────────┬────────────────────────────┘
                               │
                               ▼
┌───────────────────────────────────────────────────────────┐
│                     Google Sheets                         │
│                                                           │
│  [friendsgiving-2025]  [friendsgiving-2026]  [...]         │
└──────────────────────────────┴────────────────────────────┘
```

1.  **Frontend Hosting (GitHub Pages):** Each event lives in its own subdirectory (e.g., `friendsgiving-2026/`).
2.  **Dynamic Routing (`getEventTabName()`):** On form submission, JavaScript inspects `window.location.pathname` to extract the current folder name (e.g., `friendsgiving-2026`).
3.  **Backend API (Google Apps Script):** The site `fetch()`es your script endpoint with a `POST` request containing guest details and the target `tabName`.
4.  **Database (Google Sheets):** The Google Apps Script appends the entry to the specified tab. If the tab doesn't exist yet, it creates it dynamically with default headers.

------------------------------------------------------------------------

## 📁 Repository Structure

``` text
party-invite/
├── README.md                  # System documentation
├── friendsgiving-2025/
│   └── index.html             # 2025 Event page & RSVP form
├── friendsgiving-2026/
│   └── index.html             # 2026 Event page & RSVP form
└── [future-event-folder]/
    └── index.html             # Future event copy
```

------------------------------------------------------------------------

## ⚙️ Backend Setup (Google Sheets & Apps Script)

### 1. Google Apps Script Code (`Code.gs`)

Attached to your target Google Sheet via **Extensions \> Apps Script**:

``` javascript
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  
  var tabName = data.tabName || 'General RSVPs';
  var sheet = ss.getSheetByName(tabName);
  
  // Auto-create tab with headers if missing
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow([
      'Timestamp',
      'Guest Name',
      'Attending',
      'Plus Ones',
      'Board Idea / Notes',
      'Contact Info'
    ]);
    sheet.getRange("A1:F1").setFontWeight("bold");
  }
  
  // Append guest response
  sheet.appendRow([
    new Date(),
    data.name,
    data.attending,
    data.plusOnes,
    data.dietary,
    data.contactInfo
  ]);
  
  return ContentService
    .createTextOutput(JSON.stringify({ "result": "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### 2. Deployment Settings

- **Execute as:** `Me`
- **Who has access:** `Anyone`

------------------------------------------------------------------------

## 🚀 Adding a New Event

To spin up an invitation page for a new party, just create a new folder in the repository matching your desired Google Sheet tab name.
