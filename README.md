# Party Invite & RSVP Platform

A lightweight, self-hosted invitation and RSVP system for personal events.

It hosts event pages on GitHub Pages and routes guest responses straight into dedicated tabs in a single Google Sheet—no backend server or database setup required.

## 💡 How It Works

1. **Guest Submits Form:** A guest fills out the RSVP form on an event page (e.g., `.../friendsgiving-2026/`).

2. **Dynamic Folder Name Extraction:** The page automatically reads its URL path name (e.g., `friendsgiving-2026`).

3. **Bridge to Google Sheets:** JavaScript sends the form answers and folder name to a custom Google Apps Script.

4. **Automatic Sheet Update:** The script checks your Google Sheet for a tab matching that folder name. If the tab doesn't exist yet, it creates it with headers, then appends the guest's RSVP as a new row.

```
[ Guest Fills Out Form ]
          │
          ▼
[ Website Reads Folder Name: "friendsgiving-2026" ]
          │
          ▼
[ Script Receives Data ➔ Finds/Creates Tab ➔ Adds Row in Google Sheet ]
```

## 📁 Repository Structure

Each event lives in its own dedicated folder containing an `index.html` file. The repository also contains the Google Apps Script code used in Google Sheets.

```
party-invite/
├── README.md
├── Code.gs                 # Google Apps Script backend code
├── friendsgiving-2025/
│   └── index.html          # 2025 Event page & RSVP form
├── friendsgiving-2026/
│   └── index.html          # 2026 Event page & RSVP form
└── [new-event-folder]/
    └── index.html          # Future event folder & page
```

## ⚙️ Google Apps Script Backend

The Google Apps Script logic is saved locally in this repository at `Code.gs`.

To update or set up the script in Google Sheets:
1. Open your destination Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Copy the code from `Code.gs` in this repository into your Apps Script editor.

### Script Deployment Settings
When deploying the script from Google Sheets, use the following settings:
* **Execute as:** `Me`
* **Who has access:** `Anyone`

## 🛠️ How to Make Updates Down the Line

### Adding a New Event

1. Copy an existing event folder (e.g., duplicate `friendsgiving-2026/` and rename it `friendsgiving-2027/`).

2. Edit the text and event details inside `friendsgiving-2027/index.html`.

3. Commit and push your changes to GitHub. When the first guest submits an RSVP on the new page, the script will automatically create a tab named `friendsgiving-2027` in your Google Sheet.

### Changing Form Fields or Sheet Columns

If you add, remove, or rename questions in your HTML form:

1. Update `Code.gs` locally in your repository:
   * **Update headers:** Modify the array inside `if (!sheet)` to set your column titles for new tabs.
   * **Update output row:** Modify the array inside `sheet.appendRow([...])` at the bottom to match your new form variable names (`data.yourNewField`).
2. Copy the updated code from `Code.gs` into your Google Sheet's script editor (**Extensions > Apps Script**).
3. **Publish new deployment:** Click **Deploy > Manage Deployments**, select the edit icon next to your active deployment, select **New version**, and click **Deploy**.