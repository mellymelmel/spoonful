# Case Dashboard Customizer & Tracker

A Google Apps Script web app that tracks team cases backed by Google Sheets.
The dashboard lists cases by team-member initials and a main link, with
user-customizable columns. A Settings tab manages team members (initials, name,
role) and the column layout.

## Files

| File              | Role                                                  |
|-------------------|-------------------------------------------------------|
| `appsscript.json` | Apps Script manifest (scopes, web-app config)         |
| `Code.gs`         | Server-side logic + sheet I/O                         |
| `Index.html`      | Dashboard + Settings shell                            |
| `Stylesheet.html` | Styles (included via `<?!= include('Stylesheet') ?>`) |
| `JavaScript.html` | Client-side controller                                |

## Backend sheets (auto-created on first run)

- **Cases** — `ID | Initials | MainLink | <custom columns...>`
- **TeamMembers** — `Initials | Name | Role`
- **Columns** — `ColumnId | ColumnName | ColumnType | Options | Order`

Supported column types: `text`, `number`, `date`, `link`, `dropdown`, `checkbox`.
For `dropdown`, fill the **Options** field with a comma-separated list.

## Deploy

### Option A — Apps Script bound to a sheet

1. Create a new Google Sheet.
2. **Extensions → Apps Script**.
3. Paste the contents of `Code.gs` into `Code.gs`.
4. Add three HTML files (`Index`, `Stylesheet`, `JavaScript`) and paste each.
5. Click the gear → "Show appsscript.json" and replace it with this repo's
   `appsscript.json`.
6. **Deploy → New deployment → Web app**. Choose who can access. Authorize.
7. Open the web-app URL. Sheets are created on first load.

### Option B — Standalone script with a separate spreadsheet

1. Create the script at <https://script.google.com>.
2. Add the same files as above.
3. **Project Settings → Script Properties** → add
   `SPREADSHEET_ID` = the id of the Google Sheet you want to use.
4. Deploy as a web app.

### Option C — `clasp` push

```bash
npm i -g @google/clasp
clasp login
clasp create --type webapp --title "Case Dashboard" --rootDir .
clasp push
clasp deploy
```

## Using it

- **Cases tab** — filter by text or by team-member initials, add cases via the
  modal, edit or delete inline.
- **Settings tab**
  - *Team Members* — add/remove members. Initials are normalized to upper-case
    and used as the key.
  - *Custom Columns* — add columns of any supported type, reorder with the up/down
    arrows, edit or delete. Renaming a column updates the corresponding header
    in the Cases sheet without losing data; deleting a column removes the data.
