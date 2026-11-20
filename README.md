# RPI Project Materials Inventory

A web app for tracking project materials at Rite Products, Incorporated: stock on hand, materials received from suppliers, and materials delivered to projects. It runs as a Google Apps Script web app and stores its data in Google Sheets.

Current version: **v1.3**

## Features

- **Two projects, kept apart.** A switcher in the sidebar moves between **HVAC PROJECTS** and **SYSTEMS PROJECTS**. Each project has its own Google Sheet, so items, stock, transactions and edit requests never mix. One login works for both.
- **Stocks.** On hand = beginning balance + received − delivered, with Low and Out of stock flags against each item's reorder level.
- **Received and Delivered.** Record several items in one transaction. Transactions are grouped by PO No. (received) or DR No. (delivered), newest first.
- **Import.** Download a template (stock items, received items, delivered items), fill it in Excel or Google Sheets, and upload it as `.xlsx`, `.xls` or `.csv`. Every row is checked first (missing columns, dates, quantities, duplicate items, unknown items, stock going below zero). Nothing is saved if any row has a problem.
- **Remembered entries.** Text fields suggest values typed before and values already saved in the records (suppliers, projects, invoice, PO, JO and DR numbers, units).
- **Roles.** Admin, Editor and Viewer, set from a dropdown in the Users tab and copied to a Users tab in each Google Sheet.
- **Edit approvals.** Editors can add records, but changing an existing record needs an admin's approval.
- **Database tab (admins).** Connection status, a Refresh connection button, record counts, a link to the sheet, and a button that rebuilds the readable sheet tabs.
- Light and dark themes, and a layout that works on phones.

## Roles

| | Admin | Editor | Viewer |
|---|---|---|---|
| View Stocks, Received, Delivered | ✓ | ✓ | ✓ |
| Add records and import files | ✓ | ✓ | |
| Edit existing records | ✓ | after approval | |
| Users tab (accounts, roles, edit requests) | ✓ | | |
| Database tab and Open Google Sheet | ✓ | | |

The account set in `ADMIN_EMAIL` is the owner. It signs in with its Google account (no password), is created automatically the first time it opens the app, and always stays an admin. There is no sign-up page: admins add every other account from the Users tab.

## Files

| File | What it is |
|---|---|
| `Code.gs` | Server code: sign-in, sessions, roles, reading and writing the sheets |
| `index.html` | The whole page: layout, styles and browser code |
| `appsscript.json` | Apps Script manifest: time zone, V8 runtime, web app access settings |

## How data is stored

Each project's Google Sheet has these tabs:

| Tab | Contents |
|---|---|
| App Data | The app's own records, one row per record. Don't edit by hand. |
| Items, Received, Delivered | Readable copies the app rewrites after every save, newest first |
| Stocks, Received Summary, Delivered Summary | Formulas over the readable tabs |
| Users | Name, username, email and role of every account. Passwords are never written here. |

Accounts, password hashes and edit requests are kept in the script's Script Properties, not in the sheets.

## Security

- The web app only opens for Google accounts in the company domain (`ALLOWED_DOMAIN`), and the server checks this on every request.
- Passwords are hashed in the browser (PBKDF2, 150,000 rounds, per-user salt) and checked on the server. Hashes never leave the server.
- Every data request needs a session token issued at sign-in. Tokens last 6 hours.
- Five wrong passwords lock that account for 15 minutes.
- Roles, edit approvals and project boundaries are enforced on the server, not only hidden in the page.

## Setup

1. Create one Google Sheet per project with the tabs listed above. Copying an existing project's sheet and clearing its data rows is the quickest way.
2. Create an Apps Script project and add `Code.gs`, `index.html` and `appsscript.json`. The HTML file must be named `index`.
3. In `Code.gs`, set:
   - `PROJECTS`: the name and Google Sheet ID of each project
   - `ADMIN_EMAIL`: the owner's Google account
   - `ALLOWED_DOMAIN`: the company email domain
4. Deploy as a web app: **Execute as: Me**, **Who has access: Anyone within your domain**.
5. Open the web app link as the owner. You are signed in as admin automatically.
6. In the app, open **Database** under each project and click **Rebuild sheet tabs** once. This creates the App Data and Users tabs.

### Deploying with clasp

[clasp](https://github.com/google/clasp) pushes the files to Apps Script from the command line.

```powershell
npm install -g @google/clasp
clasp login
clasp clone <script-id>
```

Copy `Code.gs` (as `Code.js`), `index.html` and `appsscript.json` into the cloned folder, then:

```powershell
clasp push --force
clasp create-deployment --deploymentId <deployment-id> --description "what changed"
```

Updating the existing deployment keeps the web app link the same. `.clasp.json` holds the script ID and is not committed.

## Troubleshooting

- **Old screens after an update:** press Ctrl+Shift+R to reload without the browser's saved copy, then check the version under the company name in the sidebar.
- **"Not set up yet":** the owner hasn't opened the app yet. The owner's first visit creates the admin account.
- **Excel import fails to load:** the Excel reader (SheetJS) loads from cdnjs.cloudflare.com. If that's blocked, save the file as CSV and upload that.
