# SK Tech Admin Interview Portal

A Google Apps Script web app for displaying the live SK Tech interview schedule from Google Sheets.

The portal reproduces the supplied interview-panel design and automatically classifies rows as:

- **Today Interview** — green
- **Time Conflict** — red
- **Upcoming Interview** — purple

Only interviews scheduled for today or later are displayed. Rows are sorted by interview date and start time.

## Project files

- `Code.gs` — spreadsheet access, schedule preparation, conflict detection, and registration email automation
- `Html.html` — responsive interview-panel interface
- `appsscript.json` — Apps Script runtime and time-zone configuration

## Data source

The app reads the first nine columns of the `Pannel List` tab in spreadsheet:

`1VQF3CI8MTBz4vAC4R-13_XEfY774h1Q7vxhBsSxA2tU`

Expected columns:

1. Sk Tech Register ID
2. Full Name
3. Round
4. Technologies Required
5. Interview Company
6. Interview Date
7. Interview Time (From)
8. Interview Time (To)
9. Batch

The server validates the header row before returning data, so a renamed or missing column produces a useful error instead of a misaligned table.

## Deploy to Google Apps Script

1. Open the Apps Script project attached to the spreadsheet, or create a standalone Apps Script project.
2. Replace its `Code.gs` contents with this repository's `Code.gs`.
3. Add an HTML file named `Html` and paste in `Html.html`.
4. In **Project Settings**, enable viewing `appsscript.json`, then replace the manifest with this repository's `appsscript.json`.
5. Click **Deploy → New deployment → Web app**.
6. Set **Execute as** to `Me`.
7. Choose the narrowest access level that allows the intended interview team to view the schedule.
8. Deploy and open the generated `/exec` URL.

The interface reads the latest sheet values whenever the web app is opened or refreshed.

## Enable registration emails

`onFormSubmit(e)` must run as an installable spreadsheet trigger because it uses Gmail authorization.

1. In Apps Script, open **Triggers**.
2. Add a trigger for `onFormSubmit`.
3. Choose **From spreadsheet** as the event source.
4. Choose **On form submit** as the event type.
5. Authorize the requested Sheets and Gmail permissions.

The handler ignores submissions from every tab except `Form Responses 3_Candiate details`.

The sender address `support@mysktech.com` must be configured as a Gmail alias for the account executing the script. If it is not an approved alias, change or remove the `from` option in `Code.gs`.

## Local UI preview

`Html.html` contains a small, non-production preview dataset that is used only when the file is opened outside Apps Script. The deployed Apps Script page always calls `getInterviewData()` and renders live spreadsheet data.

From this directory, start any static file server and open `Html.html`. For example:

```powershell
python -m http.server 4173
```

Then visit `http://127.0.0.1:4173/Html.html`.

## Tests

The test suite runs the Apps Script server code in a mocked V8 context and verifies the sheet-header contract, date/time formatting, row filtering and sorting, conflict colors, and sequential SK Tech IDs.

```powershell
npm test
```

## Privacy

The schedule contains candidate and interview information. Keep the Apps Script web-app audience limited to the people who need access. The implementation intentionally keeps the default Apps Script frame protection instead of enabling `ALLOWALL`.
