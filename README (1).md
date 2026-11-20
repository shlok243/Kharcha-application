# Kharcha

A tiny phone-friendly expense logger that writes straight into your own Google Sheet.

Type an amount, tap a tag, tap **Save**. The date and time are added automatically, and the month's totals are calculated for you. No app store, no third-party servers: it is a Google Apps Script web app that runs inside your own Google account.

## Features

- **Fast entry:** big amount box, quick tags (Food, Chai/Coffee, Petrol, Bus, ...), and a Cash / GPay / FamPay switch that remembers your last choice.
- **Spent or Received:** log money going out or coming in.
- **Automatic date and time:** every row is stamped for you.
- **Live totals:** "Spent this month" and "Today" are shown on the main screen, plus your last few entries.
- **One-step undo:** removes only what the app just wrote.
- **Month-wise layout:** all months live in one tab, one block per month (yellow title, header row, entries, totals row).
- **Automatic month close:** on the 1st of each month the app writes the totals row for the previous month and emails you a summary, including the Cash / GPay / FamPay split.
- **Works as a home-screen app:** add it to your phone's home screen from Chrome.

## How it looks in the sheet

| Date | Received | Spent | Specification | Person/mode | Time |
|---|---|---|---|---|---|
| 11/20/2026 |  | 120 | chai | Cash | 12:20 PM |
| 11/20/2026 | 250 |  | friend returned money | FamPay | 6:05 PM |

When a new month starts, the app closes the old block with a totals row (Received, Spent, and net spent), leaves a blank line, and starts a new block with its own title and header.

## Setup (about 5 minutes)

1. Create or open the Google Sheet you want to use. Copy its ID from the URL: `https://docs.google.com/spreadsheets/d/<ID>/edit`.
2. In the sheet, open **Extensions → Apps Script**.
3. Paste `Code.gs` into the editor and set `SHEET_ID` at the top to your spreadsheet ID. Change `TZ` to your timezone if needed.
4. Add an HTML file named exactly `Index` and paste `Index.html` into it.
5. Run `setupMonthlyReport` once from the function picker and allow the permissions. This sets up the monthly close and email.
6. **Deploy → New deployment → Web app.** Set *Execute as* to **Me** and *Who has access* to **Only myself**. Copy the web app URL.
7. On your phone, open the URL in Chrome and choose **Add to Home screen**.

To update the code later: **Deploy → Manage deployments → Edit → New version → Deploy**. The URL stays the same.

## Good to know

- The app writes to the tab named `Sheet1` (or the first tab if there is no `Sheet1`).
- It finds the current month's block from the dates in your entries, not from the month title, so you can rename titles freely.
- Total rows are recognised as rows with numbers in the Received/Spent/Specification columns and nothing in Person/mode.
- Undo only works on the most recent entry, and only while nothing was added to the sheet after it.
- Column F ("Time") is added by the app.

## Privacy and safety

- Everything runs in your own Google account. Nothing is sent to any other server.
- Keep the deployment set to **Only myself**.
- **Do not commit your real spreadsheet ID** to a public repository. Keep `YOUR_SPREADSHEET_ID_HERE` in the shared code and put your real ID only in your own Apps Script copy.
- Check your sheet's sharing settings. Anyone with the sheet link can see (or edit, depending on the setting) your spending if it is shared publicly.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Backend: saves entries, builds month blocks, totals, undo, monthly close and email |
| `Index.html` | The phone screen (HTML, CSS and JavaScript in one file) |
| `appsscript.json` | Optional manifest (timezone and web app access) if you use `clasp` |

## Ideas for later

- Category breakdown per month (Food vs Petrol vs Friends)
- Budget limit with a warning
- Edit or delete older entries from the app

## License

MIT. See `LICENSE`.
