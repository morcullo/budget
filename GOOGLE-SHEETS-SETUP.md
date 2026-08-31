# Google Sheets sync setup

The app can now sync its budget, categories, transactions, archive, and frequent transactions to one Google Sheet. It keeps localStorage as a fallback.

## 1. Create the sheet
Create a Google Sheet for the app. The Apps Script backend will create a `BudgetData` tab automatically.

## 2. Add the backend
In the sheet, open **Extensions → Apps Script**. Replace the starter code with the contents of `google-apps-script.gs`.

Optional: set `SECRET` to a random string. If you do, put the same value in `GOOGLE_SHEETS_TOKEN` in `index.html`.

## 3. Deploy
Choose **Deploy → New deployment → Web app**.

- Execute as: **Me**
- Who has access: **Anyone**

Copy the generated URL ending in `/exec`.

## 4. Connect the app
Open `google-sheets-config.js` and replace the empty `window.GOOGLE_SHEETS_URL` value with your `/exec` URL. If you set a secret, put it in `window.GOOGLE_SHEETS_TOKEN`. Do not edit `index.html` for this step.

Then redeploy the app to Netlify.

## Important
This is designed for a personal/small private budget app. A Google Apps Script web app set to **Anyone** is an internet-facing endpoint. The optional shared secret adds a basic gate, but the secret is necessarily present in the browser code, so it should not be treated as strong security for sensitive or multi-user applications.

## What changed in v2
When a Google Sheets URL is configured, Google Sheets is now the source of truth. The app does not write to localStorage during normal saves. localStorage is used only if the cloud endpoint is missing or fails, so the app can still function offline. Check the browser console for the exact cloud error if syncing fails.

## iOS Shortcut: Add a transaction

The Apps Script web app also accepts a direct POST for adding a transaction.
Use the same `/exec` URL as the Budget app.

Request body:

```json
{
  "action": "addTransaction",
  "amount": 12.50,
  "category": "Food",
  "description": "Coffee",
  "date": "2026-08-31"
}
```

In Apple Shortcuts, use **Get Contents of URL**, set **Method** to `POST`,
and set the **Request Body** to JSON with those fields. If you configured a
secret, also include a `token` field containing the same secret.

A positive amount is treated as an expense and stored as a negative amount,
which matches the Budget app's existing transaction format. The endpoint
returns JSON containing `ok: true` and the newly created `transaction`.
