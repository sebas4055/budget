# Budget Tap Setup

## What this is

`app/` is a phone-friendly expense entry app. It saves expenses locally first, then sends them to Google Sheets through a Google Apps Script web app.

## Google Sheets connection

1. Create or open your Google Sheet.
2. Go to Extensions > Apps Script.
3. Paste the contents of `google-apps-script/Code.gs`.
4. Deploy > New deployment > Web app.
5. Set "Execute as" to yourself.
6. Set "Who has access" to "Anyone with the link".
7. Copy the web app URL.
8. Open Budget Tap, tap settings, and paste the URL.

The app writes rows to a sheet named `Expense Log` by default.

## iPhone Back Tap flow

1. Host or open the app URL on your phone.
2. Open Shortcuts and create a shortcut named `New Expense`.
3. Add the action "Open URLs" and paste the app URL.
4. Go to Settings > Accessibility > Touch > Back Tap.
5. Pick Double Tap and choose the `New Expense` shortcut.

After that, double tapping the back of the phone opens the expense form.
