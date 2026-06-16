# Budget Tap

Budget Tap is a small phone-friendly expense logger built for an iPhone Back Tap workflow.

The app lives in `app/` and can be hosted with GitHub Pages. It saves expenses locally first, then sends them to a Google Apps Script endpoint that appends rows to Google Sheets.

## Publish With GitHub Pages

1. Create a GitHub repository.
2. Push this project to the repository.
3. In GitHub, go to Settings > Pages.
4. Set Source to GitHub Actions.
5. The included workflow publishes the `app/` folder.

After deployment, use the Pages URL in your iPhone Shortcut:

```text
https://YOUR_USERNAME.github.io/YOUR_REPO_NAME/
```

## iPhone Back Tap

1. Open Shortcuts.
2. Create a shortcut named `New Expense`.
3. Add the action `Open URLs`.
4. Paste the GitHub Pages URL.
5. Go to Settings > Accessibility > Touch > Back Tap.
6. Set Double Tap to the `New Expense` shortcut.

## Google Sheets Sync

Paste `google-apps-script/Code.gs` into Apps Script, set `SPREADSHEET_ID`, deploy as a web app, then paste the web app URL into Budget Tap settings.
