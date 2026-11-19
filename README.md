# SeJasmin Invitation Card + Google Sheets RSVP

A floral digital invitation for Majlis Hi-Tea SeJasmin with an RSVP form that records responses in a private Google Sheet.

## Files
- `index.html` — responsive invitation page and RSVP form.
- `Code.gs` — Google Apps Script endpoint that stores RSVP submissions.

## 1. Create the Google Sheet
1. Create a Google Sheet.
2. Copy the spreadsheet ID from its URL: the value between `/d/` and `/edit`.
3. Keep the spreadsheet private.

## 2. Deploy the backend
1. Open https://script.google.com/ and create a project.
2. Paste the contents of `Code.gs`.
3. Replace `PASTE_YOUR_GOOGLE_SHEET_ID_HERE` with your Sheet ID.
4. Select **Deploy → New deployment → Web app**.
5. Set **Execute as: Me** and **Who has access: Anyone**.
6. Authorize and deploy. Copy the URL ending in `/exec`.

The endpoint is public so guests can submit. Keep the Sheet private and do not store confidential information in this simple RSVP form.

## 3. Connect the invitation
In `index.html`, replace `PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE` with the deployed Apps Script `/exec` URL.

## 4. Publish with GitHub Pages
In repository **Settings → Pages**, choose **Deploy from a branch**, select `main` and `/(root)`, then save. The site should publish at:
https://darknight1028.github.io/invitation-card-rsvp/
**Admin dashboard:** https://darknight1028.github.io/invitation-card-rsvp/admin.html

Release Version: 1.5v

## RSVP columns
The script creates a tab named `RSVP` with:
- Timestamp
- Guest Name
- Attendance (`Hadir` or `Tidak hadir`)

To download an Excel copy, use **File → Download → Microsoft Excel (.xlsx)** in Google Sheets.
