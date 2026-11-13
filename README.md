# Thyroid patient registry

Mobile web app (Google Apps Script) for recording head & neck / thyroid patients into a Google Sheet.

- `Code.gs` – server side (sheet access, allow-list check)
- `index.html` – mobile UI

Deploy: open the Google Sheet > Extensions > Apps Script, paste both files, then
Deploy > Web app (Execute as: User accessing the web app, Access: Anyone with Google account).
Add team emails in the `Users` sheet tab.

Do not commit patient data to this repository.
