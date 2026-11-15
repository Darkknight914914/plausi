# Plausi browser extension

Plausi checks spreadsheets and web tables against business rules. It works in Chrome, Edge, Brave, Opera and other Chromium browsers. Everything runs on your computer; files and tables are never uploaded.

**What it adds to your browser**
- **Toolbar button**: opens Plausi in a new tab. Drop any .xlsx, .xlsm, .xlsb, .xls, .ods or .csv file on it.
- **Right-click a web page → "Check tables on this page with Plausi"**: reads every table or data grid on the page (reports in web apps, admin panels, bank portals) and checks them. Each table becomes a sheet.
- **Right-click a link → "Check linked file with Plausi"**: downloads the Excel/CSV file behind the link and checks it, using your sign-in for that site. Plausi asks permission once per site.

## Install (about 3 minutes)

1. Unzip `plausi.zip`. Keep the `extension` folder somewhere permanent (e.g. Documents\Plausi\extension). The browser loads it from there.
2. **Add the Excel reader** (one file):
   open https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js, press Ctrl+S (Cmd+S on Mac), and save it into `extension\lib` as `xlsx.full.min.js`.
   Skip this and Plausi still checks CSV files and web tables.
3. Open `chrome://extensions` (Edge: `edge://extensions`).
4. Switch on **Developer mode** (top right; in Edge, bottom left).
5. Click **Load unpacked** and choose the `extension` folder.
6. Click the puzzle icon in the toolbar and pin **Plausi**.

Test it: click the Plausi button. You should see the sample payouts with 10 errors and 3 warnings. Then right-click any page with a table, such as a Wikipedia list, and choose **Check tables on this page with Plausi**.

**After an update**: replace the folder's files and click the ↻ reload icon on Plausi's card in `chrome://extensions`.

## Publishing to the Chrome Web Store (later)

1. Register at https://chrome.google.com/webstore/devconsole (one-time USD 5 fee).
2. Zip the *contents* of the `extension` folder, including `lib/xlsx.full.min.js`.
3. Upload it, add screenshots (1280×800), a privacy statement ("no data leaves the device") and justification for each permission:
   `contextMenus` (right-click items), `scripting` + `activeTab` (read tables on the page you right-click), `storage` (pass the table to the Plausi tab), optional host access (download a linked file you choose).
4. Review usually takes a few days. Microsoft Edge Add-ons accepts the same zip for free.
