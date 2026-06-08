# School Dress Inventory & Sales Management System

This is a mobile-friendly inventory and sales web app for a school uniform store. Staff can enter stock and sales from an Android phone, while the owner can monitor everything directly in Google Sheets.

## Files

- `index.html` - app screen and forms
- `style.css` - mobile-first design
- `script.js` - frontend logic and report export
- `Code.gs` - Google Apps Script backend

## Google Sheets Setup

1. Create a new Google Sheet.
2. Rename it if you want, for example `School Dress Inventory Database`.
3. In the Google Sheet, open `Extensions` → `Apps Script`.
4. Delete any starter code.
5. Paste the full code from `Code.gs`.
6. Save the Apps Script project.
7. Run the function named `setupSpreadsheet` once.
8. Google will ask for permission. Allow it.

The script automatically creates these tabs:

1. Inventory
2. Sales
3. Stock Summary
4. Credit Sales
5. Pending Deliveries
6. Notifications
7. Dashboard Data

## Deploy Apps Script

1. In Apps Script, click `Deploy` → `New deployment`.
2. Choose type `Web app`.
3. Set `Execute as` to `Me`.
4. Set `Who has access` to `Anyone`.
5. Click `Deploy`.
6. Copy the Web App URL. It should end with `/exec`.

## Connect The Web App

1. Open `index.html` in your browser, or open the GitHub Pages link after hosting.
2. Paste the Apps Script Web App URL.
3. Tap `Save Connection`.
4. Start entering inventory and sales.

The URL is saved on that phone or laptop, so staff do not need to paste it every day.

## GitHub Pages Hosting

1. Create a GitHub repository.
2. Upload these files to the repository:
   - `index.html`
   - `style.css`
   - `script.js`
   - `Code.gs`
   - `README.md`
3. Open repository `Settings`.
4. Go to `Pages`.
5. Under `Build and deployment`, choose:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/root`
6. Save.
7. GitHub will show the live website link after a short wait.

## Daily Use

Inventory page:

- Add supplier stock received.
- Stock automatically increases in `Stock Summary`.

Sales page:

- Record student sale, payment, phone number, credit amount, and pending delivery item.
- Stock automatically decreases.
- Credit sales appear in `Credit Sales`.
- Pending items appear in `Pending Deliveries`.

Dashboard:

- Total inventory
- Total sales
- Outstanding credit
- Pending items
- Low stock items
- Out of stock items

Reports:

- Daily sales
- Weekly sales
- Monthly sales
- Inventory report
- Credit report
- Pending delivery report

Use `Export PDF` to print or save a PDF from the browser. Use `Export Excel` to download an Excel-compatible `.xls` report.

## Low Stock Alerts

When available stock is `5` or less, it is marked `LOW STOCK`.

When available stock is `0`, it is marked `OUT OF STOCK`.

Each low or zero stock update is also saved in the `Notifications` tab.

## Notes

- No paid service is required.
- All main data is stored in Google Sheets.
- The frontend is compatible with GitHub Pages.
- The app is designed for simple mobile use by a non-technical accountant.
