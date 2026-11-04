# SubTrack

A mobile-friendly subcontractor compliance tracker prototype for small general contractors.

## Current prototype features
- Dashboard summary cards and clearance overview
- Subcontractor directory with add, edit and delete
- Project creation
- Document status and expiry tracking
- Follow-up queue with contact logging
- CSV export
- Browser local storage for prototype data

## Run locally
Requires Node.js 18 or newer.

```bash
npm install
npm run dev
```

## Build
```bash
npm run build
```

## Important prototype limitations
- All data is stored in the current browser using localStorage.
- There is no login, shared workspace, cloud database or server-side authorization yet.
- Sample records are fictional.
- Do not enter real customer information or upload sensitive documents.
- Firebase integration and production Firestore security rules are the next development stage.

## Technology
React, TypeScript, Vite and Lucide icons.


## Spreadsheet import and export

- Export Reports creates an Excel workbook with four tabs: Subcontractors, Projects, Documents and Follow-ups.
- CSV export downloads a separate CSV file for each section.
- Import Data accepts `.xlsx`, `.xls` and `.csv` files.
- Google Sheets users can download a spreadsheet as Microsoft Excel (`.xlsx`) or CSV, then import it into SubTrack. This is file-based import, not live Google Sheets synchronisation.
- Import can merge matching company/project records or replace the current local demo data.
- Import/export currently runs in the browser and stores records in local browser storage. Firebase cloud sync and user accounts are not connected yet.
