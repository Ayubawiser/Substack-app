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


## Firebase cloud setup (SubTrack project: subtrack-app1)

1. In Firebase Console, open **Authentication → Sign-in method** and enable **Anonymous**.
2. In **Firestore Database → Rules**, publish the rules from `firestore.rules` in this project. These rules restrict each signed-in anonymous user to their own workspace.
3. Install and run: `npm install` then `npm run dev`.
4. The app signs in anonymously and stores its workspace at `users/{uid}/appData/workspace`. Changes are saved to Firestore automatically.

**Important limitations:** anonymous sign-in creates a separate workspace identity per browser/device. It does not yet provide email/password accounts, team sharing, or cross-device account recovery. Add permanent sign-in before relying on this for customer data. The current prototype stores business contact and document metadata only; it does not upload document files. Do not use real sensitive customer records until authentication, privacy, backup and production security have been reviewed.

The Firebase web API key in `src/firebase.ts` is a public client configuration value, not a service-account secret. Security is enforced by Firebase Authentication and Firestore rules, not by hiding this key.
