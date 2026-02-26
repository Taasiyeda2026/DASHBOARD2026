# Attendance System - תעשיידע

## Overview
מערכת ניהול נוכחות למדריכים ומנהלי צוות - דיווחים, מעקב ודוחות בזמן אמת.
PWA (Progressive Web App) סטטי עם ממשק בעברית (RTL).

## Architecture
- **Frontend**: Single-page static HTML app (`index.html`) with embedded CSS and JS
- **Backend**: Azure Logic App connected via Azure Function proxy
- **Data**: SharePoint Online lists (Employees table, Attendance table, MonthlyApprovals table)
- **API URL**: Configured in `CONFIG.LOGIC_APP_URL` pointing to Azure Function gateway

## Project Structure
- `index.html` — Main application (login, dashboard, records, reports, approvals)
- `manifest.json` — PWA manifest
- `sw.js` — Service Worker for offline support
- `icons/` — PWA icons
- `data/` — Static configuration data (schools, Logic App definitions)
- `screenshots/` — App screenshots for PWA manifest

## API Actions
All API calls go through `apiRequest()` function to Azure Function:
- `auth` — Login with employeeId + personalCode
- `submit` / `updaterecord` / `deleterecord` — CRUD for attendance records
- `getemployeerecords` / `getsummary` / `exportdata` — Data retrieval
- `getallemployees` — Admin: list all employees
- `setinstructorapproval` / `setmanagerapproval` / `cancelinstructorapproval` — Approval workflow
- `getmonthlyapproval` / `getteamapprovalstatus` — Approval status
- `resetteammonth` — Admin: reset month
- `uploadfile` — File attachments

## Security Notes
- Auth response is sanitized client-side via `sanitizeEmployeeData()` to strip SharePoint metadata
- No sensitive fields (personalCode, Author, Editor, etc.) are stored in localStorage
- localStorage stores only: `currentUser` (JSON), `employeeName`, `team`
- All debug console.log statements exposing sensitive data have been removed
- All API calls routed through `apiRequest()` — no direct fetch to Logic App URL

## Dev Server
- Uses `npx serve . -l 5000` to serve static files on port 5000

## User Preferences
- Respond in Hebrew
