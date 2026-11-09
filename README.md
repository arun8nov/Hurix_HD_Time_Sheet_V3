# Hurix Digital - Attendance & Productivity Hub

Enterprise-grade Apps Script Web Application for tracking daily shift clock-ins/outs, simplified work logging, and consolidated manager analytics for 80+ employees.

---

## Key Features & Business Rules

1. **Mandatory Shift Selection**:
   - Upon logging in, associates select their designated shift: **General (8h)**, **1st (7h)**, **Afternoon (7h)**, or **2nd (7h)**.
   - Selected shifts are persisted with attendance records and directly drive headcount capacity and production hour benchmarks.
2. **Independent Leave Hours Tracking**:
   - Leave hours (Full day: 8h/7h based on shift; Half day: 4h/3.5h; Permission: exact approved hours) are tracked independently from worked hours and production hours across backend logs, KPI scorecards, roster views, and CSV exports.
3. **Automatic Fixed 1-Hour Break Policy**:
   - Break-time logging is removed; users do not manually log or classify breaks.
   - A fixed 1-hour break is automatically accounted for per worked day and deducted from gross shift time ($$\text{Net Shift Hours} = \max(0, \text{Gross Shift Hours} - 1.0)$$).
4. **Simplified Work Logs (AHT Removed)**:
   - AHT metrics, benchmarks, and unit tracking have been completely removed.
   - Work logs capture: **Client**, **Project** (filtered by Client), **Activity** (`Production`, `Client Meeting`, `Internal Meeting`, `Training (Material Go-Through)`, `Idle`, `Down Time`), **Hours Worked**, and **Remarks**.
5. **Executive Productivity & Headcount Summary Table**:
   - Structured and styled to match the executive spreadsheet benchmark (`#92cddc` teal header, crisp borders, warm yellow `#fff2cc` total row).
   - Columns: `Month` (or Period), `Project Name`, `Prod Head Count`, `Support Head Count`, `Total Available Production Hrs`, `Active production Hours`, `Training (Client + Internal)`, `Idle`, `Down Time`, `Meeting (Client + Internal)`, `Leave`, `Production Utilization %`, `Non Productive Hours %`.
   - Production benchmark formula:
     - General: $\text{Average Prod HC} \times 8 \text{ hours} \times \text{Days}$
     - 1st, Afternoon, 2nd: $\text{Average Prod HC} \times 7 \text{ hours} \times \text{Days}$
   - Production Utilization %: $(\text{Active Production Hours} / \text{Total Available Production Hours}) \times 100$
   - Non-Productive Hours %: $((\text{Training} + \text{Idle} + \text{Down Time} + \text{Meeting}) / \text{Total Available Production Hours}) \times 100$
6. **Support Headcount & Production Hours Benchmark Exclusion**:
   - `Support` included as a standard activity type across work logs and reporting.
   - Support Head Count is calculated from records where `activity_type === 'Support'`.
   - Support Head Count does **not** count toward Production Hours.
   - Total Available Production Hours are calculated using **only** employees assigned to `Production`, `Idle`, `Down Time`, `Client Meeting`, or `Internal Meeting`. `Support` and `Training` are strictly excluded from Production Hours.
7. **Average Headcount Calculation Across All Days**:
   - Average headcount is calculated across all calendar days in the selected date range.
   - Days with 0 employee logins contribute 0 headcount and are **not omitted** from the average divisor.
8. **Week Off (WO) & Holiday Status Management**:
   - `Week Off (WO)` and `Holiday (HOLIDAY)` status options added to employee schedules.
   - Managers can set, toggle, or clear a resource's `WO` or `Holiday` status with a single click directly from the Associate Performance Roster.
   - Resources marked `WO` or `Holiday` are strictly excluded from Production Head Count and Available Production Hours for the affected date.
9. **Single-Row Compact KPI Scorecards**:
   - Bulletproof flexbox (`flex-row flex-nowrap`) layout ensures all 8 KPI scorecards fit neatly in a single horizontal row across desktop screens without wrapping into vertical columns.
10. **Consolidated Manager Dashboard**:
   - Single unified dashboard providing **Day**, **Week**, and **Month** views.
   - Filters by Date, Day/Week/Month, Client, Project, Team, Shift, and Associate search.
   - Dual CSV exports: Consolidated Associate Roster CSV and Executive Productivity Summary CSV.

---

## Deployment Instructions (Fixes Permission Denied Error)

### Step 1: Open Apps Script Editor
1. Open your Google Sheet database (`User_Master`, `Attendance_Shift_Logs`, etc.).
2. Click **Extensions** $\rightarrow$ **Apps Script**.

### Step 2: Configure Web App Deployment Settings
1. Click the blue **Deploy** button (top right) $\rightarrow$ **Manage deployments**.
2. Click the **Edit (Pencil Icon)** on your active deployment.
3. Configure the settings exactly as follows:
   * **Execute as**: **`Me (your email address)`** *(CRITICAL: Runs script under admin permissions so employees can sign in without needing direct edit rights to the raw Google Sheet)*
   * **Who has access**: **`Anyone`** *(or `Anyone with Google Account` / `Anyone within Hurix Digital`)*
4. Click **Deploy**.
5. Select **New version** when saving the deployment.

---

## Technical File Architecture

- **`Code.gs`**: Complete Apps Script backend logic for authentication, shift state machine, 1-hour break auto-deduction, work logging, and Day/Week/Month dashboard analytics.
- **`DriveHandler.gs`**: Handles proof document uploads directly to Shared Admin Google Drive Folder.
- **`Index.html`**: Core client-side controller, UI state manager, cascading dropdowns, modal interactions, and CSV exporter.
- **`UserView.html`**: Employee portal view with live shift timer, clock in/out controls, shift status ribbon, and simplified work log form.
- **`AdminView.html`**: Consolidated manager dashboard layout with period switcher, KPI scorecards, activity progress meters, client/project breakdown tables, and associate roster.
- **`Styles.html`**: Premium Hurix design system CSS tokens and components.
- **`appsscript.json`**: Manifest configuration specifying V8 runtime and `USER_DEPLOYING` execution mode.
