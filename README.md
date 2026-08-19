# Hurix Digital - Attendance & Productivity Hub

Enterprise-grade Apps Script Web Application for tracking daily shift clock-ins/outs, task productivity, break management, and manager analytics for 80+ employees.

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

- **`Code.gs`**: Complete Apps Script backend logic for authentication, shift state machine, task logging, and admin metrics.
- **`DriveHandler.gs`**: Handles proof document uploads directly to Shared Admin Google Drive Folder `1pvJQ9q6FuqRObkebq8GizkWOQ_QbX2RH`.
- **`Index.html`**: Main HTML layout with Bootstrap 5, FontAwesome 6, Google SSO auto-detection, and optional email & password sign-in card.
- **`UserView.html`**: Employee portal view with live shift timer, clock in/out controls, daily scorecards, and task logging form.
- **`AdminView.html`**: Manager dashboard with KPI scorecards, live employee status monitor, date/team/mode filters, and CSV exporter.
- **`Styles.html`**: Premium Hurix design system CSS tokens and components.
- **`appsscript.json`**: Manifest configuration specifying V8 runtime and `USER_DEPLOYING` execution mode.
