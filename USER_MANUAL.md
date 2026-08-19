# Hurix Digital - Attendance & Productivity Hub
## Management Operations Manual & System Specification (V3 Simplified)

---

## 1. User Portal View Guide (`UserView.html`)

### 1.1 Simplified Shift & Activity Overview Ribbon
* **Automatic 1-Hour Default Break:** Employees do not need to pause or log breaks manually. A default 1.0-hour (60 mins) break deduction is automatically calculated for all completed shifts.
* **Shift Duration Timer:** Displays real-time elapsed shift timer with an indicator badge for the 1-hour automatic break.
* **Shift Clock Badges:** Clear visual indicators for Clock-In Time and Clock-Out Time.
* **Total Tasks Logged & Units Completed:** Clean counters providing immediate feedback on tasks and deliverables logged during the active shift.
* **Streamlined Experience:** Complex employee-facing KPI meters (leakage, speed variance, bench formulas) are hidden from the employee portal to ensure a clean, distraction-free logging interface.

### 1.2 Simplified Task Entry Form — Field-by-Field Dictionary
| Field Name | Input Type | Operational Purpose (Why It Exists) | User Guidelines (How to Enter) |
| :--- | :--- | :--- | :--- |
| **Active Client (`client`)** | Dropdown | Identifies client account for billing and project attribution. | Select assigned client (e.g. *iMerit, Innodata, Hurix*). |
| **Active Project (`project`)** | Dropdown | Links work directly to specific workflow or project. | Select assigned project under client dropdown. |
| **Activity Type (`activity_type`)** | Dropdown | Categorizes Production Work, QC, Training, Bench, Downtime. | Select relevant activity category. |
| **Hours Worked (`hours_worked`)** | Decimal Number | Direct entry of hours dedicated to this task batch. | Enter exact decimal hours worked (e.g. `1.5`, `2.25`, `4.0`). |
| **Units Done (`task_count`)** | Positive Integer | Quantity of deliverables or items completed in this duration. | Enter count of completed units (e.g. `15`, `50`). |
| **Task Status (`task_status`)** | Dropdown | Current state of task execution. | Default is `COMPLETED` (or `IN_PROGRESS` / `ON_HOLD`). |
| **Remarks / Notes (`remarks`)** | Text Area | Optional contextual audit notes, batch numbers, or tickets. | Enter batch ID, document title, or general notes. |

---

## 2. Unified Management Hub Guide (`AdminView.html`)

Instead of multiple disparate administrative portals, all supervisor, managerial, and executive oversight tools are unified into a single **Management Hub** with seamless sub-tab navigation.

### 2.1 Sub-Tab 1: Live Shift & Staff Monitor
* **Real-time Attendance Tracking:** Visual status indicators for Clocked In (`OPEN`), On Leave (`LEAVE_CLOSED`), and Pending (`NOT_CLOCKED_IN`).
* **1-Hour Break Normalization:** Standardized automatic 1-hour break deduction applied to floor shift durations.
* **Staff Performance Matrix:** Real-time visibility into Worked Hours, Earned Production Hours, Performance Leakage, Downtime/Idle, and Bench Training.
* **One-Click Actions:** Inspect associate task timelines and mark leaves on behalf of team members.
* **EOD Operations Reports:** 1-Click Markdown clipboard export, Executive Print PDF, and structured CSV reports.

### 2.2 Sub-Tab 2: Executive Analytics & Reports
* **Multi-Period Analytics:** Analyze floor delivery across Today, Weekly, Monthly, or Custom Date Ranges.
* **Client-Wise Executive Matrix:** Headcount utilization, billable production hours, and client shrinkage rates.
* **Team-Wise & PM Portfolios:** Delivery capacity across Annotation, SME, and Project Manager portfolios.
* **Associate Performance Roster:** Comprehensive historical appraisal table with benchmark pacing speed badges (⚡ $\ge 100\%$, ⏱️ $80-99\%$, 🐢 $<80\%$).

---

## 3. Bench vs Idle Governance & Master Formula Matrix

> [!IMPORTANT]
> **Core Management Rule:** **Bench Hours** and **Idle Hours** represent completely distinct operational states and must **NEVER** be mixed up in reporting or client billing:
> * **🟡 1st: Idle Hours (Client Project Downtime):** The associate is assigned to an active client production project, but is stalled due to client file/batch allocation delays, tool/server crashes, or network outages. It is **Client Project Downtime** with strictly **`0.00 hrs` Earned Hours**, **`0.00 hrs` Bench Hours**, and **`0.00 hrs` Performance Leakage** (100% exempt from leakage penalties).
> * **🟣 2nd: Bench Hours (No Client Project):** The associate has **no assigned client project** and is engaged in internal training, upskilling, floor management, or waiting for project onboarding. It is **100% Non-Billable** with strictly **`0.00 hrs` Earned Hours**, **`0.00 hrs` Idle Hours**, and **`0.00 hrs` Performance Leakage**.

### The Operational Work States Compared (Management Matrix)

| Operational State | Exact Definition | Billability Status | Earned Hours Rule | Performance Leakage Impact |
| :--- | :--- | :--- | :--- | :--- |
| **🟢 Earned Hours** | Active deliverable execution on a client project | **100% Billable** | $\frac{\text{Task Count} \times \text{AHT}}{60}$ | $\max(0, \text{Hours Worked} - \text{Earned})$ |
| **🟡 1st: Idle Hours** | **ON project** but stalled (Allocation Delay, Tool Crash, Outage) | **Client Downtime** | **`0.00 hrs`** *(Zero output credit)* | **`0.00 hrs`** *(100% Exempt)* |
| **🟣 2nd: Bench Hours** | **NO project assigned** (Internal Training, Upskilling, Floor Mgmt) | **Non-Billable Internal** | **`0.00 hrs`** *(Cannot bill client)* | **`0.00 hrs`** *(100% Exempt)* |

---

### Master Mathematical Formula Index

| Metric Name | Description | Exact Business Logic & Mathematical Formula |
| :--- | :--- | :--- |
| **Net Worked Shift Hours** | Shift duration minus default 1-hr break | $\max\left(0, (\text{Shift End} - \text{Shift Start}) - 1.0\text{ hr}\right)$ |
| **Actual Worked Task Hours** | Direct logged task duration | Directly entered by user in Task Form (`actual_worked_hours`) |
| **Billable Earned Hours** | Standard billable output generated | $\frac{\text{Units Completed} \times \text{AHT Benchmark Mins}}{60}$ *(Bench & Downtime = 0.00 hrs)* |
| **Technical Idle Hours** | Technical downtime on client tasks | $\sum \text{Task Hours}(\text{Activity Type = 'Downtime' / 'Waiting'})$ |
| **Non-Billable Bench Hours** | Time on Bench, Training, Floor Mgmt | $\sum \text{Task Hours}(\text{Project = 'Bench' / 'Training'})$ |
| **Performance Leakage** | Output gap between work and earned | $\max(0, \text{Hours Worked} - \text{Earned Hours})$ |
| **Speed Efficiency %** | Benchmark pacing velocity | $\text{round}\left(\frac{\text{Earned Hours}}{\text{Hours Worked}} \times 100\right)$ *(⚡ $\ge 100\%$, ⏱️ $80-99\%$, 🐢 $<80\%$)* |
| **Floor Shrinkage %** | Capacity lost to leave / absence | $\left(\frac{\text{PL} + \text{UL} + (0.5 \times \text{HL}) + \text{ABSENT}}{\text{Total Floor Resources}}\right) \times 100$ |

---

## 4. Database Schemas & Data Model

* **`User_Master`:** Employee ID, Name, Email, Password, App Role, Team, Work Mode, Client, Project, PM, Is Active.
* **`Attendance_Shift_Logs`:** Log ID, Emp ID, Date, Start Time, End Time, Total Break Mins (Default: 60), Total Shift Hours, Attendance Status, Shift Status, Permission Hours, Proof URL.
* **`Project_Task_Logs`:** Task ID, Emp ID, Client, Project, Activity Type, Start Time, End Time, Actual Worked Hours, Task Count, AHT Mins, Productive Earned Hours, Leakage Hours, Bench Hours, Idle Hours.
* **`Client_Project_Master`:** Client Name, Project Name, Is Billable, Default AHT Mins, Activity Categories.
