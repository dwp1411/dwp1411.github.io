# Standard Operating Procedure (SOP): Automating Payroll & Zoned Hours

**Purpose:** This document outlines the daily procedure for synchronizing associate payroll hours (actual clocked hours) with zoned hours (function hours assigned by supervisors) using the automated Google Sheets scripts.

**Prerequisites:**
1. You must have access to the master Calculator spreadsheet.
2. The daily 'Crate' and 'Temp' tabs must be populated with the current payroll data.

---

### Step 1: Set the Target Date
Before running any scripts, you must tell the system which day you are processing.

1. Open the master **Calculator** spreadsheet.
2. Go to the **Crate** tab.
3. In cell **A1**, enter the date you are processing (e.g., `3/18/2026`).

---

### Step 2: Submit Payroll Hours
This step takes the hours from your 'Crate' and 'Temp' tabs and automatically drops them into the 'Payroll Drop' tab of the correct weekly Timecard file.

1. Click on the **Payroll** menu at the top of the Google Sheet (next to "Help").
2. Click **Submit Hours**.
3. *What to expect:*
   - The script will find the correct Timecard file based on your date.
   - It will update the Regular and Overtime hours for every associate listed on the 'Payroll Drop' tab.
   - **Important:** If an associate's name is on the 'Payroll Drop' tab but they are *not* listed in your 'Crate' or 'Temp' tabs today, their hours will automatically be set to **0**.

---

### Step 3: Sync Zoned Hours
This step ensures that the total "Payroll Hours" (actual time clocked) perfectly matches the "Zoned Hours" (the departments they worked in).

1. Click on the **Payroll** menu again.
2. Click **Sync Zoned Hours**.
3. *What the script does behind the scenes:*
   - It finds the actual daily tab (e.g., `Monday_`) for the date you entered.
   - If an associate worked more than **0.5 hours**, it automatically deducts **30 minutes (0.5 hours)** for their break.
   - It then takes their total Payroll Hours and proportionally splits them across any departments (zones) they were assigned to that day.
   - If an associate's Payroll Hours are 0, or if the supervisor forgot to zone them entirely, the script skips them.

---

### Step 4: Handle Spelling/Name Alerts
At the end of the Zoned Hours Sync, the script cross-references the hours in the 'Payroll Drop' tab against the hours in the daily tab (e.g., `Monday_`).

1. If the sync is successful and everything matches perfectly, you will see a small green "Toast" notification pop up in the bottom-right corner saying **Sync Complete**. You are done!
2. **If an alert box pops up:** This means an associate had hours dropped into the 'Payroll Drop' tab, but those hours did *not* make it to the daily tab because their name is spelled differently between the two locations (e.g., "DELGADO, BRENDA" vs "DELGADO  BRENDA").
3. **How to fix:**
   - Read the names listed in the alert box.
   - Open the weekly Timecard file.
   - Fix the spelling of their name on the daily tab (e.g., `Monday_`) so it exactly matches the 'Overall Staff' list.
   - Go back to the Calculator and re-run **Payroll > Sync Zoned Hours**.

---

### Troubleshooting
* **"Error: Could not find a timecard file..."**: Ensure the date in A1 is correct, and that the weekly Timecard file is named correctly (e.g., `191 Week 3 3/16/26`).
* **"Date mismatch!"**: The date in cell B1 of the daily tab inside the Timecard does not match the date you entered in A1 of the Calculator. Correct the date and try again.