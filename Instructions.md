# Job Safety Observation (JSO) Tracker Web App

This project creates a mobile-friendly web application for Leaders to log Job Safety Observations on their tablets or phones. It automatically pulls Leader/Associate relationships from your existing Attendance Tracker sheet, saves the observations to a new Google Sheet, and emails Managers when an associate reaches a multiple of 3 Positive or Negative JSOs.

## Installation Instructions

### Step 1: Create the Project
1. Open Google Drive and go to a folder where you want to save this script.
2. Click **New > More > Google Apps Script**.
3. Name the project "JSO Tracker App" (or similar).

### Step 2: Add the Code
1. In the Apps Script editor, you will see a default file named `Code.gs`.
2. Delete all the default code in `Code.gs` and paste the entire contents of the provided `Code.gs` file.
3. Next, click the `+` icon next to "Files" in the left sidebar.
4. Select **HTML** and name the file exactly `Index` (with a capital I, don't type `.html`, it will add it automatically).
5. Delete the default HTML code and paste the entire contents of the provided `Index.html` file.
6. Click the Save icon (or press Ctrl+S / Cmd+S).

### Step 3: Initialize the Database
This script needs a place to store the JSO records. I have written a script to create this for you automatically.

1. Go back to `Code.gs`.
2. At the top of the editor, look for the function dropdown (it usually says `onOpen`). Change it to `setupDatabase`.
3. Click the **Run** button.
4. Google will ask for authorization. Click **Review Permissions**, select your account, click **Advanced**, and then click **Go to JSO Tracker App (unsafe)**. Allow the permissions.
5. The script will run, and it will log two things at the bottom of the screen in the Execution Log:
   - `New JSO Database ID: [SOME_LONG_ID_HERE]`
   - `New JSO Database URL: [LINK]`
6. **IMPORTANT:** Copy that long Database ID from the log.
7. Scroll up to Line 9 of `Code.gs` where it says:
   `const JSO_DATA_SHEET_ID = 'YOUR_NEW_JSO_SHEET_ID_HERE';`
8. Replace `'YOUR_NEW_JSO_SHEET_ID_HERE'` with the actual ID you just copied. Keep the single quotes around it.
9. Click **Save** again.

*Note: You can click the URL in the execution log to open your new database spreadsheet and bookmark it.*

### Step 4: Deploy the Web App
Now you need to publish the app so your leaders can access it on their tablets.

1. In the top right corner of the Apps Script editor, click the blue **Deploy** button.
2. Select **New deployment**.
3. Click the gear icon ⚙️ next to "Select type" and choose **Web app**.
4. Fill out the form:
   - **Description:** Version 1 (or whatever you like)
   - **Execute as:** Me (your email address)
   - **Who has access:** Anyone within [Your Organization Name] (or "Anyone" depending on your Google Workspace settings).
5. Click **Deploy**.
6. Google will provide you with a Web app URL. Click **Copy**.

### Step 5: Test and Share!
1. Paste that URL into your browser (or email it to yourself and open it on your tablet).
2. You should see a loading spinner briefly while it pulls the Leader/Associate lists from your Attendance sheet, and then the form will appear.
3. Select a Leader, select an Associate, toggle Positive or Negative, add a note, and hit Submit.
4. Check your new "JSO Tracking Data" Google Sheet to confirm the record was saved!
5. Bookmark the Web App URL on your leaders' tablets for easy access on the floor.