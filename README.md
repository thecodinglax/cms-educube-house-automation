
CMS HOUSE AUTO FILLER
=====================

What it does
------------
Reads an Excel file containing:
  Student Name
  House Name

and attempts to fill the matching House dropdown for each student in EduCube.

IMPORTANT
---------
This is an automation starter built around the page shown in your screenshot.
The exact HTML/controls on your EduCube account may differ. The program therefore
uses several detection methods, but you should test with Dry Run first.

Excel format
------------
Row 1 must contain headers similar to:
  Student Name | House Name

Valid houses:
  Hope
  Peace
  Unity
  Love

Installation
------------
1. Install Python 3.11+.
2. Open PowerShell in this folder.
3. Run:
       py -m pip install -r requirements.txt
4. Install the Playwright browser:
       py -m playwright install chromium

Run
---
       py app.py

Usage
-----
1. Select your Excel file.
2. Confirm the EduCube URL.
3. First tick "Dry run" and test.
4. Click START.
5. Chrome opens.
6. Log in manually if needed.
7. Open the page/list where the students are visible.
8. Click OK in the Ready popup.
9. Watch the log and progress.
10. Review the website before final submission.

Do not store your EduCube password in this application.
