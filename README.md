Canvas Tools for Google Sheets
A Google Sheets integration with Canvas LMS that allows educators to fetch student data, download grades, and upload grades directly between Google Sheets and Canvas.
Features

Fetch student data directly from Canvas into your spreadsheet
Download grades for individual assignments or the complete gradebook
Upload grades to Canvas individually or in bulk
Progress tracking with helpful toast notifications
Configurable setup through the spreadsheet (no code editing required)

Installation
Option 1: Use the Template (Easiest)

Open the template spreadsheet
A copy will be created in your Google Drive with all code included
Continue to the Setup section below

Option 2: Add to Your Existing Sheet

Create or open a Google Sheet
Go to Extensions > Apps Script
Create the following files and copy the code from this repository:

CanvasTools.gs
UserFetch.gs
GradeFetch.gs
GradeUpload.gs
ToastUtilities.gs


Save the project
Refresh your spreadsheet
You should now see a "Canvas Tools" menu
Select "Canvas Tools > Setup Spreadsheet" from the menu

Setup

After installation, select "Canvas Tools > Setup Spreadsheet" from the menu
Enter your Canvas Course ID in cell B3

Find this in the URL when viewing your course: https://canvas.yourinstitution.edu/courses/12345


Enter your Canvas Domain in cell B4 (e.g., https://canvas.chapman.edu)
For the API Key, you have two options:

Enter it directly in cell B5 (will be stored securely afterward)
Leave B5 blank and you'll be prompted when needed



Generating a Canvas API Key

Log into Canvas
Go to Account > Settings
Scroll to "Approved Integrations"
Click "New Access Token"
Enter a purpose (e.g., "Google Sheets Integration")
Set an expiration date if desired
Click "Generate Token"
IMPORTANT: Copy the token immediately - you cannot view it again!

Usage
Fetching Student Data

Ensure your Course ID and Canvas Domain are set
Click "Canvas Tools > Fetch Course Users"
Student data will populate in columns A-D starting at row 7

Fetching Assignment Grades

In row 6 of column E (or any empty column), enter the Canvas Assignment ID

Find the Assignment ID in the URL when viewing the assignment:
Example: https://canvas.chapman.edu/courses/12345/assignments/67890 → ID is 67890


Click "Canvas Tools > Fetch Assignment Grades"
Enter the column letter when prompted (e.g., "E")
Grades will populate in that column aligned with student rows

Fetching the Complete Gradebook

Click "Canvas Tools > Fetch Complete Gradebook"
The spreadsheet will be populated with:

All assignments from Canvas (column E onwards)
Points possible for each assignment (row 3)
Average scores and percentages (rows 4-5)
Assignment IDs (row 6)
All student grades



Uploading Individual Assignment Grades

Ensure the column contains the Assignment ID in row 6
Enter grades in the column aligned with student rows
Click "Canvas Tools > Upload Assignment Grades to Canvas"
When prompted, enter the column letter
A summary will appear when complete

Uploading a Range of Grades

Ensure Assignment IDs are in row 6 for all columns you want to upload
Enter grades in those columns
Click "Canvas Tools > Upload Grade Range to Canvas"
Enter the starting and ending column letters when prompted
A summary will appear when complete

Uploading the Complete Gradebook

Click "Canvas Tools > Upload Complete Gradebook to Canvas"
Confirm the action when prompted
All grades in the sheet will be uploaded to Canvas
A summary will appear when complete

Security Considerations

Your Canvas API key provides access to your Canvas account
The key is stored securely in the Script Properties and not visible to others
When entered in cell B5, it will be read once and then stored securely
Before sharing your spreadsheet, use "Canvas Tools > Clear API Key" to remove it

Troubleshooting
ProblemPossible Solution"API Key not found"Enter your Canvas API key in cell B5 or when prompted"Course ID not found"Make sure you entered your Canvas Course ID in cell B3"Canvas Domain missing"Enter your institution's Canvas URL in cell B4"Assignment ID not found"Ensure row 6 contains the correct Assignment IDNo grades appearVerify students have submissions in CanvasUpload errorsCheck you have permission to edit grades in Canvas
License
This project is licensed under the MIT License - see the LICENSE file for details.
Author
Your Name
Contributing
Contributions are welcome! Please feel free to submit a Pull Request.
