# Grade Tracking with Canvas API for Google Sheets

A Google Sheets integration with Canvas LMS that allows educators to fetch student data, download grades, and upload grades directly between Google Sheets and Canvas.

[![License](https://img.shields.io/github/license/jlouisbru/canvas-tools-for-sheets)](LICENSE)
[![Latest Release](https://img.shields.io/github/v/release/jlouisbru/canvas-tools-for-sheets)](https://github.com/jlouisbru/canvas-tools-for-sheets/releases/latest)
[![Platform](https://img.shields.io/badge/platform-Google%20Sheets-green)](https://docs.google.com/spreadsheets/d/18ZggFU-2xBdbl3pVPY3dXR-U5DYdroxvYaZJXGcvIPA/edit?usp=sharing)
[![Support on Ko-fi](https://img.shields.io/badge/Support-Ko--fi-ff5f5f)](https://ko-fi.com/louisfr)

## 🚀 Quick Start

[![Use Google Sheet Template](https://img.shields.io/badge/Use_Template-4285F4?style=for-the-badge&logo=google&logoColor=white)](https://docs.google.com/spreadsheets/d/18ZggFU-2xBdbl3pVPY3dXR-U5DYdroxvYaZJXGcvIPA/copy)

Click the button above to create your own copy of the Canvas Tools template

## ✨ Features

- **Fetch student data** directly from Canvas into your spreadsheet
- **Download grades** for individual assignments or the complete gradebook
- **Upload grades** to Canvas individually or in bulk
- **Progress tracking** with helpful toast notifications
- **Configurable setup** through the spreadsheet (no code editing required)

## 💾 Installation

### Use the Template

1. [Open the template spreadsheet](https://docs.google.com/spreadsheets/d/18ZggFU-2xBdbl3pVPY3dXR-U5DYdroxvYaZJXGcvIPA/edit?usp=sharing)
2. Copy the spreadsheet to your Google Drive with all code included
3. Continue to the Setup section below

## 🔧 Setup

1. Enter your Canvas Course ID in cell B3
   - Find this in the URL when viewing your course: `https://canvas.yourinstitution.edu/courses/12345`
2. Enter your Canvas Domain in cell B4 (e.g., https://canvas.chapman.edu)
3. For the API Key, you have two options:
   - Leave B5 blank and you'll be prompted when needed
   - Enter it directly in cell B5 (not recommended, less secured)

## 🔑 Generating a Canvas API Key

1. Log into Canvas
2. Go to Account > Settings
3. Scroll to "Approved Integrations"
4. Click "New Access Token"
5. Enter a purpose (e.g., "Google Sheets Integration")
6. Set an expiration date if desired
7. Click "Generate Token"
8. **IMPORTANT:** Copy the token immediately - you cannot view it again!

## 📝 Usage

### Fetching Student Data

1. Ensure your Course ID and Canvas Domain are set
2. Click "Canvas Tools > Fetch Course Users"
3. Student data will populate in columns A-D starting at row 7

### Fetching the Complete Gradebook

1. Click "Canvas Tools > Fetch Complete Gradebook"
2. The spreadsheet will be populated with:
   - All assignments from Canvas (column E onwards)
   - Points possible for each assignment (row 3)
   - Average scores and percentages (rows 4-5)
   - Assignment IDs (row 6)
   - All student grades

### Fetching Assignment Grades

1. In row 6 of column E (or any empty column), enter the Canvas Assignment ID
   - Find the Assignment ID in the URL when viewing the assignment:
   - Example: `https://canvas.chapman.edu/courses/12345/assignments/67890` → ID is 67890
2. Click "Canvas Tools > Fetch Assignment Grades"
3. Enter the column letter when prompted (e.g., "E")
4. Grades will populate in that column aligned with student rows

### Uploading Individual Assignment Grades

1. Ensure the column contains the Assignment ID in row 6
2. Enter grades in the column aligned with student rows
3. Click "Canvas Tools > Upload Assignment Grades to Canvas"
4. When prompted, enter the column letter
5. A summary will appear when complete

### Uploading a Range of Grades

1. Ensure Assignment IDs are in row 6 for all columns you want to upload
2. Enter grades in those columns
3. Click "Canvas Tools > Upload Grade Range to Canvas"
4. Enter the starting and ending column letters when prompted
5. A summary will appear when complete

### Uploading the Complete Gradebook

1. Click "Canvas Tools > Upload Complete Gradebook to Canvas"
2. Confirm the action when prompted
3. All grades in the sheet will be uploaded to Canvas
4. A summary will appear when complete

## 🔒 Security Considerations

- Your Canvas API key provides access to your Canvas account
- The key is stored securely in the Script Properties and not visible to others
- When entered in cell B5, it will be read once and then stored securely

## 🐞 Troubleshooting

| Problem | Possible Solution |
| --- | --- |
| "API Key not found" | Enter your Canvas API key in cell B5 or when prompted |
| "Course ID not found" | Make sure you entered your Canvas Course ID in cell B3 |
| "Canvas Domain missing" | Enter your institution's Canvas URL in cell B4 |
| "Assignment ID not found" | Ensure row 6 contains the correct Assignment ID |
| No grades appear | Verify students have submissions in Canvas |
| Upload errors | Check you have permission to edit grades in Canvas |

## 📄 Licensing

Grade Tracking with Canvas API for Google Sheets uses dual licensing:

- **Code**: All source code is licensed under the [MIT License](LICENSE)
- **Documentation**: All documentation, screenshots, and educational content are licensed under [Creative Commons Attribution-ShareAlike 4.0 International License](docs/LICENSE-DOCS.md)

## 👤 Author

[Jean-Louis Bru, Ph.D.](https://www.jlouisbru.com/)

Instructional Assistant Professor at [Chapman University](https://www.chapman.edu/)

## 🤝 Contributing

Contributions are welcome! [Thank you so much for your support!](https://ko-fi.com/louisfr)

## 🐛 Issues and Feedback

Found a bug or have a suggestion to improve Canvas Tools? We'd love to hear from you!

[![Report Bug](https://img.shields.io/badge/Report-Bug-red?style=for-the-badge&logo=github)](https://github.com/jlouisbru/grade-tracking-Canvas-API/issues/new?template=bug_report.yml)
[![Request Feature](https://img.shields.io/badge/Request-Feature-blue?style=for-the-badge&logo=github)](https://github.com/jlouisbru/grade-tracking-Canvas-API/issues/new?template=feature_request.yml)
[![View Issues](https://img.shields.io/badge/View-Issues-green?style=for-the-badge&logo=github)](https://github.com/jlouisbru/grade-tracking-Canvas-API/issues)

---

*Note: This tool is not affiliated with or endorsed by Instructure (Canvas) or Google.*
