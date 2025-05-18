# Frequently Asked Questions

## General Questions

### What is Grade Tracking with Canvas API?
Canvas Tools is a Google Sheets integration that connects with Canvas LMS to help educators manage student grades and data more efficiently.

### Is Grade Tracking with Canvas API free to use?
Yes, Grade Tracking with Canvas API is completely free and open-source under the MIT License.

## Setup Questions

### Do I need to know programming to use Grade Tracking with Canvas API?
No, Grade Tracking with Canvas API is designed for educators with no programming experience. Simply make a copy of our template spreadsheet to get started.

### How do I find my Canvas Course ID?
Your Course ID appears in the URL when viewing your course. For example, in `https://canvas.university.edu/courses/12345`, the ID is `12345`.

## Usage Questions

### Can I use Grade Tracking with Canvas API with multiple courses?
Yes, but you'll need a separate spreadsheet for each course. Create a new copy of the template for each course.

### Will students see when I update their grades using Grade Tracking with Canvas API?
This depends on your Canvas course settings. If "Post Grades Immediately" is enabled for an assignment, students will see updates right away.

## Technical Questions

### How secure is my Canvas API key?
Your API key is stored securely in the Script Properties of your spreadsheet and is not visible to others, even if you share the spreadsheet.

### Does Grade Tracking with Canvas API access any student data beyond grades?
Canvas Tools only accesses the data needed for the functions you use: student names, SIS IDs, email addresses, and assignment grades.

## Privacy & Compliance Questions

### Is Grade Tracking with Canvas API FERPA compliant?
Grade Tracking with Canvas API is designed to work within FERPA guidelines by:
1. Only transferring data between systems you already have authorized access to (Canvas and Google Workspace)
2. Not storing any student data on third-party servers
3. Using secure API connections for all data transfers
4. Maintaining the same level of privacy protection provided by Canvas and Google Workspace

Remember that you as an educator are still responsible for ensuring your use of any tool, including Grade Tracking with Canvas API, complies with your institution's policies and applicable laws.

### Who can see the student data I import into Google Sheets?
Anyone you share your Google Sheet with will be able to see the student data. Make sure you only share your spreadsheet with authorized individuals who have a legitimate educational interest in the information, as required by FERPA.

### Do I need special permission from my institution to use Canvas Tools?
While Grade Tracking with Canvas API is designed to work within existing data processing agreements your institution likely has with both Canvas and Google, you should check your institution's policies regarding the use of third-party tools and APIs with student data. Some institutions may require approval for tools that access their LMS data.
