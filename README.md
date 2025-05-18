<h1 align="center">Grade Tracking with Canvas API for Google Sheets</h1>

<p align="center">
  A Google Sheets integration with Canvas LMS that allows educators to fetch student data, download grades, and upload grades directly between Google Sheets and Canvas.
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/jlouisbru/canvas-tools-for-sheets" alt="License"></a>
  <a href="https://github.com/jlouisbru/canvas-tools-for-sheets/releases/latest"><img src="https://img.shields.io/github/v/release/jlouisbru/canvas-tools-for-sheets" alt="Latest Release"></a>
  <a href="https://docs.google.com/spreadsheets/d/18ZggFU-2xBdbl3pVPY3dXR-U5DYdroxvYaZJXGcvIPA/edit?usp=sharing"><img src="https://img.shields.io/badge/platform-Google%20Sheets-green" alt="Platform"></a>
  <a href="https://ko-fi.com/louisfr"><img src="https://img.shields.io/badge/Support-Ko--fi-ff5f5f" alt="Support on Ko-fi"></a>
</p>

<h2>🚀 Quick Start</h2>

<p align="center">
  <a href="https://docs.google.com/spreadsheets/d/18ZggFU-2xBdbl3pVPY3dXR-U5DYdroxvYaZJXGcvIPA/copy"><img src="https://img.shields.io/badge/Use_Template-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Use Google Sheet Template"></a>
</p>

<p align="center">Click the button above to create your own copy of the Canvas Tools template</p>

<h2>✨ Features</h2>

<ul>
  <li><strong>Fetch student data</strong> directly from Canvas into your spreadsheet</li>
  <li><strong>Download grades</strong> for individual assignments or the complete gradebook</li>
  <li><strong>Upload grades</strong> to Canvas individually or in bulk</li>
  <li><strong>Progress tracking</strong> with helpful toast notifications</li>
  <li><strong>Configurable setup</strong> through the spreadsheet (no code editing required)</li>
</ul>

<h2>💾 Installation</h2>

<h3>Use the Template</h3>

<ol>
  <li><a href="https://docs.google.com/spreadsheets/d/18ZggFU-2xBdbl3pVPY3dXR-U5DYdroxvYaZJXGcvIPA/edit?usp=sharing">Open the template spreadsheet</a></li>
  <li>Copy the spreadsheet to your Google Drive with all code included</li>
  <li>Continue to the Setup section below</li>
</ol>

<h2>🔧 Setup</h2>

<ol>
  <li>Enter your Canvas Course ID in cell B3
    <ul>
      <li>Find this in the URL when viewing your course: <code>https://canvas.yourinstitution.edu/courses/12345</code></li>
    </ul>
  </li>
  <li>Enter your Canvas Domain in cell B4 (e.g., https://canvas.chapman.edu)</li>
  <li>For the API Key, you have two options:
    <ul>
      <li>Leave B5 blank and you'll be prompted when needed</li>
      <li>Enter it directly in cell B5 (not recommended, less secured)</li>
    </ul>
  </li>
</ol>

<h2>🔑 Generating a Canvas API Key</h2>

<ol>
  <li>Log into Canvas</li>
  <li>Go to Account > Settings</li>
  <li>Scroll to "Approved Integrations"</li>
  <li>Click "New Access Token"</li>
  <li>Enter a purpose (e.g., "Google Sheets Integration")</li>
  <li>Set an expiration date if desired</li>
  <li>Click "Generate Token"</li>
  <li><strong>IMPORTANT:</strong> Copy the token immediately - you cannot view it again!</li>
</ol>

<h2>📝 Usage</h2>

<h3>Fetching Student Data</h3>

<ol>
  <li>Ensure your Course ID and Canvas Domain are set</li>
  <li>Click "Canvas Tools > Fetch Course Users"</li>
  <li>Student data will populate in columns A-D starting at row 7</li>
</ol>

<h3>Fetching the Complete Gradebook</h3>

<ol>
  <li>Click "Canvas Tools > Fetch Complete Gradebook"</li>
  <li>The spreadsheet will be populated with:
    <ul>
      <li>All assignments from Canvas (column E onwards)</li>
      <li>Points possible for each assignment (row 3)</li>
      <li>Average scores and percentages (rows 4-5)</li>
      <li>Assignment IDs (row 6)</li>
      <li>All student grades</li>
    </ul>
  </li>
</ol>

<h3>Fetching Assignment Grades</h3>

<ol>
  <li>In row 6 of column E (or any empty column), enter the Canvas Assignment ID
    <ul>
      <li>Find the Assignment ID in the URL when viewing the assignment:</li>
      <li>Example: <code>https://canvas.chapman.edu/courses/12345/assignments/67890</code> → ID is 67890</li>
    </ul>
  </li>
  <li>Click "Canvas Tools > Fetch Assignment Grades"</li>
  <li>Enter the column letter when prompted (e.g., "E")</li>
  <li>Grades will populate in that column aligned with student rows</li>
</ol>

<h3>Uploading Individual Assignment Grades</h3>

<ol>
  <li>Ensure the column contains the Assignment ID in row 6</li>
  <li>Enter grades in the column aligned with student rows</li>
  <li>Click "Canvas Tools > Upload Assignment Grades to Canvas"</li>
  <li>When prompted, enter the column letter</li>
  <li>A summary will appear when complete</li>
</ol>

<h3>Uploading a Range of Grades</h3>

<ol>
  <li>Ensure Assignment IDs are in row 6 for all columns you want to upload</li>
  <li>Enter grades in those columns</li>
  <li>Click "Canvas Tools > Upload Grade Range to Canvas"</li>
  <li>Enter the starting and ending column letters when prompted</li>
  <li>A summary will appear when complete</li>
</ol>

<h3>Uploading the Complete Gradebook</h3>

<ol>
  <li>Click "Canvas Tools > Upload Complete Gradebook to Canvas"</li>
  <li>Confirm the action when prompted</li>
  <li>All grades in the sheet will be uploaded to Canvas</li>
  <li>A summary will appear when complete</li>
</ol>

<h2>🔒 Security Considerations</h2>

<ul>
  <li>Your Canvas API key provides access to your Canvas account</li>
  <li>The key is stored securely in the Script Properties and not visible to others</li>
  <li>When entered in cell B5, it will be read once and then stored securely</li>
</ul>

<h2>🐞 Troubleshooting</h2>

<table>
  <thead>
    <tr>
      <th>Problem</th>
      <th>Possible Solution</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>"API Key not found"</td>
      <td>Enter your Canvas API key in cell B5 or when prompted</td>
    </tr>
    <tr>
      <td>"Course ID not found"</td>
      <td>Make sure you entered your Canvas Course ID in cell B3</td>
    </tr>
    <tr>
      <td>"Canvas Domain missing"</td>
      <td>Enter your institution's Canvas URL in cell B4</td>
    </tr>
    <tr>
      <td>"Assignment ID not found"</td>
      <td>Ensure row 6 contains the correct Assignment ID</td>
    </tr>
    <tr>
      <td>No grades appear</td>
      <td>Verify students have submissions in Canvas</td>
    </tr>
    <tr>
      <td>Upload errors</td>
      <td>Check you have permission to edit grades in Canvas</td>
    </tr>
  </tbody>
</table>

<h2>📄 Licensing</h2>

<p>Grade Tracking with Canvas API for Google Sheets uses dual licensing:</p>

<ul>
  <li><strong>Code</strong>: All source code is licensed under the <a href="LICENSE">MIT License</a></li>
  <li><strong>Documentation</strong>: All documentation, screenshots, and educational content are licensed under <a href="docs/LICENSE-DOCS.md">Creative Commons Attribution-ShareAlike 4.0 International License</a></li>
</ul>

<h2>👤 Author</h2>

<p><a href="https://www.jlouisbru.com/">Jean-Louis Bru, Ph.D.</a></p>
<p>Instructional Assistant Professor at <a href="https://www.chapman.edu/">Chapman University</a></p>

<h2>🤝 Contributing</h2>

<p>Contributions are welcome! <a href="https://ko-fi.com/louisfr">Thank you so much for your support!</a></p>

<h2>🐛 Issues and Feedback</h2>

<p>Found a bug or have a suggestion to improve Canvas Tools? We'd love to hear from you!</p>

<p>
  <a href="https://github.com/jlouisbru/canvas-tools-for-sheets/issues/new?template=bug_report.yml"><img src="https://img.shields.io/badge/Report-Bug-red?style=for-the-badge&logo=github" alt="Report Bug"></a>
  <a href="https://github.com/jlouisbru/canvas-tools-for-sheets/issues/new?template=feature_request.yml"><img src="https://img.shields.io/badge/Request-Feature-blue?style=for-the-badge&logo=github" alt="Request Feature"></a>
  <a href="https://github.com/jlouisbru/canvas-tools-for-sheets/issues"><img src="https://img.shields.io/badge/View-Issues-green?style=for-the-badge&logo=github" alt="View Issues"></a>
</p>

<hr>

<p align="center">
  <i>Note: This tool is not affiliated with or endorsed by Instructure (Canvas) or Google.</i>
</p>
