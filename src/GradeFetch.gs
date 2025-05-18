/**
 * Fetches grades for a specific assignment from Canvas and populates them
 * into the active Google Sheet, matching students based on SIS User ID.
 * Retrieves API Key from Script Properties for security.
 * Relies on constants defined in a separate configuration file (e.g., Config.gs).
 * NOTE: This script assumes it's part of a project that includes a separate
 * onOpen function to add its menu item.
 */

// --- User Interaction ---
/**
 * Prompts the user to enter the column letter where grades should be written.
 * This function should be called by a menu item created in a central onOpen function.
 * @private
 */
function promptForGradeColumn_() { // Renamed function
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Fetch Canvas Grades (by SIS ID)',
    'Enter the column letter where grades should be written (e.g., "F"):',
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    var gradeColumn = response.getResponseText().toUpperCase().trim();
    // Validate that it's a valid column letter
    if (/^[A-Z]+$/.test(gradeColumn)) {
      fetchCanvasGradesBySisId_(gradeColumn); // Call the main fetching function
    } else {
      ui.alert('Invalid input. Please enter a valid column letter (e.g., "F").');
    }
  }
}

// --- Core Logic ---
/**
 * Fetches Canvas assignment grades and writes them to the specified column,
 * matching users by SIS User ID.
 * @param {string} gradeOutputColumn The letter of the column to write grades to.
 * @private
 */
function fetchCanvasGradesBySisId_(gradeOutputColumn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi(); // Get UI for alerts

  // --- Get Canvas Domain from cell B4 ---
  const canvasDomain = getCanvasDomain();
  if (!canvasDomain) {
    // getCanvasDomain() already showed an alert, just exit
    return;
  }

  // --- Get API Key (from cell B5, Script Properties, or user prompt) ---
  const apiKey = getCanvasApiKey();
  if (!apiKey) {
    // getCanvasApiKey() already showed an alert, just exit
    return;
  }

  // --- 1. Get Course ID ---
  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert(`Error: Course ID not found in cell ${COURSE_ID_CELL}.`);
    Logger.log(`Error: Course ID not found in cell ${COURSE_ID_CELL}`);
    return;
  }
  Logger.log(`Using Course ID: ${courseId}`);

  // --- 2. Get Assignment ID ---
  const assignmentIdCell = `${gradeOutputColumn}${ASSIGNMENT_ID_HEADER_ROW}`;
  const assignmentId = sheet.getRange(assignmentIdCell).getValue();
  if (!assignmentId) {
    ui.alert(`Error: Assignment ID not found in cell ${assignmentIdCell}. Please ensure the Canvas Assignment ID is in this cell.`);
    Logger.log(`Error: Assignment ID not found in cell ${assignmentIdCell}`);
    return;
  }
  Logger.log(`Using Assignment ID: ${assignmentId}`);

  // --- 3. Fetch Submissions from Canvas ---
  const url = `${canvasDomain}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions?include[]=user&per_page=100`;
  Logger.log(`Fetching submissions from: ${url}`);

  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + apiKey // Use the key retrieved from properties
    },
    muteHttpExceptions: true
  };

  let allSubmissions = [];
  let nextPage = url;
  let pageCount = 0;

  try {
    while (nextPage) {
      pageCount++;
      Logger.log(`Fetching page ${pageCount}: ${nextPage}`);
      const response = UrlFetchApp.fetch(nextPage, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const submissions = JSON.parse(responseBody);
        if (submissions && submissions.length > 0) {
            allSubmissions = allSubmissions.concat(submissions);
        } else {
            // No more submissions on this page, stop pagination
            nextPage = null;
        }
        // Check for pagination link even if submissions array is empty, just in case
        const linkHeader = response.getHeaders()['Link'] || response.getHeaders()['link'];
        nextPage = parseLinkHeader_(linkHeader); // Update nextPage based on header
        Logger.log("Next page URL from header: " + (nextPage || 'None'));
      } else {
        Logger.log(`API Error - Status: ${responseCode}, Response: ${responseBody.substring(0, 500)}`);
        // Provide more specific feedback for common errors
        let errorDetails = `Canvas API request failed with status ${responseCode}. Check Course/Assignment IDs and API permissions. URL: ${nextPage || url}`; // Use original URL if nextPage is null
        if (responseCode === 401) {
            errorDetails = `Canvas API request failed with status 401 (Unauthorized). Please verify that the API Key in Script Properties ('${API_KEY_PROPERTY_NAME}') is correct and has the necessary permissions.`;
        } else if (responseCode === 404) {
            errorDetails = `Canvas API request failed with status 404 (Not Found). Please verify the Course ID (${courseId}) and Assignment ID (${assignmentId}) are correct and exist in Canvas.`;
        }
        throw new Error(errorDetails);
      }
    } // end while(nextPage)

    Logger.log(`Successfully fetched ${allSubmissions.length} total submissions.`);

    // --- 4. Get SIS IDs from Spreadsheet ---
    const lastRow = sheet.getLastRow();
    // Uses FIRST_DATA_ROW defined in the configuration file
    if (lastRow < FIRST_DATA_ROW) {
        ui.alert("No student data found in the sheet starting from row " + FIRST_DATA_ROW);
        Logger.log("No student data rows found.");
        return;
    }
    // Uses SIS_ID_COLUMN and FIRST_DATA_ROW defined in the configuration file
    const sisIdRange = sheet.getRange(`${SIS_ID_COLUMN}${FIRST_DATA_ROW}:${SIS_ID_COLUMN}${lastRow}`);
    const sheetSisIds = sisIdRange.getValues();
    Logger.log(`Read ${sheetSisIds.length} SIS IDs from column ${SIS_ID_COLUMN}.`);
    if (sheetSisIds.length > 0) {
        Logger.log(`First few SIS IDs from Sheet (Column ${SIS_ID_COLUMN}): [${sheetSisIds.slice(0, 5).map(r => `'${String(r[0]).trim()}'`).join(', ')}]`);
    }

    // --- 5. Match Submissions to Sheet Rows and Write Grades ---
    let updatedCount = 0;
    let notFoundCount = 0;
    let submissionsWithoutSisId = 0;
    let gradesToWrite = [];

    allSubmissions.forEach((submission, index) => {
      const sisUserIdFromCanvas = submission.user && submission.user.sis_user_id ? String(submission.user.sis_user_id).trim() : null;
      const grade = submission.score; // Can be null, number, or string depending on grading type

      if (index < 5) {
          Logger.log(`Processing Submission #${index + 1}: Canvas User ID: ${submission.user_id}, SIS User ID: '${sisUserIdFromCanvas}', Score: ${grade}`);
      }

      if (sisUserIdFromCanvas) {
        // Find the row index in the sheetSisIds array
        const studentRowIndex = sheetSisIds.findIndex(row => String(row[0]).trim() === sisUserIdFromCanvas);

        if (studentRowIndex !== -1) {
          // Calculate the actual row number in the sheet
          // Uses FIRST_DATA_ROW defined in the configuration file
          const outputSheetRow = studentRowIndex + FIRST_DATA_ROW;
          gradesToWrite.push({
              row: outputSheetRow,
              col: columnLetterToNumber_(gradeOutputColumn),
              value: grade === null ? '' : grade // Write empty string if grade is null/ungraded
          });
          updatedCount++;
        } else {
          // Log only if it wasn't one of the first few already logged
          if (index >= 5) {
              Logger.log(`SIS User ID '${sisUserIdFromCanvas}' (Canvas User: ${submission.user_id}) from submission not found in spreadsheet column ${SIS_ID_COLUMN}.`);
          }
          notFoundCount++;
        }
      } else {
        Logger.log(`Submission for Canvas User ID ${submission.user_id} does not have an SIS User ID.`);
        submissionsWithoutSisId++;
      }
    }); // end forEach(submission)

    // --- 6. Perform Batch Update ---
    if (gradesToWrite.length > 0) {
        Logger.log(`Attempting to write ${gradesToWrite.length} grades.`);
        // Write grades individually for simplicity and better error isolation,
        // though batchUpdate could be used for performance on very large sheets.
        gradesToWrite.forEach(item => {
            try {
                sheet.getRange(item.row, item.col).setValue(item.value);
            } catch (e) {
                Logger.log(`Error writing grade ${item.value} to row ${item.row}, col ${item.col}: ${e}`);
            }
        });
        Logger.log("Finished writing grades.");
    } else {
        Logger.log("No grades to write (no matching SIS IDs found or submissions were empty).");
    }

    // --- 7. Report Results ---
    let message = `Grade fetch complete.\nUpdated grades for ${updatedCount} students.\n`;
    if (notFoundCount > 0) {
      message += `${notFoundCount} SIS User IDs from Canvas submissions were not found in column ${SIS_ID_COLUMN}.\n`;
    }
    if (submissionsWithoutSisId > 0) {
      message += `${submissionsWithoutSisId} submissions lacked an SIS User ID.\n`;
    }
    message += `Check script logs (View > Logs) for details.`;
    ui.alert(message);

  } catch (error) {
    Logger.log('Error during grade fetch: ' + error);
    Logger.log('Stack Trace: ' + error.stack);
    ui.alert('An error occurred: ' + error.message + '\nCheck script logs (View > Logs) for more details.');
  }
}


// --- Helper Functions ---

/**
 * Parses the 'Link' header from Canvas API responses to find the 'next' page URL.
 * @param {string} linkHeader The 'Link' header string.
 * @return {string|null} The URL for the next page, or null if not found.
 * @private
 */
function parseLinkHeader_(linkHeader) {
  if (!linkHeader) {
    return null;
  }
  // Example header: <url1>; rel="current", <url2>; rel="next", <url3>; rel="first", <url4>; rel="last"
  const links = linkHeader.split(',');
  for (let i = 0; i < links.length; i++) {
      const link = links[i].trim();
      const parts = link.split(';');
      if (parts.length >= 2) {
          const urlPart = parts[0].trim();
          const relPart = parts[1].trim();
          if (relPart === 'rel="next"') {
              const match = urlPart.match(/<(.*?)>/);
              if (match && match[1]) {
                  return match[1]; // Return the URL inside <>
              }
          }
      }
  }
  return null; // No 'next' link found
}

/**
 * Converts a column letter (A, B, ..., Z, AA, AB, ...) to its 1-based column number.
 * @param {string} letter The column letter(s).
 * @return {number} The 1-based column number.
 * @private
 */
function columnLetterToNumber_(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    // Ensure uppercase for calculation
    column += (letter.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * ==========================================================================
 * SCRIPT TO FETCH COMPLETE CANVAS GRADEBOOK
 * Fetches all assignments and grades from Canvas for a course and 
 * populates them into the "Gradebook" spreadsheet.
 * Requires global constants (CANVAS_DOMAIN, API_KEY_PROPERTY_NAME,
 * COURSE_ID_CELL, SIS_ID_COLUMN, FIRST_DATA_ROW) defined elsewhere.
 * ==========================================================================
 */

/**
 * Main function to trigger fetching the complete Canvas gradebook.
 * Can be run from the Apps Script editor or a custom menu.
 */
function fetchCompleteCanvasGradebook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // --- Get API Key from Script Properties ---
  const apiKey = getCanvasApiKey();
  if (!apiKey) {
    return;
}
  
  // --- Get Canvas Domain from cell B4 ---
  const canvasDomain = getCanvasDomain();
  if (!canvasDomain) {
    return;
  }
  
  // Get or create the Gradebook sheet
  let gradebookSheet = ss.getSheetByName("Gradebook");
  if (!gradebookSheet) {
    gradebookSheet = ss.insertSheet("Gradebook");
    Logger.log("Created new 'Gradebook' sheet");
  }
  
  // --- 1. Get Course ID ---
  const courseId = gradebookSheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert('Error: Course ID not found in cell ' + COURSE_ID_CELL);
    Logger.log('Error: Course ID not found in cell ' + COURSE_ID_CELL);
    return;
  }
  Logger.log('Fetching gradebook for Course ID: ' + courseId);
  
  try {
    // Show toast notifications
    ToastManager.showToast('Fetching users from Canvas...', 'Gradebook Fetch: Step 1/4', 30);
    
    // --- 2. Fetch Users from Canvas ---
    Logger.log('Fetching users from Canvas...');
    const users = fetchAllCanvasUsers_(courseId, token, CANVAS_DOMAIN);
    if (users.length === 0) {
      ui.alert('No users found for Course ID ' + courseId + '. Check the Course ID and API permissions.');
      return;
    }
    Logger.log(`Successfully fetched ${users.length} users from Canvas.`);
    
    // Show toast notifications
    ToastManager.showToast('Fetching assignments from Canvas...', 'Gradebook Fetch: Step 2/4', 30);
    
    // --- 3. Fetch Assignments from Canvas ---
    Logger.log('Fetching assignments from Canvas...');
    const assignments = fetchCanvasAssignments_(courseId, token, CANVAS_DOMAIN);
    if (assignments.length === 0) {
      ui.alert('No assignments found for Course ID ' + courseId + '. Check the Course ID and API permissions.');
      return;
    }
    Logger.log(`Successfully fetched ${assignments.length} assignments from Canvas.`);
    
    // Show toast notifications
    ToastManager.showToast('Fetching grades from Canvas...', 'Gradebook Fetch: Step 3/4', 30);
    
    // --- 4. Fetch Complete Gradebook from Canvas ---
    Logger.log('Fetching complete gradebook from Canvas...');
    const gradebook = fetchCanvasGradebook_(courseId, token, CANVAS_DOMAIN);
    Logger.log(`Successfully fetched gradebook data for ${Object.keys(gradebook).length} students.`);
    
    // Show toast notifications
    ToastManager.showToast('Writing data to sheet...', 'Gradebook Fetch: Step 4/4', 30);
    
    // --- 5. Prepare and Write Data to Sheet ---
    writeGradebookToSheet_(gradebookSheet, users, assignments, gradebook);
    
    // Show completion toast
    ToastManager.showToast(
      `Success! Gradebook updated with data for ${users.length} users and ${assignments.length} assignments.`,
      'Gradebook Fetch: Complete!',
      10
    );
    
    ui.alert(`Success! Gradebook sheet updated with data for ${users.length} users and ${assignments.length} assignments.`);
    
  } catch (error) {
    Logger.log('Error fetching or processing Canvas data: ' + error);
    Logger.log('Stack Trace: ' + error.stack);
    ToastManager.showToast('Error: ' + error.message, 'Gradebook Fetch: Failed', 10);
    ui.alert('An error occurred: ' + error.message + '\n\nPlease check the Script Execution Logs (View > Logs) for more details.');
  }
}

/**
 * Fetches all users for a given course ID from the Canvas API, handling pagination.
 * @param {string|number} courseId The Canvas Course ID.
 * @param {string} apiToken The Canvas API token.
 * @param {string} canvasDomain The base Canvas domain URL.
 * @return {Array<Object>} An array of Canvas user objects.
 * @private
 */
function fetchAllCanvasUsers_(courseId, apiToken, canvasDomain) {
  let allUsers = [];
  let nextPageUrl = `${canvasDomain}/api/v1/courses/${courseId}/users?include[]=email&per_page=100`;

  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiToken
    },
    'muteHttpExceptions': true
  };

  while (nextPageUrl) {
    Logger.log('Fetching URL: ' + nextPageUrl);
    const response = UrlFetchApp.fetch(nextPageUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const users = JSON.parse(responseBody);
      if (users && users.length > 0) {
         allUsers = allUsers.concat(users);
      } else {
         nextPageUrl = null;
      }
      const linkHeader = response.getHeaders()['Link'] || response.getHeaders()['link'];
      nextPageUrl = parseLinkHeader_(linkHeader);
      Logger.log("Next page URL: " + (nextPageUrl || 'None'));
    } else {
      Logger.log('API Error - Response Code: ' + responseCode);
      Logger.log('API Error - Response Body: ' + responseBody);
      throw new Error(`Canvas API request failed with status code ${responseCode}. Check Course ID (${courseId}), API Token validity, and Permissions. Response: ${responseBody.substring(0, 500)}`);
    }
  }

  return allUsers;
}

/**
 * Fetches all assignments for a given course ID from the Canvas API, handling pagination.
 * @param {string|number} courseId The Canvas Course ID.
 * @param {string} apiToken The Canvas API token.
 * @param {string} canvasDomain The base Canvas domain URL.
 * @return {Array<Object>} An array of Canvas assignment objects.
 * @private
 */
function fetchCanvasAssignments_(courseId, apiToken, canvasDomain) {
  let allAssignments = [];
  let nextPageUrl = `${canvasDomain}/api/v1/courses/${courseId}/assignments?per_page=100`;

  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiToken
    },
    'muteHttpExceptions': true
  };

  while (nextPageUrl) {
    Logger.log('Fetching URL: ' + nextPageUrl);
    const response = UrlFetchApp.fetch(nextPageUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const assignments = JSON.parse(responseBody);
      if (assignments && assignments.length > 0) {
         allAssignments = allAssignments.concat(assignments);
      } else {
         nextPageUrl = null;
      }
      const linkHeader = response.getHeaders()['Link'] || response.getHeaders()['link'];
      nextPageUrl = parseLinkHeader_(linkHeader);
      Logger.log("Next page URL: " + (nextPageUrl || 'None'));
    } else {
      Logger.log('API Error - Response Code: ' + responseCode);
      Logger.log('API Error - Response Body: ' + responseBody);
      throw new Error(`Canvas API request failed with status code ${responseCode}. Check Course ID (${courseId}), API Token validity, and Permissions. Response: ${responseBody.substring(0, 500)}`);
    }
  }

  // Sort assignments by position (if available) or ID
  allAssignments.sort((a, b) => {
    if (a.position !== undefined && b.position !== undefined) {
      return a.position - b.position;
    }
    return a.id - b.id;
  });

  return allAssignments;
}

/**
 * Fetches the complete gradebook for a given course ID from the Canvas API.
 * @param {string|number} courseId The Canvas Course ID.
 * @param {string} apiToken The Canvas API token.
 * @param {string} canvasDomain The base Canvas domain URL.
 * @return {Object} An object mapping student IDs to their gradebook entries.
 * @private
 */
function fetchCanvasGradebook_(courseId, apiToken, canvasDomain) {
  // We'll build a map of student ID -> { assignment ID -> grade }
  const gradebook = {};
  
  // Fetch the detailed grade data from the Canvas Progress tab
  let nextPageUrl = `${canvasDomain}/api/v1/courses/${courseId}/students/submissions?include[]=user&student_ids[]=all&per_page=100`;

  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiToken
    },
    'muteHttpExceptions': true
  };

  while (nextPageUrl) {
    Logger.log('Fetching URL: ' + nextPageUrl);
    const response = UrlFetchApp.fetch(nextPageUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const submissions = JSON.parse(responseBody);
      if (submissions && submissions.length > 0) {
        submissions.forEach(submission => {
          const userId = submission.user_id;
          const assignmentId = submission.assignment_id;
          const score = submission.score;
          
          if (!gradebook[userId]) {
            gradebook[userId] = {};
          }
          
          gradebook[userId][assignmentId] = score;
        });
      } else {
        nextPageUrl = null;
      }
      const linkHeader = response.getHeaders()['Link'] || response.getHeaders()['link'];
      nextPageUrl = parseLinkHeader_(linkHeader);
      Logger.log("Next page URL: " + (nextPageUrl || 'None'));
    } else {
      Logger.log('API Error - Response Code: ' + responseCode);
      Logger.log('API Error - Response Body: ' + responseBody);
      throw new Error(`Canvas API request failed with status code ${responseCode}. Check Course ID (${courseId}), API Token validity, and Permissions. Response: ${responseBody.substring(0, 500)}`);
    }
  }

  return gradebook;
}

/**
 * Writes the gradebook data to the specified sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to write to.
 * @param {Array<Object>} users The users fetched from Canvas.
 * @param {Array<Object>} assignments The assignments fetched from Canvas.
 * @param {Object} gradebook The gradebook data (userId -> assignmentId -> grade).
 * @private
 */
function writeGradebookToSheet_(sheet, users, assignments, gradebook) {
  Logger.log('Preparing to write gradebook data to sheet...');
  
  // Don't clear the entire sheet - preserve columns A-D and row 2
  const firstAssignmentCol = 5; // Column E
  const lastCol = sheet.getLastColumn();
  if (lastCol >= firstAssignmentCol) {
    // Clear all rows except row 2
    sheet.getRange(1, firstAssignmentCol, 1, lastCol - firstAssignmentCol + 1).clear(); // Row 1
    sheet.getRange(3, firstAssignmentCol, sheet.getLastRow() - 2, lastCol - firstAssignmentCol + 1).clear(); // Row 3 onwards
  }
  
  // Define row structure according to requirements
  const assignmentNameRow = 1;
  // Row 2 is left untouched as requested
  const totalPointsRow = 3;
  const avgScoreRow = 4;
  const avgPercentRow = 5;
  const assignmentIdRow = 6;
  const firstDataRow = 7; // Student data starts at row 7
  
  // Set up assignment headers
  let currentCol = firstAssignmentCol;
  
  // Group assignments by group_id/category and sort them
  const assignmentGroups = {};
  assignments.forEach(assignment => {
    const groupId = assignment.assignment_group_id || 0;
    if (!assignmentGroups[groupId]) {
      assignmentGroups[groupId] = [];
    }
    assignmentGroups[groupId].push(assignment);
  });
  
  // Track assignment columns for statistics calculation
  const assignmentColumns = [];
  
  // Set up assignment headers
  for (const groupId in assignmentGroups) {
    const groupAssignments = assignmentGroups[groupId];
    
    if (groupAssignments.length > 0) {
      // For each assignment in the group
      groupAssignments.forEach(assignment => {
        const name = assignment.name;
        const points = assignment.points_possible;
        const id = assignment.id;
        
        // Store assignment column information for statistics calculation later
        assignmentColumns.push({
          column: currentCol,
          id: id,
          points: points
        });
        
        // Write assignment info according to the required rows
        sheet.getRange(assignmentNameRow, currentCol).setValue(name);
        sheet.getRange(totalPointsRow, currentCol).setValue(points);
        sheet.getRange(assignmentIdRow, currentCol).setValue(id);
        
        // No notes added to cells in row 1 as requested
        
        currentCol++;
      });
    }
  }
  
  // Create a map of SIS User IDs to user objects for easy lookup
  const usersBySisId = {};
  users.forEach(user => {
    const sisUserId = user.sis_user_id;
    if (sisUserId) {
      usersBySisId[sisUserId] = user;
    }
  });
  
  // Get SIS User IDs from the spreadsheet to match the existing order
  const sisIdRange = sheet.getRange(`C${firstDataRow}:C${sheet.getLastRow()}`);
  const sheetSisIds = sisIdRange.getValues();
  
  // Track statistics for each assignment
  const statsData = {};
  assignmentColumns.forEach(info => {
    statsData[info.column] = {
      scores: [],
      totalPoints: info.points
    };
  });
  
  // Match grades to existing students by SIS ID
  for (let i = 0; i < sheetSisIds.length; i++) {
    const sisId = sheetSisIds[i][0];
    if (!sisId) continue; // Skip empty rows
    
    const rowIndex = i + firstDataRow;
    const user = usersBySisId[sisId];
    
    if (user && gradebook[user.id]) {
      // For each assignment, add the score to the appropriate column
      assignmentColumns.forEach(assignmentInfo => {
        const score = gradebook[user.id][assignmentInfo.id];
        if (score !== null && score !== undefined) {
          sheet.getRange(rowIndex, assignmentInfo.column).setValue(score);
          
          // Track this score for statistics calculation
          statsData[assignmentInfo.column].scores.push(score);
        }
      });
    }
  }
  
  // Calculate and add statistics for each assignment
  assignmentColumns.forEach(assignmentInfo => {
    const col = assignmentInfo.column;
    const maxPoints = assignmentInfo.points || 0;
    const scores = statsData[col].scores;
    
    if (scores.length > 0) {
      // Calculate average score
      const sum = scores.reduce((a, b) => a + b, 0);
      const avgScore = sum / scores.length;
      sheet.getRange(avgScoreRow, col).setValue(avgScore);
      
      // Calculate average percentage if possible
      if (maxPoints > 0) {
        const avgPercent = (avgScore / maxPoints) * 100;
        sheet.getRange(avgPercentRow, col).setValue(avgPercent + '%');
      }
    } else {
      // No scores available
      sheet.getRange(avgScoreRow, col).setValue("No data");
      sheet.getRange(avgPercentRow, col).setValue("No data");
    }
  });
  
  // Format the header rows
  const headerRows = [assignmentNameRow, totalPointsRow, avgScoreRow, avgPercentRow, assignmentIdRow];
  headerRows.forEach(row => {
    if (lastCol >= firstAssignmentCol) {
      sheet.getRange(row, firstAssignmentCol, 1, currentCol - firstAssignmentCol)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
    }
  });
  
  // Freeze the header rows and user info columns
  sheet.setFrozenRows(6);  // Freeze through row 6
  sheet.setFrozenColumns(4); // Freeze columns A-D
  
  Logger.log('Finished writing gradebook data to sheet.');
}
