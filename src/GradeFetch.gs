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
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Fetch Canvas Grades (by SIS ID)',
    'Enter the column letter where grades should be written (e.g., "F"):',
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    const gradeColumn = response.getResponseText().toUpperCase().trim();
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
  ToastManager.showToast(`Fetching grades for Assignment ID: ${assignmentId}`, 'Grade Fetch: In Progress', 10);

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
        }
        const linkHeader = response.getHeaders()['Link'] || response.getHeaders()['link'];
        nextPage = parseLinkHeader_(linkHeader);
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
    const nonEmptySisIds = sheetSisIds.filter(r => String(r[0]).trim()).length;
    Logger.log(`Read ${sheetSisIds.length} rows from column ${SIS_ID_COLUMN}; ${nonEmptySisIds} have SIS IDs.`);

    // --- 5. Match Submissions to Sheet Rows and Write Grades ---
    let updatedCount = 0;
    let notFoundCount = 0;
    let submissionsWithoutSisId = 0;
    const gradeColNum = columnLetterToNumber_(gradeOutputColumn);
    const numStudentRows = lastRow - FIRST_DATA_ROW + 1;
    const gradesArray = Array.from({length: numStudentRows}, () => ['']);

    allSubmissions.forEach((submission, index) => {
      const sisUserIdFromCanvas = submission.user && submission.user.sis_user_id ? String(submission.user.sis_user_id).trim() : null;
      const grade = submission.score; // Can be null, number, or string depending on grading type

      if (index < 5) {
          Logger.log(`Processing Submission #${index + 1}: has SIS User ID: ${!!sisUserIdFromCanvas}, has Score: ${grade !== null && grade !== undefined}`);
      }

      if (sisUserIdFromCanvas) {
        // Find the row index in the sheetSisIds array
        const studentRowIndex = sheetSisIds.findIndex(row => String(row[0]).trim() === sisUserIdFromCanvas);

        if (studentRowIndex !== -1) {
          gradesArray[studentRowIndex][0] = grade === null ? '' : grade;
          updatedCount++;
        } else {
          // Log only if it wasn't one of the first few already logged
          if (index >= 5) {
              Logger.log(`Submission #${index + 1} SIS User ID not matched in spreadsheet column ${SIS_ID_COLUMN}.`);
          }
          notFoundCount++;
        }
      } else {
        Logger.log(`Submission #${index + 1} does not have an SIS User ID.`);
        submissionsWithoutSisId++;
      }
    }); // end forEach(submission)

    // --- 6. Perform Batch Update ---
    if (updatedCount > 0) {
        Logger.log(`Attempting to write ${updatedCount} grades.`);
        sheet.getRange(FIRST_DATA_ROW, gradeColNum, numStudentRows, 1).setValues(gradesArray);
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
    ToastManager.showCompletionToast('Grades have been fetched and updated in the sheet.', 'Grade Fetch', 5);
    ui.alert(message);

  } catch (error) {
    Logger.log('Error during grade fetch: ' + error);
    Logger.log('Stack Trace: ' + error.stack);
    ToastManager.showErrorToast(error.message, 10);
    ui.alert('An error occurred: ' + error.message + '\nCheck script logs (View > Logs) for more details.');
  }
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
  
  // Use the active sheet instead of looking for "Gradebook"
  let gradebookSheet = ss.getActiveSheet();
  Logger.log("Using the active sheet for gradebook operations.");
  
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
    // Pass apiKey instead of undefined 'token'
    const users = fetchAllCanvasUsers_(courseId, apiKey, canvasDomain);
    if (users.length === 0) {
      ui.alert('No users found for Course ID ' + courseId + '. Check the Course ID and API permissions.');
      return;
    }
    Logger.log(`Successfully fetched ${users.length} users from Canvas.`);
    
    // Show toast notifications
    ToastManager.showToast('Fetching assignments from Canvas...', 'Gradebook Fetch: Step 2/4', 30);
    
    // --- 3. Fetch Assignments from Canvas ---
    Logger.log('Fetching assignments from Canvas...');
    // Pass apiKey instead of undefined 'token'
    const assignments = fetchCanvasAssignments_(courseId, apiKey, canvasDomain);
    if (assignments.length === 0) {
      ui.alert('No assignments found for Course ID ' + courseId + '. Check the Course ID and API permissions.');
      return;
    }
    Logger.log(`Successfully fetched ${assignments.length} assignments from Canvas.`);
    
    // Show toast notifications
    ToastManager.showToast('Fetching grades from Canvas...', 'Gradebook Fetch: Step 3/4', 30);
    
    // --- 4. Fetch Complete Gradebook from Canvas ---
    Logger.log('Fetching complete gradebook from Canvas...');
    // Pass apiKey instead of undefined 'token'
    const gradebook = fetchCanvasGradebook_(courseId, apiKey, canvasDomain);
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

  // Clear existing assignment columns, preserving columns A-D and row 2
  const lastCol = sheet.getLastColumn();
  if (lastCol >= FIRST_ASSIGNMENT_COL) {
    sheet.getRange(1, FIRST_ASSIGNMENT_COL, 1, lastCol - FIRST_ASSIGNMENT_COL + 1).clear(); // Row 1
    sheet.getRange(3, FIRST_ASSIGNMENT_COL, sheet.getLastRow() - 2, lastCol - FIRST_ASSIGNMENT_COL + 1).clear(); // Row 3 onwards
  }

  // Row layout (rows 1-6 are headers; FIRST_DATA_ROW+ is student data)
  const assignmentNameRow = 1;
  // Row 2 is left untouched
  const totalPointsRow = 3;
  const avgScoreRow = 4;
  const avgPercentRow = 5;
  // ASSIGNMENT_ID_HEADER_ROW (6) holds assignment IDs

  // Group assignments by assignment_group_id
  const assignmentGroups = {};
  assignments.forEach(assignment => {
    const groupId = assignment.assignment_group_id || 0;
    if (!assignmentGroups[groupId]) assignmentGroups[groupId] = [];
    assignmentGroups[groupId].push(assignment);
  });

  // Build ordered list of assignments and collect header data
  const assignmentColumns = []; // {column, id, points}
  const headerNames = [];
  const headerPoints = [];
  const headerIds = [];
  let currentCol = FIRST_ASSIGNMENT_COL;

  for (const groupId in assignmentGroups) {
    assignmentGroups[groupId].forEach(assignment => {
      assignmentColumns.push({ column: currentCol, id: assignment.id, points: assignment.points_possible });
      headerNames.push(assignment.name);
      headerPoints.push(assignment.points_possible);
      headerIds.push(assignment.id);
      currentCol++;
    });
  }

  const numAssignmentCols = assignmentColumns.length;
  if (numAssignmentCols === 0) {
    Logger.log('No assignments to write.');
    return;
  }

  // Batch write all assignment header rows at once
  sheet.getRange(assignmentNameRow, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols).setValues([headerNames]);
  sheet.getRange(totalPointsRow, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols).setValues([headerPoints]);
  sheet.getRange(ASSIGNMENT_ID_HEADER_ROW, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols).setValues([headerIds]);

  // Build SIS User ID → user object lookup
  const usersBySisId = {};
  users.forEach(user => {
    if (user.sis_user_id) usersBySisId[user.sis_user_id] = user;
  });

  // Get student SIS IDs from the sheet to match the existing row order
  const sheetSisIds = sheet.getRange(`${SIS_ID_COLUMN}${FIRST_DATA_ROW}:${SIS_ID_COLUMN}${sheet.getLastRow()}`).getValues();
  const numStudents = sheetSisIds.length;

  // Build grade matrix and collect per-assignment scores for statistics
  const gradeMatrix = Array.from({ length: numStudents }, () => new Array(numAssignmentCols).fill(''));
  const columnScores = assignmentColumns.map(() => []); // scores[j] = numeric scores for assignment j

  for (let i = 0; i < numStudents; i++) {
    const sisId = sheetSisIds[i][0];
    if (!sisId) continue;
    const user = usersBySisId[sisId];
    if (!user || !gradebook[user.id]) continue;
    for (let j = 0; j < numAssignmentCols; j++) {
      const score = gradebook[user.id][assignmentColumns[j].id];
      if (score !== null && score !== undefined) {
        gradeMatrix[i][j] = score;
        columnScores[j].push(score);
      }
    }
  }

  // Batch write all grade data at once
  if (numStudents > 0) {
    sheet.getRange(FIRST_DATA_ROW, FIRST_ASSIGNMENT_COL, numStudents, numAssignmentCols).setValues(gradeMatrix);
  }

  // Calculate statistics and batch write stat rows
  const avgScoreValues = [];
  const avgPercentValues = [];
  assignmentColumns.forEach((assignmentInfo, j) => {
    const scores = columnScores[j];
    const maxPoints = assignmentInfo.points || 0;
    if (scores.length > 0) {
      const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
      avgScoreValues.push(avgScore);
      avgPercentValues.push(maxPoints > 0 ? (avgScore / maxPoints) : 'No data');
    } else {
      avgScoreValues.push('No data');
      avgPercentValues.push('No data');
    }
  });

  sheet.getRange(avgScoreRow, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols).setValues([avgScoreValues]);
  sheet.getRange(avgPercentRow, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols).setValues([avgPercentValues]);
  sheet.getRange(avgPercentRow, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols).setNumberFormat('0.00%');

  // Format all header rows
  [assignmentNameRow, totalPointsRow, avgScoreRow, avgPercentRow, ASSIGNMENT_ID_HEADER_ROW].forEach(row => {
    sheet.getRange(row, FIRST_ASSIGNMENT_COL, 1, numAssignmentCols)
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  });

  // Freeze header rows and user info columns
  sheet.setFrozenRows(ASSIGNMENT_ID_HEADER_ROW);
  sheet.setFrozenColumns(4); // Freeze columns A-D

  Logger.log('Finished writing gradebook data to sheet.');
}
