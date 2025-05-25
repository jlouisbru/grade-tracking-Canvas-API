/**
 * Uploads grades from a Google Sheet column to a specific Canvas assignment.
 * Reads Course ID, Assignment ID, and Student IDs from the sheet.
 * Retrieves API Key from Script Properties. Prompts user for grade column.
 * Relies on constants defined in a separate configuration file (e.g., Config.gs).
 * NOTE: This script assumes it's part of a project that includes a separate
 * onOpen function to add its menu item.
 */

// --- User Interaction ---
/**
 * Prompts the user to enter the column letter containing the grades to upload.
 * This function should be called by a menu item created in a central onOpen function.
 * @private
 */
function promptForGradeUploadColumn_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Upload Grades to Canvas',
    'Enter the column letter containing the grades to upload (e.g., "F"):',
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    const gradeColumn = response.getResponseText().toUpperCase().trim();
    // Validate that it's a valid column letter
    if (/^[A-Z]+$/.test(gradeColumn)) {
      // Call the main function - it will use global constants directly
      updateCanvasGrades_(gradeColumn);
    } else {
      ui.alert('Invalid input. Please enter a valid column letter (e.g., "F").');
    }
  }
}


// --- Core Logic ---
/**
 * Reads grades from the specified sheet column and uploads them to Canvas.
 * Uses global constants defined in the configuration file for sheet layout and API details.
 * @param {string} gradeColumn The letter of the column containing grades to upload.
 * @private
 */
function updateCanvasGrades_(gradeColumn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // --- Get API Key from Script Properties ---
  const apiKey = getCanvasApiKey();
  if (!apiKey) {
    return;
  }

  // --- Get Course ID from Sheet ---
  // Uses COURSE_ID_CELL defined in the configuration file
  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert(`Error: Course ID not found in cell ${COURSE_ID_CELL}.`);
    Logger.log(`Error: Course ID not found in cell ${COURSE_ID_CELL}`);
    return;
  }
  Logger.log(`Using Course ID: ${courseId}`);

  // --- Get Assignment ID from Sheet ---
  // Uses ASSIGNMENT_ID_HEADER_ROW defined in the configuration file
  const assignmentIdCell = `${gradeColumn}${ASSIGNMENT_ID_HEADER_ROW}`;
  const assignmentId = sheet.getRange(assignmentIdCell).getValue();
  if (!assignmentId) {
    ui.alert(`Error: Assignment ID not found in cell ${assignmentIdCell}. Please ensure the Canvas Assignment ID is in this cell.`);
    Logger.log(`Error: Assignment ID not found in cell ${assignmentIdCell}`);
    return;
  }
  Logger.log(`Using Assignment ID: ${assignmentId}`);

  // --- Read Data from Sheet ---
  const lastRow = sheet.getLastRow();
  // Uses FIRST_DATA_ROW defined in the configuration file
  if (lastRow < FIRST_DATA_ROW) {
      ui.alert("No student data found in the sheet starting from row " + FIRST_DATA_ROW);
      Logger.log("No student data rows found.");
      return;
  }
  // Read only the necessary columns (SIS ID and Grade Column) for efficiency
  // Uses SIS_ID_COLUMN and FIRST_DATA_ROW from config
  const studentIdColNum = columnToNumber_(SIS_ID_COLUMN);
  const gradeColNum = columnToNumber_(gradeColumn);
  const firstCol = Math.min(studentIdColNum, gradeColNum);
  const lastCol = Math.max(studentIdColNum, gradeColNum);
  const numCols = lastCol - firstCol + 1;

  const dataRange = sheet.getRange(FIRST_DATA_ROW, firstCol, lastRow - FIRST_DATA_ROW + 1, numCols);
  const values = dataRange.getValues();
  Logger.log(`Read data from range ${dataRange.getA1Notation()}.`);

  // Calculate 0-based indices relative to the fetched range `values`
  const studentIdIndex = studentIdColNum - firstCol;
  const gradeIndex = gradeColNum - firstCol;

  let successCount = 0;
  let failCount = 0;
  let invalidGradeCount = 0;
  let skippedMissingIdCount = 0;
  let failedStudents = []; // Stores {id: studentId, reason: message}

  // --- Process Each Student Row ---
  // Uses FIRST_DATA_ROW from config
  const startRowIndex = 0; // Start from the first row of the fetched `values` array
  const sheetStartRow = FIRST_DATA_ROW; // Keep track of the actual sheet row number for logging
  Logger.log(`Starting grade upload process. Processing ${values.length} potential student rows from sheet row ${sheetStartRow}.`);

  for (let i = startRowIndex; i < values.length; i++) {
    const currentRowInSheet = sheetStartRow + i; // Actual row number in the spreadsheet
    const studentId = values[i][studentIdIndex] ? String(values[i][studentIdIndex]).trim() : null; // Trim SIS ID
    let grade = values[i][gradeIndex];

    if (!studentId) {
        Logger.log(`Skipping sheet row ${currentRowInSheet}: Missing or empty Student ID in column ${SIS_ID_COLUMN}.`);
        skippedMissingIdCount++;
        continue; // Skip this row if SIS ID is missing or empty
    }

    let gradeToUpload;
    // Check for empty/null/undefined explicitly
    if (grade === '' || grade === null || grade === undefined) {
        gradeToUpload = ''; // Canvas API accepts empty string to clear grade
        Logger.log(`Processing sheet row ${currentRowInSheet}: Student ID ${studentId}, Grade: [Empty/Cleared]`);
    }
    // Check if it's a number or a string that can be converted to a finite number
    else if (!isNaN(parseFloat(grade)) && isFinite(grade)) {
        // Allow decimals for upload, Canvas handles rounding/precision based on assignment settings
        gradeToUpload = Number(grade);
        Logger.log(`Processing sheet row ${currentRowInSheet}: Student ID ${studentId}, Grade: ${gradeToUpload}`);
    } else {
        // If it's not empty and not a valid number, skip it
        Logger.log(`Skipping sheet row ${currentRowInSheet}: Invalid grade format '${grade}' for Student ID ${studentId}. Grade must be numeric or empty.`);
        invalidGradeCount++;
        failedStudents.push({ id: studentId, reason: `Invalid grade format ('${grade}')` });
        continue;
    }

    // Call the helper function to update Canvas for this student
    // It will use the global CANVAS_DOMAIN
    const updateResult = updateGradeInCanvas_(apiKey, courseId, assignmentId, studentId, gradeToUpload);

    if (updateResult.success) {
      successCount++;
    } else {
      failCount++;
      failedStudents.push({ id: studentId, reason: updateResult.message });
      // Log detailed failure reason from the API call
      Logger.log(`Failed to update grade for Student ID ${studentId} (Sheet Row ${currentRowInSheet}). Reason: ${updateResult.message}`);
    }
    // Utilities.sleep(100); // Optional: Small delay to avoid hitting API rate limits on very large courses
  } // End of loop

  // --- Log Summary ---
  let summaryMessage = `Grade upload complete for Assignment ID ${assignmentId}.\n`;
  summaryMessage += `------------------------------------------\n`;
  summaryMessage += `Rows Processed: ${values.length}\n`;
  summaryMessage += `Successful Updates: ${successCount}\n`;
  summaryMessage += `Failed Updates: ${failCount}\n`;
  summaryMessage += `Skipped (Invalid Grade Format): ${invalidGradeCount}\n`;
  summaryMessage += `Skipped (Missing Student ID): ${skippedMissingIdCount}\n`;

  if (failedStudents.length > 0) {
    summaryMessage += `------------------------------------------\n`;
    summaryMessage += `Details for Failures/Skipped Grades:\n`;
    // Log only the first 10-15 failures to avoid huge alert boxes
    const failuresToShow = failedStudents.slice(0, 15);
    failuresToShow.forEach(student => {
      summaryMessage += `- ID: ${student.id}, Reason: ${student.reason}\n`;
    });
    if (failedStudents.length > 15) {
        summaryMessage += `... and ${failedStudents.length - 15} more. Check logs for full details.\n`;
    }
  }
  summaryMessage += `------------------------------------------\n`;
  summaryMessage += `Check script logs (View > Logs) for full details.`;

  Logger.log(summaryMessage); // Log the full summary
  // Display a slightly shorter version in the alert
  ui.alert(summaryMessage.replace(/------------------------------------------\n/g, ''));

}

/**
 * Sends the grade update request to the Canvas API for a single student.
 * Uses the global CANVAS_DOMAIN constant.
 * @param {string} accessToken    Canvas API access token.
 * @param {string|number} courseId     Canvas Course ID.
 * @param {string|number} assignmentId Canvas Assignment ID.
 * @param {string} studentId      The student SIS User ID.
 * @param {number|string} grade        The grade to post (numeric or empty string).
 * @return {object} An object {success: boolean, message: string}.
 * @private
 */
function updateGradeInCanvas_(accessToken, courseId, assignmentId, studentId, grade) {
  // Get Canvas Domain
  const canvasDomain = getCanvasDomain();
  if (!canvasDomain) {
    return { success: false, message: "Canvas domain not configured in cell " + CANVAS_DOMAIN_CELL };
  }
  
  const url = `${canvasDomain}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions/sis_user_id:${studentId}`;


  // Construct payload - Canvas expects grade under submission.posted_grade
  const payload = {
    submission: {
      posted_grade: grade
    }
    // You could add comments here too, e.g.:
    // comment: {
    //   text_comment: "Grade updated via Google Sheet script."
    // }
  };

  const options = {
    method: 'put',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Prevent script termination on API errors, handle them manually
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    // Check for successful response codes (200 OK, 201 Created, 204 No Content)
    if (responseCode >= 200 && responseCode < 300) {
      return { success: true, message: `OK (Code: ${responseCode})` };
    } else {
      // Provide specific error messages based on common Canvas API responses
      let errorMessage = `API Error (Code: ${responseCode}). Response: ${responseBody.substring(0, 500)}`; // Default
      if (responseCode === 404) {
          errorMessage = `API Error 404: Submission/User not found for SIS ID '${studentId}' and Assignment ID ${assignmentId}. Verify ID, enrollment, and assignment existence.`;
      } else if (responseCode === 401 || responseCode === 403) {
          errorMessage = `API Error ${responseCode}: Unauthorized/Forbidden. Check API Key permissions for grading.`;
      } else if (responseCode === 400) {
          // Try to parse Canvas error message if available
          let canvasErrorDetail = '';
          try {
              const errorJson = JSON.parse(responseBody);
              if (errorJson.errors && errorJson.errors.length > 0) {
                  canvasErrorDetail = errorJson.errors.map(e => e.message || JSON.stringify(e)).join('; ');
              } else if (errorJson.message) {
                  canvasErrorDetail = errorJson.message;
              }
          } catch (e) { /* Ignore parsing error */ }
          errorMessage = `API Error 400: Bad Request. Check payload format or grade value. Canvas message: ${canvasErrorDetail || responseBody.substring(0, 200)}`;
      } else if (responseCode === 409) {
           errorMessage = `API Error 409: Conflict. Potentially a concurrent edit occurred.`;
      }
      // Log the full response body for debugging if needed, but don't return it all in the message
      Logger.log(`Failed API call to ${url} for student ${studentId}. Code: ${responseCode}. Response: ${responseBody}`);
      return { success: false, message: errorMessage };
    }
  } catch (error) {
    // Catch network errors or other UrlFetchApp issues
    Logger.log(`Network or script error during API call for student ${studentId}: ${error} \nStack: ${error.stack}`);
    return { success: false, message: `Network/Script Error: ${error.message}` };
  }
}

/**
 * Converts a column letter (A, B, ..., Z, AA, AB, ...) to its 1-based column number.
 * @param {string} column The column letter(s).
 * @return {number} The 1-based column number.
 * @private
 */
function columnToNumber_(column) { // Renamed slightly to indicate private use
  let result = 0;
  column = column.toUpperCase(); // Ensure uppercase
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    // Character code for 'A' is 65. Subtract 64 to get 1 for 'A', 2 for 'B', etc.
    result += column.charCodeAt(i) - 64;
  }
  return result;
}

/**
 * ==========================================================================
 * SCRIPT TO UPLOAD CANVAS GRADES
 * Contains functions for uploading grades to Canvas.
 * Requires global constants (CANVAS_DOMAIN, API_KEY_PROPERTY_NAME,
 * COURSE_ID_CELL) defined elsewhere.
 * ==========================================================================
 */

/**
 * Prompts the user to upload all grades from the gradebook to Canvas.
 * This function should be called by a menu item created in the onOpen function.
 */
function promptForGradebookUpload() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Upload Gradebook to Canvas',
    'This will upload ALL grades in the current sheet to Canvas. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    uploadGradebookToCanvas_();
  }
}

/**
 * Uploads all the grades from the gradebook sheet to Canvas.
 * @private
 */
function uploadGradebookToCanvas_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // --- Get API Key from Script Properties ---
  const apiKey = getCanvasApiKey();
  if (!apiKey) {
    return;
  }
  
  // Get Course ID
  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert(`Error: Course ID not found in cell ${COURSE_ID_CELL}.`);
    Logger.log(`Error: Course ID not found in cell ${COURSE_ID_CELL}`);
    return;
  }
  Logger.log(`Using Course ID: ${courseId}`);
  
  try {
    // Define row structure 
    const assignmentIdRow = 6;
    const firstDataRow = 7;
    const firstAssignmentCol = 5; // Column E
    const lastCol = sheet.getLastColumn();
    
    // Exit if there are no assignment columns
    if (lastCol < firstAssignmentCol) {
      ui.alert('No assignment columns found. Please fetch the gradebook first.');
      return;
    }
    
    // Get assignment IDs from row 6
    const assignmentIds = sheet.getRange(assignmentIdRow, firstAssignmentCol, 1, lastCol - firstAssignmentCol + 1).getValues()[0];
    
    // Get SIS User IDs from column C and corresponding Canvas grades
    const lastRow = sheet.getLastRow();
    if (lastRow < firstDataRow) {
      ui.alert('No student data found.');
      return;
    }
    
    const sisUserIds = sheet.getRange(`C${firstDataRow}:C${lastRow}`).getValues();
    const gradesData = sheet.getRange(firstDataRow, firstAssignmentCol, lastRow - firstDataRow + 1, lastCol - firstAssignmentCol + 1).getValues();
    
    // Fetch Canvas users to get their IDs
    ToastManager.showToast('Fetching Canvas users...', 'Gradebook Upload: Step 1/3', 30);
    Logger.log('Fetching Canvas users to get IDs...');
    const users = fetchAllCanvasUsers_(courseId, apiKey, CANVAS_DOMAIN);
    
    // Create lookup map of SIS User ID to Canvas User ID
    const userIdLookup = {};
    users.forEach(user => {
      if (user.sis_user_id) {
        userIdLookup[user.sis_user_id] = user.id;
      }
    });
    
    // Prepare to track statistics
    let totalSubmissions = 0;
    let successfulSubmissions = 0;
    let errorSubmissions = 0;
    let skippedSubmissions = 0;
    
    // Calculate total expected submissions for progress tracking
    let totalExpectedSubmissions = 0;
    for (let i = 0; i < sisUserIds.length; i++) {
      const sisUserId = sisUserIds[i][0];
      if (!sisUserId) continue;
      
      const canvasUserId = userIdLookup[sisUserId];
      if (!canvasUserId) continue;
      
      for (let j = 0; j < assignmentIds.length; j++) {
        const assignmentId = assignmentIds[j];
        if (!assignmentId) continue;
        
        const grade = gradesData[i][j];
        if (grade === null || grade === undefined || grade === '') continue;
        
        totalExpectedSubmissions++;
      }
    }
    
    // Show initial progress toast
    ToastManager.showToast(`Preparing to upload ${totalExpectedSubmissions} grades...`, 'Gradebook Upload: Step 2/3', 30);
    
    // Track upload progress
    let progressCounter = 0;
    let lastProgressPercentage = 0;
    
    // Loop through each student and their grades
    for (let i = 0; i < sisUserIds.length; i++) {
      const sisUserId = sisUserIds[i][0];
      if (!sisUserId) {
        continue; // Skip rows without SIS User ID
      }
      
      const canvasUserId = userIdLookup[sisUserId];
      if (!canvasUserId) {
        Logger.log(`Warning: No Canvas User ID found for SIS User ID: ${sisUserId}`);
        skippedSubmissions += assignmentIds.length;
        continue;
      }
      
      // Loop through each assignment for this student
      for (let j = 0; j < assignmentIds.length; j++) {
        const assignmentId = assignmentIds[j];
        if (!assignmentId) {
          continue; // Skip columns without Assignment ID
        }
        
        const grade = gradesData[i][j];
        if (grade === null || grade === undefined || grade === '') {
          skippedSubmissions++;
          continue; // Skip empty grades
        }
        
        totalSubmissions++;
        
        // Upload the grade to Canvas
        try {
          // Define the endpoint for updating a grade for a specific assignment and user
          const url = `${CANVAS_DOMAIN}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions/${canvasUserId}`;
          
          const payload = {
            'submission': {
              'posted_grade': grade
            }
          };
          
          const options = {
            'method': 'put',
            'contentType': 'application/json',
            'payload': JSON.stringify(payload),
            'headers': {
              'Authorization': 'Bearer ' + apiKey
            },
            'muteHttpExceptions': true
          };
          
          // Make the API call to update the grade
          const response = UrlFetchApp.fetch(url, options);
          const responseCode = response.getResponseCode();
          
          if (responseCode === 200 || responseCode === 201) {
            successfulSubmissions++;
          } else {
            Logger.log(`Error updating grade for user ${canvasUserId}, assignment ${assignmentId}: Status ${responseCode}, Response: ${response.getContentText().substring(0, 200)}`);
            errorSubmissions++;
          }
          
          // Update progress counter and show toast notifications periodically
          progressCounter++;
          const progressPercentage = Math.floor((progressCounter / totalExpectedSubmissions) * 100);
          
          // Update toast every 5% progress to avoid too many notifications
          if (progressPercentage >= lastProgressPercentage + 5 || progressCounter === totalExpectedSubmissions) {
            lastProgressPercentage = progressPercentage;
            const studentName = sheet.getRange(firstDataRow + i, 1, 1, 2).getValues()[0];
            const lastName = studentName[0] || '';
            const firstName = studentName[1] || '';
            const displayName = `${firstName} ${lastName}`.trim();
            const assignmentName = sheet.getRange(1, firstAssignmentCol + j).getValue() || `Assignment ${assignmentId}`;
            
            ToastManager.showToast(
              `Progress: ${progressCounter}/${totalExpectedSubmissions} (${progressPercentage}%)\n` +
              `Last uploaded: ${displayName} - ${assignmentName}: ${grade}`,
              'Gradebook Upload: Step 2/3', 
              30
            );
          }
          
          // Add a small delay to avoid rate limiting
          Utilities.sleep(100);
          
        } catch (error) {
          Logger.log(`Exception updating grade for user ${canvasUserId}, assignment ${assignmentId}: ${error}`);
          errorSubmissions++;
        }
      }
    }
    
    // Show completion message as toast
    ToastManager.showToast(
      `Upload Complete!\n` +
      `Total: ${totalSubmissions}\n` +
      `Success: ${successfulSubmissions}\n` +
      `Errors: ${errorSubmissions}\n` +
      `Skipped: ${skippedSubmissions}`,
      'Gradebook Upload: Complete!', 
      10
    );
    
    // Show completion message as dialog
    ui.alert(
      'Gradebook Upload Complete',
      `Total grades processed: ${totalSubmissions}\n` +
      `Successful uploads: ${successfulSubmissions}\n` +
      `Errors: ${errorSubmissions}\n` +
      `Skipped (empty/not found): ${skippedSubmissions}\n\n` +
      `See logs for details.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error during gradebook upload: ' + error);
    Logger.log('Stack Trace: ' + error.stack);
    ToastManager.showToast('Upload failed: ' + error.message, 'Error', 10);
    ui.alert('An error occurred: ' + error.message + '\n\nPlease check the Script Execution Logs (View > Logs) for more details.');
  }
}

/**
 * Prompts the user to upload grades from a specific range of columns to Canvas.
 * This function should be called by a menu item created in the onOpen function.
 */
function promptForGradeRangeUpload() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask for starting column
  const startResponse = ui.prompt(
    'Upload Grade Range to Canvas',
    'Enter the STARTING column letter (e.g., "F"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startResponse.getSelectedButton() != ui.Button.OK) {
    return; // User cancelled
  }
  
  const startCol = startResponse.getResponseText().toUpperCase().trim();
  if (!/^[A-Z]+$/.test(startCol)) {
    ui.alert('Invalid input. Please enter a valid column letter (e.g., "F").');
    return;
  }
  
  // Ask for ending column
  const endResponse = ui.prompt(
    'Upload Grade Range to Canvas',
    'Enter the ENDING column letter (e.g., "H"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endResponse.getSelectedButton() != ui.Button.OK) {
    return; // User cancelled
  }
  
  const endCol = endResponse.getResponseText().toUpperCase().trim();
  if (!/^[A-Z]+$/.test(endCol)) {
    ui.alert('Invalid input. Please enter a valid column letter (e.g., "H").');
    return;
  }
  
  // Convert column letters to numbers
  const startColNum = columnLetterToNumber_(startCol);
  const endColNum = columnLetterToNumber_(endCol);
  
  if (startColNum > endColNum) {
    ui.alert('Invalid range. Starting column must come before ending column.');
    return;
  }
  
  // Confirm upload
  const confirmation = ui.alert(
    'Confirm Grade Upload',
    `This will upload grades from column ${startCol} to column ${endCol} to Canvas. Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation == ui.Button.YES) {
    uploadGradeRangeToCanvas_(startColNum, endColNum);
  }
}

/**
 * Uploads grades from a specific range of columns to Canvas.
 * @param {number} startColNum The 1-based starting column number.
 * @param {number} endColNum The 1-based ending column number.
 * @private
 */
function uploadGradeRangeToCanvas_(startColNum, endColNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // --- Get API Key from Script Properties ---
  const apiKey = getCanvasApiKey();
  if (!apiKey) {
    return;
  }
  
  // Get Course ID
  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert(`Error: Course ID not found in cell ${COURSE_ID_CELL}.`);
    Logger.log(`Error: Course ID not found in cell ${COURSE_ID_CELL}`);
    return;
  }
  Logger.log(`Using Course ID: ${courseId}`);
  
  try {
    // Define row structure 
    const assignmentIdRow = 6;
    const firstDataRow = 7;
    
    // Calculate the number of columns to upload
    const numCols = endColNum - startColNum + 1;
    
    // Exit if startColNum is less than 5 (Column E)
    if (startColNum < 5) {
      ui.alert('Error: Starting column must be column E or later (student data is in columns A-D).');
      return;
    }
    
    // Get assignment IDs from row 6 for the specified range
    const assignmentIds = sheet.getRange(assignmentIdRow, startColNum, 1, numCols).getValues()[0];
    
    // Check if there are valid assignment IDs
    let hasValidIds = false;
    for (let i = 0; i < assignmentIds.length; i++) {
      if (assignmentIds[i]) {
        hasValidIds = true;
        break;
      }
    }
    
    if (!hasValidIds) {
      ui.alert('No valid assignment IDs found in row 6 of the selected columns. Please ensure the gradebook is properly set up.');
      return;
    }
    
    // Get SIS User IDs from column C and corresponding Canvas grades
    const lastRow = sheet.getLastRow();
    if (lastRow < firstDataRow) {
      ui.alert('No student data found.');
      return;
    }
    
    const sisUserIds = sheet.getRange(`C${firstDataRow}:C${lastRow}`).getValues();
    const gradesData = sheet.getRange(firstDataRow, startColNum, lastRow - firstDataRow + 1, numCols).getValues();
    
    // Fetch Canvas users to get their IDs
    ToastManager.showToast('Fetching Canvas users...', 'Grade Range Upload: Step 1/3', 30);
    Logger.log('Fetching Canvas users to get IDs...');
    const users = fetchAllCanvasUsers_(courseId, apiKey, CANVAS_DOMAIN);
    
    // Create lookup map of SIS User ID to Canvas User ID
    const userIdLookup = {};
    users.forEach(user => {
      if (user.sis_user_id) {
        userIdLookup[user.sis_user_id] = user.id;
      }
    });
    
    // Prepare to track statistics
    let totalSubmissions = 0;
    let successfulSubmissions = 0;
    let errorSubmissions = 0;
    let skippedSubmissions = 0;
    
    // Calculate total expected submissions for progress tracking
    let totalExpectedSubmissions = 0;
    for (let i = 0; i < sisUserIds.length; i++) {
      const sisUserId = sisUserIds[i][0];
      if (!sisUserId) continue;
      
      const canvasUserId = userIdLookup[sisUserId];
      if (!canvasUserId) continue;
      
      for (let j = 0; j < assignmentIds.length; j++) {
        const assignmentId = assignmentIds[j];
        if (!assignmentId) continue;
        
        const grade = gradesData[i][j];
        if (grade === null || grade === undefined || grade === '') continue;
        
        totalExpectedSubmissions++;
      }
    }
    
    // Show initial progress toast
    ToastManager.showToast(`Preparing to upload ${totalExpectedSubmissions} grades...`, 'Grade Range Upload: Step 2/3', 30);
    
    // Track upload progress
    let progressCounter = 0;
    let lastProgressPercentage = 0;
    
    // Loop through each student and their grades
    for (let i = 0; i < sisUserIds.length; i++) {
      const sisUserId = sisUserIds[i][0];
      if (!sisUserId) {
        continue; // Skip rows without SIS User ID
      }
      
      const canvasUserId = userIdLookup[sisUserId];
      if (!canvasUserId) {
        Logger.log(`Warning: No Canvas User ID found for SIS User ID: ${sisUserId}`);
        skippedSubmissions += assignmentIds.length;
        continue;
      }
      
      // Loop through each assignment for this student
      for (let j = 0; j < assignmentIds.length; j++) {
        const assignmentId = assignmentIds[j];
        if (!assignmentId) {
          continue; // Skip columns without Assignment ID
        }
        
        const grade = gradesData[i][j];
        if (grade === null || grade === undefined || grade === '') {
          skippedSubmissions++;
          continue; // Skip empty grades
        }
        
        totalSubmissions++;
        
        // Upload the grade to Canvas
        try {
          // Define the endpoint for updating a grade for a specific assignment and user
          const url = `${CANVAS_DOMAIN}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions/${canvasUserId}`;
          
          const payload = {
            'submission': {
              'posted_grade': grade
            }
          };
          
          const options = {
            'method': 'put',
            'contentType': 'application/json',
            'payload': JSON.stringify(payload),
            'headers': {
              'Authorization': 'Bearer ' + apiKey
            },
            'muteHttpExceptions': true
          };
          
          // Make the API call to update the grade
          const response = UrlFetchApp.fetch(url, options);
          const responseCode = response.getResponseCode();
          
          if (responseCode === 200 || responseCode === 201) {
            successfulSubmissions++;
          } else {
            Logger.log(`Error updating grade for user ${canvasUserId}, assignment ${assignmentId}: Status ${responseCode}, Response: ${response.getContentText().substring(0, 200)}`);
            errorSubmissions++;
          }
          
          // Update progress counter and show toast notifications periodically
          progressCounter++;
          const progressPercentage = Math.floor((progressCounter / totalExpectedSubmissions) * 100);
          
          // Update toast every 5% progress to avoid too many notifications
          if (progressPercentage >= lastProgressPercentage + 5 || progressCounter === totalExpectedSubmissions) {
            lastProgressPercentage = progressPercentage;
            const studentName = sheet.getRange(firstDataRow + i, 1, 1, 2).getValues()[0];
            const lastName = studentName[0] || '';
            const firstName = studentName[1] || '';
            const displayName = `${firstName} ${lastName}`.trim();
            const assignmentName = sheet.getRange(1, startColNum + j).getValue() || `Assignment ${assignmentId}`;
            
            ToastManager.showToast(
              `Progress: ${progressCounter}/${totalExpectedSubmissions} (${progressPercentage}%)\n` +
              `Last uploaded: ${displayName} - ${assignmentName}: ${grade}`,
              'Grade Range Upload: Step 2/3', 
              30
            );
          }
          
          // Add a small delay to avoid rate limiting
          Utilities.sleep(100);
          
        } catch (error) {
          Logger.log(`Exception updating grade for user ${canvasUserId}, assignment ${assignmentId}: ${error}`);
          errorSubmissions++;
        }
      }
    }
    
    // Show completion message as toast
    ToastManager.showToast(
      `Upload Complete!\n` +
      `Total: ${totalSubmissions}\n` +
      `Success: ${successfulSubmissions}\n` +
      `Errors: ${errorSubmissions}\n` +
      `Skipped: ${skippedSubmissions}`,
      'Grade Range Upload: Complete!', 
      10
    );
    
    // Show completion message as dialog
    ui.alert(
      'Grade Range Upload Complete',
      `Total grades processed: ${totalSubmissions}\n` +
      `Successful uploads: ${successfulSubmissions}\n` +
      `Errors: ${errorSubmissions}\n` +
      `Skipped (empty/not found): ${skippedSubmissions}\n\n` +
      `See logs for details.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error during grade range upload: ' + error);
    Logger.log('Stack Trace: ' + error.stack);
    ToastManager.showToast('Upload failed: ' + error.message, 'Error', 10);
    ui.alert('An error occurred: ' + error.message + '\n\nPlease check the Script Execution Logs (View > Logs) for more details.');
  }
}
