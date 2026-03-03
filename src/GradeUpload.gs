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
 * Reads grades from the specified sheet column and uploads them to Canvas
 * using parallel requests via UrlFetchApp.fetchAll.
 * @param {string} gradeColumn The letter of the column containing grades to upload.
 * @private
 */
function updateCanvasGrades_(gradeColumn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const apiKey = getCanvasApiKey();
  if (!apiKey) return;

  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert(`Error: Course ID not found in cell ${COURSE_ID_CELL}.`);
    Logger.log(`Error: Course ID not found in cell ${COURSE_ID_CELL}`);
    return;
  }
  Logger.log(`Using Course ID: ${courseId}`);

  const assignmentIdCell = `${gradeColumn}${ASSIGNMENT_ID_HEADER_ROW}`;
  const assignmentId = sheet.getRange(assignmentIdCell).getValue();
  if (!assignmentId) {
    ui.alert(`Error: Assignment ID not found in cell ${assignmentIdCell}. Please ensure the Canvas Assignment ID is in this cell.`);
    Logger.log(`Error: Assignment ID not found in cell ${assignmentIdCell}`);
    return;
  }
  Logger.log(`Using Assignment ID: ${assignmentId}`);
  ToastManager.showToast(`Uploading grades for Assignment ID: ${assignmentId}`, 'Grade Upload: Starting', 10);

  const canvasDomain = getCanvasDomain();
  if (!canvasDomain) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) {
    ui.alert("No student data found in the sheet starting from row " + FIRST_DATA_ROW);
    Logger.log("No student data rows found.");
    return;
  }

  const studentIdColNum = columnLetterToNumber_(SIS_ID_COLUMN);
  const gradeColNum = columnLetterToNumber_(gradeColumn);
  const firstCol = Math.min(studentIdColNum, gradeColNum);
  const lastCol = Math.max(studentIdColNum, gradeColNum);

  const dataRange = sheet.getRange(FIRST_DATA_ROW, firstCol, lastRow - FIRST_DATA_ROW + 1, lastCol - firstCol + 1);
  const values = dataRange.getValues();
  Logger.log(`Read data from range ${dataRange.getA1Notation()}.`);

  const studentIdIndex = studentIdColNum - firstCol;
  const gradeIndex = gradeColNum - firstCol;

  let successCount = 0;
  let failCount = 0;
  let invalidGradeCount = 0;
  let skippedMissingIdCount = 0;
  let failedStudents = [];

  // --- Build requests ---
  const requests = [];
  const requestMeta = [];
  Logger.log(`Building grade upload requests from ${values.length} potential student rows.`);

  for (let i = 0; i < values.length; i++) {
    const currentRowInSheet = FIRST_DATA_ROW + i;
    const studentId = values[i][studentIdIndex] ? String(values[i][studentIdIndex]).trim() : null;
    const grade = values[i][gradeIndex];

    if (!studentId) {
      Logger.log(`Skipping sheet row ${currentRowInSheet}: Missing or empty Student ID in column ${SIS_ID_COLUMN}.`);
      skippedMissingIdCount++;
      continue;
    }

    let gradeToUpload;
    if (grade === '' || grade === null || grade === undefined) {
      gradeToUpload = '';
      Logger.log(`Processing sheet row ${currentRowInSheet}: grade is empty/cleared.`);
    } else if (!isNaN(parseFloat(grade)) && isFinite(grade)) {
      gradeToUpload = Number(grade);
      Logger.log(`Processing sheet row ${currentRowInSheet}: valid numeric grade.`);
    } else {
      Logger.log(`Skipping sheet row ${currentRowInSheet}: invalid grade format.`);
      invalidGradeCount++;
      failedStudents.push({ id: studentId, row: currentRowInSheet, reason: `Invalid grade format` });
      continue;
    }

    requests.push({
      url: `${canvasDomain}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions/sis_user_id:${studentId}`,
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify({ submission: { posted_grade: gradeToUpload } }),
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true
    });
    requestMeta.push({ id: studentId, row: currentRowInSheet });
  }

  // --- Fire all requests in parallel ---
  if (requests.length > 0) {
    ToastManager.showToast(`Uploading ${requests.length} grades...`, 'Grade Upload: In Progress', 30);
    const responses = UrlFetchApp.fetchAll(requests);
    responses.forEach((response, idx) => {
      const { id: studentId, row: rowNum } = requestMeta[idx];
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode >= 200 && responseCode < 300) {
        successCount++;
      } else {
        failCount++;
        let errorMessage = `API Error (Code: ${responseCode}). Response: ${responseBody.substring(0, 500)}`;
        if (responseCode === 404) {
          errorMessage = `API Error 404: Submission/User not found for Assignment ID ${assignmentId}. Verify student enrollment and assignment existence.`;
        } else if (responseCode === 401 || responseCode === 403) {
          errorMessage = `API Error ${responseCode}: Unauthorized/Forbidden. Check API Key permissions for grading.`;
        } else if (responseCode === 400) {
          let canvasErrorDetail = '';
          try {
            const errorJson = JSON.parse(responseBody);
            if (errorJson.errors && errorJson.errors.length > 0) {
              canvasErrorDetail = errorJson.errors.map(e => e.message || JSON.stringify(e)).join('; ');
            } else if (errorJson.message) {
              canvasErrorDetail = errorJson.message;
            }
          } catch (e) { /* Ignore parsing error */ }
          errorMessage = `API Error 400: Bad Request. Canvas message: ${canvasErrorDetail || responseBody.substring(0, 200)}`;
        } else if (responseCode === 409) {
          errorMessage = `API Error 409: Conflict. Potentially a concurrent edit occurred.`;
        }
        failedStudents.push({ id: studentId, row: rowNum, reason: errorMessage });
        Logger.log(`Failed to update grade for sheet row ${rowNum} (status ${responseCode}): ${errorMessage}`);
      }
    });
  }

  // --- Summary ---
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
    failedStudents.slice(0, 15).forEach((student, idx) => {
      summaryMessage += `- Row ${student.row}: ${student.reason}\n`;
    });
    if (failedStudents.length > 15) {
      summaryMessage += `... and ${failedStudents.length - 15} more. Check logs for full details.\n`;
    }
  }
  summaryMessage += `------------------------------------------\n`;
  summaryMessage += `Check script logs (View > Logs) for full details.`;

  Logger.log(summaryMessage);
  ToastManager.showCompletionToast('Grade upload complete. See dialog for details.', 'Grade Upload', 5);
  ui.alert(summaryMessage.replace(/------------------------------------------\n/g, ''));
}

/**
 * ==========================================================================
 * SCRIPT TO UPLOAD CANVAS GRADES
 * Contains functions for uploading grades to Canvas.
 * Requires global constants (API_KEY_PROPERTY_NAME, COURSE_ID_CELL,
 * CANVAS_DOMAIN_CELL, ASSIGNMENT_ID_HEADER_ROW, FIRST_DATA_ROW) defined elsewhere.
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
  const lastCol = sheet.getLastColumn();
  if (lastCol < FIRST_ASSIGNMENT_COL) {
    ui.alert('No assignment columns found. Please fetch the gradebook first.');
    return;
  }
  uploadGradeColumnsToCanvas_(FIRST_ASSIGNMENT_COL, lastCol, 'Gradebook Upload');
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
  if (startColNum < FIRST_ASSIGNMENT_COL) {
    ui.alert('Error: Starting column must be column E or later (student data is in columns A-D).');
    return;
  }
  const numCols = endColNum - startColNum + 1;
  const assignmentIds = sheet.getRange(ASSIGNMENT_ID_HEADER_ROW, startColNum, 1, numCols).getValues()[0];
  if (!assignmentIds.some(id => !!id)) {
    ui.alert(`No valid assignment IDs found in row ${ASSIGNMENT_ID_HEADER_ROW} of the selected columns. Please ensure the gradebook is properly set up.`);
    return;
  }
  uploadGradeColumnsToCanvas_(startColNum, endColNum, 'Grade Range Upload');
}

/**
 * Core function for uploading grades from a column range to Canvas.
 * Groups requests by assignment and fires all student updates in parallel
 * using UrlFetchApp.fetchAll, replacing the previous one-request-per-student loop.
 * @param {number} startColNum 1-based starting column number (must be >= 5).
 * @param {number} endColNum 1-based ending column number.
 * @param {string} operationLabel Label used in toast messages (e.g. 'Gradebook Upload').
 * @private
 */
function uploadGradeColumnsToCanvas_(startColNum, endColNum, operationLabel) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const apiKey = getCanvasApiKey();
  if (!apiKey) return;

  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    ui.alert(`Error: Course ID not found in cell ${COURSE_ID_CELL}.`);
    Logger.log(`Error: Course ID not found in cell ${COURSE_ID_CELL}`);
    return;
  }
  Logger.log(`Using Course ID: ${courseId}`);

  const canvasDomain = getCanvasDomain();
  if (!canvasDomain) return;

  try {
    const numCols = endColNum - startColNum + 1;
    const lastRow = sheet.getLastRow();
    if (lastRow < FIRST_DATA_ROW) {
      ui.alert('No student data found.');
      return;
    }

    const assignmentIds = sheet.getRange(ASSIGNMENT_ID_HEADER_ROW, startColNum, 1, numCols).getValues()[0];
    const sisUserIds = sheet.getRange(`${SIS_ID_COLUMN}${FIRST_DATA_ROW}:${SIS_ID_COLUMN}${lastRow}`).getValues();
    const gradesData = sheet.getRange(FIRST_DATA_ROW, startColNum, lastRow - FIRST_DATA_ROW + 1, numCols).getValues();

    // Pre-read assignment names for progress toasts
    const assignmentNames = sheet.getRange(1, startColNum, 1, numCols).getValues()[0];

    let totalSubmissions = 0;
    let successfulSubmissions = 0;
    let errorSubmissions = 0;
    let skippedSubmissions = 0;
    let progressCounter = 0;

    // Count total non-empty grade cells for progress tracking
    let totalExpectedSubmissions = 0;
    for (let i = 0; i < sisUserIds.length; i++) {
      const sisUserId = sisUserIds[i][0];
      if (!sisUserId) continue;
      for (let j = 0; j < assignmentIds.length; j++) {
        if (!assignmentIds[j]) continue;
        const grade = gradesData[i][j];
        if (grade === null || grade === undefined || grade === '') continue;
        totalExpectedSubmissions++;
      }
    }

    ToastManager.showToast(`Preparing to upload ${totalExpectedSubmissions} grades...`, `${operationLabel}: Step 1/2`, 30);

    // Upload per assignment: collect all student requests then fire in parallel with fetchAll
    for (let j = 0; j < assignmentIds.length; j++) {
      const assignmentId = assignmentIds[j];
      if (!assignmentId) continue;

      const requests = [];
      for (let i = 0; i < sisUserIds.length; i++) {
        const sisUserId = sisUserIds[i][0];
        if (!sisUserId) continue;

        const grade = gradesData[i][j];
        if (grade === null || grade === undefined || grade === '') {
          skippedSubmissions++;
          continue;
        }

        requests.push({
          url: `${canvasDomain}/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions/sis_user_id:${sisUserId}`,
          method: 'put',
          contentType: 'application/json',
          payload: JSON.stringify({ submission: { posted_grade: grade } }),
          headers: { 'Authorization': 'Bearer ' + apiKey },
          muteHttpExceptions: true
        });
        totalSubmissions++;
      }

      if (requests.length === 0) continue;

      // Fire all student requests for this assignment simultaneously
      const responses = UrlFetchApp.fetchAll(requests);
      responses.forEach((response, idx) => {
        const responseCode = response.getResponseCode();
        if (responseCode === 200 || responseCode === 201) {
          successfulSubmissions++;
        } else {
          Logger.log(`Error uploading grade for request ${idx}, assignment ${assignmentId}: Status ${responseCode}, Response: ${response.getContentText().substring(0, 200)}`);
          errorSubmissions++;
        }
        progressCounter++;
      });

      // Show progress after each assignment batch completes
      const progressPercentage = Math.floor((progressCounter / totalExpectedSubmissions) * 100);
      const assignmentName = assignmentNames[j] || `Assignment ${assignmentId}`;
      ToastManager.showToast(
        `Progress: ${progressCounter}/${totalExpectedSubmissions} (${progressPercentage}%)\nCompleted: ${assignmentName}`,
        `${operationLabel}: Step 2/2`,
        30
      );
    }

    ToastManager.showToast(
      `Upload Complete!\nTotal: ${totalSubmissions}\nSuccess: ${successfulSubmissions}\nErrors: ${errorSubmissions}\nSkipped: ${skippedSubmissions}`,
      `${operationLabel}: Complete!`,
      10
    );

    ui.alert(
      `${operationLabel} Complete`,
      `Total grades processed: ${totalSubmissions}\n` +
      `Successful uploads: ${successfulSubmissions}\n` +
      `Errors: ${errorSubmissions}\n` +
      `Skipped (empty/not found): ${skippedSubmissions}\n\n` +
      `See logs for details.`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`Error during ${operationLabel}: ` + error);
    Logger.log('Stack Trace: ' + error.stack);
    ToastManager.showToast('Upload failed: ' + error.message, 'Error', 10);
    ui.alert('An error occurred: ' + error.message + '\n\nPlease check the Script Execution Logs (View > Logs) for more details.');
  }
}
