/**
 * ==========================================================================
 * SCRIPT TO FETCH CANVAS USERS
 * Assumes global constants (CANVAS_DOMAIN, API_KEY_PROPERTY_NAME,
 * COURSE_ID_CELL, SIS_ID_COLUMN, FIRST_DATA_ROW, USER_INFO_HEADER_ROW)
 * are defined elsewhere in the project.
 * ==========================================================================
 */

/**
 * Main function to trigger the Canvas user fetching process.
 * Can be run from the Apps Script editor or the custom menu.
 */
function fetchAndPopulateCanvasUsers() {
  // --- Get Canvas Domain from cell B4 ---
  const canvasDomain = getCanvasDomain();
  if (!canvasDomain) {
    // getCanvasDomain() already showed an alert, just exit
    return;
  }

  // --- Get API Key (from cell B5, Script Properties, or user prompt) ---
  const token = getCanvasApiKey();
  if (!token) {
    // getCanvasApiKey() already showed an alert, just exit
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // --- 1. Get Course ID ---
  const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
  if (!courseId) {
    SpreadsheetApp.getUi().alert('Error: Course ID not found in cell ' + COURSE_ID_CELL);
    Logger.log('Error: Course ID not found in cell ' + COURSE_ID_CELL);
    return;
  }
  Logger.log('Fetching users for Course ID: ' + courseId);

  try {
    // --- 2. Fetch Users from Canvas (Handles Pagination) ---
    // Pass the retrieved token and canvasDomain to the helper function
    const users = fetchAllCanvasUsers_(courseId, token, canvasDomain);
    Logger.log('Fetched ' + users.length + ' users from Canvas.');

    // --- 3. Prepare Data for Sheet ---
    const dataForSheet = users.map(user => {
      let lastName = '';
      let firstName = '';
      if (user.sortable_name) {
        const nameParts = user.sortable_name.split(', ');
        lastName = nameParts[0] ? nameParts[0].trim() : '';
        firstName = nameParts[1] ? nameParts[1].trim() : '';
      } else if (user.name) {
        const nameParts = user.name.split(' ');
        firstName = nameParts[0] ? nameParts[0].trim() : '';
        lastName = nameParts.slice(1).join(' ').trim();
      }
      if (firstName) {
        firstName = firstName.split(' ')[0];
      }

      const sisUserId = user.sis_user_id || '';
      if (!sisUserId) {
           Logger.log(`User ${user.name || user.sortable_name} (Canvas ID: ${user.id}) has no sis_user_id returned.`);
      }
      const email = user.email || '';

      return [
        lastName,   // Column A
        firstName,  // Column B
        sisUserId,  // Column C (Uses global SIS_ID_COLUMN implicitly via writeData_)
        email       // Column D
      ];
    });

    // --- 4. Write Data to Sheet ---
    clearSheetData_(sheet, FIRST_DATA_ROW, USER_INFO_HEADER_ROW); // Pass relevant constants
    writeHeaders_(sheet, USER_INFO_HEADER_ROW); // Pass relevant constant
    writeData_(sheet, dataForSheet, FIRST_DATA_ROW); // Pass relevant constant

    Logger.log('Successfully populated sheet with user data.');
    SpreadsheetApp.getUi().alert('Success! Sheet updated with ' + users.length + ' users from Canvas.');

  } catch (error) {
    Logger.log('Error fetching or processing Canvas data: ' + error);
    Logger.log('Stack Trace: ' + error.stack);
    SpreadsheetApp.getUi().alert('An error occurred: ' + error.message + '\n\nPlease check the Script Execution Logs (View > Logs) for more details.');
  }
}

/**
 * Fetches all users for a given course ID from the Canvas API, handling pagination.
 * @param {string|number} courseId The Canvas Course ID.
 * @param {string} apiToken The Canvas API token.
 * @param {string} canvasDomain The base Canvas domain URL. // *** ADDED parameter ***
 * @return {Array<Object>} An array of Canvas user objects.
 * @private
 */
function fetchAllCanvasUsers_(courseId, apiToken, canvasDomain) { // *** ADDED canvasDomain ***
  let allUsers = [];
  // *** MODIFIED: Use canvasDomain parameter ***
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
      nextPageUrl = parseLinkHeader_(linkHeader); // Assumes parseLinkHeader_ is defined elsewhere or below
      Logger.log("Next page URL: " + (nextPageUrl || 'None'));
    } else {
      Logger.log('API Error - Response Code: ' + responseCode);
      Logger.log('API Error - Response Body: ' + responseBody);
      throw new Error(`Canvas API request failed with status code ${responseCode}. Check Course ID (${courseId}), API Token validity, and Permissions. Response: ${responseBody.substring(0, 500)}`);
    }
  } // end while

  return allUsers;
}

/**
 * Parses the 'Link' header from Canvas API responses to find the 'next' page URL.
 * (Ensure this function exists ONCE in your project)
 * @param {string} linkHeader The 'Link' header string.
 * @return {string|null} The URL for the next page, or null if not found.
 * @private
 */
function parseLinkHeader_(linkHeader) {
  if (!linkHeader) {
    return null;
  }
  const links = linkHeader.split(',');
  const nextLink = links.find(link => link.includes('rel="next"'));
  if (nextLink) {
    const match = nextLink.match(/<(.*?)>/);
    if (match && match[1]) {
      return match[1];
    }
  }
  return null;
}


/**
 * Clears data and formatting from the first data row downwards in columns A to D.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number} firstDataRow The row number where student data begins. // *** ADDED parameter ***
 * @param {number} headerRow The row number where headers are located. // *** ADDED parameter ***
 * @private
 */
function clearSheetData_(sheet, firstDataRow, headerRow) { // *** ADDED parameters ***
  const lastRow = sheet.getLastRow();
  if (lastRow >= firstDataRow) {
    const numRowsToClear = lastRow - firstDataRow + 1;
    sheet.getRange(firstDataRow, 1, numRowsToClear, 4).clear({contentsOnly: false, formatOnly: false});
    Logger.log(`Cleared data and formatting from row ${firstDataRow} to ${lastRow} in columns A:D.`);
  } else {
     Logger.log(`No data to clear starting from row ${firstDataRow}. Last data row is ${lastRow}.`);
  }
  if (sheet.getMaxRows() >= headerRow) {
     try {
       sheet.getRange(headerRow, 1, 1, 4).clear({contentsOnly: false, formatOnly: false});
       Logger.log(`Cleared header row ${headerRow}.`);
     } catch (e) {
       Logger.log(`Could not clear header row ${headerRow}: ${e}`);
     }
  }
}

/**
 * Writes headers to the specified header row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number} headerRow The row number where headers should be written. // *** ADDED parameter ***
 * @private
 */
function writeHeaders_(sheet, headerRow) { // *** ADDED parameter ***
  const headers = [['Last Name', 'First Name', 'SIS User ID', 'Email']];
  sheet.getRange(headerRow, 1, 1, 4) // Use parameter
       .setValues(headers)
       .setFontWeight('bold')
       .setHorizontalAlignment('center');
  Logger.log('Wrote headers to row ' + headerRow);
}

/**
 * Writes the fetched user data array to the sheet starting from the specified row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {Array<Array<string|number>>} data The 2D array of user data.
 * @param {number} firstDataRow The row number where writing should begin. // *** ADDED parameter ***
 * @private
 */
function writeData_(sheet, data, firstDataRow) { // *** ADDED parameter ***
  if (data && data.length > 0) {
    const targetRange = sheet.getRange(firstDataRow, 1, data.length, data[0].length); // Use parameter
    targetRange.setValues(data);
    Logger.log('Wrote ' + data.length + ' rows of user data starting at row ' + firstDataRow);
  } else {
     Logger.log('No user data provided to write.');
  }
}
