// --- Menu Setup ---
/**
 * Adds custom menu items to the spreadsheet UI.
 * NOTE: If you have other scripts in this project with an onOpen,
 * combine them into a single onOpen function.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Set up the Canvas Tools menu
  ui.createMenu('Canvas Tools')
      .addItem('Fetch Course Users', 'fetchAndPopulateCanvasUsers')
      .addSeparator()
      .addItem('Fetch Assignment Grades', 'promptForGradeColumn_')
      .addItem('Fetch Complete Gradebook', 'fetchCompleteCanvasGradebook')
      .addSeparator()
      .addItem('Upload Assignment Grades to Canvas', 'promptForGradeUploadColumn_')
      .addItem('Upload Grade Range to Canvas', 'promptForGradeRangeUpload')
      .addItem('Upload Complete Gradebook to Canvas', 'promptForGradebookUpload')
      .addToUi();
  
  // Check if Canvas domain is set
  try {
    const canvasDomain = sheet.getRange(CANVAS_DOMAIN_CELL).getValue();
    if (!canvasDomain) {
      // Only show this on new sheet creation or first run
      ui.alert(
        'Canvas Tools Setup',
        `Please enter your Canvas domain in cell ${CANVAS_DOMAIN_CELL} (e.g., https://canvas.chapman.edu)\n\n` +
        `You'll also need to set up your Canvas API Key in Script Properties.`,
        ui.ButtonSet.OK
      );
      
      // Add a helpful label and placeholder in cell A4
      if (!sheet.getRange('A4').getValue()) {
        sheet.getRange('A4').setValue('Canvas Domain:');
        sheet.getRange(CANVAS_DOMAIN_CELL).setValue('');
      }
    }
  } catch (e) {
    // Ignore errors, this is just a helper
    Logger.log('Error checking Canvas domain: ' + e);
  }
}

/**
 * ==========================================================================
 * GLOBAL CONFIGURATION CONSTANTS
 * Define these ONCE in your project (e.g., in Config.gs or Canvas Tools.gs)
 * Remove duplicate declarations from other script files.
 * ==========================================================================
 */

// --- Canvas API Settings ---

/**
 * The cell containing the Canvas domain for the current sheet.
 * User should enter their institution's Canvas URL (e.g., 'https://canvas.chapman.edu').
 */
const CANVAS_DOMAIN_CELL = 'B4';

/**
 * The cell containing the Canvas API Key.
 * This is optional - if present, the key will be read from here.
 * If not present, the user will be prompted and the key will be stored in Script Properties.
 */
const API_KEY_CELL = 'B5';

/**
 * The name of the Script Property used to store your Canvas API Key.
 * This property is used if the key is not found in API_KEY_CELL.
 */
const API_KEY_PROPERTY_NAME = 'CANVAS_API_KEY';


// --- Spreadsheet Layout Settings ---

/**
 * The cell containing the Canvas Course ID for the current sheet.
 */
const COURSE_ID_CELL = 'B3';

/**
 * The column letter containing the Student Information System (SIS) User ID.
 * This ID is used for matching students when fetching or uploading grades.
 */
const SIS_ID_COLUMN = 'C';

/**
 * The first row number containing actual student data (below headers).
 */
const FIRST_DATA_ROW = 7;

/**
 * The row number where general data headers (like Last Name, First Name, SIS User ID) are located.
 * Used by the script that fetches user information.
 */
const USER_INFO_HEADER_ROW = 6;

/**
 * The row number where the specific Canvas Assignment ID should be placed
 * in the header of the column used for fetching or uploading grades for that assignment.
 */
const ASSIGNMENT_ID_HEADER_ROW = 6;


// --- Notes ---
// - The actual API key is NOT stored here; it's retrieved using API_KEY_PROPERTY_NAME.
// - Canvas domain is now retrieved from cell B4 in the active sheet.
// - Scripts that need these values will refer to these globally defined constants.
// - Ensure you only have ONE `onOpen()` function defined across all your .gs files.


/**
 * Gets the Canvas domain from the spreadsheet cell B4.
 * If the cell is empty or invalid, uses a default domain or shows an error.
 * @return {string} The Canvas domain URL.
 */
function getCanvasDomain() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const domainFromCell = sheet.getRange(CANVAS_DOMAIN_CELL).getValue();
  
  // Validate the domain
  if (!domainFromCell) {
    // If empty, alert the user and provide a default
    SpreadsheetApp.getUi().alert(
      'Canvas Domain Missing',
      `Please enter your Canvas domain URL in cell ${CANVAS_DOMAIN_CELL} (e.g., https://canvas.chapman.edu)`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return ''; // Return empty to cause function to exit in calling code
  }
  
  // Ensure domain starts with https://
  let formattedDomain = domainFromCell.trim();
  if (!formattedDomain.startsWith('http://') && !formattedDomain.startsWith('https://')) {
    formattedDomain = 'https://' + formattedDomain;
  }
  
  // Remove trailing slash if present
  if (formattedDomain.endsWith('/')) {
    formattedDomain = formattedDomain.slice(0, -1);
  }
  
  return formattedDomain;
}

/**
 * Gets the Canvas API key, checking in the following order:
 * 1. Cell B5 in the spreadsheet
 * 2. Script Properties
 * 3. If not found in either place, prompts the user to enter it
 * 
 * @return {string|null} The Canvas API key if found or entered, null if canceled
 */
function getCanvasApiKey() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // First, check cell B5
  const apiKeyFromCell = sheet.getRange(API_KEY_CELL).getValue();
  if (apiKeyFromCell && apiKeyFromCell.trim()) {
    Logger.log("Using API key from cell " + API_KEY_CELL);
    // If API key is in the cell, save it to Script Properties for future use
    PropertiesService.getScriptProperties().setProperty(API_KEY_PROPERTY_NAME, apiKeyFromCell.trim());
    return apiKeyFromCell.trim();
  }
  
  // Second, check Script Properties
  const apiKeyFromProperties = PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY_NAME);
  if (apiKeyFromProperties) {
    Logger.log("Using API key from Script Properties");
    return apiKeyFromProperties;
  }
  
  // If not found in either place, prompt the user
  const response = ui.prompt(
    'Canvas API Key Required',
    'Please enter your Canvas API Key. This key will be securely stored ' +
    'in the Script Properties and not visible in the spreadsheet.\n\n' +
    'To generate an API key:\n' +
    '1. Log into Canvas\n' +
    '2. Go to Account > Settings\n' +
    '3. Scroll to "Approved Integrations"\n' +
    '4. Click "New Access Token"\n' +
    '5. Enter a purpose (e.g., "Google Sheets Integration")\n' +
    '6. Copy the generated token',
    ui.ButtonSet.OK_CANCEL
  );
  
  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      // Store the API key in Script Properties
      PropertiesService.getScriptProperties().setProperty(API_KEY_PROPERTY_NAME, apiKey);
      
      // Confirm to the user but don't show the key
      ui.alert(
        'API Key Saved',
        'Your Canvas API key has been securely stored in the Script Properties. ' +
        'You won\'t need to enter it again on this spreadsheet.',
        ui.ButtonSet.OK
      );
      
      return apiKey;
    } else {
      ui.alert('Error', 'No API key was entered. Canvas Tools requires an API key to function.', ui.ButtonSet.OK);
      return null;
    }
  } else {
    // User canceled
    ui.alert('Canvas API Access Required', 'Canvas Tools requires an API key to access Canvas data.', ui.ButtonSet.OK);
    return null;
  }
}

/**
 * Gets the Canvas domain from the spreadsheet cell B4.
 * If the cell is empty or invalid, shows an error.
 * @return {string|null} The Canvas domain URL or null if not properly configured
 */
function getCanvasDomain() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const domainFromCell = sheet.getRange(CANVAS_DOMAIN_CELL).getValue();
  
  // Validate the domain
  if (!domainFromCell) {
    // If empty, alert the user
    SpreadsheetApp.getUi().alert(
      'Canvas Domain Missing',
      `Please enter your Canvas domain URL in cell ${CANVAS_DOMAIN_CELL} (e.g., https://canvas.chapman.edu)`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
  
  // Ensure domain starts with https://
  let formattedDomain = domainFromCell.trim();
  if (!formattedDomain.startsWith('http://') && !formattedDomain.startsWith('https://')) {
    formattedDomain = 'https://' + formattedDomain;
  }
  
  // Remove trailing slash if present
  if (formattedDomain.endsWith('/')) {
    formattedDomain = formattedDomain.slice(0, -1);
  }
  
  return formattedDomain;
}
