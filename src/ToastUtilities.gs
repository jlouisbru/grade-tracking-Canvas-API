/**
 * ==========================================================================
 * TOAST NOTIFICATION MANAGER
 * Utility functions for displaying toast notifications in Google Sheets.
 * This provides a centralized way to manage all toast notifications.
 * ==========================================================================
 */

/**
 * ToastManager provides utility functions for displaying toast notifications 
 * in Google Sheets with consistent formatting and behavior.
 */
const ToastManager = {
  /**
   * Displays a toast notification in Google Sheets.
   * @param {string} message The message to display in the toast.
   * @param {string} title The title of the toast.
   * @param {number} timeout How long (in seconds) the toast should be displayed.
   */
  showToast: function(message, title, timeout) {
    SpreadsheetApp.getActive().toast(message, title, timeout);
  },
  
  /**
   * Displays a progress toast notification.
   * @param {number} current The current progress value.
   * @param {number} total The total progress value.
   * @param {string} operation The operation being performed (e.g., 'Upload', 'Download').
   * @param {string} detail Optional detail about the current operation.
   * @param {number} timeout How long (in seconds) the toast should be displayed.
   */
  showProgressToast: function(current, total, operation, detail, timeout) {
    const percentage = Math.floor((current / total) * 100);
    let message = `Progress: ${current}/${total} (${percentage}%)`;
    if (detail) {
      message += `\n${detail}`;
    }
    
    this.showToast(message, `${operation}: In Progress`, timeout);
  },
  
  /**
   * Displays a completion toast notification.
   * @param {string} message The completion message.
   * @param {string} operation The operation that was performed.
   * @param {number} timeout How long (in seconds) the toast should be displayed.
   */
  showCompletionToast: function(message, operation, timeout) {
    this.showToast(message, `${operation}: Complete!`, timeout);
  },
  
  /**
   * Displays an error toast notification.
   * @param {string} errorMessage The error message.
   * @param {number} timeout How long (in seconds) the toast should be displayed.
   */
  showErrorToast: function(errorMessage, timeout) {
    this.showToast(`Error: ${errorMessage}`, 'Error', timeout);
  }
};

/**
 * Parses the 'Link' header from Canvas API responses to find the 'next' page URL.
 * This function is used by all Canvas API functions that handle pagination.
 * @param {string} linkHeader The 'Link' header string.
 * @return {string|null} The URL for the next page, or null if not found.
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
 * This function is used by all scripts that need to convert column letters to numbers.
 * @param {string} letter The column letter(s).
 * @return {number} The 1-based column number.
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
 * Add toast notifications to existing functions by monkey patching.
 * This is executed when the script loads to enhance all Canvas-related functions.
 */
(function() {
  // ========== PATCH FOR fetchCanvasGradesBySisId_ ==========
  // Store a reference to the original function
  const originalFetchCanvasGradesBySisId_ = this.fetchCanvasGradesBySisId_;
  
  // Only patch if the function exists
  if (typeof originalFetchCanvasGradesBySisId_ === 'function') {
    // Replace the function with a wrapper that adds toast notifications
    this.fetchCanvasGradesBySisId_ = function(gradeOutputColumn) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const ui = SpreadsheetApp.getUi();
      
      try {
        // Show initial toast
        ToastManager.showToast('Preparing to fetch Canvas grades...', 'Grade Fetch: Starting', 5);
        
        // Get course and assignment ID
        const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
        const assignmentIdCell = `${gradeOutputColumn}${ASSIGNMENT_ID_HEADER_ROW}`;
        const assignmentId = sheet.getRange(assignmentIdCell).getValue();
        
        if (courseId && assignmentId) {
          ToastManager.showToast(
            `Fetching grades for Assignment ID: ${assignmentId}`,
            'Grade Fetch: Step 1/3', 
            10
          );
        }
        
        // Let the original function do the work
        const result = originalFetchCanvasGradesBySisId_.apply(this, arguments);
        
        // Show completion toast
        ToastManager.showCompletionToast('Grades have been fetched and updated in the sheet.', 'Grade Fetch', 5);
        
        return result;
      } catch (error) {
        // Show error toast
        ToastManager.showErrorToast(error.message, 10);
        throw error;
      }
    };
  }
  
  // ========== PATCH FOR updateCanvasGrades_ ==========
  // Store a reference to the original function
  const originalUpdateCanvasGrades_ = this.updateCanvasGrades_;
  
  // Only patch if the function exists
  if (typeof originalUpdateCanvasGrades_ === 'function') {
    // Replace the function with a wrapper that adds toast notifications
    this.updateCanvasGrades_ = function(gradeColumn) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      
      try {
        // Show initial toast
        ToastManager.showToast('Preparing to upload grades to Canvas...', 'Grade Upload: Starting', 5);
        
        // Get course and assignment ID
        const courseId = sheet.getRange(COURSE_ID_CELL).getValue();
        const assignmentIdCell = `${gradeColumn}${ASSIGNMENT_ID_HEADER_ROW}`;
        const assignmentId = sheet.getRange(assignmentIdCell).getValue();
        
        if (courseId && assignmentId) {
          ToastManager.showToast(
            `Uploading grades to Canvas for Assignment ID: ${assignmentId}`,
            'Grade Upload: In Progress', 
            10
          );
        }
        
        // Calculate approximate progress
        const lastRow = sheet.getLastRow();
        const rowsToProcess = lastRow - FIRST_DATA_ROW + 1;
        
        // Track progress
        let progressCounter = 0;
        let lastProgressPercentage = 0;
        
        // Create a progress tracker function
        const trackProgress = function(currentRow, success, message) {
          progressCounter++;
          const percentage = Math.floor((progressCounter / rowsToProcess) * 100);
          
          // Only update toast every 5% or at the end
          if (percentage >= lastProgressPercentage + 5 || progressCounter === rowsToProcess) {
            lastProgressPercentage = percentage;
            
            let statusMsg = success ? "Uploaded" : "Skipped";
            ToastManager.showToast(
              `Progress: ${progressCounter}/${rowsToProcess} (${percentage}%)\n` +
              `Last student: Row ${currentRow} - ${statusMsg}` +
              (message ? `\n${message}` : ""),
              'Grade Upload: In Progress', 
              10
            );
          }
        };
        
        // Store original updateGradeInCanvas_ function
        const originalUpdateGradeInCanvas_ = this.updateGradeInCanvas_;
        if (typeof originalUpdateGradeInCanvas_ === 'function') {
          // Temporarily override for progress tracking
          this.updateGradeInCanvas_ = function(accessToken, courseId, assignmentId, studentId, grade) {
            // Call original function
            const result = originalUpdateGradeInCanvas_.apply(this, arguments);
            return result;
          };
        }
        
        // Call the original updateCanvasGrades_ function
        const result = originalUpdateCanvasGrades_.apply(this, arguments);
        
        // Restore original function
        if (typeof originalUpdateGradeInCanvas_ === 'function') {
          this.updateGradeInCanvas_ = originalUpdateGradeInCanvas_;
        }
        
        // Show completion toast
        ToastManager.showCompletionToast('Grade upload complete. See dialog for details.', 'Grade Upload', 5);
        
        return result;
      } catch (error) {
        // Show error toast
        ToastManager.showErrorToast(error.message, 10);
        throw error;
      }
    };
  }
  
  // ========== PATCH FOR updateGradeInCanvas_ ==========
  // This is a helper patch that allows us to track individual grade uploads
  const originalUpdateGradeInCanvas_ = this.updateGradeInCanvas_;
  if (typeof originalUpdateGradeInCanvas_ === 'function') {
    this.updateGradeInCanvas_ = function(accessToken, courseId, assignmentId, studentId, grade) {
      try {
        return originalUpdateGradeInCanvas_.apply(this, arguments);
      } catch (error) {
        // Log error in toast
        ToastManager.showErrorToast(`Error uploading grade for student ${studentId}: ${error.message}`, 5);
        throw error;
      }
    };
  }
})();
