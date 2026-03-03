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

