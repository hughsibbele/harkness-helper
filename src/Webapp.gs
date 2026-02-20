/**
 * Web App for Harkness Recorder
 *
 * Serves a mobile-friendly recording page that captures audio,
 * asks for section + date, and uploads directly to the Drive upload folder.
 *
 * Deploy: Publish > Deploy as web app
 *   - Execute as: Me
 *   - Who has access: Anyone with a Google account (or just yourself)
 *
 * The resulting URL can be bookmarked or added to the phone's home screen.
 */

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================

/**
 * Serve the recorder HTML page
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('RecorderApp')
    .setTitle('Harkness Recorder')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ============================================================================
// CLIENT-CALLABLE FUNCTIONS
// ============================================================================

/**
 * Get configuration needed by the recorder UI.
 * Returns known sections from the Students sheet so the teacher can pick one.
 * @returns {Object} { sections: string[] }
 */
function getRecorderConfig() {
  try {
    const students = getAllRows(CONFIG.SHEETS.STUDENTS);
    const sections = [...new Set(students.map(s => s.section).filter(Boolean))];
    return { sections: sections.sort() };
  } catch (e) {
    Logger.log('getRecorderConfig error: ' + e.message);
    return { sections: [] };
  }
}

/**
 * Receive an audio recording from the client and save it to the upload folder.
 *
 * @param {string} base64Data - Base64-encoded audio data (no data URL prefix)
 * @param {string} fileName - Desired filename (e.g., "Section 1 - 2025-02-20.webm")
 * @param {string} mimeType - Audio MIME type (e.g., "audio/webm")
 * @returns {Object} { fileId, fileName }
 */
function uploadAudioFile(base64Data, fileName, mimeType) {
  const folderId = getAudioFolderId();
  const folder = DriveApp.getFolderById(folderId);

  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = folder.createFile(blob);

  Logger.log('Uploaded ' + fileName + ' (' + decoded.length + ' bytes) to upload folder');
  return { fileId: file.getId(), fileName: file.getName() };
}
