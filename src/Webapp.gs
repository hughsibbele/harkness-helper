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
 * Serve the recorder HTML page, or handle API requests when called with an action parameter.
 * @param {Object} e - The event parameter from Apps Script
 */
function doGet(e) {
  const action = e?.parameter?.action;

  if (action === "getConfig") {
    try {
      const config = getRecorderConfig();
      return ContentService.createTextOutput(JSON.stringify(config))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error", error: error.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default: serve the HTML page
  return HtmlService.createHtmlOutputFromFile('RecorderApp')
    .setTitle('Harkness Recorder')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Handle POST requests for audio upload from the GitHub Pages frontend.
 * @param {Object} e - The event parameter from Apps Script
 */
function doPost(e) {
  try {
    const action = e?.parameter?.action;

    if (action === "uploadAudio") {
      const body = JSON.parse(e.postData.contents);
      const result = uploadAudioFile(body.base64Data, body.fileName, body.mimeType);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "error", error: "Unknown action"
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error", error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================================
// CLIENT-CALLABLE FUNCTIONS
// ============================================================================

/**
 * Get configuration needed by the recorder UI.
 * Returns known sections from the Students sheet so the teacher can pick one.
 * In multi-course mode, also returns courses and a course→sections mapping.
 * @returns {Object} { sections, multiCourse, courses, coursesSections }
 */
function getRecorderConfig() {
  try {
    const students = getAllRows(CONFIG.SHEETS.STUDENTS);
    const sections = [...new Set(students.map(s => s.section).filter(Boolean))];
    const multiCourse = isMultiCourseMode();

    const config = {
      sections: sections.sort(),
      multiCourse: multiCourse
    };

    if (multiCourse) {
      const courses = getAllCourses();
      config.courses = courses.map(c => c.course_name);

      // Build course → sections mapping from student data
      const coursesSections = {};
      for (const course of config.courses) {
        const courseSections = [...new Set(
          students
            .filter(s => s.course === course && s.section)
            .map(s => s.section)
        )];
        coursesSections[course] = courseSections.sort();
      }
      config.coursesSections = coursesSections;
    }

    return config;
  } catch (e) {
    Logger.log('getRecorderConfig error: ' + e.message);
    return { sections: [], multiCourse: false };
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
