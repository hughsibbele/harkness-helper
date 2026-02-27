/**
 * Google Sheets helper functions for Harkness Helper
 *
 * Provides CRUD operations for seven sheets:
 * - Settings: Key-value global configuration
 * - Discussions: Track discussion recordings and their status
 * - Students: Student roster with Canvas IDs
 * - Transcripts: Raw and processed transcripts
 * - SpeakerMap: Per-speaker rows for teacher review
 * - StudentReports: Individual student reports (individual mode)
 * - Prompts: Teacher-editable prompt templates
 */

// ============================================================================
// SPREADSHEET ACCESS
// ============================================================================

/**
 * Get the main spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

/**
 * Get a sheet by name, creating it if it doesn't exist
 * @param {string} sheetName - Name of the sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initializeSheetHeaders(sheet, sheetName);
  }
  return sheet;
}

/**
 * Initialize headers for a new sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} sheetName
 */
function initializeSheetHeaders(sheet, sheetName) {
  const headers = {
    [CONFIG.SHEETS.SETTINGS]: [
      'setting_key', 'setting_value'
    ],
    [CONFIG.SHEETS.DISCUSSIONS]: [
      'status', 'next_step', 'date', 'section', 'course',
      'grade', 'approved', 'group_feedback',
      'canvas_assignment_id', 'canvas_item_type',
      'discussion_id', 'audio_file_id', 'error_message',
      'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.STUDENTS]: [
      'name', 'email', 'section', 'course', 'canvas_user_id', 'student_id'
    ],
    [CONFIG.SHEETS.TRANSCRIPTS]: [
      'discussion_id', 'raw_transcript', 'raw_transcript_2',
      'speaker_map', 'named_transcript', 'named_transcript_2',
      'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.SPEAKER_MAP]: [
      'discussion_id', 'speaker_label', 'suggested_name', 'student_name', 'confirmed'
    ],
    [CONFIG.SHEETS.STUDENT_REPORTS]: [
      'student_name', 'grade', 'approved', 'sent', 'feedback',
      'discussion_id', 'transcript_contributions', 'participation_summary',
      'student_id', 'report_id', 'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.PROMPTS]: [
      'prompt_name', 'prompt_text'
    ],
    [CONFIG.SHEETS.COURSES]: [
      'course_name', 'canvas_course_id', 'canvas_base_url', 'canvas_item_type'
    ]
  };

  if (headers[sheetName]) {
    sheet.getRange(1, 1, 1, headers[sheetName].length).setValues([headers[sheetName]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers[sheetName].length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');
  }
}

// ============================================================================
// GENERIC CRUD HELPERS
// ============================================================================

/**
 * Get all data from a sheet as objects (using header row as keys)
 * @param {string} sheetName
 * @returns {Object[]} Array of row objects
 */
function getAllRows(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  return data.slice(1).map((row, index) => {
    const obj = { _rowIndex: index + 2 }; // 1-indexed, skip header
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

/**
 * Find a row by a column value
 * @param {string} sheetName
 * @param {string} columnName
 * @param {*} value
 * @returns {Object|null}
 */
function findRow(sheetName, columnName, value) {
  const rows = getAllRows(sheetName);
  return rows.find(row => row[columnName] === value) || null;
}

/**
 * Find all rows matching a column value
 * @param {string} sheetName
 * @param {string} columnName
 * @param {*} value
 * @returns {Object[]}
 */
function findRows(sheetName, columnName, value) {
  const rows = getAllRows(sheetName);
  return rows.filter(row => row[columnName] === value);
}

/**
 * Insert a new row into a sheet
 * @param {string} sheetName
 * @param {Object} data - Object with column names as keys
 * @returns {number} Row index of inserted row
 */
function insertRow(sheetName, data) {
  const sheet = getOrCreateSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const rowData = headers.map(header => data[header] ?? '');
  sheet.appendRow(rowData);

  return sheet.getLastRow();
}

/**
 * Update a row by row index
 * @param {string} sheetName
 * @param {number} rowIndex - 1-indexed row number
 * @param {Object} data - Object with column names as keys (partial updates OK)
 */
function updateRow(sheetName, rowIndex, data) {
  const sheet = getOrCreateSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  headers.forEach((header, colIndex) => {
    if (header in data) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(data[header]);
    }
  });
}

/**
 * Get column index by name (1-indexed)
 * @param {string} sheetName
 * @param {string} columnName
 * @returns {number}
 */
function getColumnIndex(sheetName, columnName) {
  const sheet = getOrCreateSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(columnName);
  return index === -1 ? -1 : index + 1;
}

// ============================================================================
// SETTINGS CRUD
// ============================================================================

/**
 * Get a setting value from the Settings sheet
 * @param {string} key
 * @returns {string} The setting value, or empty string if not found
 */
function getSetting(key) {
  const row = findRow(CONFIG.SHEETS.SETTINGS, 'setting_key', key);
  return row ? String(row.setting_value) : '';
}

/**
 * Set a setting value in the Settings sheet
 * @param {string} key
 * @param {string} value
 */
function setSetting(key, value) {
  const row = findRow(CONFIG.SHEETS.SETTINGS, 'setting_key', key);
  if (row) {
    updateRow(CONFIG.SHEETS.SETTINGS, row._rowIndex, { setting_value: value });
  } else {
    insertRow(CONFIG.SHEETS.SETTINGS, { setting_key: key, setting_value: value });
  }
}

/**
 * Initialize Settings sheet with default values (only adds missing keys)
 */
function initializeSettings() {
  const defaults = {
    'mode': CONFIG.MODES.GROUP,
    'distribute_email': 'true',
    'distribute_canvas': 'false',
    'grade_scale': '0-100',
    'teacher_email': '',
    'teacher_name': '',
    'email_subject_template': 'Harkness Discussion Report - {date}',
    'gemini_model': 'gemini-2.0-flash',
    'elevenlabs_model': 'scribe_v2',
    'canvas_course_id': '',
    'canvas_base_url': '',
    'canvas_item_type': 'assignment'
  };

  for (const [key, value] of Object.entries(defaults)) {
    const existing = findRow(CONFIG.SHEETS.SETTINGS, 'setting_key', key);
    if (!existing) {
      insertRow(CONFIG.SHEETS.SETTINGS, { setting_key: key, setting_value: value });
    }
  }

  Logger.log('Settings initialized with defaults.');
}

// ============================================================================
// DISCUSSIONS CRUD
// ============================================================================

/**
 * Create a new discussion record
 * @param {Object} data
 * @returns {string} The discussion ID
 */
function createDiscussion(data) {
  const discussionId = `disc_${Date.now()}`;
  const now = new Date().toISOString();

  insertRow(CONFIG.SHEETS.DISCUSSIONS, {
    discussion_id: discussionId,
    date: data.date || new Date().toISOString().split('T')[0],
    section: data.section || '',
    course: data.course || '',
    audio_file_id: data.audio_file_id || '',
    status: CONFIG.STATUS.UPLOADED,
    next_step: 'Waiting for transcription...',
    grade: '',
    group_feedback: '',
    approved: false,
    canvas_assignment_id: data.canvas_assignment_id || '',
    canvas_item_type: data.canvas_item_type || '',
    error_message: '',
    created_at: now,
    updated_at: now
  });

  return discussionId;
}

/**
 * Get a discussion by ID
 * @param {string} discussionId
 * @returns {Object|null}
 */
function getDiscussion(discussionId) {
  return findRow(CONFIG.SHEETS.DISCUSSIONS, 'discussion_id', discussionId);
}

/**
 * Get all discussions with a specific status
 * @param {string} status
 * @returns {Object[]}
 */
function getDiscussionsByStatus(status) {
  return findRows(CONFIG.SHEETS.DISCUSSIONS, 'status', status);
}

/**
 * Update a discussion record. Appends to error_message if provided.
 * @param {string} discussionId
 * @param {Object} data
 */
function updateDiscussion(discussionId, data) {
  const discussion = getDiscussion(discussionId);
  if (!discussion) return;

  // Append error messages instead of overwriting
  if (data.error_message && discussion.error_message) {
    data.error_message = discussion.error_message + '\n' + data.error_message;
  }

  // Auto-set next_step on error
  if (data.status === CONFIG.STATUS.ERROR && data.error_message) {
    const brief = data.error_message.split('\n').pop().substring(0, 80);
    data.next_step = `ERROR: ${brief}`;
  }

  data.updated_at = new Date().toISOString();
  updateRow(CONFIG.SHEETS.DISCUSSIONS, discussion._rowIndex, data);
}

// ============================================================================
// STUDENTS CRUD
// ============================================================================

/**
 * Get all students for a section
 * @param {string} section
 * @returns {Object[]}
 */
function getStudentsBySection(section) {
  return findRows(CONFIG.SHEETS.STUDENTS, 'section', section);
}

/**
 * Get a student by ID
 * @param {string} studentId
 * @returns {Object|null}
 */
function getStudent(studentId) {
  return findRow(CONFIG.SHEETS.STUDENTS, 'student_id', studentId);
}

/**
 * Get a student by name (case-insensitive)
 * @param {string} name
 * @param {string} section - Optional section filter
 * @param {string} course - Optional course filter
 * @returns {Object|null}
 */
function getStudentByName(name, section = null, course = null) {
  const rows = getAllRows(CONFIG.SHEETS.STUDENTS);
  const normalizedName = name.toLowerCase().trim();

  return rows.find(row => {
    const nameMatch = String(row.name).toLowerCase().trim() === normalizedName;
    const sectionMatch = section ? row.section === section : true;
    const courseMatch = course ? row.course === course : true;
    return nameMatch && sectionMatch && courseMatch;
  }) || null;
}

/**
 * Add or update a student
 * @param {Object} data
 * @returns {string} Student ID
 */
function upsertStudent(data) {
  const existing = getStudentByName(data.name, data.section, data.course || null);

  if (existing) {
    updateRow(CONFIG.SHEETS.STUDENTS, existing._rowIndex, data);
    return existing.student_id;
  } else {
    const studentId = `stu_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`;
    insertRow(CONFIG.SHEETS.STUDENTS, {
      student_id: studentId,
      ...data
    });
    return studentId;
  }
}

// ============================================================================
// TRANSCRIPTS CRUD
// ============================================================================

/**
 * Split a string into chunks that fit within the Sheets cell character limit.
 * @param {string} text
 * @returns {string[]} Array of [chunk1, chunk2]
 */
function splitForCells(text) {
  const limit = CONFIG.LIMITS.CELL_CHAR_LIMIT;
  if (!text || text.length <= limit) return [text || '', ''];
  return [text.substring(0, limit), text.substring(limit)];
}

/**
 * Join primary and overflow cell values back into a single string.
 * @param {string} primary
 * @param {string} overflow
 * @returns {string}
 */
function joinFromCells(primary, overflow) {
  const p = primary ? String(primary) : '';
  const o = overflow ? String(overflow) : '';
  return p + o;
}

/**
 * Create or update a transcript record.
 * Automatically splits raw_transcript and named_transcript across overflow
 * columns if they exceed the Sheets cell character limit.
 * @param {string} discussionId
 * @param {Object} data
 */
function upsertTranscript(discussionId, data) {
  const existing = findRow(CONFIG.SHEETS.TRANSCRIPTS, 'discussion_id', discussionId);
  const now = new Date().toISOString();

  // Split long transcript fields into primary + overflow columns
  if ('raw_transcript' in data) {
    const parts = splitForCells(data.raw_transcript);
    data.raw_transcript = parts[0];
    data.raw_transcript_2 = parts[1];
  }
  if ('named_transcript' in data) {
    const parts = splitForCells(data.named_transcript);
    data.named_transcript = parts[0];
    data.named_transcript_2 = parts[1];
  }

  if (existing) {
    data.updated_at = now;
    updateRow(CONFIG.SHEETS.TRANSCRIPTS, existing._rowIndex, data);
  } else {
    insertRow(CONFIG.SHEETS.TRANSCRIPTS, {
      discussion_id: discussionId,
      raw_transcript: data.raw_transcript || '',
      raw_transcript_2: data.raw_transcript_2 || '',
      speaker_map: data.speaker_map || '{}',
      named_transcript: data.named_transcript || '',
      named_transcript_2: data.named_transcript_2 || '',
      created_at: now,
      updated_at: now
    });
  }
}

/**
 * Get transcript for a discussion.
 * Transparently joins overflow columns so callers always get the full text.
 * @param {string} discussionId
 * @returns {Object|null}
 */
function getTranscript(discussionId) {
  const row = findRow(CONFIG.SHEETS.TRANSCRIPTS, 'discussion_id', discussionId);
  if (!row) return null;

  // Join primary + overflow columns so downstream code sees full transcripts
  row.raw_transcript = joinFromCells(row.raw_transcript, row.raw_transcript_2);
  row.named_transcript = joinFromCells(row.named_transcript, row.named_transcript_2);
  return row;
}

// ============================================================================
// SPEAKER MAP CRUD
// ============================================================================

/**
 * Get all speaker map entries for a discussion
 * @param {string} discussionId
 * @returns {Object[]}
 */
function getSpeakerMapForDiscussion(discussionId) {
  return findRows(CONFIG.SHEETS.SPEAKER_MAP, 'discussion_id', discussionId);
}

/**
 * Insert or update a speaker mapping row
 * @param {string} discussionId
 * @param {string} speakerLabel - e.g. "Speaker 0"
 * @param {string} suggestedName - Gemini's guess
 */
function upsertSpeakerMapping(discussionId, speakerLabel, suggestedName) {
  const rows = getSpeakerMapForDiscussion(discussionId);
  const existing = rows.find(r => r.speaker_label === speakerLabel);

  if (existing) {
    updateRow(CONFIG.SHEETS.SPEAKER_MAP, existing._rowIndex, {
      suggested_name: suggestedName,
      student_name: suggestedName,
      confirmed: isGroupMode()
    });
  } else {
    insertRow(CONFIG.SHEETS.SPEAKER_MAP, {
      discussion_id: discussionId,
      speaker_label: speakerLabel,
      suggested_name: suggestedName,
      student_name: suggestedName,
      confirmed: isGroupMode()
    });
  }
}

/**
 * Check if all speakers are confirmed for a discussion
 * @param {string} discussionId
 * @returns {boolean}
 */
function isSpeakerMapConfirmed(discussionId) {
  const rows = getSpeakerMapForDiscussion(discussionId);
  if (rows.length === 0) return false;
  return rows.every(r => r.confirmed === true);
}

/**
 * Build a speaker map object from the SpeakerMap sheet
 * @param {string} discussionId
 * @returns {Object} Map of speaker labels to student names
 */
function buildSpeakerMapObject(discussionId) {
  const rows = getSpeakerMapForDiscussion(discussionId);
  const map = {};
  for (const row of rows) {
    map[row.speaker_label] = row.student_name || row.suggested_name || row.speaker_label;
  }
  return map;
}

// ============================================================================
// STUDENT REPORTS CRUD
// ============================================================================

/**
 * Create a student report
 * @param {Object} data
 * @returns {string} Report ID
 */
function createStudentReport(data) {
  const reportId = `rep_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`;
  const now = new Date().toISOString();

  insertRow(CONFIG.SHEETS.STUDENT_REPORTS, {
    report_id: reportId,
    discussion_id: data.discussion_id,
    student_id: data.student_id,
    student_name: data.student_name || '',
    transcript_contributions: data.transcript_contributions || '',
    participation_summary: data.participation_summary || '',
    grade: data.grade ?? '',
    feedback: data.feedback || '',
    approved: false,
    sent: false,
    created_at: now,
    updated_at: now
  });

  return reportId;
}

/**
 * Get all reports for a discussion
 * @param {string} discussionId
 * @returns {Object[]}
 */
function getReportsForDiscussion(discussionId) {
  return findRows(CONFIG.SHEETS.STUDENT_REPORTS, 'discussion_id', discussionId);
}

/**
 * Get a specific student's report for a discussion
 * @param {string} discussionId
 * @param {string} studentId
 * @returns {Object|null}
 */
function getStudentReport(discussionId, studentId) {
  const reports = getReportsForDiscussion(discussionId);
  return reports.find(r => r.student_id === studentId) || null;
}

/**
 * Update a student report
 * @param {string} reportId
 * @param {Object} data
 */
function updateStudentReport(reportId, data) {
  const report = findRow(CONFIG.SHEETS.STUDENT_REPORTS, 'report_id', reportId);
  if (report) {
    data.updated_at = new Date().toISOString();
    updateRow(CONFIG.SHEETS.STUDENT_REPORTS, report._rowIndex, data);
  }
}

/**
 * Get all approved but unsent reports for a discussion
 * @param {string} discussionId
 * @returns {Object[]}
 */
function getApprovedUnsentReports(discussionId) {
  const reports = getReportsForDiscussion(discussionId);
  return reports.filter(r => r.approved === true && r.sent !== true);
}

// ============================================================================
// PROMPTS CRUD (sheet-based)
// ============================================================================

/**
 * Get a prompt template from the Prompts sheet
 * @param {string} name - The prompt_name
 * @returns {string} The prompt text, or empty string if not found
 */
function getPromptFromSheet(name) {
  const row = findRow(CONFIG.SHEETS.PROMPTS, 'prompt_name', name);
  return row ? String(row.prompt_text) : '';
}

/**
 * Insert or update a prompt in the Prompts sheet
 * @param {string} name
 * @param {string} text
 */
function upsertPrompt(name, text) {
  const existing = findRow(CONFIG.SHEETS.PROMPTS, 'prompt_name', name);
  if (existing) {
    updateRow(CONFIG.SHEETS.PROMPTS, existing._rowIndex, { prompt_text: text });
  } else {
    insertRow(CONFIG.SHEETS.PROMPTS, { prompt_name: name, prompt_text: text });
  }
}

// ============================================================================
// COURSES CRUD & MULTI-COURSE HELPERS
// ============================================================================

/**
 * Get all courses from the Courses sheet
 * @returns {Object[]} Array of course row objects
 */
function getAllCourses() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COURSES);
  if (!sheet) return [];
  return getAllRows(CONFIG.SHEETS.COURSES).filter(c => c.course_name);
}

/**
 * Get a course by its friendly name
 * @param {string} name - The course_name value
 * @returns {Object|null}
 */
function getCourseByName(name) {
  if (!name) return null;
  const courses = getAllCourses();
  return courses.find(c => c.course_name === name) || null;
}

/**
 * Check if multi-course mode is active.
 * Returns true if the Courses sheet exists and has at least one row.
 * @returns {boolean}
 */
function isMultiCourseMode() {
  return getAllCourses().length > 0;
}

/**
 * Get the Canvas course ID for a discussion, using the fallback chain:
 * 1. Look up discussion.course in Courses sheet → canvas_course_id
 * 2. Fall back to global getSetting('canvas_course_id')
 * 3. If neither → error
 *
 * @param {Object} discussion - Discussion row object
 * @returns {string} Canvas course ID
 */
function getCanvasCourseIdForDiscussion(discussion) {
  // 1. Per-course lookup
  if (discussion.course) {
    const course = getCourseByName(discussion.course);
    if (course && course.canvas_course_id) {
      return String(course.canvas_course_id);
    }
  }

  // 2. Global fallback
  const globalId = getSetting('canvas_course_id');
  if (globalId) return globalId;

  // 3. Error
  throw new Error(
    'No Canvas course ID found. Set it on the Courses sheet (for multi-course) or in the Settings sheet.'
  );
}

/**
 * Get the effective Canvas item type for a discussion.
 * 3-tier fallback: per-discussion → per-course → global setting.
 *
 * @param {Object} discussion - Discussion row object
 * @returns {string} 'assignment' or 'discussion'
 */
function getCanvasItemTypeForDiscussion(discussion) {
  // 1. Per-discussion override
  const perDiscussion = discussion.canvas_item_type;
  if (perDiscussion === 'assignment' || perDiscussion === 'discussion') {
    return perDiscussion;
  }

  // 2. Per-course override
  if (discussion.course) {
    const course = getCourseByName(discussion.course);
    if (course) {
      const perCourse = course.canvas_item_type;
      if (perCourse === 'assignment' || perCourse === 'discussion') {
        return perCourse;
      }
    }
  }

  // 3. Global setting
  const global = getSetting('canvas_item_type');
  return global === 'discussion' ? 'discussion' : 'assignment';
}

/**
 * Get the Canvas base URL for a discussion.
 * Checks per-course override first, then falls back to global.
 *
 * @param {Object} discussion - Discussion row object
 * @returns {string} Canvas base URL
 */
function getCanvasBaseUrlForDiscussion(discussion) {
  if (discussion.course) {
    const course = getCourseByName(discussion.course);
    if (course && course.canvas_base_url) {
      return String(course.canvas_base_url).replace(/\/$/, '');
    }
  }
  return getCanvasBaseUrl();
}

/**
 * Get students filtered by section and optionally by course.
 * In multi-course mode, filters by both section and course.
 * In single-course mode, filters by section only (backward compatible).
 *
 * @param {string} section - Section name
 * @param {string} course - Optional course name
 * @returns {Object[]}
 */
function getStudentsBySectionAndCourse(section, course = null) {
  const rows = getAllRows(CONFIG.SHEETS.STUDENTS);
  return rows.filter(row => {
    const sectionMatch = section ? row.section === section : true;
    const courseMatch = course ? row.course === course : true;
    return sectionMatch && courseMatch;
  });
}

/**
 * Migrate existing spreadsheet to multi-course mode.
 * - Inserts 'course' columns into Discussions and Students sheets if missing
 * - Creates Courses sheet if missing
 * - Pre-populates with current canvas_course_id if set
 */
function migrateToMultiCourse() {
  const ss = getSpreadsheet();

  // Add 'course' column to Discussions if missing
  const discSheet = ss.getSheetByName(CONFIG.SHEETS.DISCUSSIONS);
  if (discSheet) {
    const discHeaders = discSheet.getRange(1, 1, 1, discSheet.getLastColumn()).getValues()[0];
    if (!discHeaders.includes('course')) {
      // Insert after 'section' column
      const sectionIdx = discHeaders.indexOf('section');
      const insertAt = sectionIdx >= 0 ? sectionIdx + 2 : discHeaders.length + 1;
      discSheet.insertColumnAfter(insertAt - 1);
      discSheet.getRange(1, insertAt).setValue('course')
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('white');
      Logger.log('Added "course" column to Discussions sheet');
    }
  }

  // Add 'course' column to Students if missing
  const studSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  if (studSheet) {
    const studHeaders = studSheet.getRange(1, 1, 1, studSheet.getLastColumn()).getValues()[0];
    if (!studHeaders.includes('course')) {
      // Insert after 'section' column
      const sectionIdx = studHeaders.indexOf('section');
      const insertAt = sectionIdx >= 0 ? sectionIdx + 2 : studHeaders.length + 1;
      studSheet.insertColumnAfter(insertAt - 1);
      studSheet.getRange(1, insertAt).setValue('course')
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('white');
      Logger.log('Added "course" column to Students sheet');
    }
  }

  // Create Courses sheet if it doesn't exist
  let coursesSheet = ss.getSheetByName(CONFIG.SHEETS.COURSES);
  if (!coursesSheet) {
    coursesSheet = ss.insertSheet(CONFIG.SHEETS.COURSES);
    initializeSheetHeaders(coursesSheet, CONFIG.SHEETS.COURSES);
    Logger.log('Created Courses sheet');
  }

  // Pre-populate with current canvas_course_id if set
  const existingCourses = getAllCourses();
  if (existingCourses.length === 0) {
    const courseId = getSetting('canvas_course_id');
    if (courseId) {
      insertRow(CONFIG.SHEETS.COURSES, {
        course_name: 'My Course',
        canvas_course_id: courseId,
        canvas_base_url: '',
        canvas_item_type: ''
      });
      Logger.log(`Pre-populated Courses sheet with course ID ${courseId}`);
    }
  }

  // Re-apply formatting (checkboxes, wrapping, dropdowns)
  formatSheets();

  Logger.log('Multi-course migration complete');
}

// ============================================================================
// SHEET INITIALIZATION
// ============================================================================

/**
 * Initialize all sheets with proper structure
 */
function initializeAllSheets() {
  Object.values(CONFIG.SHEETS).forEach(sheetName => {
    getOrCreateSheet(sheetName);
  });

  Logger.log('All sheets initialized successfully.');
}

/**
 * Reorder columns in existing sheets to match the canonical header order
 * defined in initializeSheetHeaders(). Moves data along with headers.
 * Safe to run multiple times — no-ops if columns already in order.
 * Sheets with no data rows are skipped (headers-only sheets are rewritten in place).
 */
function reorderColumns() {
  const ss = getSpreadsheet();

  // Canonical header order (must match initializeSheetHeaders)
  const canonicalHeaders = {
    [CONFIG.SHEETS.DISCUSSIONS]: [
      'status', 'next_step', 'date', 'section', 'course',
      'grade', 'approved', 'group_feedback',
      'canvas_assignment_id', 'canvas_item_type',
      'discussion_id', 'audio_file_id', 'error_message',
      'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.STUDENTS]: [
      'name', 'email', 'section', 'course', 'canvas_user_id', 'student_id'
    ],
    [CONFIG.SHEETS.STUDENT_REPORTS]: [
      'student_name', 'grade', 'approved', 'sent', 'feedback',
      'discussion_id', 'transcript_contributions', 'participation_summary',
      'student_id', 'report_id', 'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.SPEAKER_MAP]: [
      'discussion_id', 'speaker_label', 'suggested_name', 'student_name', 'confirmed'
    ],
    [CONFIG.SHEETS.TRANSCRIPTS]: [
      'discussion_id', 'raw_transcript', 'raw_transcript_2',
      'speaker_map', 'named_transcript', 'named_transcript_2',
      'created_at', 'updated_at'
    ]
  };

  for (const [sheetName, targetOrder] of Object.entries(canonicalHeaders)) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) continue;

    // Read current headers
    const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

    // Build column mapping: for each target column, find its current position
    const colMap = [];  // array of current column indices (0-based) in target order
    for (const header of targetOrder) {
      const idx = currentHeaders.indexOf(header);
      if (idx >= 0) {
        colMap.push(idx);
      }
    }
    // Append any extra columns not in the canonical order (preserves unknown columns)
    for (let i = 0; i < currentHeaders.length; i++) {
      if (!colMap.includes(i)) {
        colMap.push(i);
      }
    }

    // Check if already in order
    const alreadyOrdered = colMap.every((val, i) => val === i);
    if (alreadyOrdered) {
      Logger.log(`${sheetName}: columns already in order, skipping.`);
      continue;
    }

    // Read all data (headers + rows)
    const allData = lastRow > 0
      ? sheet.getRange(1, 1, lastRow, lastCol).getValues()
      : [];

    // Reorder each row according to colMap
    const reordered = allData.map(row => colMap.map(idx => row[idx]));

    // Write back
    sheet.clear();
    if (reordered.length > 0) {
      sheet.getRange(1, 1, reordered.length, reordered[0].length).setValues(reordered);
    }

    // Re-apply header formatting
    const numCols = reordered[0] ? reordered[0].length : targetOrder.length;
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, numCols)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');

    Logger.log(`${sheetName}: columns reordered successfully.`);
  }

  // Re-apply all formatting (widths, highlighting, validation, etc.)
  formatSheets();

  Logger.log('Column reordering complete.');
}

/**
 * Add data validation, formatting, column widths, and tab ordering to sheets.
 * Safe to run multiple times — applies formatting idempotently.
 */
function formatSheets() {
  const ss = getSpreadsheet();
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  // --- Discussions sheet ---
  const discussionsSheet = ss.getSheetByName(CONFIG.SHEETS.DISCUSSIONS);
  if (discussionsSheet) {
    // Status dropdown
    const statusCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'status');
    if (statusCol > 0) {
      const statusValues = Object.values(CONFIG.STATUS);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statusValues)
        .build();
      discussionsSheet.getRange(2, statusCol, 1000, 1).setDataValidation(rule);
      // Highlight status column with light blue
      discussionsSheet.getRange(2, statusCol, 1000, 1).setBackground('#e3f2fd');
      discussionsSheet.setColumnWidth(statusCol, 100);
    }

    // Next step: wrap, highlight, and widen
    const nextStepCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'next_step');
    if (nextStepCol > 0) {
      discussionsSheet.getRange(2, nextStepCol, 1000, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
        .setBackground('#fff9c4');
      discussionsSheet.setColumnWidth(nextStepCol, 280);
    }

    // Freeze status + next_step columns so they're always visible
    discussionsSheet.setFrozenColumns(2);

    // Approved checkbox
    const approvedCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'approved');
    if (approvedCol > 0) {
      discussionsSheet.getRange(2, approvedCol, 1000, 1).setDataValidation(checkboxRule);
    }

    // canvas_item_type dropdown
    const itemTypeCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'canvas_item_type');
    if (itemTypeCol > 0) {
      const itemTypeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['assignment', 'discussion'])
        .setAllowInvalid(true)
        .build();
      discussionsSheet.getRange(2, itemTypeCol, 1000, 1).setDataValidation(itemTypeRule);
    }

    // Column widths for readability
    const colWidths = {
      'date': 100, 'section': 100, 'course': 120, 'grade': 65,
      'approved': 80, 'group_feedback': 300,
      'canvas_assignment_id': 140, 'canvas_item_type': 120,
      'discussion_id': 140, 'audio_file_id': 120,
      'error_message': 200, 'created_at': 140, 'updated_at': 140
    };
    for (const [col, width] of Object.entries(colWidths)) {
      const idx = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, col);
      if (idx > 0) discussionsSheet.setColumnWidth(idx, width);
    }

    // Wrap group_feedback and error_message
    const feedbackCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'group_feedback');
    if (feedbackCol > 0) {
      discussionsSheet.getRange(2, feedbackCol, 1000, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    const errorCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'error_message');
    if (errorCol > 0) {
      discussionsSheet.getRange(2, errorCol, 1000, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }

    // Conditional formatting: error rows red, sent rows green
    const lastCol = Math.max(discussionsSheet.getLastColumn(), 1);
    if (statusCol > 0) {
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(CONFIG.STATUS.ERROR)
        .setBackground('#ffcdd2')
        .setRanges([discussionsSheet.getRange(2, 1, 1000, lastCol)])
        .build();
      const sentRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(CONFIG.STATUS.SENT)
        .setBackground('#c8e6c9')
        .setRanges([discussionsSheet.getRange(2, 1, 1000, lastCol)])
        .build();
      discussionsSheet.setConditionalFormatRules([errorRule, sentRule]);
    }

    discussionsSheet.setTabColor('#4285f4');
  }

  // --- Students sheet ---
  const studentsSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  if (studentsSheet) {
    const studentColWidths = {
      'name': 160, 'email': 200, 'section': 100,
      'course': 120, 'canvas_user_id': 110, 'student_id': 140
    };
    for (const [col, width] of Object.entries(studentColWidths)) {
      const idx = getColumnIndex(CONFIG.SHEETS.STUDENTS, col);
      if (idx > 0) studentsSheet.setColumnWidth(idx, width);
    }
    studentsSheet.setTabColor('#34a853');
  }

  // --- SpeakerMap sheet ---
  const speakerMapSheet = ss.getSheetByName(CONFIG.SHEETS.SPEAKER_MAP);
  if (speakerMapSheet) {
    const confirmedCol = getColumnIndex(CONFIG.SHEETS.SPEAKER_MAP, 'confirmed');
    if (confirmedCol > 0) {
      speakerMapSheet.getRange(2, confirmedCol, 1000, 1).setDataValidation(checkboxRule);
    }
    speakerMapSheet.setTabColor('#ff9800');
  }

  // --- StudentReports sheet ---
  const reportsSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENT_REPORTS);
  if (reportsSheet) {
    const approvedCol = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, 'approved');
    const sentCol = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, 'sent');

    if (approvedCol > 0) {
      reportsSheet.getRange(2, approvedCol, 1000, 1).setDataValidation(checkboxRule);
    }
    if (sentCol > 0) {
      reportsSheet.getRange(2, sentCol, 1000, 1).setDataValidation(checkboxRule);
    }

    // Column widths
    const reportColWidths = {
      'student_name': 140, 'grade': 65, 'approved': 80, 'sent': 60,
      'feedback': 300, 'discussion_id': 140,
      'transcript_contributions': 250, 'participation_summary': 200
    };
    for (const [col, width] of Object.entries(reportColWidths)) {
      const idx = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, col);
      if (idx > 0) reportsSheet.setColumnWidth(idx, width);
    }

    // Wrap feedback and contributions
    const feedbackCol = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, 'feedback');
    if (feedbackCol > 0) {
      reportsSheet.getRange(2, feedbackCol, 1000, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    const contribCol = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, 'transcript_contributions');
    if (contribCol > 0) {
      reportsSheet.getRange(2, contribCol, 1000, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }

    reportsSheet.setTabColor('#9c27b0');
  }

  // --- Transcripts sheet ---
  const transcriptsSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSCRIPTS);
  if (transcriptsSheet) {
    transcriptsSheet.setTabColor('#9e9e9e');
  }

  // --- Courses sheet ---
  const coursesSheet = ss.getSheetByName(CONFIG.SHEETS.COURSES);
  if (coursesSheet) {
    coursesSheet.setTabColor('#00acc1');
  }

  // --- Prompts sheet ---
  const promptsSheet = ss.getSheetByName(CONFIG.SHEETS.PROMPTS);
  if (promptsSheet) {
    promptsSheet.setTabColor('#fdd835');
    const textCol = getColumnIndex(CONFIG.SHEETS.PROMPTS, 'prompt_text');
    if (textCol > 0) {
      promptsSheet.getRange(2, textCol, 100, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      promptsSheet.setColumnWidth(textCol, 600);
    }
  }

  // --- Settings sheet ---
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (settingsSheet) {
    settingsSheet.setTabColor('#9e9e9e');
  }

  // --- Tab ordering (teacher-facing first, config last) ---
  const tabOrder = [
    CONFIG.SHEETS.DISCUSSIONS,
    CONFIG.SHEETS.STUDENTS,
    CONFIG.SHEETS.SPEAKER_MAP,
    CONFIG.SHEETS.STUDENT_REPORTS,
    CONFIG.SHEETS.TRANSCRIPTS,
    CONFIG.SHEETS.COURSES,
    CONFIG.SHEETS.PROMPTS,
    CONFIG.SHEETS.SETTINGS
  ];
  for (let i = 0; i < tabOrder.length; i++) {
    const sheet = ss.getSheetByName(tabOrder[i]);
    if (sheet) {
      sheet.activate();
      ss.moveActiveSheet(i + 1);
    }
  }
  // Return to Discussions tab
  const discTab = ss.getSheetByName(CONFIG.SHEETS.DISCUSSIONS);
  if (discTab) discTab.activate();

  Logger.log('Sheet formatting applied.');
}
