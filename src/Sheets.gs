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
      'discussion_id', 'date', 'section', 'audio_file_id',
      'status', 'next_step', 'grade', 'group_feedback', 'approved',
      'canvas_assignment_id', 'canvas_item_type',
      'error_message', 'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.STUDENTS]: [
      'student_id', 'name', 'email', 'section', 'canvas_user_id'
    ],
    [CONFIG.SHEETS.TRANSCRIPTS]: [
      'discussion_id', 'raw_transcript', 'speaker_map', 'named_transcript',
      'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.SPEAKER_MAP]: [
      'discussion_id', 'speaker_label', 'suggested_name', 'student_name', 'confirmed'
    ],
    [CONFIG.SHEETS.STUDENT_REPORTS]: [
      'report_id', 'discussion_id', 'student_id', 'student_name',
      'transcript_contributions', 'participation_summary',
      'grade', 'feedback', 'approved', 'sent',
      'created_at', 'updated_at'
    ],
    [CONFIG.SHEETS.PROMPTS]: [
      'prompt_name', 'prompt_text'
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
    'gemini_model': 'gemini-1.5-flash',
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
    audio_file_id: data.audio_file_id || '',
    status: CONFIG.STATUS.UPLOADED,
    next_step: 'Waiting for transcription...',
    grade: '',
    group_feedback: '',
    approved: false,
    canvas_assignment_id: data.canvas_assignment_id || '',
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
 * @returns {Object|null}
 */
function getStudentByName(name, section = null) {
  const rows = getAllRows(CONFIG.SHEETS.STUDENTS);
  const normalizedName = name.toLowerCase().trim();

  return rows.find(row => {
    const nameMatch = String(row.name).toLowerCase().trim() === normalizedName;
    const sectionMatch = section ? row.section === section : true;
    return nameMatch && sectionMatch;
  }) || null;
}

/**
 * Add or update a student
 * @param {Object} data
 * @returns {string} Student ID
 */
function upsertStudent(data) {
  const existing = getStudentByName(data.name, data.section);

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
 * Create or update a transcript record
 * @param {string} discussionId
 * @param {Object} data
 */
function upsertTranscript(discussionId, data) {
  const existing = findRow(CONFIG.SHEETS.TRANSCRIPTS, 'discussion_id', discussionId);
  const now = new Date().toISOString();

  if (existing) {
    data.updated_at = now;
    updateRow(CONFIG.SHEETS.TRANSCRIPTS, existing._rowIndex, data);
  } else {
    insertRow(CONFIG.SHEETS.TRANSCRIPTS, {
      discussion_id: discussionId,
      raw_transcript: data.raw_transcript || '',
      speaker_map: data.speaker_map || '{}',
      named_transcript: data.named_transcript || '',
      created_at: now,
      updated_at: now
    });
  }
}

/**
 * Get transcript for a discussion
 * @param {string} discussionId
 * @returns {Object|null}
 */
function getTranscript(discussionId) {
  return findRow(CONFIG.SHEETS.TRANSCRIPTS, 'discussion_id', discussionId);
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
 * Add data validation and formatting to sheets
 */
function formatSheets() {
  const ss = getSpreadsheet();

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
    }

    // Approved checkbox (validation only — insertCheckboxes pre-fills FALSE which
    // makes appendRow skip past 1000 rows)
    const approvedCol = getColumnIndex(CONFIG.SHEETS.DISCUSSIONS, 'approved');
    if (approvedCol > 0) {
      const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
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

    // Conditional formatting: error rows turn red
    const lastCol = discussionsSheet.getLastColumn();
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
  }

  // --- SpeakerMap sheet ---
  const speakerMapSheet = ss.getSheetByName(CONFIG.SHEETS.SPEAKER_MAP);
  if (speakerMapSheet) {
    const confirmedCol = getColumnIndex(CONFIG.SHEETS.SPEAKER_MAP, 'confirmed');
    if (confirmedCol > 0) {
      const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      speakerMapSheet.getRange(2, confirmedCol, 1000, 1).setDataValidation(checkboxRule);
    }
  }

  // --- StudentReports sheet ---
  const reportsSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENT_REPORTS);
  if (reportsSheet) {
    const approvedCol = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, 'approved');
    const sentCol = getColumnIndex(CONFIG.SHEETS.STUDENT_REPORTS, 'sent');

    if (approvedCol > 0) {
      const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      reportsSheet.getRange(2, approvedCol, 1000, 1).setDataValidation(checkboxRule);
    }
    if (sentCol > 0) {
      const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      reportsSheet.getRange(2, sentCol, 1000, 1).setDataValidation(checkboxRule);
    }
  }

  Logger.log('Sheet formatting applied.');
}
