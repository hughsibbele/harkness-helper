/**
 * Harkness Discussion Helper — Main Entry Point
 *
 * Pipeline:
 * 1. Monitor Drive folder for uploaded audio recordings
 * 2. Transcribe with ElevenLabs (synchronous, with speaker diarization)
 * 3. Gemini auto-suggests speaker mapping from introductions
 * 4. Teacher reviews speaker map (individual mode) or auto-advances (group mode)
 * 5. Teacher enters grade(s), clicks "Generate Feedback" → Gemini generates
 * 6. Teacher approves → clicks "Send" → email and/or Canvas distribution
 *
 * Status machine: uploaded → transcribing → mapping → review → approved → sent
 *                                                                        ↘ error
 */

// ============================================================================
// MENU & UI
// ============================================================================

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Harkness Helper')
    .addItem('Setup Folders', 'setupDriveFolders')
    .addItem('Initialize Sheets', 'initializeAllSheets')
    .addItem('Check Configuration', 'showConfigStatus')
    .addSeparator()
    .addItem('Process New Files', 'menuProcessNewFiles')
    .addItem('Generate Feedback', 'menuGenerateFeedback')
    .addItem('Send Approved Feedback', 'menuSendApprovedFeedback')
    .addSeparator()
    .addItem('Sync Canvas Roster', 'menuSyncCanvasRoster')
    .addItem('Fetch Canvas Course Data', 'menuFetchCanvasCourseData')
    .addSeparator()
    .addItem('Setup Automatic Triggers', 'setupTriggers')
    .addItem('Remove All Triggers', 'removeAllTriggers')
    .addToUi();
}

/**
 * Show configuration status dialog
 */
function showConfigStatus() {
  const ui = SpreadsheetApp.getUi();
  const validation = validateConfiguration();

  let message = '';
  if (validation.valid) {
    message = 'All required configuration is set!\n\n';
    message += 'Mode: ' + getMode() + '\n';
    message += 'Email distribution: ' + getSetting('distribute_email') + '\n';
    message += 'Canvas distribution: ' + getSetting('distribute_canvas') + '\n';
    message += 'Canvas integration: ' + (isCanvasConfigured() ? 'Configured' : 'Not configured');
  } else {
    message = 'Missing required configuration:\n\n';
    message += validation.missing.map(k => '- ' + k).join('\n');
    message += '\n\nPlease add these in Project Settings > Script Properties.';
  }

  ui.alert('Configuration Status', message, ui.ButtonSet.OK);
}

// ============================================================================
// MAIN PROCESSING LOOP (runs on 10-minute trigger)
// ============================================================================

/**
 * Main trigger function — runs periodically to process work
 */
function mainProcessingLoop() {
  Logger.log('=== Starting main processing loop ===');

  try {
    // Step 1: Check for new audio files and transcribe
    checkAndProcessNewFiles();

    // Step 2: Detect stuck transcriptions
    checkStuckTranscriptions();

    // Step 3: In group mode, auto-advance confirmed discussions from mapping → review
    advanceMappingStatus();

    Logger.log('=== Main processing loop complete ===');
  } catch (e) {
    Logger.log(`Error in main processing loop: ${e.message}`);
    Logger.log(e.stack);
  }
}

// ============================================================================
// FILE DETECTION & TRANSCRIPTION
// ============================================================================

/**
 * Check for new audio files, transcribe them, and run speaker mapping
 */
function checkAndProcessNewFiles() {
  const validation = validateConfiguration();
  if (!validation.valid) {
    Logger.log('Configuration incomplete, skipping file check');
    return;
  }

  const newFiles = checkForNewAudioFiles();
  Logger.log(`Found ${newFiles.length} new audio files`);

  for (const fileInfo of newFiles) {
    const discussionId = processNewAudioFile(fileInfo);
    Logger.log(`Created discussion: ${discussionId}`);

    try {
      // Synchronous transcription with ElevenLabs
      startTranscription(discussionId);

      // Run speaker identification and populate SpeakerMap
      processSpeakerMapping(discussionId);
    } catch (e) {
      const timestamp = new Date().toISOString();
      Logger.log(`Failed to process ${discussionId}: ${e.message}`);
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.ERROR,
        error_message: `[${timestamp}] Transcription/mapping: ${e.message}`
      });
    }
  }
}

// ============================================================================
// SPEAKER MAPPING
// ============================================================================

/**
 * Run Gemini speaker identification and populate the SpeakerMap sheet.
 * Also generates named transcript, extracts teacher feedback, and creates summary.
 *
 * @param {string} discussionId
 */
function processSpeakerMapping(discussionId) {
  Logger.log(`Processing speaker mapping for ${discussionId}`);

  const transcript = getTranscript(discussionId);
  if (!transcript || !transcript.raw_transcript) {
    throw new Error('No transcript found for speaker mapping');
  }

  const rawTranscript = String(transcript.raw_transcript);

  // Get first ~3 minutes of transcript for speaker ID
  const lines = rawTranscript.split('\n');
  const excerptLines = [];
  for (const line of lines) {
    excerptLines.push(line);
    // Check if we've reached ~3 minutes by parsing timestamps
    const timeMatch = line.match(/\[(\d+):/);
    if (timeMatch && parseInt(timeMatch[1]) >= 3) break;
  }
  const excerpt = excerptLines.join('\n') || rawTranscript.substring(0, 5000);

  // Gemini speaker identification
  const speakerMap = identifySpeakers(excerpt);

  // Also get all speaker labels from transcript (in case Gemini missed some)
  const allLabels = extractSpeakerLabels(rawTranscript);

  // Populate SpeakerMap sheet
  for (const label of allLabels) {
    const suggestedName = speakerMap[label] || '?';
    upsertSpeakerMapping(discussionId, label, suggestedName);
  }

  // Build map from SpeakerMap sheet and generate named transcript
  const confirmedMap = buildSpeakerMapObject(discussionId);
  const namedTranscript = applySpeakerNames(rawTranscript, confirmedMap);

  // Extract teacher feedback
  Utilities.sleep(300);
  const teacherFeedback = extractTeacherFeedback(namedTranscript);

  // Generate discussion summary
  Utilities.sleep(300);
  const summary = generateDiscussionSummary(namedTranscript);

  // Save to transcript and discussion records
  upsertTranscript(discussionId, {
    speaker_map: JSON.stringify(confirmedMap),
    named_transcript: namedTranscript,
    teacher_feedback: teacherFeedback
  });

  updateDiscussion(discussionId, {
    status: CONFIG.STATUS.MAPPING,
    discussion_summary: summary,
    next_step: isGroupMode()
      ? 'Enter grade on this row, then click "Generate Feedback"'
      : 'Review speaker map in SpeakerMap sheet, confirm all speakers, then click "Generate Feedback"'
  });

  Logger.log(`Speaker mapping complete for ${discussionId}`);
}

// ============================================================================
// STATUS ADVANCEMENT
// ============================================================================

/**
 * Auto-advance discussions from mapping → review where appropriate.
 * In group mode: auto-advance since speaker map is auto-confirmed.
 * In individual mode: only advance if teacher has confirmed all speakers.
 */
function advanceMappingStatus() {
  const mappingDiscussions = getDiscussionsByStatus(CONFIG.STATUS.MAPPING);

  for (const discussion of mappingDiscussions) {
    const discussionId = discussion.discussion_id;

    if (isGroupMode()) {
      // Group mode: auto-advance to review
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.REVIEW,
        next_step: 'Enter grade on this row, then click "Generate Feedback"'
      });
      Logger.log(`Auto-advanced ${discussionId} to review (group mode)`);
    } else if (isSpeakerMapConfirmed(discussionId)) {
      // Individual mode: advance only when speakers confirmed
      createStudentReportsFromSpeakerMap(discussionId);
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.REVIEW,
        next_step: 'Enter grades in StudentReports, then click "Generate Feedback"'
      });
      Logger.log(`Advanced ${discussionId} to review (speakers confirmed)`);
    }
  }
}

/**
 * Create StudentReport rows from confirmed SpeakerMap entries (individual mode).
 * @param {string} discussionId
 */
function createStudentReportsFromSpeakerMap(discussionId) {
  const discussion = getDiscussion(discussionId);
  const transcript = getTranscript(discussionId);
  const namedTranscript = String(transcript.named_transcript || '');
  const speakerMap = buildSpeakerMapObject(discussionId);
  const studentNames = getStudentNames(speakerMap);

  for (const studentName of studentNames) {
    // Find or create student record
    let student = getStudentByName(studentName, discussion.section);
    if (!student) {
      const studentId = upsertStudent({
        name: studentName,
        email: '',
        section: discussion.section || ''
      });
      student = { student_id: studentId };
    }

    // Skip if report already exists
    const existing = getStudentReport(discussionId, student.student_id);
    if (existing) continue;

    // Extract this student's contributions
    const contributions = extractStudentContributions(namedTranscript, studentName);

    createStudentReport({
      discussion_id: discussionId,
      student_id: student.student_id,
      student_name: studentName,
      transcript_contributions: contributions
    });
  }

  Logger.log(`Created student reports for ${studentNames.length} students`);
}

// ============================================================================
// FEEDBACK GENERATION (menu action)
// ============================================================================

/**
 * Generate feedback for discussions in review status.
 * Group mode: generates group_feedback on the Discussion row.
 * Individual mode: generates per-student feedback on StudentReport rows.
 */
function generateFeedbackForDiscussions() {
  const discussions = [
    ...getDiscussionsByStatus(CONFIG.STATUS.REVIEW),
    ...getDiscussionsByStatus(CONFIG.STATUS.MAPPING)
  ];

  let generated = 0;

  for (const discussion of discussions) {
    const discussionId = discussion.discussion_id;

    try {
      const transcript = getTranscript(discussionId);
      if (!transcript) continue;

      const namedTranscript = String(transcript.named_transcript || transcript.raw_transcript || '');
      const teacherFeedback = String(transcript.teacher_feedback || '');

      if (isGroupMode()) {
        // Group mode: need grade on Discussion row
        const grade = discussion.grade;
        if (!grade) {
          updateDiscussion(discussionId, {
            next_step: 'Enter grade on this row, then click "Generate Feedback" again'
          });
          continue;
        }

        const feedback = generateGroupFeedback(namedTranscript, String(grade), teacherFeedback);

        updateDiscussion(discussionId, {
          group_feedback: feedback,
          status: CONFIG.STATUS.REVIEW,
          next_step: 'Review feedback, check approved, then click "Send Approved Feedback"'
        });

        generated++;
      } else {
        // Individual mode: need per-student grades
        if (!isSpeakerMapConfirmed(discussionId)) {
          updateDiscussion(discussionId, {
            next_step: 'Confirm all speakers in SpeakerMap sheet first'
          });
          continue;
        }

        // Ensure StudentReports exist
        const reports = getReportsForDiscussion(discussionId);
        if (reports.length === 0) {
          createStudentReportsFromSpeakerMap(discussionId);
        }

        const updatedReports = getReportsForDiscussion(discussionId);
        let studentsNeedGrades = false;

        for (const report of updatedReports) {
          if (!report.grade) {
            studentsNeedGrades = true;
            continue;
          }

          // Skip if feedback already generated
          if (report.feedback) continue;

          try {
            Utilities.sleep(500);
            const contributions = String(report.transcript_contributions || '');
            const feedback = generateIndividualFeedback(
              report.student_name,
              contributions,
              namedTranscript,
              String(report.grade),
              teacherFeedback
            );

            updateStudentReport(report.report_id, { feedback: feedback });
            generated++;
          } catch (e) {
            Logger.log(`Error generating feedback for ${report.student_name}: ${e.message}`);
            updateStudentReport(report.report_id, {
              feedback: `ERROR: ${e.message}`
            });
          }
        }

        if (studentsNeedGrades) {
          updateDiscussion(discussionId, {
            next_step: 'Enter grades for all students in StudentReports, then click "Generate Feedback" again'
          });
        } else {
          updateDiscussion(discussionId, {
            status: CONFIG.STATUS.REVIEW,
            next_step: 'Review feedback, approve students, then click "Send Approved Feedback"'
          });
        }
      }
    } catch (e) {
      const timestamp = new Date().toISOString();
      Logger.log(`Error generating feedback for ${discussionId}: ${e.message}`);
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.ERROR,
        error_message: `[${timestamp}] Feedback generation: ${e.message}`
      });
    }
  }

  return generated;
}

// ============================================================================
// SEND APPROVED FEEDBACK (menu action)
// ============================================================================

/**
 * Send feedback for approved discussions via enabled channels.
 * Checks distribute_email and distribute_canvas settings.
 */
function sendApprovedFeedback() {
  const emailEnabled = getSetting('distribute_email') === 'true';
  const canvasEnabled = getSetting('distribute_canvas') === 'true' && isCanvasConfigured();

  const discussions = [
    ...getDiscussionsByStatus(CONFIG.STATUS.REVIEW),
    ...getDiscussionsByStatus(CONFIG.STATUS.APPROVED)
  ];

  let totalEmailsSent = 0;
  let totalCanvasPosted = 0;
  let totalFailed = 0;

  for (const discussion of discussions) {
    const discussionId = discussion.discussion_id;

    // Check if approved
    if (isGroupMode()) {
      if (discussion.approved !== true) continue;
    } else {
      const reports = getApprovedUnsentReports(discussionId);
      if (reports.length === 0) continue;
    }

    // Send emails
    if (emailEnabled) {
      try {
        const emailResult = sendAllReportsForDiscussion(discussionId);
        totalEmailsSent += emailResult.sent;
        totalFailed += emailResult.failed;
      } catch (e) {
        Logger.log(`Email error for ${discussionId}: ${e.message}`);
        totalFailed++;
      }
    }

    // Post to Canvas
    if (canvasEnabled && discussion.canvas_assignment_id) {
      try {
        const canvasResult = postGradesForDiscussion(discussionId);
        totalCanvasPosted += canvasResult.success;
        totalFailed += canvasResult.failed;
      } catch (e) {
        Logger.log(`Canvas error for ${discussionId}: ${e.message}`);
        totalFailed++;
      }
    }

    // Update status
    updateDiscussion(discussionId, {
      status: CONFIG.STATUS.SENT,
      next_step: 'Feedback sent!'
    });
  }

  return {
    emailsSent: totalEmailsSent,
    canvasPosted: totalCanvasPosted,
    failed: totalFailed,
    emailEnabled: emailEnabled,
    canvasEnabled: canvasEnabled
  };
}

// ============================================================================
// FAILSAFES
// ============================================================================

/**
 * Detect discussions stuck in transcribing status for too long.
 * Marks them as error with a helpful message.
 */
function checkStuckTranscriptions() {
  const transcribing = getDiscussionsByStatus(CONFIG.STATUS.TRANSCRIBING);
  const now = Date.now();

  for (const discussion of transcribing) {
    const updatedAt = new Date(discussion.updated_at).getTime();
    const elapsed = now - updatedAt;

    if (elapsed > CONFIG.LIMITS.TRANSCRIPTION_TIMEOUT_MS) {
      const timestamp = new Date().toISOString();
      Logger.log(`Discussion ${discussion.discussion_id} stuck in transcribing for ${Math.round(elapsed / 60000)} minutes`);
      updateDiscussion(discussion.discussion_id, {
        status: CONFIG.STATUS.ERROR,
        error_message: `[${timestamp}] Transcription timed out after ${Math.round(elapsed / 60000)} minutes. Try splitting the audio file into shorter segments.`
      });
    }
  }
}

// ============================================================================
// MENU ACTION WRAPPERS (with UI feedback)
// ============================================================================

function menuProcessNewFiles() {
  const ui = SpreadsheetApp.getUi();
  try {
    checkAndProcessNewFiles();
    ui.alert('Process New Files', 'File processing complete. Check the Discussions sheet for results.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function menuGenerateFeedback() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = generateFeedbackForDiscussions();
    ui.alert('Generate Feedback',
      `Generated feedback for ${count} item(s).\n\nCheck the Discussions and StudentReports sheets.`,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function menuSendApprovedFeedback() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = sendApprovedFeedback();
    let message = '';
    if (result.emailEnabled) {
      message += `Emails sent: ${result.emailsSent}\n`;
    } else {
      message += 'Email distribution: disabled in Settings\n';
    }
    if (result.canvasEnabled) {
      message += `Canvas grades posted: ${result.canvasPosted}\n`;
    } else {
      message += 'Canvas distribution: disabled in Settings\n';
    }
    if (result.failed > 0) {
      message += `\nFailed: ${result.failed} (check Execution log for details)`;
    }
    ui.alert('Send Feedback', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function menuSyncCanvasRoster() {
  const ui = SpreadsheetApp.getUi();

  if (!isCanvasConfigured()) {
    ui.alert('Canvas Not Configured',
      'Please add CANVAS_API_TOKEN and CANVAS_BASE_URL to Script Properties.',
      ui.ButtonSet.OK);
    return;
  }

  const courseId = getSetting('canvas_course_id');
  if (!courseId) {
    ui.alert('No Course ID',
      'Please set canvas_course_id in the Settings sheet.',
      ui.ButtonSet.OK);
    return;
  }

  try {
    // Fetch sections to auto-assign students
    const sections = getCanvasSections(courseId);
    const sectionMap = {};
    for (const section of sections) {
      sectionMap[section.id] = section.name;
    }

    const fallbackSection = sections.length === 1 ? sections[0].name : '';
    const count = syncCanvasStudents(courseId, fallbackSection, sectionMap);

    let message = `Synced ${count} students from Canvas.\n\n`;
    message += `Sections found: ${sections.map(s => s.name).join(', ')}\n`;
    message += 'Students auto-assigned to their Canvas section.';
    ui.alert('Roster Synced', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function menuFetchCanvasCourseData() {
  const ui = SpreadsheetApp.getUi();

  if (!isCanvasConfigured()) {
    ui.alert('Canvas Not Configured',
      'Please add CANVAS_API_TOKEN and CANVAS_BASE_URL to Script Properties.',
      ui.ButtonSet.OK);
    return;
  }

  let courseId = getSetting('canvas_course_id');
  if (!courseId) {
    const response = ui.prompt('Canvas Course ID',
      'Enter the Canvas course ID:',
      ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    courseId = response.getResponseText().trim();
    setSetting('canvas_course_id', courseId);
  }

  try {
    const result = fetchCanvasCourseData(courseId);

    // Build info for dialog
    let info = `Course: ${result.courseName}\nStudents synced: ${result.studentCount}\n\n`;

    info += 'SECTIONS:\n';
    for (const s of result.sections) {
      info += `  ${s.name} (ID: ${s.id})\n`;
    }
    info += '\nStudents have been auto-assigned to their Canvas section.\n';
    info += 'Check the Students sheet to verify.\n\n';

    info += 'ASSIGNMENTS (set canvas_item_type to "assignment" in Settings):\n\n';
    for (const a of result.assignments) {
      info += `ID: ${a.id} | ${a.name} | Due: ${a.due_at} | Points: ${a.points_possible}\n`;
    }

    if (result.discussionTopics.length > 0) {
      info += '\nGRADED DISCUSSION TOPICS (set canvas_item_type to "discussion" in Settings):\n\n';
      for (const d of result.discussionTopics) {
        info += `ID: ${d.id} | ${d.name} | Due: ${d.due_at} | Points: ${d.points_possible}\n`;
      }
    }

    info += '\nEnter the ID in the canvas_assignment_id column on the Discussions sheet.\n';
    info += 'Set canvas_item_type in Settings to match your item type.\n';

    // Show in a dialog (HTML for scrollability)
    const html = HtmlService.createHtmlOutput(
      '<pre style="font-family: monospace; font-size: 12px; white-space: pre-wrap;">' +
      info.replace(/</g, '&lt;') +
      '</pre>'
    ).setWidth(700).setHeight(500);

    ui.showModalDialog(html, 'Canvas Course Data');
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// TRIGGER MANAGEMENT
// ============================================================================

/**
 * Setup time-based triggers for automatic processing
 */
function setupTriggers() {
  removeAllTriggers();

  ScriptApp.newTrigger('mainProcessingLoop')
    .timeBased()
    .everyMinutes(CONFIG.TRIGGERS.MAIN_LOOP)
    .create();

  Logger.log('Triggers created successfully');
  SpreadsheetApp.getUi().alert(
    'Triggers Setup',
    `Automatic processing enabled.\nMain loop runs every ${CONFIG.TRIGGERS.MAIN_LOOP} minutes.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Remove all project triggers
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  Logger.log(`Removed ${triggers.length} triggers`);
}

// ============================================================================
// SETUP & INITIALIZATION
// ============================================================================

/**
 * One-time setup function — run this first!
 */
function initialSetup() {
  Logger.log('=== Running initial setup ===');

  // 1. Setup script properties
  setupScriptProperties();

  // 2. Create Drive folders
  const folders = setupDriveFolders();
  Logger.log(`Created folders: ${JSON.stringify(folders)}`);

  // 3. Create spreadsheet structure
  try {
    initializeAllSheets();
    formatSheets();
    initializeSettings();
    initializeDefaultPrompts();
  } catch (e) {
    Logger.log('Could not initialize sheets (run from spreadsheet): ' + e.message);
  }

  Logger.log('=== Initial setup complete ===');
  Logger.log('Next steps:');
  Logger.log('1. Add API keys in Project Settings > Script Properties');
  Logger.log('2. Set SPREADSHEET_ID to your tracking spreadsheet');
  Logger.log('3. Configure Settings sheet (mode, distribution, etc.)');
  Logger.log('4. Run setupTriggers() to enable automatic processing');
  Logger.log(`5. Upload audio files to: ${getUploadFolderUrl()}`);
}

/**
 * Test the configuration by making test API calls
 */
function testConfiguration() {
  const results = [];

  // Test ElevenLabs
  try {
    getElevenLabsKey();
    results.push('ElevenLabs API key: configured');
  } catch (e) {
    results.push('ElevenLabs: ' + e.message);
  }

  // Test Gemini
  try {
    getGeminiKey();
    const response = callGemini('Say "Hello" in exactly one word.');
    results.push('Gemini API: working');
  } catch (e) {
    results.push('Gemini: ' + e.message);
  }

  // Test Canvas (if configured)
  if (isCanvasConfigured()) {
    try {
      const courseId = getSetting('canvas_course_id');
      if (courseId) {
        canvasRequest(`/courses/${courseId}`);
        results.push('Canvas API: working');
      } else {
        results.push('Canvas: Course ID not set in Settings');
      }
    } catch (e) {
      results.push('Canvas: ' + e.message);
    }
  } else {
    results.push('Canvas: Not configured (optional)');
  }

  // Test Drive folders
  try {
    const uploadId = getAudioFolderId();
    const processingId = getProcessingFolderId();
    DriveApp.getFolderById(uploadId);
    DriveApp.getFolderById(processingId);
    results.push('Drive folders: accessible');
  } catch (e) {
    results.push('Drive folders: ' + e.message);
  }

  // Test Settings sheet
  try {
    const mode = getMode();
    results.push('Settings sheet: OK (mode=' + mode + ')');
  } catch (e) {
    results.push('Settings sheet: ' + e.message);
  }

  Logger.log('=== Configuration Test Results ===');
  results.forEach(r => Logger.log(r));

  return results;
}
