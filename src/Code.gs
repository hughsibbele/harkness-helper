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
// CONSTANTS
// ============================================================================

/** Auto-stop processing after this many minutes */
const PROCESSING_TIMEOUT_MINUTES = 60;

// ============================================================================
// MENU & UI
// ============================================================================

/**
 * Create custom menu when spreadsheet opens.
 * Shows minimal menu before setup, full menu after.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  if (!isSetupComplete()) {
    ui.createMenu('Harkness Helper')
      .addItem('Setup Wizard (start here)', 'showSetupWizard')
      .addToUi();
    return;
  }

  // Check multi-course mode safely — onOpen() runs as a simple trigger
  // where PropertiesService may not be available
  let multiCourse = false;
  try {
    multiCourse = isMultiCourseMode();
  } catch (e) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Courses');
      if (sheet && sheet.getLastRow() > 1) multiCourse = true;
    } catch (e2) {
      // Ignore — default to single-course menu
    }
  }

  const menu = ui.createMenu('Harkness Helper')
    .addItem('Start Processing', 'menuStartProcessing')
    .addItem('Stop Processing', 'menuStopProcessing')
    .addItem('Process New Files Now', 'menuProcessNewFiles')
    .addSeparator()
    .addItem('Generate Feedback', 'menuGenerateFeedback')
    .addItem('Send Approved Feedback', 'menuSendApprovedFeedback')
    .addSeparator()
    .addItem('Configure Canvas Course', 'showCanvasConfigDialog')
    .addItem('Sync Canvas Roster', 'menuSyncCanvasRoster')
    .addItem('Fetch Canvas Course Data', 'menuFetchCanvasCourseData');

  if (multiCourse) {
    menu.addItem('Sync All Course Rosters', 'menuSyncAllCourseRosters');
  }

  menu.addSeparator()
    .addItem('Reorder & Format Sheets', 'menuReorderColumns')
    .addItem('Check Configuration', 'showConfigStatus');

  if (!multiCourse) {
    menu.addSeparator()
      .addItem('Enable Multi-Course', 'menuEnableMultiCourse');
  }

  menu.addSeparator()
    .addItem('Re-run Setup Wizard', 'showSetupWizard');

  menu.addToUi();
}

// ============================================================================
// SETUP WIZARD
// ============================================================================

/**
 * Show the Setup Wizard dialog.
 * Collects ElevenLabs and Gemini API keys, then auto-configures everything.
 */
function showSetupWizard() {
  const props = PropertiesService.getScriptProperties().getProperties();

  // Escape HTML special characters to prevent XSS via stored property values
  const escapeHtml = (str) => String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');

  const existingElevenlabs = escapeHtml(props['ELEVENLABS_API_KEY']);
  const existingGemini = escapeHtml(props['GEMINI_API_KEY']);

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; color: #333; }
      h2 { color: #1a73e8; margin-top: 0; }
      label { display: block; margin-top: 16px; font-weight: bold; font-size: 14px; }
      input { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ccc;
              border-radius: 4px; font-size: 14px; box-sizing: border-box; }
      .help { font-size: 12px; color: #666; margin-top: 2px; }
      button { margin-top: 24px; padding: 10px 24px; background: #1a73e8; color: white;
               border: none; border-radius: 4px; font-size: 14px; cursor: pointer; width: 100%; }
      button:hover { background: #1557b0; }
      button:disabled { background: #ccc; cursor: default; }
      .error { color: #d93025; margin-top: 8px; font-size: 13px; }
      .success { color: #188038; margin-top: 8px; font-size: 13px; }
    </style>

    <h2>Harkness Helper Setup</h2>
    <p>Enter your API keys below. Everything else is configured automatically.</p>

    <label for="elevenlabs">ElevenLabs API Key</label>
    <input type="password" id="elevenlabs" value="${existingElevenlabs}"
           placeholder="Enter your ElevenLabs API key">
    <div class="help">Get one at <b>elevenlabs.io</b> → Profile → API Keys</div>

    <label for="gemini">Gemini API Key</label>
    <input type="password" id="gemini" value="${existingGemini}"
           placeholder="Enter your Gemini API key">
    <div class="help">Get one at <b>aistudio.google.com</b> → API Keys</div>

    <div id="status"></div>

    <button id="btn" onclick="runSetup()">Run Setup</button>

    <script>
      function runSetup() {
        var elevenlabs = document.getElementById('elevenlabs').value.trim();
        var gemini = document.getElementById('gemini').value.trim();
        var status = document.getElementById('status');
        var btn = document.getElementById('btn');

        if (!elevenlabs || !gemini) {
          status.className = 'error';
          status.textContent = 'Both API keys are required.';
          return;
        }

        btn.disabled = true;
        btn.textContent = 'Setting up...';
        status.className = '';
        status.textContent = 'Creating folders, initializing sheets...';

        google.script.run
          .withSuccessHandler(function(result) {
            status.className = 'success';
            status.innerHTML = '<b>Setup complete!</b><br><br>' +
              'Next steps:<br>' +
              '1. Upload audio files to the <b>Harkness Helper / Upload</b> folder in your Google Drive<br>' +
              '2. Click <b>Harkness Helper → Start Processing</b><br>' +
              '3. Customize the <b>Settings</b> sheet (mode, distribution, etc.)<br><br>' +
              'Refresh this page to see the full menu.';
            btn.textContent = 'Done!';
          })
          .withFailureHandler(function(err) {
            status.className = 'error';
            status.textContent = 'Error: ' + err.message;
            btn.disabled = false;
            btn.textContent = 'Run Setup';
          })
          .runSetupWizard({ elevenlabs: elevenlabs, gemini: gemini });
      }
    </script>
  `).setWidth(480).setHeight(480);

  SpreadsheetApp.getUi().showModalDialog(html, 'Setup Wizard');
}

/**
 * Server-side handler for the Setup Wizard.
 * Auto-captures spreadsheet ID, stores API keys, creates Drive folders,
 * and initializes all sheets/settings/prompts.
 *
 * @param {Object} config - {elevenlabs, gemini}
 * @returns {Object} {success: true}
 */
function runSetupWizard(config) {
  // 1. Auto-capture spreadsheet ID
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  setProperty('SPREADSHEET_ID', ssId);

  // 2. Store API keys
  setProperty('ELEVENLABS_API_KEY', config.elevenlabs);
  setProperty('GEMINI_API_KEY', config.gemini);

  // 3. Create Drive folders (stores AUDIO_FOLDER_ID and PROCESSING_FOLDER_ID)
  setupDriveFolders();

  // 4. Initialize spreadsheet structure
  initializeAllSheets();
  formatSheets();
  initializeSettings();
  initializeDefaultPrompts();

  Logger.log('Setup Wizard completed successfully');
  return { success: true };
}

// ============================================================================
// CONFIGURATION STATUS
// ============================================================================

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
    message += 'Gemini model: ' + (getSetting('gemini_model') || 'gemini-2.0-flash') + '\n';
    message += 'Email distribution: ' + getSetting('distribute_email') + '\n';
    message += 'Canvas distribution: ' + getSetting('distribute_canvas') + '\n';
    message += 'Canvas integration: ' + (isCanvasConfigured() ? 'Configured' : 'Not configured') + '\n';
    if (isMultiCourseMode()) {
      const courses = getAllCourses();
      message += '\nMulti-course mode: ENABLED (' + courses.length + ' course' + (courses.length !== 1 ? 's' : '') + ')\n';
      courses.forEach(c => {
        message += '  - ' + c.course_name + ' (ID: ' + c.canvas_course_id + ')\n';
      });
    } else {
      message += '\nMulti-course mode: disabled (single-course)\n';
      const courseId = getSetting('canvas_course_id');
      if (courseId) message += 'Canvas course ID: ' + courseId + '\n';
    }
  } else {
    message = 'Missing required configuration:\n\n';
    message += validation.missing.map(k => '- ' + k).join('\n');
    message += '\n\nPlease run the Setup Wizard from the Harkness Helper menu.';
  }

  ui.alert('Configuration Status', message, ui.ButtonSet.OK);
}

// ============================================================================
// AUTO-OFF TRIGGER SYSTEM
// ============================================================================

/**
 * Start processing: install trigger and run immediately.
 * Removes any existing processing triggers first to prevent duplicates.
 */
function menuStartProcessing() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Remove any existing processing triggers
    removeProcessingTriggers();

    // Record start time
    setProperty('PROCESSING_START_TIME', new Date().toISOString());

    // Install 10-minute trigger
    ScriptApp.newTrigger('mainProcessingLoop')
      .timeBased()
      .everyMinutes(CONFIG.TRIGGERS.MAIN_LOOP)
      .create();

    // Run immediately
    mainProcessingLoop();

    ui.alert('Processing Started',
      'Processing is now running and will check for new files every 10 minutes.\n\n' +
      'It will automatically stop after ' + PROCESSING_TIMEOUT_MINUTES + ' minutes.\n\n' +
      'Upload audio files to your Harkness Helper / Upload folder in Google Drive.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Stop processing: remove trigger and clear start time.
 */
function menuStopProcessing() {
  const ui = SpreadsheetApp.getUi();
  removeProcessingTriggers();

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('PROCESSING_START_TIME');

  ui.alert('Processing Stopped',
    'Automatic processing has been stopped.',
    ui.ButtonSet.OK);
}

/**
 * Remove only mainProcessingLoop triggers (not all project triggers).
 */
function removeProcessingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'mainProcessingLoop') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  Logger.log(`Removed ${removed} processing trigger(s)`);
}

// ============================================================================
// MAIN PROCESSING LOOP (runs on 10-minute trigger)
// ============================================================================

/**
 * Main trigger function — runs periodically to process work.
 * Auto-stops after PROCESSING_TIMEOUT_MINUTES.
 */
function mainProcessingLoop() {
  // Check elapsed time — auto-stop if past timeout
  const props = PropertiesService.getScriptProperties();
  const startTime = props.getProperty('PROCESSING_START_TIME');
  if (!startTime) {
    // Orphaned trigger — no start time recorded. Remove it.
    Logger.log('No PROCESSING_START_TIME found. Removing orphaned trigger.');
    removeProcessingTriggers();
    return;
  }
  const elapsed = (Date.now() - new Date(startTime).getTime()) / 60000;
  if (elapsed >= PROCESSING_TIMEOUT_MINUTES) {
    Logger.log(`Processing timeout reached (${Math.round(elapsed)} minutes). Auto-stopping.`);
    removeProcessingTriggers();
    props.deleteProperty('PROCESSING_START_TIME');
    return;
  }

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

  // Save to transcript and discussion records
  upsertTranscript(discussionId, {
    speaker_map: JSON.stringify(confirmedMap),
    named_transcript: namedTranscript
  });

  updateDiscussion(discussionId, {
    status: CONFIG.STATUS.MAPPING,
    next_step: isGroupMode()
      ? 'Enter grade in the Grade column on this row, then click "Generate Feedback"'
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
        next_step: 'Enter grade in the Grade column on this row, then click "Generate Feedback"'
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
    let student = getStudentByName(studentName, discussion.section, discussion.course || null);
    if (!student) {
      const studentData = {
        name: studentName,
        email: '',
        section: discussion.section || ''
      };
      if (discussion.course) studentData.course = discussion.course;
      const studentId = upsertStudent(studentData);
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

      if (isGroupMode()) {
        // Group mode: need grade on Discussion row
        const grade = discussion.grade;
        if (!grade) {
          updateDiscussion(discussionId, {
            next_step: 'Enter grade in the Grade column on this row, then click "Generate Feedback" again'
          });
          continue;
        }

        const feedback = generateGroupFeedback(namedTranscript, String(grade));

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
              String(report.grade)
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
    let discussionFailed = 0;
    const discussionErrors = [];

    if (emailEnabled) {
      try {
        const emailResult = sendAllReportsForDiscussion(discussionId);
        totalEmailsSent += emailResult.sent;
        discussionFailed += emailResult.failed;
        discussionErrors.push(...emailResult.errors);
      } catch (e) {
        Logger.log(`Email error for ${discussionId}: ${e.message}`);
        discussionFailed++;
        discussionErrors.push(`Email error: ${e.message}`);
      }
    }

    // Post to Canvas
    if (canvasEnabled && discussion.canvas_assignment_id) {
      try {
        const canvasResult = postGradesForDiscussion(discussionId);
        totalCanvasPosted += canvasResult.success;
        discussionFailed += canvasResult.failed;
        discussionErrors.push(...canvasResult.errors);
      } catch (e) {
        Logger.log(`Canvas error for ${discussionId}: ${e.message}`);
        discussionFailed++;
        discussionErrors.push(`Canvas error: ${e.message}`);
      }
    }

    totalFailed += discussionFailed;

    // Only mark 'sent' if all deliveries succeeded
    if (discussionFailed === 0) {
      updateDiscussion(discussionId, {
        status: CONFIG.STATUS.SENT,
        next_step: 'Feedback sent!'
      });
    } else {
      const timestamp = new Date().toISOString();
      const existingErrors = discussion.error_message || '';
      const newError = `[${timestamp}] Partial send failure (${discussionFailed} failed): ${discussionErrors.join('; ')}`;
      updateDiscussion(discussionId, {
        next_step: `Send incomplete — ${discussionFailed} failed. Re-run Send to retry.`,
        error_message: existingErrors ? existingErrors + '\n' + newError : newError
      });
    }
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
// CANVAS CONFIGURATION
// ============================================================================

/**
 * 3-step prompt dialog for Canvas course configuration.
 * Stores credentials, sets course ID, and auto-syncs all rosters.
 */
function showCanvasConfigDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties().getProperties();

  // Step 1: Canvas base URL
  const existingUrl = props['CANVAS_BASE_URL'] || getSetting('canvas_base_url') || '';
  const urlResponse = ui.prompt('Configure Canvas Course (Step 1 of 3)',
    'Enter your Canvas base URL:\n(e.g., https://yourschool.instructure.com)' +
    (existingUrl ? '\n\nCurrent: ' + existingUrl : ''),
    ui.ButtonSet.OK_CANCEL);
  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;
  const baseUrl = urlResponse.getResponseText().trim().replace(/\/$/, '') || existingUrl;
  if (!baseUrl) {
    ui.alert('Error', 'Canvas base URL is required.', ui.ButtonSet.OK);
    return;
  }

  // Step 2: API token
  const existingToken = props['CANVAS_API_TOKEN'] ? '(already set)' : '';
  const tokenResponse = ui.prompt('Configure Canvas Course (Step 2 of 3)',
    'Enter your Canvas API token:\n(Generate at Canvas → Account → Settings → New Access Token)' +
    (existingToken ? '\n\nLeave blank to keep existing token.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (tokenResponse.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenResponse.getResponseText().trim();
  if (!token && !props['CANVAS_API_TOKEN']) {
    ui.alert('Error', 'Canvas API token is required.', ui.ButtonSet.OK);
    return;
  }

  // Step 3: Course ID
  const existingCourseId = getSetting('canvas_course_id') || '';
  const courseResponse = ui.prompt('Configure Canvas Course (Step 3 of 3)',
    'Enter the Canvas course ID:\n(Find it in the URL: canvas.school.edu/courses/12345 → 12345)' +
    (existingCourseId ? '\n\nCurrent: ' + existingCourseId : ''),
    ui.ButtonSet.OK_CANCEL);
  if (courseResponse.getSelectedButton() !== ui.Button.OK) return;
  const courseId = courseResponse.getResponseText().trim() || existingCourseId;
  if (!courseId) {
    ui.alert('Error', 'Course ID is required.', ui.ButtonSet.OK);
    return;
  }

  // Store configuration
  setProperty('CANVAS_BASE_URL', baseUrl);
  if (token) {
    setProperty('CANVAS_API_TOKEN', token);
  }
  setSetting('canvas_base_url', baseUrl);
  setSetting('canvas_course_id', courseId);
  setSetting('distribute_canvas', 'true');

  // Auto-sync all rosters
  try {
    const sections = getCanvasSections(courseId);
    const sectionMap = {};
    for (const section of sections) {
      sectionMap[section.id] = section.name;
    }

    const fallbackSection = sections.length === 1 ? sections[0].name : '';
    const count = syncCanvasStudents(courseId, fallbackSection, sectionMap);

    ui.alert('Canvas Configured',
      `Canvas course configured and roster synced!\n\n` +
      `Synced ${count} students from ${sections.length} section(s).\n` +
      `Sections: ${sections.map(s => s.name).join(', ')}\n\n` +
      `To switch courses or re-sync, run this dialog again.`,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Partial Success',
      'Canvas credentials saved, but roster sync failed:\n' + e.message +
      '\n\nCheck your API token and course ID, then try again.',
      ui.ButtonSet.OK);
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
      'Please configure Canvas first via Harkness Helper → Configure Canvas Course.',
      ui.ButtonSet.OK);
    return;
  }

  // Multi-course mode: prompt teacher to pick a course
  if (isMultiCourseMode()) {
    const courses = getAllCourses();
    const courseNames = courses.map(c => c.course_name);
    const response = ui.prompt('Select Course',
      'Which course roster to sync?\n\nAvailable courses: ' + courseNames.join(', ') +
      '\n\nEnter course name:',
      ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;

    const selectedName = response.getResponseText().trim();
    const course = getCourseByName(selectedName);
    if (!course) {
      ui.alert('Error', 'Course "' + selectedName + '" not found on the Courses sheet.', ui.ButtonSet.OK);
      return;
    }

    try {
      const courseId = String(course.canvas_course_id);
      const sections = getCanvasSections(courseId);
      const sectionMap = {};
      for (const section of sections) {
        sectionMap[section.id] = section.name;
      }
      const fallbackSection = sections.length === 1 ? sections[0].name : '';
      const count = syncCanvasStudents(courseId, fallbackSection, sectionMap, selectedName);

      let message = `Synced ${count} students for "${selectedName}".\n\n`;
      message += `Sections found: ${sections.map(s => s.name).join(', ')}\n`;
      message += 'Students tagged with course "' + selectedName + '".';
      ui.alert('Roster Synced', message, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Error', e.message, ui.ButtonSet.OK);
    }
    return;
  }

  // Single-course mode
  const courseId = getSetting('canvas_course_id');
  if (!courseId) {
    ui.alert('No Course ID',
      'Please set canvas_course_id in the Settings sheet or use Configure Canvas Course.',
      ui.ButtonSet.OK);
    return;
  }

  try {
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
      'Please configure Canvas first via Harkness Helper → Configure Canvas Course.',
      ui.ButtonSet.OK);
    return;
  }

  let courseId;
  let courseName = null;

  if (isMultiCourseMode()) {
    // Multi-course mode: prompt to pick or enter a course
    const courses = getAllCourses();
    const courseNames = courses.map(c => c.course_name);
    const response = ui.prompt('Select Course',
      'Which course to fetch data for?\n\nAvailable courses: ' + courseNames.join(', ') +
      '\n\nEnter course name (or enter a Canvas course ID to add a new course):',
      ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;

    const input = response.getResponseText().trim();
    const existingCourse = getCourseByName(input);

    if (existingCourse) {
      courseId = String(existingCourse.canvas_course_id);
      courseName = existingCourse.course_name;
    } else if (/^\d+$/.test(input)) {
      // Looks like a course ID — offer to add it
      courseId = input;
      const addResponse = ui.prompt('Add Course',
        'Course ID ' + courseId + ' is not in the Courses sheet yet.\n\nEnter a friendly name for this course (e.g., "AP English"):',
        ui.ButtonSet.OK_CANCEL);
      if (addResponse.getSelectedButton() === ui.Button.OK) {
        courseName = addResponse.getResponseText().trim();
        if (courseName) {
          insertRow(CONFIG.SHEETS.COURSES, {
            course_name: courseName,
            canvas_course_id: courseId,
            canvas_base_url: '',
            canvas_item_type: ''
          });
        }
      }
    } else {
      ui.alert('Error', 'Course "' + input + '" not found on the Courses sheet.', ui.ButtonSet.OK);
      return;
    }
  } else {
    // Single-course mode
    courseId = getSetting('canvas_course_id');
    if (!courseId) {
      const response = ui.prompt('Canvas Course ID',
        'Enter the Canvas course ID:',
        ui.ButtonSet.OK_CANCEL);
      if (response.getSelectedButton() !== ui.Button.OK) return;
      courseId = response.getResponseText().trim();
      setSetting('canvas_course_id', courseId);
    }
  }

  try {
    const result = fetchCanvasCourseData(courseId, courseName);

    // Build info for dialog
    let info = '';
    info += `Course: ${result.courseName}\n`;
    info += `Students synced: ${result.studentCount}\n`;
    info += `Sections: ${result.sections.map(s => s.name).join(', ')}\n`;
    info += '\nStudents have been added to the Students sheet, auto-assigned to their Canvas section.\n';

    info += '\n==================================================\n';
    info += 'WHAT TO DO NEXT\n';
    info += '==================================================\n\n';

    info += '1. CHECK THE STUDENTS SHEET\n';
    info += '   Verify students are listed with the correct section.\n\n';

    info += '2. SET YOUR CANVAS ITEM TYPE (one-time)\n';
    info += '   In the Settings sheet, set canvas_item_type to:\n';
    info += '   - "discussion" if your Canvas grades are on Discussion Topics\n';
    info += '   - "assignment" if your Canvas grades are on regular Assignments\n\n';

    info += '3. SET THE CANVAS ID ON EACH DISCUSSION ROW\n';
    info += '   Each Harkness discussion maps to a Canvas item. Find the\n';
    info += '   matching item below and enter its ID in the canvas_assignment_id\n';
    info += '   column on the Discussions sheet. If both sections did the same\n';
    info += '   Harkness, both rows get the same ID — grades are sent\n';
    info += '   per-section automatically based on each row\'s section column.\n\n';

    info += '4. MAKE SURE THE SECTION IS SET\n';
    info += '   Each Discussion row has a "section" column. When grades are sent,\n';
    info += '   only students in that section receive them. If you name your audio\n';
    info += '   files like "Section 1 - 2024-01-15.m4a", this is set automatically.\n';
    info += '   Otherwise, type the section name manually on the Discussions sheet.\n';

    info += '\n==================================================\n';

    if (result.assignments.length > 0) {
      info += 'ASSIGNMENTS\n';
      info += '==================================================\n\n';
      for (const a of result.assignments) {
        info += `ID: ${a.id} | ${a.name} | Due: ${a.due_at} | Points: ${a.points_possible}\n`;
      }
    }

    if (result.discussionTopics.length > 0) {
      info += '\n==================================================\n';
      info += 'GRADED DISCUSSION TOPICS\n';
      info += '==================================================\n\n';
      for (const d of result.discussionTopics) {
        info += `ID: ${d.id} | ${d.name} | Due: ${d.due_at} | Points: ${d.points_possible}\n`;
      }
    }

    // Show in a dialog (HTML for scrollability)
    const html = HtmlService.createHtmlOutput(
      '<pre style="font-family: monospace; font-size: 12px; white-space: pre-wrap;">' +
      info.replace(/</g, '&lt;') +
      '</pre>'
    ).setWidth(700).setHeight(600);

    ui.showModalDialog(html, 'Canvas Course Data');
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Sync rosters for all courses in the Courses sheet
 */
function menuSyncAllCourseRosters() {
  const ui = SpreadsheetApp.getUi();

  if (!isCanvasConfigured()) {
    ui.alert('Canvas Not Configured',
      'Please configure Canvas first via Harkness Helper → Configure Canvas Course.',
      ui.ButtonSet.OK);
    return;
  }

  const courses = getAllCourses();
  if (courses.length === 0) {
    ui.alert('No Courses', 'No courses found on the Courses sheet.', ui.ButtonSet.OK);
    return;
  }

  try {
    let totalStudents = 0;
    const results = [];

    for (const course of courses) {
      const courseId = String(course.canvas_course_id);
      const sections = getCanvasSections(courseId);
      const sectionMap = {};
      for (const section of sections) {
        sectionMap[section.id] = section.name;
      }
      const fallbackSection = sections.length === 1 ? sections[0].name : '';
      const count = syncCanvasStudents(courseId, fallbackSection, sectionMap, course.course_name);
      totalStudents += count;
      results.push(`${course.course_name}: ${count} students`);
    }

    let message = `Synced ${totalStudents} students across ${courses.length} courses:\n\n`;
    message += results.join('\n');
    ui.alert('All Rosters Synced', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Enable multi-course mode by running the migration
 */
function menuEnableMultiCourse() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert('Enable Multi-Course',
    'This will:\n' +
    '1. Add a "course" column to Discussions and Students sheets\n' +
    '2. Create a Courses sheet\n' +
    '3. Pre-populate with your current canvas_course_id (if set)\n\n' +
    'Your existing data will not be affected. Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  try {
    migrateToMultiCourse();
    ui.alert('Multi-Course Enabled',
      'Migration complete!\n\n' +
      'Next steps:\n' +
      '1. Go to the Courses sheet\n' +
      '2. Rename "My Course" to your actual course name\n' +
      '3. Add additional courses with their Canvas course IDs\n' +
      '4. Run "Sync Canvas Roster" for each course',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Reorder columns in all sheets to the canonical order and re-apply formatting
 */
function menuReorderColumns() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert('Reorder & Format Sheets',
    'This will:\n' +
    '1. Reorder columns to put teacher-facing columns first\n' +
    '2. Apply highlighting, column widths, and text wrapping\n' +
    '3. Set tab colors and logical tab order\n\n' +
    'Your data will be preserved — only column positions change. Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  try {
    reorderColumns();
    ui.alert('Done', 'Columns reordered and formatting applied.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}
