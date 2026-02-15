/**
 * ElevenLabs Speech-to-Text integration for transcription with speaker diarization
 *
 * Uses ElevenLabs Scribe v2 for synchronous transcription. The API returns
 * a words[] array with speaker_id, which we group into speaker-labeled lines.
 *
 * Two upload paths:
 * - Files <= 50MB: direct blob upload (multipart form data)
 * - Files > 50MB: cloud_storage_url with temporary public Drive link
 */

// ============================================================================
// TRANSCRIPTION SUBMISSION
// ============================================================================

/**
 * Submit a transcription via direct blob upload (for files <= 50MB)
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} ElevenLabs API response
 */
function submitTranscriptionBlob(fileId) {
  const apiKey = getElevenLabsKey();
  const sttModel = getSetting('elevenlabs_model') || 'scribe_v2';
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();

  const boundary = '----FormBoundary' + Utilities.getUuid();
  const contentType = 'multipart/form-data; boundary=' + boundary;

  // Build multipart form data manually for GAS compatibility
  const parts = [];

  // model_id field
  parts.push(
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="model_id"\r\n\r\n' +
    sttModel
  );

  // diarize field
  parts.push(
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="diarize"\r\n\r\n' +
    'true'
  );

  // language_code field
  parts.push(
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="language_code"\r\n\r\n' +
    'en'
  );

  // Combine text parts
  const preFileData = Utilities.newBlob(parts.join('\r\n') + '\r\n').getBytes();

  // File part header
  const fileHeader = Utilities.newBlob(
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="file"; filename="' + file.getName() + '"\r\n' +
    'Content-Type: ' + file.getMimeType() + '\r\n\r\n'
  ).getBytes();

  // File content
  const fileBytes = blob.getBytes();

  // Closing boundary
  const closing = Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes();

  // Concatenate all byte arrays
  const payload = [].concat(preFileData, fileHeader, fileBytes, closing);

  const options = {
    method: 'POST',
    headers: {
      'xi-api-key': apiKey
    },
    contentType: contentType,
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(CONFIG.ENDPOINTS.ELEVENLABS_STT, options);
  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`ElevenLabs API error (${response.getResponseCode()}): ${JSON.stringify(result)}`);
  }

  return result;
}

/**
 * Submit a transcription via cloud storage URL (for files > 50MB)
 * @param {string} audioUrl - Publicly accessible URL to the audio file
 * @returns {Object} ElevenLabs API response
 */
function submitTranscriptionUrl(audioUrl) {
  const apiKey = getElevenLabsKey();
  const sttModel = getSetting('elevenlabs_model') || 'scribe_v2';

  const boundary = '----FormBoundary' + Utilities.getUuid();
  const contentType = 'multipart/form-data; boundary=' + boundary;

  const body =
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="model_id"\r\n\r\n' +
    sttModel + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="diarize"\r\n\r\n' +
    'true\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="language_code"\r\n\r\n' +
    'en\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="cloud_storage_url"\r\n\r\n' +
    audioUrl + '\r\n' +
    '--' + boundary + '--\r\n';

  const options = {
    method: 'POST',
    headers: {
      'xi-api-key': apiKey
    },
    contentType: contentType,
    payload: body,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(CONFIG.ENDPOINTS.ELEVENLABS_STT, options);
  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`ElevenLabs API error (${response.getResponseCode()}): ${JSON.stringify(result)}`);
  }

  return result;
}

// ============================================================================
// TRANSCRIPT FORMATTING
// ============================================================================

/**
 * Format ElevenLabs response into a readable transcript with speaker labels
 *
 * ElevenLabs returns a words[] array where each word has:
 *   { text, start, end, type, speaker_id }
 *
 * We group consecutive words by speaker and format as:
 *   [MM:SS] Speaker N: sentence text
 *
 * @param {Object} result - ElevenLabs API response
 * @returns {string} Formatted transcript
 */
function formatElevenLabsTranscript(result) {
  const words = result.words;
  if (!words || words.length === 0) {
    return result.text || '';
  }

  const segments = [];
  let currentSpeaker = null;
  let currentWords = [];
  let segmentStart = 0;

  for (const word of words) {
    // Skip non-word tokens (spacing, punctuation attached to words is fine)
    if (word.type === 'spacing') continue;

    const speaker = word.speaker_id;

    if (speaker !== currentSpeaker && currentWords.length > 0) {
      // Flush previous segment
      segments.push({
        speaker: currentSpeaker,
        start: segmentStart,
        text: currentWords.join(' ').replace(/\s+([.,!?;:])/g, '$1')
      });
      currentWords = [];
      segmentStart = word.start;
    }

    if (currentWords.length === 0) {
      segmentStart = word.start;
    }

    currentSpeaker = speaker;
    currentWords.push(word.text);
  }

  // Flush last segment
  if (currentWords.length > 0) {
    segments.push({
      speaker: currentSpeaker,
      start: segmentStart,
      text: currentWords.join(' ').replace(/\s+([.,!?;:])/g, '$1')
    });
  }

  // Format segments
  return segments.map(seg => {
    const totalSeconds = Math.floor(seg.start);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    const timestamp = `${minutes}:${seconds.toString().padStart(2, '0')}`;
    // Normalize: ElevenLabs returns "speaker_1" → we want "Speaker 1"
    const label = seg.speaker != null
      ? `Speaker ${String(seg.speaker).replace(/\D+/g, '')}`
      : 'Unknown';
    return `[${timestamp}] ${label}: ${seg.text}`;
  }).join('\n\n');
}

/**
 * Extract unique speaker labels from the formatted transcript
 * @param {string} transcript - Formatted transcript
 * @returns {string[]} Array of speaker labels like ["Speaker 0", "Speaker 1"]
 */
function extractSpeakerLabels(transcript) {
  const matches = transcript.match(/Speaker \d+/g) || [];
  return [...new Set(matches)];
}

// ============================================================================
// AUDIO FILE URL GENERATION (kept from AssemblyAI.gs)
// ============================================================================

/**
 * Generate a temporary public URL for a Drive file
 * @param {string} fileId - Google Drive file ID
 * @returns {string} Direct download URL
 */
function getDriveFileUrl(fileId) {
  const file = DriveApp.getFileById(fileId);

  // Store original sharing settings to restore later
  const originalAccess = file.getSharingAccess();
  const originalPermission = file.getSharingPermission();

  // Make file publicly accessible temporarily
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const url = `https://drive.google.com/uc?export=download&id=${fileId}`;

  // Store original settings in Script Properties for restoration
  const props = PropertiesService.getScriptProperties();
  props.setProperty(`file_permissions_${fileId}`, JSON.stringify({
    access: originalAccess.toString(),
    permission: originalPermission.toString()
  }));

  return url;
}

/**
 * Restore original file permissions after transcription
 * @param {string} fileId - Google Drive file ID
 */
function restoreFilePermissions(fileId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const savedPerms = props.getProperty(`file_permissions_${fileId}`);

    if (savedPerms) {
      const file = DriveApp.getFileById(fileId);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      props.deleteProperty(`file_permissions_${fileId}`);
      Logger.log(`Restored permissions for file ${fileId}`);
    }
  } catch (e) {
    Logger.log(`Warning: Could not restore permissions for ${fileId}: ${e.message}`);
  }
}

// ============================================================================
// HIGH-LEVEL TRANSCRIPTION
// ============================================================================

/**
 * Start and complete transcription for a discussion (synchronous).
 *
 * ElevenLabs Scribe is synchronous — one API call returns the full result.
 * Chooses blob upload vs URL upload based on file size.
 *
 * @param {string} discussionId
 * @returns {string} Formatted transcript
 */
function startTranscription(discussionId) {
  const discussion = getDiscussion(discussionId);
  if (!discussion) {
    throw new Error(`Discussion not found: ${discussionId}`);
  }

  const fileId = discussion.audio_file_id;
  const file = DriveApp.getFileById(fileId);
  const sizeMB = file.getSize() / (1024 * 1024);
  let usedUrl = false;
  let result;

  // Update status
  updateDiscussion(discussionId, {
    status: CONFIG.STATUS.TRANSCRIBING,
    next_step: 'Transcribing audio with ElevenLabs...'
  });

  try {
    if (sizeMB <= CONFIG.LIMITS.MAX_AUDIO_SIZE_MB) {
      Logger.log(`Transcribing ${file.getName()} via blob upload (${sizeMB.toFixed(1)} MB)`);
      result = submitTranscriptionBlob(fileId);
    } else {
      Logger.log(`Transcribing ${file.getName()} via URL upload (${sizeMB.toFixed(1)} MB > ${CONFIG.LIMITS.MAX_AUDIO_SIZE_MB} MB threshold)`);
      const audioUrl = getDriveFileUrl(fileId);
      usedUrl = true;
      result = submitTranscriptionUrl(audioUrl);
    }
  } catch (e) {
    // Restore permissions if we made the file public
    if (usedUrl) {
      restoreFilePermissions(fileId);
    }
    throw e;
  }

  // Restore permissions if URL approach was used
  if (usedUrl) {
    restoreFilePermissions(fileId);
  }

  // Format and store transcript
  const transcript = formatElevenLabsTranscript(result);

  upsertTranscript(discussionId, {
    raw_transcript: transcript
  });

  Logger.log(`Transcription completed for ${discussionId} (${transcript.length} chars)`);
  return transcript;
}
