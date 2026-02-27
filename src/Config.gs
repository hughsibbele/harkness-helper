/**
 * Configuration management for Harkness Helper
 *
 * API keys and sensitive configuration are stored in Script Properties
 * (set automatically by the Setup Wizard).
 * Runtime settings (mode, distribution flags) live in the Settings sheet.
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const CONFIG = {
  // Sheet names in the spreadsheet
  SHEETS: {
    SETTINGS: 'Settings',
    DISCUSSIONS: 'Discussions',
    STUDENTS: 'Students',
    TRANSCRIPTS: 'Transcripts',
    SPEAKER_MAP: 'SpeakerMap',
    STUDENT_REPORTS: 'StudentReports',
    PROMPTS: 'Prompts',
    COURSES: 'Courses'
  },

  // Discussion status values
  STATUS: {
    UPLOADED: 'uploaded',
    TRANSCRIBING: 'transcribing',
    MAPPING: 'mapping',
    REVIEW: 'review',
    APPROVED: 'approved',
    SENT: 'sent',
    ERROR: 'error'
  },

  // Discussion modes
  MODES: {
    GROUP: 'group',
    INDIVIDUAL: 'individual'
  },

  // Trigger intervals (in minutes)
  TRIGGERS: {
    MAIN_LOOP: 10
  },

  // API endpoints
  ENDPOINTS: {
    ELEVENLABS_STT: 'https://api.elevenlabs.io/v1/speech-to-text',
    GEMINI: 'https://generativelanguage.googleapis.com/v1beta'
  },

  // Processing limits
  LIMITS: {
    MAX_AUDIO_SIZE_MB: 50,           // Blob vs URL threshold
    GAS_TIMEOUT_MS: 300000,          // 5 minutes (GAS limit is 6)
    TRANSCRIPTION_TIMEOUT_MS: 600000 // 10 minutes before marking stuck
  }
};

// ============================================================================
// SCRIPT PROPERTIES HELPERS
// ============================================================================

/**
 * Get a required configuration value from Script Properties
 * @param {string} key - The property key
 * @returns {string} The property value
 * @throws {Error} If the property is not set
 */
function getRequiredProperty(key) {
  const props = PropertiesService.getScriptProperties();
  const value = props.getProperty(key);
  if (!value) {
    throw new Error(`Required configuration '${key}' is not set. Please run the Setup Wizard from the Harkness Helper menu.`);
  }
  return value;
}

/**
 * Get an optional configuration value from Script Properties
 * @param {string} key - The property key
 * @param {string} defaultValue - Default if not set
 * @returns {string} The property value or default
 */
function getOptionalProperty(key, defaultValue = '') {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(key) || defaultValue;
}

/**
 * Set a configuration value in Script Properties
 * @param {string} key - The property key
 * @param {string} value - The value to set
 */
function setProperty(key, value) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(key, value);
}

// ============================================================================
// CONFIGURATION GETTERS
// ============================================================================

/**
 * Get ElevenLabs API key
 */
function getElevenLabsKey() {
  return getRequiredProperty('ELEVENLABS_API_KEY');
}

/**
 * Get Gemini API key
 */
function getGeminiKey() {
  return getRequiredProperty('GEMINI_API_KEY');
}

/**
 * Get Canvas API token
 */
function getCanvasToken() {
  return getRequiredProperty('CANVAS_API_TOKEN');
}

/**
 * Get Canvas base URL (checks Settings sheet first, then Script Properties)
 */
function getCanvasBaseUrl() {
  const fromSettings = getSetting('canvas_base_url');
  const url = fromSettings || getOptionalProperty('CANVAS_BASE_URL', '');
  if (!url) {
    throw new Error('Canvas base URL not set. Add it in the Settings sheet or Script Properties.');
  }
  return url.replace(/\/$/, '');
}

/**
 * Get main spreadsheet ID
 */
function getSpreadsheetId() {
  return getRequiredProperty('SPREADSHEET_ID');
}

/**
 * Get audio upload folder ID
 */
function getAudioFolderId() {
  return getRequiredProperty('AUDIO_FOLDER_ID');
}

/**
 * Get processing folder ID
 */
function getProcessingFolderId() {
  return getRequiredProperty('PROCESSING_FOLDER_ID');
}

// ============================================================================
// MODE HELPERS (read from Settings sheet at runtime)
// ============================================================================

/**
 * Get current discussion mode from Settings sheet
 * @returns {string} 'group' or 'individual'
 */
function getMode() {
  const mode = getSetting('mode');
  return mode === CONFIG.MODES.INDIVIDUAL ? CONFIG.MODES.INDIVIDUAL : CONFIG.MODES.GROUP;
}

/**
 * @returns {boolean} true if group mode
 */
function isGroupMode() {
  return getMode() === CONFIG.MODES.GROUP;
}

/**
 * @returns {boolean} true if individual mode
 */
function isIndividualMode() {
  return getMode() === CONFIG.MODES.INDIVIDUAL;
}

// ============================================================================
// SETUP FUNCTIONS
// ============================================================================

/**
 * Check if the Setup Wizard has been completed.
 * Returns true only when all 5 required Script Properties are set.
 * @returns {boolean}
 */
function isSetupComplete() {
  const required = [
    'SPREADSHEET_ID',
    'ELEVENLABS_API_KEY',
    'GEMINI_API_KEY',
    'AUDIO_FOLDER_ID',
    'PROCESSING_FOLDER_ID'
  ];
  const props = PropertiesService.getScriptProperties().getProperties();
  return required.every(key => !!props[key]);
}

/**
 * Verify all required configuration is present
 * @returns {Object} Validation result with {valid: boolean, missing: string[]}
 */
function validateConfiguration() {
  const required = [
    'ELEVENLABS_API_KEY',
    'GEMINI_API_KEY',
    'SPREADSHEET_ID',
    'AUDIO_FOLDER_ID',
    'PROCESSING_FOLDER_ID'
  ];

  const props = PropertiesService.getScriptProperties().getProperties();
  const missing = required.filter(key => !props[key]);

  return {
    valid: missing.length === 0,
    missing: missing
  };
}

/**
 * Check if Canvas integration is configured
 * @returns {boolean}
 */
function isCanvasConfigured() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return !!(props['CANVAS_API_TOKEN'] && props['CANVAS_BASE_URL']);
}
