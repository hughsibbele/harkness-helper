/**
 * Google Drive folder monitoring for audio uploads
 *
 * Watches a designated folder for new audio files and
 * initiates the processing pipeline.
 */

// ============================================================================
// FOLDER MONITORING
// ============================================================================

/**
 * Check for new audio files in the upload folder
 * @returns {Object[]} Array of new file info objects
 */
function checkForNewAudioFiles() {
  const uploadFolderId = getAudioFolderId();
  const processingFolderId = getProcessingFolderId();

  const uploadFolder = DriveApp.getFolderById(uploadFolderId);
  const processingFolder = DriveApp.getFolderById(processingFolderId);

  const newFiles = [];

  // Supported audio formats
  const audioMimeTypes = [
    'audio/mpeg',           // .mp3
    'audio/mp4',            // .m4a
    'audio/x-m4a',          // .m4a (alternate)
    'audio/wav',            // .wav
    'audio/x-wav',          // .wav (alternate)
    'audio/ogg',            // .ogg
    'audio/webm',           // .webm
    'audio/aac',            // .aac
    'audio/flac',           // .flac
    'video/mp4',            // Some recorders save as .mp4
    'video/quicktime'       // .mov (iPhone videos with audio)
  ];

  // Get all files in upload folder
  const files = uploadFolder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    const fileName = file.getName();

    // Check if it's an audio file
    const isAudio = audioMimeTypes.includes(mimeType) ||
      fileName.match(/\.(mp3|m4a|wav|ogg|webm|aac|flac|mp4|mov)$/i);

    if (!isAudio) {
      Logger.log(`Skipping non-audio file: ${fileName} (${mimeType})`);
      continue;
    }

    // Log file size (files > 50MB will use URL-based upload instead of blob)
    const sizeMB = file.getSize() / (1024 * 1024);
    if (sizeMB > CONFIG.LIMITS.MAX_AUDIO_SIZE_MB) {
      Logger.log(`Large file detected: ${fileName} (${sizeMB.toFixed(1)} MB) — will use URL-based upload`);
    }

    newFiles.push({
      id: file.getId(),
      name: fileName,
      mimeType: mimeType,
      size: file.getSize(),
      created: file.getDateCreated()
    });

    // Move file to processing folder
    file.moveTo(processingFolder);
    Logger.log(`Found and moved new audio file: ${fileName}`);
  }

  return newFiles;
}

/**
 * Get distinct section names from the Students sheet.
 * Optionally filtered by course for multi-course mode.
 * @param {string} [course] - Course name to filter by
 * @returns {string[]} Unique section names, sorted longest-first
 */
function getKnownSections(course) {
  const rows = getAllRows(CONFIG.SHEETS.STUDENTS);
  const sections = new Set();
  for (const row of rows) {
    if (row.section) {
      if (course && row.course && row.course !== course) continue;
      sections.add(row.section);
    }
  }
  // Sort longest-first so exact substring matching prefers longer matches
  return Array.from(sections).sort((a, b) => b.length - a.length);
}

/**
 * Parse section, date, and course from filename.
 *
 * Section matching strategy (in order):
 * 1. Exact substring match against known sections from Students sheet
 * 2. Token-based fuzzy match against known sections
 * 3. Regex fallback (Section N, Block A, etc.) for empty rosters
 *
 * @param {string} fileName
 * @returns {Object} {section, date, course}
 */
function parseFileName(fileName) {
  let section = '';
  let date = new Date().toISOString().split('T')[0];  // Default to today
  let course = '';

  // Strip file extension before parsing
  fileName = fileName.replace(/\.\w+$/, '');

  // In multi-course mode, check if filename starts with a known course name
  if (isMultiCourseMode()) {
    const courses = getAllCourses();
    for (const c of courses) {
      const name = c.course_name;
      if (name && fileName.startsWith(name)) {
        course = name;
        // Strip the course prefix and separator for further parsing
        fileName = fileName.substring(name.length).replace(/^\s*[-–—_]\s*/, '');
        break;
      }
    }
  }

  // Try to extract date first (before section matching, since date is independent)
  // Format: YYYY-MM-DD or YYYYMMDD
  const dateMatch = fileName.match(/(\d{4})[-_]?(\d{2})[-_]?(\d{2})/);
  if (dateMatch) {
    date = `${dateMatch[1]}-${dateMatch[2]}-${dateMatch[3]}`;
  }

  // --- Section matching ---
  const knownSections = getKnownSections(course);

  if (knownSections.length > 0) {
    // Strategy 1: Exact substring match (longest match wins, already sorted)
    const fileNameLower = fileName.toLowerCase();
    for (const s of knownSections) {
      if (fileNameLower.includes(s.toLowerCase())) {
        section = s;
        break;
      }
    }

    // Strategy 2: Token-based fuzzy match
    if (!section) {
      // Tokenize filename: split on non-alphanumeric characters, drop date tokens and empties
      const fileTokens = fileName.split(/[^a-zA-Z0-9]+/).filter(t => t && !t.match(/^\d{4,}$/));

      let bestSection = '';
      let bestScore = 0;
      let bestRatio = 0;

      for (const s of knownSections) {
        const sectionTokens = s.split(/[^a-zA-Z0-9]+/).filter(t => t);
        let matchCount = 0;

        for (const ft of fileTokens) {
          const ftLower = ft.toLowerCase();
          for (const st of sectionTokens) {
            const stLower = st.toLowerCase();
            // A token matches if either contains the other (case-insensitive)
            if (stLower.includes(ftLower) || ftLower.includes(stLower)) {
              matchCount++;
              break; // Count each file token at most once
            }
          }
        }

        if (matchCount > 0) {
          // Ratio: what fraction of file tokens matched
          const ratio = matchCount / fileTokens.length;
          if (matchCount > bestScore || (matchCount === bestScore && ratio > bestRatio)) {
            bestScore = matchCount;
            bestRatio = ratio;
            bestSection = s;
          }
        }
      }

      if (bestSection) {
        section = bestSection;
      }
    }
  }

  // Strategy 3: Regex fallback for when roster is empty
  if (!section) {
    const sectionMatch = fileName.match(/[SsPp](?:ection\s*|eriod\s*)?(\d+)/);
    if (sectionMatch) {
      section = `Section ${sectionMatch[1]}`;
    } else {
      const blockMatch = fileName.match(/[Bb]lock\s*([A-Ga-g])/) ||
                         fileName.match(/([A-Ga-g])\s*[Bb]lock/) ||
                         fileName.match(/(?:^|\s|[-–—_])([A-Ga-g])(?:\s|[-–—_.]|$)/);
      if (blockMatch) {
        section = `Block ${blockMatch[1].toUpperCase()}`;
      }
    }
  }

  return { section, date, course };
}

/**
 * Process a newly detected audio file
 * @param {Object} fileInfo - File info from checkForNewAudioFiles
 * @returns {string} Discussion ID
 */
function processNewAudioFile(fileInfo) {
  const parsed = parseFileName(fileInfo.name);

  // Create discussion record
  const discussionData = {
    date: parsed.date,
    section: parsed.section,
    audio_file_id: fileInfo.id
  };
  if (parsed.course) {
    discussionData.course = parsed.course;
  }
  const discussionId = createDiscussion(discussionData);

  Logger.log(`Created discussion ${discussionId} for file ${fileInfo.name}`);

  return discussionId;
}

// ============================================================================
// FOLDER SETUP
// ============================================================================

/**
 * Create the required Drive folders if they don't exist
 * @returns {Object} {uploadFolderId, processingFolderId}
 */
function setupDriveFolders() {
  const rootFolder = DriveApp.getRootFolder();

  // Check for existing Harkness Helper folder
  let mainFolder;
  const mainFolderIterator = rootFolder.getFoldersByName('Harkness Helper');

  if (mainFolderIterator.hasNext()) {
    mainFolder = mainFolderIterator.next();
  } else {
    mainFolder = rootFolder.createFolder('Harkness Helper');
    Logger.log('Created "Harkness Helper" folder');
  }

  // Create subfolders
  let uploadFolder, processingFolder;

  const uploadIterator = mainFolder.getFoldersByName('Upload');
  if (uploadIterator.hasNext()) {
    uploadFolder = uploadIterator.next();
  } else {
    uploadFolder = mainFolder.createFolder('Upload');
    Logger.log('Created "Upload" subfolder');
  }

  const processingIterator = mainFolder.getFoldersByName('Processing');
  if (processingIterator.hasNext()) {
    processingFolder = processingIterator.next();
  } else {
    processingFolder = mainFolder.createFolder('Processing');
    Logger.log('Created "Processing" subfolder');
  }

  // Store folder IDs in Script Properties
  setProperty('AUDIO_FOLDER_ID', uploadFolder.getId());
  setProperty('PROCESSING_FOLDER_ID', processingFolder.getId());

  Logger.log(`Upload folder ID: ${uploadFolder.getId()}`);
  Logger.log(`Processing folder ID: ${processingFolder.getId()}`);

  return {
    uploadFolderId: uploadFolder.getId(),
    processingFolderId: processingFolder.getId(),
    mainFolderId: mainFolder.getId()
  };
}

/**
 * Get the URL for the upload folder (for sharing with teacher)
 * @returns {string}
 */
function getUploadFolderUrl() {
  const folderId = getAudioFolderId();
  return `https://drive.google.com/drive/folders/${folderId}`;
}

// ============================================================================
// COMPLETED FILE MANAGEMENT
// ============================================================================

/**
 * Move a processed file to a "Completed" folder
 * @param {string} fileId
 * @param {string} discussionId - For organizing
 */
function moveToCompleted(fileId, discussionId) {
  try {
    const processingFolder = DriveApp.getFolderById(getProcessingFolderId());
    const file = DriveApp.getFileById(fileId);

    // Get or create Completed folder
    let completedFolder;
    const completedIterator = processingFolder.getParents().next().getFoldersByName('Completed');

    if (completedIterator.hasNext()) {
      completedFolder = completedIterator.next();
    } else {
      completedFolder = processingFolder.getParents().next().createFolder('Completed');
    }

    file.moveTo(completedFolder);
    Logger.log(`Moved ${file.getName()} to Completed folder`);
  } catch (e) {
    Logger.log(`Error moving file to completed: ${e.message}`);
  }
}
