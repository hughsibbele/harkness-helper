/**
 * Google Gemini API integration for LLM-powered analysis
 *
 * Uses Gemini 2.0 Flash for cost efficiency. Supports both group and
 * individual feedback modes. All prompts read from the Prompts sheet.
 */

// ============================================================================
// GEMINI API CORE
// ============================================================================

/**
 * Call Gemini API with a prompt
 * @param {string} prompt - The prompt to send
 * @param {Object} options - Optional settings
 * @returns {string} The generated text response
 */
function callGemini(prompt, options = {}) {
  const apiKey = getGeminiKey();
  const model = options.model || getSetting('gemini_model') || 'gemini-2.0-flash';

  const url = `${CONFIG.ENDPOINTS.GEMINI}/models/${model}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: options.temperature ?? 0.3,
      maxOutputTokens: options.maxTokens || 2048,
      topP: options.topP || 0.8,
      ...(options.responseMimeType && { responseMimeType: options.responseMimeType })
    }
  };

  const requestOptions = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, requestOptions);
  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    const error = result.error?.message || 'Unknown Gemini API error';
    throw new Error(`Gemini API error: ${error}`);
  }

  const text = result.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    throw new Error('No text in Gemini response');
  }

  return text.trim();
}

/**
 * Call Gemini and parse JSON response
 * @param {string} prompt - The prompt (should request JSON output)
 * @param {Object} options - Optional settings
 * @returns {Object} Parsed JSON object
 */
function callGeminiJSON(prompt, options = {}) {
  let fullPrompt = prompt;
  if (!prompt.toLowerCase().includes('json')) {
    fullPrompt += '\n\nRespond with valid JSON only.';
  }

  const response = callGemini(fullPrompt, {
    ...options,
    temperature: options.temperature ?? 0.1,
    responseMimeType: 'application/json'
  });

  // Extract JSON from response (handle markdown code blocks)
  // Strip code fences first — regex match can fail on certain Gemini formats
  let jsonStr = response.trim()
    .replace(/^```(?:json)?\s*\n?/g, '')
    .replace(/\n?```\s*$/g, '')
    .trim();

  // Extract the JSON object or array
  const objectMatch = jsonStr.match(/\{[\s\S]*\}/);
  const arrayMatch = jsonStr.match(/\[[\s\S]*\]/);

  if (objectMatch) {
    jsonStr = objectMatch[0];
  } else if (arrayMatch) {
    jsonStr = arrayMatch[0];
  }

  try {
    return JSON.parse(jsonStr);
  } catch (e) {
    // Fix common Gemini JSON issues: smart quotes, single quotes, trailing commas,
    // unescaped newlines inside string values, control characters
    let fixed = jsonStr
      .replace(/[\u201C\u201D\u201E]/g, '"')
      .replace(/[\u2018\u2019\u201A]/g, "'")
      .replace(/'/g, '"')
      .replace(/,\s*([}\]])/g, '$1');

    // Fix unescaped newlines/tabs inside JSON string values
    fixed = fixed.replace(/"([^"]*?)"/g, function(match, content) {
      return '"' + content
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '\\r')
        .replace(/\t/g, '\\t') + '"';
    });

    // Remove trailing commas that might appear after fixing
    fixed = fixed.replace(/,\s*([}\]])/g, '$1');

    try {
      return JSON.parse(fixed);
    } catch (e2) {
      // Last resort: try to extract key-value pairs manually for simple objects
      Logger.log(`Failed to parse Gemini JSON response: ${response}`);
      const pairPattern = /"([^"]+)"\s*:\s*"([^"]*?)"/g;
      let match;
      const result = {};
      let found = false;
      while ((match = pairPattern.exec(response)) !== null) {
        result[match[1]] = match[2];
        found = true;
      }
      if (found) {
        Logger.log(`Recovered JSON via regex extraction: ${JSON.stringify(result)}`);
        return result;
      }
      throw new Error(`Invalid JSON from Gemini: ${e.message}`);
    }
  }
}

// ============================================================================
// SPEAKER IDENTIFICATION
// ============================================================================

/**
 * Identify speakers from transcript introductions
 * @param {string} transcriptExcerpt - First few minutes of transcript
 * @returns {Object} Map of speaker labels to names (e.g. {"Speaker 0": "Maria"})
 */
function identifySpeakers(transcriptExcerpt) {
  const prompt = getPrompt('SPEAKER_IDENTIFICATION', {
    transcript: transcriptExcerpt
  });

  const speakerMap = callGeminiJSON(prompt);

  Logger.log(`Identified speakers: ${JSON.stringify(speakerMap)}`);
  return speakerMap;
}

/**
 * Apply speaker names to a transcript
 * @param {string} transcript - Raw transcript with speaker labels
 * @param {Object} speakerMap - Map of labels to names
 * @returns {string} Transcript with real names
 */
function applySpeakerNames(transcript, speakerMap) {
  let namedTranscript = transcript;

  for (const [label, name] of Object.entries(speakerMap)) {
    if (!name || name === '?') continue;

    // Replace "Speaker N:" with "Name:" in the formatted lines
    const escapedLabel = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(escapedLabel + ':', 'g');
    namedTranscript = namedTranscript.replace(pattern, `${name}:`);
  }

  return namedTranscript;
}

/**
 * Get list of unique student names from speaker map
 * @param {Object} speakerMap
 * @returns {string[]} Array of student names (excluding "Teacher" and "?")
 */
function getStudentNames(speakerMap) {
  return Object.values(speakerMap)
    .filter(name => name && name !== '?' && name.toLowerCase() !== 'teacher')
    .filter((name, index, arr) => arr.indexOf(name) === index);
}

// ============================================================================
// GROUP MODE FEEDBACK
// ============================================================================

/**
 * Generate group feedback for the whole class
 * @param {string} transcript - Named transcript
 * @param {string} grade - Teacher's grade for the discussion
 * @returns {string} Group feedback paragraph
 */
function generateGroupFeedback(transcript, grade) {
  const prompt = getPrompt('GROUP_FEEDBACK', {
    transcript: transcript,
    grade: grade || 'not yet assigned'
  });

  return callGemini(prompt, { temperature: 0.5, maxTokens: 4096 });
}

// ============================================================================
// INDIVIDUAL MODE FEEDBACK
// ============================================================================

/**
 * Generate personalized feedback for a single student
 * @param {string} studentName
 * @param {string} contributions - This student's extracted lines
 * @param {string} transcript - Full named transcript (for context)
 * @param {string} grade - Teacher's grade for this student
 * @returns {string} Personalized feedback paragraph
 */
function generateIndividualFeedback(studentName, contributions, transcript, grade) {
  const prompt = getPrompt('INDIVIDUAL_FEEDBACK', {
    student_name: studentName,
    contributions: contributions || 'No specific contributions found in transcript.',
    transcript: transcript,
    grade: grade || 'not yet assigned'
  });

  return callGemini(prompt, { temperature: 0.5, maxTokens: 4096 });
}

/**
 * Extract a specific student's contributions from a named transcript.
 * Finds all lines where the student is the speaker.
 *
 * @param {string} namedTranscript - Transcript with real names applied
 * @param {string} studentName - The student's name to filter for
 * @returns {string} The student's lines concatenated
 */
function extractStudentContributions(namedTranscript, studentName) {
  const lines = namedTranscript.split('\n');
  const contributions = [];

  for (const line of lines) {
    // Match lines like "[0:42] Maria: some text" or "Maria: some text"
    if (line.includes(studentName + ':')) {
      contributions.push(line.trim());
    }
  }

  return contributions.join('\n');
}
