/**
 * LLM Prompts for Harkness Helper
 *
 * Prompts live in the Prompts sheet so teachers can edit them.
 * DEFAULT_PROMPTS provides seed values that initializeDefaultPrompts()
 * writes into the sheet on first run (won't overwrite existing edits).
 *
 * Use getPrompt(name, {vars}) everywhere — it reads from the sheet
 * and fills {placeholder} variables.
 */

// ============================================================================
// DEFAULT PROMPT TEMPLATES
// ============================================================================

const DEFAULT_PROMPTS = {

  SPEAKER_IDENTIFICATION: `You are analyzing the beginning of a classroom Harkness discussion recording.

Students typically introduce themselves at the start. Listen for patterns like:
- "Hi, I'm [name]"
- "My name is [name]"
- "This is [name]"
- "[Name] here"
- Other natural introductions

Analyze this transcript excerpt and identify which speaker label corresponds to which student name.

IMPORTANT RULES:
1. Only include speakers you can confidently identify from explicit introductions
2. If a speaker cannot be identified, map them to "?" instead
3. Be case-sensitive with names as they appear
4. The teacher may also speak — if identified, include them as "Teacher"

Return ONLY a valid JSON object mapping speaker labels to names.
Example format: {"Speaker 0": "Maria", "Speaker 1": "James", "Speaker 2": "Teacher", "Speaker 3": "?"}

Transcript excerpt:
{transcript}

JSON mapping:`,

  DISCUSSION_SUMMARY: `Summarize this Harkness discussion for the teacher's records.

Include:
1. **Main themes/topics** discussed (2-3 bullet points)
2. **Key insights** that emerged from the discussion
3. **Discussion quality** assessment (engagement level, depth of analysis)
4. **Notable moments** (breakthrough insights, good collaboration, etc.)

Keep the summary to about 150-200 words.

Transcript:
{transcript}

Discussion Summary:`,

  GROUP_FEEDBACK: `Generate feedback for the whole class based on their Harkness discussion.

The teacher gave this discussion a grade of {grade}.

Teacher's notes/feedback:
{teacher_feedback}

Discussion transcript:
{transcript}

Write a supportive but honest group feedback paragraph (4-6 sentences) that:
1. Summarizes the overall quality of the discussion
2. Highlights specific strong moments or contributions
3. Identifies areas where the class can improve their discussion skills
4. Connects to Harkness discussion values (collaboration, evidence-based reasoning, building on others' ideas)
5. Ends on an encouraging note

Group feedback:`,

  INDIVIDUAL_FEEDBACK: `Generate personalized feedback for {student_name} based on their Harkness discussion participation.

The teacher gave this student a grade of {grade}.

Teacher's notes/feedback:
{teacher_feedback}

{student_name}'s contributions from the transcript:
{contributions}

Full discussion transcript (for context):
{transcript}

Write a supportive but honest feedback paragraph (3-5 sentences) that:
1. Acknowledges specific strengths from their participation
2. References actual points they made when notable
3. Provides one concrete suggestion for growth
4. Ends on an encouraging note

Feedback for {student_name}:`,

  TEACHER_FEEDBACK_EXTRACTION: `You are analyzing a Harkness discussion transcript.

At the end of many discussions, the teacher provides verbal feedback to the class. This typically includes:
- Overall assessment of the discussion quality
- Highlights of good contributions
- Areas for improvement
- Specific student callouts (positive or constructive)

Extract ONLY the teacher's feedback section from this transcript.
This is usually near the end, after the main discussion concludes.

If there is no clear teacher feedback section, return "NO_FEEDBACK_FOUND".

Return ONLY the teacher feedback text, nothing else.

Full transcript:
{transcript}

Teacher feedback:`

};

// ============================================================================
// PROMPT HELPERS
// ============================================================================

/**
 * Get a prompt with placeholders filled in.
 * Reads from the Prompts sheet first, falls back to DEFAULT_PROMPTS.
 * @param {string} promptName - Name of the prompt (e.g. 'GROUP_FEEDBACK')
 * @param {Object} variables - Object with placeholder values
 * @returns {string} Filled prompt
 */
function getPrompt(promptName, variables = {}) {
  // Try sheet first
  let prompt = getPromptFromSheet(promptName);

  // Fall back to defaults
  if (!prompt) {
    prompt = DEFAULT_PROMPTS[promptName];
  }

  if (!prompt) {
    throw new Error(`Unknown prompt: ${promptName}`);
  }

  for (const [key, value] of Object.entries(variables)) {
    prompt = prompt.replace(new RegExp(`\\{${key}\\}`, 'g'), value || '');
  }

  return prompt;
}

/**
 * Seed the Prompts sheet with default prompts (only adds missing ones)
 */
function initializeDefaultPrompts() {
  for (const [name, text] of Object.entries(DEFAULT_PROMPTS)) {
    const existing = getPromptFromSheet(name);
    if (!existing) {
      upsertPrompt(name, text);
    }
  }

  Logger.log('Default prompts initialized.');
}
