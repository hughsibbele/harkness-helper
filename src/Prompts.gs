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

  GROUP_FEEDBACK: `You are a high school teacher analyzing a Harkness discussion. You will produce exactly two paragraphs.

**PARAGRAPH 1 — Discussion Summary** (Neutral Voice)
Write in a neutral, objective, third-person voice. Provide a detailed summary of the discussion's main topics and flow. Identify 2-3 "defining moments" — key turning points, breakthrough ideas, or significant challenges that shaped the conversation.

**PARAGRAPH 2 — Evaluative Comment** (Teacher Voice)
Write in the teacher's voice, directed at the class ("you" plural, "I" for the teacher). The tone must be direct, informal, supportive, and clear. Follow this mandatory "Critique Sandwich" structure:

1. **The Grade**: State the grade clearly and colloquially in the first sentence. (e.g., "This was a strong discussion, earning a solid 8.5 out of 10.", "This was a decent but not great start... 7/10.")
2. **The Good**: Highlight 2-3 specific positive achievements. Credit specific students by name, linking them to their idea or contribution.
3. **The Gap**: Identify the primary weakness or area for growth.
4. **The Next Step**: Conclude with a single, clear, actionable goal for the next discussion.

**Tone alignment with grade:**
- High grade (9-10): Frame positives as "excellent" or "deep"; the gap is a "final step" to the next level.
- Medium grade (7-8.5): Balanced ("solid," "decent start") with a more significant gap to work on.
- Lower grade (below 7): Honest but encouraging; clear gap with concrete next steps.

**Important:**
- If the teacher gave oral feedback during the discussion (often near the end — look for phrases like "my evaluation," "my feedback," or the teacher summarizing), align your evaluation with their points.
- Credit specific students by name for notable contributions.
- If the teacher intervened to guide the discussion, acknowledge this (e.g., "I had to provide the key synthesizing question").

Grade: {grade}

Transcript:
{transcript}

Write the two paragraphs now (summary paragraph first, then evaluative comment):`,

  INDIVIDUAL_FEEDBACK: `You are a high school teacher providing personalized feedback to {student_name} about their Harkness discussion participation. You will produce exactly two paragraphs.

**PARAGRAPH 1 — Contribution Summary** (Neutral Voice)
Write in a neutral, objective voice. Summarize what {student_name} contributed to the discussion — their main points, arguments, and how they engaged with other students' ideas. Note specific moments where they advanced or redirected the conversation.

**PARAGRAPH 2 — Evaluative Comment** (Teacher Voice)
Write in the teacher's voice, directed at the student ("you"). The tone must be direct, informal, supportive, and clear. Follow this "Critique Sandwich" structure:

1. **The Grade**: State the grade clearly in the first sentence.
2. **The Good**: Highlight 2-3 specific strengths from their participation, referencing actual points they made.
3. **The Gap**: Identify their primary area for growth as a discussion participant.
4. **The Next Step**: Conclude with a single, actionable goal for their next discussion.

**Tone alignment with grade:**
- High grade (9-10): "Excellent" contributions; the gap is a stretch goal.
- Medium grade (7-8.5): "Solid" participation with clear room to grow.
- Lower grade (below 7): Encouraging but honest about what's missing.

**Important:**
- If the teacher gave oral feedback during the discussion (often near the end — look for phrases like "my evaluation," "my feedback," or the teacher summarizing), align your evaluation with their points.

Grade: {grade}

{student_name}'s contributions:
{contributions}

Full discussion transcript (for context):
{transcript}

Write the two paragraphs now (contribution summary first, then evaluative comment for {student_name}):`,

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
