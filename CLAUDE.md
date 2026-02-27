# Harkness Helper

An automated Harkness discussion workflow system built as a Google Apps Script project. Teachers upload audio recordings, and the system handles transcription, speaker identification, AI feedback generation, and distribution via email and Canvas.

## Architecture

- **Google Apps Script** (V8 runtime) — no build system, no package manager
- **Google Sheets** as the database (8 sheets), config store, and prompt store
- **ElevenLabs Scribe v2** for synchronous transcription with speaker diarization
- **Gemini 2.0 Flash** for speaker identification and feedback generation
- **Canvas LMS API** (optional) for roster sync and grade posting; supports multi-course

## Files

| File | Role |
|------|------|
| `Code.gs` | Entry point: conditional menu, Setup Wizard, auto-off trigger system, Canvas config dialog, processing pipeline, feedback generation, send orchestration, multi-course menu items |
| `Config.gs` | `CONFIG` constant (including `COURSES` sheet), `isSetupComplete()`, Script Properties access, mode helpers, `validateConfiguration()` |
| `Sheets.gs` | Generic CRUD layer over Google Sheets, plus domain-specific helpers for all 8 sheets; includes Courses CRUD, `reorderColumns()`, multi-course helpers |
| `Prompts.gs` | Sheet-based prompt system with `DEFAULT_PROMPTS` fallback; `getPrompt(name, {vars})` |
| `ElevenLabs.gs` | Synchronous transcription via Scribe v2, blob vs URL upload based on file size |
| `Gemini.gs` | Gemini API calls (default: gemini-2.0-flash), speaker ID, group/individual feedback generation |
| `Canvas.gs` | Canvas LMS API: paginated requests, student sync, dual-mode grade posting, `stripHtml()`, `fetchCanvasCourseData()` |
| `DriveMonitor.gs` | Upload folder monitoring, filename parsing (section + course + date), folder setup, `moveToCompleted()` |
| `Email.gs` | Dual-mode HTML/plaintext email templates with HTML escaping, distribution, `sendTestEmail()` |
| `Webapp.gs` | API backend for the GitHub Pages recorder: `doGet(?action=getConfig)`, `doPost(?action=uploadAudio)` |
| `docs/index.html` | Mobile-friendly recording UI hosted on GitHub Pages (not in GAS); timer, pause/resume, section/course selection, base64 upload via `fetch()` |

## Data Flow

1. Teacher uploads audio to `Harkness Helper / Upload` folder in Drive
2. `mainProcessingLoop()` (10-min trigger) detects file, moves to Processing folder
3. ElevenLabs transcribes synchronously with speaker diarization
4. Gemini identifies speakers from introductions, populates SpeakerMap sheet
5. **Group mode**: auto-advances to review; **Individual mode**: waits for teacher to confirm speakers
6. Teacher enters grade(s), clicks "Generate Feedback" → Gemini generates
7. Teacher reviews, approves, clicks "Send" → email and/or Canvas distribution

## Key Design Patterns

### Setup Wizard (self-configuring template)

- `isSetupComplete()` checks 5 required Script Properties: `SPREADSHEET_ID`, `ELEVENLABS_API_KEY`, `GEMINI_API_KEY`, `AUDIO_FOLDER_ID`, `PROCESSING_FOLDER_ID`
- `onOpen()` shows minimal menu (just "Setup Wizard") until setup is complete, then shows full menu
- `runSetupWizard()` auto-captures spreadsheet ID via `SpreadsheetApp.getActiveSpreadsheet().getId()`, stores API keys, creates Drive folders, initializes sheets/settings/prompts
- No `setupScriptProperties()`, `initialSetup()`, or `testConfiguration()` — the wizard replaces all of these

### Auto-Off Trigger System

- `menuStartProcessing()` removes existing triggers, records `PROCESSING_START_TIME`, installs 10-min trigger, runs loop immediately
- `mainProcessingLoop()` checks elapsed time at the top — auto-removes trigger and returns if ≥ `PROCESSING_TIMEOUT_MINUTES` (60 min)
- `removeProcessingTriggers()` only removes `mainProcessingLoop` triggers (not all project triggers)
- No always-on triggers — processing is teacher-initiated and self-terminating

### Canvas Config Dialog

- `showCanvasConfigDialog()` uses 3-step `ui.prompt()` flow: base URL → API token → course ID
- Automatically syncs all rosters (all sections) as the final step
- Stores token/URL in PropertiesService, course ID + `distribute_canvas=true` in Settings sheet
- Idempotent — pre-fills existing values, safe to re-run

### Conditional Menu

Before setup:
```
Harkness Helper → Setup Wizard (start here)
```

After setup:
```
Harkness Helper
  Start Processing
  Stop Processing
  Process New Files Now
  ───
  Generate Feedback
  Send Approved Feedback
  ───
  Configure Canvas Course
  Sync Canvas Roster
  Sync All Course Rosters (multi-course only)
  Fetch Canvas Course Data
  ───
  Enable Multi-Course (multi-course only)
  Check Configuration
  Reorder / Format Sheets
  Re-run Setup Wizard
```

## Discussion Status Machine

Defined in `Config.gs` as `CONFIG.STATUS`:
```
uploaded → transcribing → mapping → review → approved → sent
                                                        ↘ error
```

## Two Modes

Controlled by `mode` setting in the Settings sheet:
- **Group mode** (`group`): One grade + one feedback for the whole class. Stored on the Discussion row.
- **Individual mode** (`individual`): Per-student grades + feedback. Stored in StudentReports rows.

## Multi-Course Support

Optional. When the **Courses** sheet has entries:
- Each course has its own Canvas course ID and base URL
- Students can be tagged with a `course` column
- Audio filenames can include a course prefix (e.g., `AP Lit - Section 2 - 2025-01-15.m4a`)
- `menuSyncAllCourseRosters()` syncs students from all configured courses
- Discussion rows have a `course` column for routing

## Mobile Recorder (Optional)

- Frontend (`docs/index.html`) hosted on GitHub Pages — not in GAS (Apps Script iframes block microphone access)
- Backend (`Webapp.gs`) deployed as a GAS web app, acts as API only (no HTML serving)
- Frontend communicates via `fetch()` to `?action=getConfig` (GET) and `?action=uploadAudio` (POST)
- `Content-Type: text/plain` used intentionally to avoid CORS preflight requests
- Provides in-browser audio recording with section/course/date selection
- Multi-course aware — shows course selector when `isMultiCourseMode()` is true

## Google Sheets Structure (8 sheets)

| Sheet | Purpose |
|-------|---------|
| Settings | Key-value config (mode, distribution flags, grade scale, teacher info) |
| Discussions | One row per discussion session with status, grade, group_feedback, course |
| Students | Student roster with email, section, course, and Canvas IDs |
| Transcripts | Raw/named transcripts, speaker map JSON |
| SpeakerMap | One row per speaker per discussion for teacher review |
| StudentReports | Individual reports with contributions, grades, feedback (individual mode) |
| Prompts | Teacher-editable prompt templates read at runtime |
| Courses | Multi-course Canvas configuration (course_name, canvas_course_id, canvas_base_url, canvas_item_type) |

## Development Notes

- This is NOT a Node.js project. No `npm`, `build`, `lint`, or `test` commands
- Deploy `.gs` files by copying into the Apps Script editor; no HTML files go in GAS
- The recorder frontend (`docs/index.html`) is deployed via GitHub Pages, separate from GAS
- Debug via `Logger.log()` → View > Executions in the script editor
- Secrets stored in `PropertiesService.getScriptProperties()` (set by Setup Wizard)
- `getSpreadsheetId()` reads from Script Properties (auto-set by wizard), not hardcoded
- All sheet initialization is idempotent (`getOrCreateSheet` checks before creating)
- GAS 6-minute execution limit; `CONFIG.LIMITS.GAS_TIMEOUT_MS` = 5 minutes safety margin
- Default Gemini model: `gemini-2.0-flash` (configurable via Settings sheet)

## Style Conventions

- JSDoc comments on functions
- `CONFIG` constant for sheet names, status values, modes, limits
- `getSetting(key)` / `setSetting(key, value)` for Settings sheet
- `getRequiredProperty(key)` / `setProperty(key, value)` for Script Properties
- Error messages append to `Discussions.error_message` with timestamps
- HTML escaping via `escapeHtmlForEmail()` for all dynamic values in email templates
