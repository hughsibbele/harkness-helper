# Harkness Discussion Helper

A Google Apps Script application that automates the workflow for Harkness discussions: recording transcription with speaker identification, AI-powered feedback generation, Canvas integration for grade posting, and email distribution to students.

## Features

- **Automatic Transcription**: Upload audio recordings to Google Drive and get them automatically transcribed with speaker diarization (ElevenLabs Scribe v2)
- **Speaker Identification**: Gemini AI identifies students from their introductions ("Hi, I'm...")
- **Two Grading Modes**: Group mode (one grade for the whole class) or Individual mode (per-student grades and feedback)
- **AI Feedback Generation**: Gemini generates 2-paragraph feedback using a critique sandwich format, editable by the teacher before sending
- **Canvas Integration**: Sync student roster, post grades and feedback (supports assignments and discussion topics, multi-section)
- **Email Distribution**: Send personalized HTML reports to students via Gmail
- **Teacher Review**: Review and approve everything in Google Sheets before sending
- **Editable Prompts**: All AI prompts live in a Prompts sheet — teachers can customize tone, criteria, and style without touching code

## Architecture

```
Phone Recording → Google Drive (Upload folder)
    → ElevenLabs Scribe v2 (synchronous transcription + speaker diarization)
    → Google Gemini (speaker ID from introductions)
    → Teacher reviews speaker map in SpeakerMap sheet
    → MODE BRANCH:
        Group:      Teacher enters ONE grade → Gemini group feedback → approve → send
        Individual: Teacher enters per-student grades → Gemini per-student feedback → approve → send
    → Gmail (report emails) + Canvas API (grade posting)
```

### Status Machine

```
uploaded → transcribing → mapping → review → approved → sent
                                                        ↘ error (from any step)
```

## Setup Instructions

### 1. Create the Google Apps Script Project

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Delete the default `Code.gs` content
3. Create files for each `.gs` file in the `src/` folder:
   - `Code.gs` — Entry point: menu, triggers, processing loop, feedback generation, send orchestration
   - `Config.gs` — Global config, Script Properties access, mode helpers, validation
   - `Sheets.gs` — Generic CRUD layer over Google Sheets (the "database")
   - `Prompts.gs` — Sheet-based prompt system with default fallbacks
   - `ElevenLabs.gs` — Synchronous transcription via ElevenLabs Scribe v2
   - `Gemini.gs` — Gemini API: speaker ID, feedback generation, contribution extraction
   - `Canvas.gs` — Canvas LMS API: roster sync, grade posting, course data fetch
   - `DriveMonitor.gs` — Upload folder monitoring, filename parsing
   - `Email.gs` — HTML/plaintext email templates and distribution
   - `Webapp.gs` — Web app entry point and audio upload handler
4. Create one HTML file:
   - `RecorderApp.html` — Mobile recording UI (create via File > New > HTML file)
5. Copy the content from each file in this repository
6. Update `appsscript.json` (View > Show manifest file) with the contents from this repo

### 2. Create the Tracking Spreadsheet

1. Import `template.xlsx` into Google Sheets (File > Import), or create a new spreadsheet manually
2. Copy the spreadsheet ID from the URL (the long string between `/d/` and `/edit`)
3. Save this ID for configuration

### 3. Get API Keys

#### ElevenLabs
1. Sign up at [elevenlabs.io](https://elevenlabs.io/)
2. Go to your Profile and copy your API key
3. Used for: Scribe v2 speech-to-text with speaker diarization

#### Google Gemini
1. Go to [AI Studio](https://aistudio.google.com/app/apikey)
2. Create an API key
3. Cost: Free tier is generous for this use case

#### Canvas (Optional)
1. In Canvas, go to Account > Settings
2. Click "+ New Access Token"
3. Give it a purpose (e.g., "Harkness Helper")
4. Copy the token (you won't see it again!)

### 4. Configure Script Properties

In the Apps Script editor:
1. Click ⚙️ Project Settings
2. Scroll to "Script Properties"
3. Add these properties:

| Property | Description | Required |
|----------|-------------|----------|
| `ELEVENLABS_API_KEY` | Your ElevenLabs API key | Yes |
| `GEMINI_API_KEY` | Your Gemini API key | Yes |
| `SPREADSHEET_ID` | ID of your tracking spreadsheet | Yes |
| `AUDIO_FOLDER_ID` | Set automatically by setup | Yes |
| `PROCESSING_FOLDER_ID` | Set automatically by setup | Yes |
| `CANVAS_API_TOKEN` | Your Canvas access token | Optional |
| `CANVAS_BASE_URL` | Your Canvas URL (e.g., `https://school.instructure.com`) | Optional |

### 5. Run Initial Setup

1. In the Apps Script editor, select `initialSetup` from the function dropdown
2. Click ▶️ Run
3. Authorize the app when prompted
4. Check the logs for the folder URLs

### 6. Setup Automatic Triggers

1. Open your tracking spreadsheet
2. You should see a "Harkness Helper" menu
3. Click "Setup Automatic Triggers"

### 7. Deploy the Recording Web App (Optional)

The recorder lets you record and upload discussions directly from your phone's browser.

1. In the Apps Script editor, click **Deploy > New deployment**
2. Select type: **Web app**
3. Execute as: **Me**
4. Who has access: **Only myself** (or anyone with a Google account in your org)
5. Click **Deploy** and copy the URL
6. On your phone, open the URL in Safari/Chrome and use **Share > Add to Home Screen** for app-like access

## Usage

### Recording Discussions

**Option A: Use the Recorder web app (recommended)**

1. Open the Recorder on your phone
2. Tap the record button and have students introduce themselves at the beginning
3. Use pause/resume as needed
4. Tap stop when the discussion ends
5. Select the section, confirm the date, and tap Upload
6. The file lands in the Drive upload folder with the correct filename — the processing pipeline takes it from there

**Option B: Manual upload**

1. Record your Harkness discussion on your phone
2. **Important**: Have students introduce themselves at the beginning ("Hi, I'm [name]")
3. Upload the audio file to the "Upload" folder in Google Drive
4. Name the file like `Section 1 - 2024-01-15.m4a` so section and date are auto-detected

### Processing Workflow

1. **Automatic**: System detects new audio file (checks every 10 minutes)
2. **Automatic**: ElevenLabs transcribes with speaker diarization (synchronous — no polling needed)
3. **Automatic**: Gemini identifies speakers from introductions, populates SpeakerMap sheet
4. **Manual** (individual mode): Review/correct speaker names in the SpeakerMap sheet
5. **Manual**: Enter grade(s) — one grade on the Discussion row (group mode) or per-student in StudentReports (individual mode)
6. **Manual**: Click "Generate Feedback" from the menu — Gemini writes feedback
7. **Manual**: Review/edit feedback, check the approved box(es)
8. **Manual**: Click "Send Approved Feedback" from the menu

### Google Sheets Structure

| Sheet | Purpose |
|-------|---------|
| **Settings** | Key-value global config (mode, distribution flags, grade scale, teacher info) |
| **Discussions** | One row per discussion session with status, grade, group_feedback, next_step, course |
| **Students** | Student roster with email, Canvas IDs, and course assignment |
| **Transcripts** | Raw/named transcripts, speaker map JSON |
| **SpeakerMap** | One row per speaker per discussion for teacher review/confirmation |
| **StudentReports** | Individual student reports with contributions, grades, feedback (individual mode) |
| **Prompts** | Teacher-editable prompt templates read at runtime |
| **Courses** | Multi-course lookup table (optional): course_name, canvas_course_id, overrides |

### Two Modes

Controlled by the `mode` setting in the Settings sheet:

- **Group mode** (`group`): One grade and one feedback for the whole class. Stored on the Discussion row. Speaker map auto-confirms.
- **Individual mode** (`individual`): Per-student grades and feedback. Stored in StudentReports rows. Teacher must confirm speaker map before proceeding.

### Multi-Course Support

If you teach multiple courses (e.g., AP English and World History), you can manage them all from one spreadsheet.

**Setup:**
1. Click **Enable Multi-Course** from the Harkness Helper menu
2. This creates a **Courses** sheet and adds a `course` column to Discussions and Students
3. Rename "My Course" to your actual course name, add additional courses with their Canvas course IDs
4. Run **Sync Canvas Roster** for each course (or **Sync All Course Rosters** to sync all at once)

**Courses sheet columns:**
| Column | Purpose |
|--------|---------|
| `course_name` | Friendly name (e.g., "AP English") |
| `canvas_course_id` | The Canvas course ID for this course |
| `canvas_base_url` | Optional override (falls back to global Settings) |
| `canvas_item_type` | Optional override: "assignment" or "discussion" |

**How it works:**
- Each Discussion and Student row has a `course` column linking to the Courses sheet
- Grade posting, roster sync, and email distribution are course-aware
- The Recorder web app shows a course picker when multi-course mode is enabled
- Filenames include the course name: `AP English - Section 1 - 2025-02-20.webm`
- If a discussion's `course` column is empty, the system falls back to the global `canvas_course_id` in Settings

**Backward compatible:** If you don't enable multi-course, everything works exactly as before.

## File Naming Convention

For best results, name your audio files like:
- `Section 1 - 2024-01-15.m4a`
- `S1_Discussion_20240115.mp3`
- `Section1.m4a` (date defaults to today)

Multi-course format (course name prefix is auto-detected):
- `AP English - Section 1 - 2024-01-15.m4a`

## Customizing Prompts

All AI prompts live in the **Prompts sheet**. You can customize:
- Speaker identification instructions
- How participation is evaluated
- Grading criteria and scale
- Feedback tone and style (default: 2-paragraph critique sandwich)

Edit prompts directly in the sheet — no code changes needed.

## Estimated Costs

| Service | Cost |
|---------|------|
| ElevenLabs | Scribe v2 pricing (check current rates at elevenlabs.io) |
| Google Gemini | Free tier usually sufficient |
| Canvas API | Free |
| Google Apps Script | Free |

## Troubleshooting

### "Configuration incomplete" error
Run `testConfiguration()` from the script editor to see which API keys are missing.

### Transcription stuck in "transcribing"
- The system auto-detects stuck transcriptions after 10 minutes and marks them as errors
- Check the error_message column on the Discussions sheet
- Try splitting very long audio files into shorter segments

### Speaker identification wrong
- Students need clear introductions at the start
- You can manually edit speaker names in the SpeakerMap sheet
- Click "Generate Feedback" again after correcting

### Emails not sending
- Check students have email addresses in the Students sheet
- Check Gmail sending limits (500/day for regular accounts)
- Verify `distribute_email` is set to `true` in Settings

### Canvas API errors
- Token might have expired (create a new one)
- Check CANVAS_BASE_URL doesn't have a trailing slash
- Verify canvas_course_id is set in Settings
- Set canvas_item_type to "assignment" or "discussion" in Settings

## Privacy & Data

- Audio files are temporarily made public for ElevenLabs to access (files > 50MB only; smaller files are uploaded directly as blobs)
- Permissions are automatically restored after transcription completes
- Transcripts and reports are stored in your Google Sheets (your control)
- API keys are stored in Script Properties (encrypted by Google)

## License

MIT License - feel free to modify for your needs!

## Support

If you encounter issues:
1. Check the Execution log in Apps Script (View > Executions)
2. Run `testConfiguration()` to verify setup
3. Open an issue on GitHub with the error message
