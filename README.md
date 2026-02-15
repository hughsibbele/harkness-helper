# Harkness Discussion Helper

A Google Apps Script application that automates the workflow for Harkness discussions: recording transcription with speaker identification, AI-powered summaries and evaluations, Canvas integration for reflections and grade posting, and email distribution to students.

## Features

- **Automatic Transcription**: Upload audio recordings to Google Drive and get them automatically transcribed with speaker diarization (AssemblyAI)
- **Speaker Identification**: AI identifies students from their introductions ("Hi, I'm...")
- **Participation Analysis**: Generate individual participation summaries for each student
- **Canvas Integration**: Pull student reflections and post grades/feedback
- **Email Distribution**: Send personalized reports to students
- **Teacher Review**: Review and approve everything in Google Sheets before sending

## Architecture

```
Phone Recording → Google Drive → AssemblyAI (transcription + diarization)
                                        ↓
                              Google Gemini (speaker ID, summaries, evaluation)
                                        ↓
                              Canvas API (pull reflections, post grades)
                                        ↓
                              Google Sheets (teacher review)
                                        ↓
                              Gmail + Canvas (distribution)
```

## Setup Instructions

### 1. Create the Google Apps Script Project

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Delete the default `Code.gs` content
3. Create files for each `.gs` file in the `src/` folder:
   - `Code.gs`
   - `Config.gs`
   - `Sheets.gs`
   - `Prompts.gs`
   - `AssemblyAI.gs`
   - `Gemini.gs`
   - `Canvas.gs`
   - `DriveMonitor.gs`
   - `Email.gs`
4. Copy the content from each file in this repository
5. Update `appsscript.json` (View > Show manifest file) with the contents from this repo

### 2. Create the Tracking Spreadsheet

1. Create a new Google Spreadsheet
2. Copy the spreadsheet ID from the URL (the long string between `/d/` and `/edit`)
3. Save this ID for configuration

### 3. Get API Keys

#### AssemblyAI
1. Sign up at [assemblyai.com](https://www.assemblyai.com/)
2. Go to Dashboard and copy your API key
3. Cost: ~$0.37/hour of audio (first hours free)

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
| `ASSEMBLYAI_API_KEY` | Your AssemblyAI API key | ✅ |
| `GEMINI_API_KEY` | Your Gemini API key | ✅ |
| `SPREADSHEET_ID` | ID of your tracking spreadsheet | ✅ |
| `AUDIO_FOLDER_ID` | Set automatically by setup | ✅ |
| `PROCESSING_FOLDER_ID` | Set automatically by setup | ✅ |
| `CANVAS_API_TOKEN` | Your Canvas access token | Optional |
| `CANVAS_BASE_URL` | Your Canvas URL (e.g., `https://school.instructure.com`) | Optional |
| `CANVAS_COURSE_ID` | The Canvas course ID | Optional |

### 5. Run Initial Setup

1. In the Apps Script editor, select `initialSetup` from the function dropdown
2. Click ▶️ Run
3. Authorize the app when prompted
4. Check the logs for the folder URLs

### 6. Setup Automatic Triggers

1. Open your tracking spreadsheet
2. You should see a "🎓 Harkness Helper" menu
3. Click "Setup Automatic Triggers"

## Usage

### Recording Discussions

1. Record your Harkness discussion on your phone
2. **Important**: Have students introduce themselves at the beginning ("Hi, I'm [name]")
3. **Important**: Give verbal feedback at the end of the discussion
4. Upload the audio file to the "Upload" folder in Google Drive

### Processing Workflow

1. **Automatic**: System detects new audio file
2. **Automatic**: Transcription begins (2-10 minutes depending on length)
3. **Automatic**: Speakers are identified, summaries generated
4. **Manual**: Review reports in the StudentReports sheet
5. **Manual**: Edit grades/feedback as needed
6. **Manual**: Check the "approved" box for each student
7. **Manual**: Click "Send Approved Reports" from the menu

### Google Sheets Structure

#### Discussions Sheet
Track each discussion session with status progression:
`uploaded → transcribing → processing → review → approved → sent`

#### Students Sheet
Student roster with email and Canvas IDs. Auto-populated from:
- Speaker identification
- Canvas course roster (if configured)

#### Transcripts Sheet
Raw and processed transcripts with speaker maps.

#### StudentReports Sheet
Individual student reports with:
- Participation summary
- Reflection (from Canvas)
- Grade (editable)
- Feedback (editable)
- Approval checkbox

## File Naming Convention

For best results, name your audio files like:
- `Period 3 - 2024-01-15.m4a`
- `P3_Discussion_20240115.mp3`
- `Period3.m4a` (date defaults to today)

## Customizing Prompts

All AI prompts are in `Prompts.gs`. You can customize:
- How participation is evaluated
- What makes a "good" reflection
- Grading criteria and scale
- Feedback tone and style

## Estimated Costs

| Service | Cost |
|---------|------|
| AssemblyAI | ~$0.37/hour of audio |
| Google Gemini | Free tier usually sufficient |
| Canvas API | Free |
| Google Apps Script | Free |
| **Monthly estimate** | **$5-10** (for ~6 hours of discussions) |

## Troubleshooting

### "Configuration incomplete" error
Run `testConfiguration()` from the script editor to see which API keys are missing.

### Transcription stuck in "transcribing"
- Check AssemblyAI dashboard for job status
- File might be too large (>50MB limit)
- Audio format might not be supported

### Speaker identification wrong
- Students need clear introductions at the start
- You can manually edit the `speaker_map` JSON in the Transcripts sheet

### Emails not sending
- Check students have email addresses in the Students sheet
- Check Gmail sending limits (500/day for regular accounts)

### Canvas API errors
- Token might have expired (create a new one)
- Check CANVAS_BASE_URL doesn't have a trailing slash
- Verify CANVAS_COURSE_ID is correct

## Privacy & Data

- Audio files are temporarily made public for AssemblyAI to access
- Permissions are restored after transcription completes
- Transcripts and reports are stored in your Google Sheets (your control)
- API keys are stored in Script Properties (encrypted by Google)

## License

MIT License - feel free to modify for your needs!

## Support

If you encounter issues:
1. Check the Execution log in Apps Script (View > Executions)
2. Run `testConfiguration()` to verify setup
3. Open an issue on GitHub with the error message
