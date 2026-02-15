# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Harkness Discussion Helper is a **Google Apps Script** application that automates Harkness discussion workflows for teachers. There is no build system, package manager, or test framework — all `.gs` files in `src/` are manually copied into the Google Apps Script editor at [script.google.com](https://script.google.com).

## Development Workflow

- **No CLI commands**: This is a GAS project. There's no `npm`, `build`, `lint`, or `test` to run.
- **Deployment**: Copy each `src/*.gs` file into the Apps Script editor and update `appsscript.json` via View > Show manifest file.
- **Testing**: Run functions like `testConfiguration()` or `initialSetup()` directly from the Apps Script editor's function dropdown.
- **Debugging**: Use `Logger.log()` — output appears in the Apps Script Execution log (View > Executions).

## Architecture

### Processing Pipeline

```
Audio upload (Drive) → ElevenLabs (transcription + diarization, synchronous)
    → Gemini (speaker ID from introductions, auto-suggest)
    → Teacher reviews speaker map in SpeakerMap sheet
    → MODE BRANCH:
        Group:      Teacher enters ONE grade → Gemini group feedback → approve → send
        Individual: Teacher enters per-student grades → Gemini per-student feedback → approve → send
    → Gmail (report emails) + Canvas API (grade posting)
```

### Discussion Status Machine

Defined in `Config.gs` as `CONFIG.STATUS`:
```
uploaded → transcribing → mapping → review → approved → sent
                                                        ↘ error (from any step)
```

- `uploaded → transcribing`: automatic (trigger detects file, starts ElevenLabs)
- `transcribing → mapping`: automatic (transcription complete, Gemini suggests speaker names)
- `mapping → review`: group mode auto-advances; individual mode waits for teacher to confirm speaker map
- `review → approved`: teacher enters grade(s), reviews feedback, checks approved checkbox(es)
- `approved → sent`: teacher clicks "Send Approved Feedback" from menu

The `mainProcessingLoop()` in `Code.gs` runs on a 10-minute trigger and handles automatic advancement. Feedback generation and sending are manual (menu-driven).

### Two Modes

Controlled by `mode` setting in the Settings sheet:
- **Group mode** (`group`): One grade + one feedback for the whole class. Stored on the Discussion row.
- **Individual mode** (`individual`): Per-student grades + feedback. Stored in StudentReports rows.

### Source Files

| File | Role |
|------|------|
| `Code.gs` | Entry point: menu, trigger setup, `mainProcessingLoop()`, speaker mapping, feedback generation, send orchestration |
| `Config.gs` | Global `CONFIG` constant, Script Properties access, mode helpers (`isGroupMode`/`isIndividualMode`), validation |
| `Sheets.gs` | Generic CRUD layer over Google Sheets (the "database"), plus domain-specific helpers for all 7 sheets |
| `Prompts.gs` | Sheet-based prompt system with `DEFAULT_PROMPTS` fallback; `getPrompt(name, {vars})` reads from Prompts sheet |
| `ElevenLabs.gs` | Synchronous transcription via ElevenLabs Scribe v2, blob vs URL upload based on file size, Drive file sharing |
| `Gemini.gs` | Gemini API calls, speaker ID, group/individual feedback generation, contribution extraction, discussion summaries |
| `Canvas.gs` | Canvas LMS API: paginated requests, student sync, dual-mode grade posting, course data fetch |
| `DriveMonitor.gs` | Upload folder monitoring, filename parsing (section + date extraction), folder setup |
| `Email.gs` | Dual-mode HTML/plaintext email templates, distribution with settings-based channel control |

### Sheets Structure (7 sheets)

| Sheet | Purpose |
|-------|---------|
| `Settings` | Key-value global config (mode, distribution flags, grade scale, teacher info) |
| `Discussions` | One row per discussion session with status, grade, group_feedback, next_step |
| `Students` | Student roster with email and Canvas IDs |
| `Transcripts` | Raw/named transcripts, speaker map JSON, teacher feedback |
| `SpeakerMap` | One row per speaker per discussion for teacher review/confirmation |
| `StudentReports` | Individual student reports with contributions, grades, feedback (individual mode) |
| `Prompts` | Teacher-editable prompt templates read at runtime |

### Key Patterns

- **Google Sheets as database**: All data lives in 7 sheets. `Sheets.gs` provides `getAllRows()`, `findRow()`, `insertRow()`, `updateRow()` that treat the header row as column keys and return row objects with `_rowIndex` for updates.
- **Settings sheet for runtime config**: Teacher-facing settings (mode, distribution flags) live in a key-value Settings sheet, read via `getSetting(key)`.
- **Script Properties for secrets**: API keys and folder IDs stored via `PropertiesService.getScriptProperties()`. Access through typed getters in `Config.gs` (e.g., `getElevenLabsKey()`).
- **Sheet-based prompts**: All AI prompts live in the Prompts sheet. `getPrompt(name, {vars})` reads from sheet first, falls back to `DEFAULT_PROMPTS`. Teachers can edit prompts without touching code.
- **Temporary file sharing**: `getDriveFileUrl()` makes audio files temporarily public for ElevenLabs URL-based upload (files > 50MB), `restoreFilePermissions()` reverts after transcription.
- **Blob vs URL upload**: Files <= 50MB are uploaded as blobs directly to ElevenLabs. Files > 50MB use a temporary public Drive URL to avoid GAS memory limits.
- **Error logging**: Errors append to `Discussions.error_message` with timestamps. `next_step` column shows "ERROR: ..." for teacher visibility. All errors also go to `Logger.log()`.
- **Rate limiting**: `Utilities.sleep()` calls between API requests (200-500ms) to avoid hitting external API limits.
- **GAS execution limit**: 6-minute timeout for script execution. `CONFIG.LIMITS.GAS_TIMEOUT_MS` is set to 5 minutes as a safety margin. Stuck transcriptions detected after 10 minutes.

### External APIs

| API | Model/Tier | Used For |
|-----|-----------|----------|
| ElevenLabs | Scribe v2 | Synchronous transcription with speaker diarization |
| Google Gemini | `gemini-1.5-flash` | Speaker ID, summaries, feedback generation |
| Canvas LMS | REST v1 | Student roster, grade posting |

### Required OAuth Scopes (appsscript.json)

`spreadsheets`, `drive`, `gmail.send`, `script.external_request`, `script.scriptapp`
