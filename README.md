# SALT Transcript Processor

Automatically routes Gemini meeting transcript notes from Warren & Billy's daily quoting meetings into the correct client folders in Google Drive.

## How It Works

1. Every 15 minutes, checks Gmail for new emails from `gemini-notes@google.com` with subject starting with "Notes:"
2. Filters for meetings matching configured keywords (Warren, Billy, Quoting)
3. Loads all client folder names from the "Landscaping Projects" Drive folder
4. Fuzzy-matches client names and addresses mentioned in the transcript body
5. Splits the transcript into per-client segments based on mention order
6. Saves each segment as a Google Doc in that client's `03 - Site Information` subfolder
7. Labels the email as processed so it won't be picked up again

## Fuzzy Matching

The matching is token-based with weighted scoring:

| Signal | Score | Example |
|--------|-------|---------|
| Full name match | 3.0 | "Sean Bryant" found in transcript |
| Last name alone | 1.5 | "Bryant" found as whole word |
| First name alone | 1.0 | "Sean" found as whole word |
| Address/suburb word | 1.0 each | "Capalaba" or "Marine" found |

Minimum score of 2.0 required (configurable via `MIN_MATCH_SCORE`).

Possessives are normalised ("Sean's place" still matches "Sean").

## Deployment

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Replace the contents of `Code.gs` with the script
4. Run `setup()` once — this creates the Gmail label, log sheet, unmatched folder, and 15-minute trigger
5. Authorise the requested scopes when prompted
6. Done — it runs automatically from here

## Testing

- **`testMatching()`** — Dry run: loads the first unprocessed transcript and shows which clients matched and their scores, without saving anything
- **`processManually("msg-id")`** — Process a specific email by its Gmail message ID (find it via Gmail URL or GmailApp)

## Configuration

All config lives in the `CONFIG` object at the top of `Code.gs`:

| Variable | Purpose |
|----------|---------|
| `PARENT_FOLDER_ID` | Google Drive folder ID for "Landscaping Projects" |
| `GMAIL_ACCOUNT` | The Gmail address receiving Gemini notes |
| `GEMINI_SENDER` | Sender email for Gemini meeting notes |
| `MEETING_TITLE_KEYWORDS` | Words to match in meeting title (any one match suffices) |
| `LABEL_NAME` | Gmail label for processed emails |
| `CLIENT_SUBFOLDER` | Subfolder name inside each client folder |
| `DOC_NAME_PREFIX` | Prefix for created Google Doc names |
| `LOG_SHEET_NAME` | Name of the logging spreadsheet |
| `MIN_MATCH_SCORE` | Minimum fuzzy match score to consider a client mentioned |
| `MAX_EMAILS_PER_RUN` | Safety cap on emails processed per trigger execution |

## Monitoring

Check the "Transcript Processing Log" spreadsheet in your Drive root. Every run logs:

- `RUN_START` / `RUN_COMPLETE` — batch boundaries
- `PROCESSING` — started working on an email
- `SAVED` — successfully created a doc in a client folder
- `NO_MATCH` — no clients matched, saved to Unmatched folder
- `SAVE_ERROR` / `FATAL_ERROR` — something went wrong

## Required Scopes

The script will request these permissions on first run:

- `https://www.googleapis.com/auth/gmail.modify` (read emails, apply labels)
- `https://www.googleapis.com/auth/drive` (read folders, create docs)
- `https://www.googleapis.com/auth/spreadsheets` (write to log sheet)
- `https://www.googleapis.com/auth/documents` (create Google Docs)
- `https://www.googleapis.com/auth/script.scriptapp` (manage triggers)
