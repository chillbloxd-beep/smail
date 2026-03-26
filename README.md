# AI Gmail Digest with Gemini (Apps Script)

This project reads unread Gmail emails, summarizes them with Gemini, and sends two digest emails daily.

## Features

- Summarizes unread Gmail emails with Gemini
- Sends digest twice daily
- Uses batching to reduce Gemini rate-limit issues
- Shows a maximum of 50 summarized emails in the digest email body
- Attaches a `.txt` file with up to 50 additional summarized emails
- Keeps track of processed emails using the `AI-DIGESTED` Gmail label
- Does not summarize labeled emails again
- Supports manual preview mode without changing labels
- Auto-classifies:
  - `@students.edu.sg` as important
  - `@moe.edu.sg` as important
  - `@gmail.com` as important
  - `@classroom.google.com` as homework

## Script Properties Required

Set these in Apps Script > Project Settings > Script Properties:

- `GEMINI_API_KEY` = your Gemini API key
- `DIGEST_TO` = the email address that receives the digest

## Files

### `Code.gs`
Main digest system:
- reads unread unlabeled emails
- summarizes them in batches
- sends digest email
- attaches overflow `.txt` summary file
- applies processed label

### `Debug.gs`
Debug and health-check utilities:
- `testGeminiConnection()`
- `debug()`
- `createMidnightGeminiTestTrigger()`

## Main Functions

### `setup()`
Creates the Gmail label used for processed emails.

### `createDigestTriggers()`
Creates the daily morning and evening digest triggers.

### `scheduledMorningDigest()`
Scheduled AM digest job.

### `scheduledEveningDigest()`
Scheduled PM digest job.

### `manualPreviewDigestNow()`
Runs a preview digest without marking emails as processed.

### `manualProcessDigestNow()`
Runs a real manual digest and marks processed emails.

## Debug Functions

### `testGeminiConnection()`
Tests Gemini API connectivity and emails the result.

### `debug()`
Emails a detailed debug report including:
- config values
- trigger list
- unread/unprocessed sample counts
- sender domain sample stats
- estimated Gemini calls per run
- sample email metadata

### `createMidnightGeminiTestTrigger()`
Creates a daily trigger that runs `testGeminiConnection()` around midnight.

## Processing Limits

Per run:
- Up to 50 summaries in the email body
- Up to 50 more summaries in a `.txt` attachment
- Maximum total processed per run: 100 emails

Anything beyond that remains unlabeled and is handled in future runs.

## How processed emails are tracked

Processed emails are tracked using the Gmail label:

`AI-DIGESTED`

If an email thread already has that label, it is skipped in scheduled and manual processing runs.

## Trigger behavior

This project is designed to send digests before:
- 7:00 AM
- 7:00 PM

Apps Script triggers are approximate, so the project schedules them earlier to provide time buffer.

## Notes

- Apps Script runs on Google’s servers. You do not need to keep the Apps Script tab open.
- Preview mode does not change labels.
- If Gemini fails, check `debug()` and `testGeminiConnection()` first.
