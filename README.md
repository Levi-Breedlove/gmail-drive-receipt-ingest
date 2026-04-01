# gmail-receipt-to-drive-pdf

Google Apps Script workflow that ingests labeled receipt emails from Gmail, converts the email body to a PDF with inline images preserved as closely as possible, saves original attachments to Google Drive, organizes output by person and month, and generates cleaner filenames using the email date, vendor, and order or reference number.

## Overview

This project automates a receipt-ingestion flow using Gmail labels and Google Drive.

It monitors labeled Gmail emails, identifies which person a receipt belongs to based on the applied label, renders the email body into a PDF artifact, preserves inline images inside that rendered PDF as closely as Apps Script allows, saves original attachments such as PDF and image files, and stores everything in a structured Google Drive archive.

The archive is organized by:

- person
- year
- month

The output filenames are designed to be cleaner and more searchable by using:

- email date
- vendor name
- order or reference number

## What This Project Does

- Reads Gmail messages from configured labels
- Maps each label to a person folder
- Parses the email body and inline image content
- Converts the email body into a PDF
- Saves original attachments into Drive
- Organizes output into person/year/month folders
- Marks processed messages to prevent duplicate ingestion
- Supports dry-run testing before real file writes
- Generates cleaner filenames based on extracted receipt information
- Restricts processing to emails dated **2026-01-01 and onward**

## Current Output Behavior

For each processed email, the workflow saves:

1. **Email body PDF**
   - rendered from the email body
   - inline images are preserved as closely as possible
   - remote image fetching is limited to avoid bloated or looping image fetches

2. **Original attachments**
   - PDF, PNG, JPG, and other supported files
   - stored in an `_attachments` subfolder when present

The current workflow does **not** save:

- HTML body files
- raw `.eml` files
- manifest JSON files
- full rendered PNG screenshots of the email body

## Folder Structure

```text
Receipts Archive/
  _TEST/
    archana/
      2026/
        2026-03/
          31Mar-HampersAndCo-W56606-email-body.pdf
          _attachments/
            31Mar-HampersAndCo-W56606-original-file.pdf
    dan/
      2026/
        2026-03/
    rhea/
      2026/
        2026-03/
    _Needs Review/
      archana/
      dan/
      rhea/
    _Logs/
```

## File Naming

The workflow generates filenames using this pattern:

```text
DDMon-Vendor-OrderOrReference-email-body.pdf
```

Examples:

```text
31Mar-HampersAndCo-W56606-email-body.pdf
31Mar-PopMart-O1843507719879487488-email-body.pdf
31Mar-Devines-360069x2-email-body.pdf
```

Attachment filenames follow the same general naming pattern and preserve the original attachment name after the receipt context prefix.

Example:

```text
31Mar-Devines-360069x2-Trip_Receipt_360069x2.pdf
```

## Gmail Labels

The workflow currently expects Gmail labels like:

- `archana-expenses`
- `dan-expenses`
- `rhea-expenses`

It also uses processing labels such as:

- `receipt-ingested-test`
- `receipt-needs-review-test`

The source label determines which person folder the receipt is routed into.

## How to Change Label Names

If you want to use different Gmail label names, you need to update both:

1. the labels in Gmail
2. the labels expected by the script

### Step 1: Create or rename the labels in Gmail

In the Gmail account that the script uses, create the labels you want the workflow to monitor.

Examples:

- `bob-expenses`
- `sarah-expenses`
- `mike-expenses`

Apply those labels to the receipt emails you want the workflow to ingest.

### Step 2: Update the label list in the script configuration

Update the `LABELS` value in the script so it matches the Gmail labels you created.

Example:

```javascript
LABELS: 'bob-expenses,sarah-expenses,mike-expenses'
```

If you are using Script Properties directly, make sure the `LABELS` property matches the same comma-separated values.

Example:

```text
LABELS=bob-expenses,sarah-expenses,mike-expenses
```

### Step 3: Update the label-to-folder mapping in code

The current script uses a hardcoded `mapLabelToPerson_()` function.

If you change the label names, you must also change the mapping.

Example:

```javascript
function mapLabelToPerson_(labelName) {
  if (labelName === 'bob-expenses') return 'bob';
  if (labelName === 'sarah-expenses') return 'sarah';
  if (labelName === 'mike-expenses') return 'mike';
  return 'unknown';
}
```

This means:

- `bob-expenses` → `bob/`
- `sarah-expenses` → `sarah/`
- `mike-expenses` → `mike/`

## Code Changes Required for Label Mapping

This project does **not** automatically infer new Gmail label names from nowhere.

If you change the Gmail labels, you must update the script in these two places.

### 1. Change the `LABELS` value in `setupConfig()`

Find this section in the code:

```javascript
function setupConfig() {
  const props = PropertiesService.getScriptProperties();

  props.setProperties(
    {
      ROOT_FOLDER_NAME: 'Receipts Archive',
      TEST_SUBROOT_NAME: '_TEST',
      PROCESSED_LABEL: 'receipt-ingested-test',
      REVIEW_LABEL: 'receipt-needs-review-test',
      LABELS: 'archana-expenses,dan-expenses,rhea-expenses',
      TIMEZONE: 'America/Los_Angeles',
      DRY_RUN: 'true',
      MIN_EMAIL_DATE: '2026-01-01',

      SAVE_HTML: 'false',
      SAVE_PDF: 'true',
      SAVE_RAW_EMAIL: 'false',
      SAVE_MANIFEST: 'false',
      SAVE_ATTACHMENTS: 'true',
      SAVE_INLINE_IMAGE_FILES: 'false'
    },
    true
  );

  Logger.log('Config saved to Script Properties.');
}
```

Replace this line:

```javascript
LABELS: 'archana-expenses,dan-expenses,rhea-expenses',
```

with your own labels, for example:

```javascript
LABELS: 'bob-expenses,sarah-expenses,mike-expenses',
```

### 2. Change the `mapLabelToPerson_()` function

Find this function in the code:

```javascript
function mapLabelToPerson_(labelName) {
  if (labelName === 'archana-expenses') return 'archana';
  if (labelName === 'dan-expenses') return 'dan';
  if (labelName === 'rhea-expenses') return 'rhea';
  return 'unknown';
}
```

Replace it with your own mappings, for example:

```javascript
function mapLabelToPerson_(labelName) {
  if (labelName === 'bob-expenses') return 'bob';
  if (labelName === 'sarah-expenses') return 'sarah';
  if (labelName === 'mike-expenses') return 'mike';
  return 'unknown';
}
```

### 3. Make sure Gmail label names and code values match exactly

These values must line up exactly:

- the label you create in Gmail
- the label string inside `LABELS`
- the label string inside `mapLabelToPerson_()`

Example:

```text
Gmail label: bob-expenses
LABELS value: bob-expenses
mapLabelToPerson_(): if (labelName === 'bob-expenses') return 'bob';
```

If these do not match exactly, the script will not route the email into the correct folder.

### 4. Re-run setup and verification

After changing label names in Gmail and in the script, run:

```javascript
setupConfig()
verifySetup()
dryRunScan()
```

That confirms the new labels are recognized before you do a live ingestion run.

### Optional: Change processing labels too

If you also want to rename the processing labels, update:

- `PROCESSED_LABEL`
- `REVIEW_LABEL`

Example:

```javascript
PROCESSED_LABEL: 'receipt-ingested-test',
REVIEW_LABEL: 'receipt-needs-review-test'
```

If you rename those, make sure the new labels exist in Gmail too.

## Date Filtering

The current workflow only processes emails from **2026-01-01 and onward**.

This is controlled by the `MIN_EMAIL_DATE` setting.

Example:

```javascript
MIN_EMAIL_DATE: '2026-01-01',
```

The script enforces this in two ways:

1. it adds an `after:` filter to the Gmail search query
2. it checks each message date again in code before processing

### How to change the minimum date

If you want to change the date cutoff, update:

```javascript
MIN_EMAIL_DATE: '2026-01-01',
```

to something else, for example:

```javascript
MIN_EMAIL_DATE: '2027-01-01',
```

Then re-run:

```javascript
setupConfig()
dryRunScan()
```

## Required Google Apps Script Services

### Built-in services

- `GmailApp`
- `DriveApp`
- `PropertiesService`
- `LockService`
- `ScriptApp`
- `UrlFetchApp`

### Advanced services

- `Gmail`
- `Drive`

## Required Manifest Configuration

Your `appsscript.json` should include the enabled advanced services and required OAuth scopes.

Example:

```json
{
  "timeZone": "America/Los_Angeles",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v3",
        "serviceId": "drive"
      },
      {
        "userSymbol": "Gmail",
        "version": "v1",
        "serviceId": "gmail"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

## Script Configuration

The project uses Script Properties for configuration.

Typical properties include:

- `ROOT_FOLDER_NAME`
- `TEST_SUBROOT_NAME`
- `PROCESSED_LABEL`
- `REVIEW_LABEL`
- `LABELS`
- `TIMEZONE`
- `DRY_RUN`
- `MIN_EMAIL_DATE`
- `SAVE_PDF`
- `SAVE_ATTACHMENTS`

## Commands to Run

### First-time setup

Run these functions in order:

```javascript
setupConfig()
verifySetup()
dryRunScan()
```

### Real processing run

After dry run looks correct:

```javascript
resetTestProcessedState()
setDryRunFalse()
runReceiptIngestion()
```

### Re-run test emails again

If you already processed the test emails and want to test again:

```javascript
resetTestProcessedState()
dryRunScan()
setDryRunFalse()
runReceiptIngestion()
```

## Optional Scheduling

To make the script run automatically on a schedule, install the weekly trigger:

```javascript
installWeeklyTrigger()
```

This creates a recurring trigger for `runReceiptIngestion()`.

### Current trigger behavior

The current helper creates a weekly trigger for:

- `runReceiptIngestion()`
- Friday
- around 5 PM

## Running This Script in Another Person's Gmail Account

If this workflow is used in another person’s Gmail account, that person should authorize the script and create the installable trigger from their own account.

Recommended setup for another user:

1. copy or open the script in that user’s Google account
2. enable the required Apps Script services
3. authorize the script from that user’s account
4. create the Gmail labels in that user’s Gmail
5. create the Drive folders in that user’s Google Drive
6. run:
   - `setupConfig()`
   - `verifySetup()`
   - `dryRunScan()`
7. install the trigger by running:
   - `installWeeklyTrigger()`

If the trigger is created by a different account, the workflow may run under the wrong Gmail / Drive context.

## Key Functions

### `setupConfig()`

Writes initial Script Properties used by the workflow.

### `verifySetup()`

Verifies Gmail labels, Drive folders, manifest access, and required service access.

### `dryRunScan()`

Parses messages without writing files, so you can verify vendor extraction, order number extraction, attachment detection, minimum date filtering, and PDF naming safely.

### `resetTestProcessedState()`

Clears test processed-state tracking so test messages can be run again.

### `setDryRunFalse()`

Switches the workflow out of dry-run mode.

### `runReceiptIngestion()`

Processes matching labeled emails, writes files to Drive, and marks messages as processed.

### `installWeeklyTrigger()`

Creates a scheduled trigger so the workflow runs automatically instead of manually.

## How Vendor and Order Number Extraction Works

Vendor naming is based on best-effort extraction from:

1. attachment names
2. forwarded sender display name inside the email body
3. known vendor patterns inside the actual email content
4. HTML title or heading content
5. forwarded sender domain
6. cleaned subject fallback

Order or reference numbers are extracted from:

- subject line
- plain body
- HTML body text
- attachment names

This is intentionally heuristic and can be extended over time as new vendor formats appear.

## Idempotency

The workflow avoids duplicate processing by:

- adding a processed Gmail label
- storing processed message IDs in Script Properties

Deleting the generated files from Drive does not automatically make the emails eligible again. For testing, use:

```javascript
resetTestProcessedState()
```

## How Remote Images Are Handled

The script attempts to preserve images in the rendered PDF by:

1. embedding `cid:` inline images from the email MIME parts
2. fetching a limited number of remote `<img src="...">` images
3. converting those images into data URIs before PDF conversion

To avoid bloated loops or noisy ride / marketing emails, the script also:

- caps remote image fetches
- skips obvious tracking and beacon URLs
- skips many static map / marker URLs
- skips extremely long signed URLs

This helps keep PDF generation more reliable while still preserving the most important visible email images.

## Limitations

This project is intentionally Apps Script-only.

That means:

- PDF generation is based on Apps Script blob conversion
- visual fidelity is strong, but not the same as a true browser screenshot renderer
- a full rendered email-body PNG screenshot is not supported cleanly in Apps Script-only server-side automation
- vendor extraction is heuristic, not guaranteed for every merchant format
- some heavily proxied or authenticated remote images may still fail to appear in the PDF

The best artifact in this workflow is the rendered PDF plus the original saved attachments.

## Repository Purpose

This repo exists to turn labeled Gmail receipt emails into a cleaner, searchable Google Drive archive with:

- readable rendered email PDFs
- preserved receipt attachments
- structured folder organization
- better filenames for long-term retrieval
