# Gmail Drive Receipt Ingest

Google Apps Script workflow that turns labeled Gmail receipt emails into a clean, searchable Google Drive archive.

It scans receipt emails from configured Gmail labels, routes them to person-specific folders, converts the email body into a PDF, saves supported attachments, and names files using the receipt date, vendor, and order or reference number.

## Why this exists

This project was built around a real executive receipt-management use case.

The goal was to reduce manual sorting by automatically:

- pulling receipt emails from Gmail
- organizing them in Google Drive
- preserving the readable email body as a PDF
- saving original receipt attachments
- creating cleaner filenames for long-term retrieval

## What it does

- Reads receipt emails from configured Gmail labels
- Maps each source label to a person folder
- Parses HTML, plain text, inline images, and attachments
- Converts the email body into a PDF
- Preserves important inline images as closely as Apps Script allows
- Saves original attachments into Drive
- Organizes output by person, year, and month
- Marks processed messages to prevent duplicate ingestion
- Supports dry-run testing before real file writes
- Routes uncertain items to review
- Ignores attached `.eml` files
- Restricts processing to emails on or after a configured minimum date

## How the workflow works

1. A receipt email is labeled in Gmail
2. The script checks whether the message qualifies for processing
3. The source label is mapped to a person folder
4. The email body is rendered into a PDF
5. Supported attachments are saved to an `_attachments` folder
6. The files are renamed using extracted receipt context
7. The message is marked as processed so it does not run again

## Output structure

~~~text
Receipts Archive/
  _TEST/
    archana/
      2026/
        2026-03/
          31Mar-HampersAndCo-W56606-email-body.pdf
          _attachments/
            31Mar-HampersAndCo-W56606-Trip_Receipt.pdf
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
~~~

## File naming

Generated files follow this general pattern:

~~~text
DDMon-Vendor-OrderOrReference-email-body.pdf
~~~

Examples:

~~~text
31Mar-HampersAndCo-W56606-email-body.pdf
31Mar-PopMart-O1843507719879487488-email-body.pdf
31Mar-Devines-360069x2-email-body.pdf
~~~

Attachments follow the same general naming style and keep the original filename after the receipt context prefix.

Example:

~~~text
31Mar-Devines-360069x2-Trip_Receipt_360069x2.pdf
~~~

## Current behavior

For each processed email, the workflow currently saves:

1. **Email body PDF**
   - rendered from the email body
   - inline images preserved as closely as possible
   - remote image fetching handled on a best-effort basis

2. **Original attachments**
   - PDFs, images, and other supported file types
   - stored in an `_attachments` subfolder when present

The workflow currently does **not** save:

- raw `.eml` files
- HTML snapshots as separate files
- manifest JSON files
- browser-rendered PNG screenshots of the email body

## Gmail labels

The script expects source labels like:

- `archana-expenses`
- `dan-expenses`
- `rhea-expenses`

It also uses processing labels such as:

- `receipt-ingested-test`
- `receipt-needs-review-test`

The source label determines which person folder the receipt is routed into.

## Setup

### 1. Add the script files

Create a Google Apps Script project and add:

- `Code.gs`
- `appsscript.json`

### 2. Enable required services

Built-in Apps Script services used:

- `GmailApp`
- `DriveApp`
- `PropertiesService`
- `LockService`
- `ScriptApp`
- `UrlFetchApp`

Advanced services required:

- `Gmail`
- `Drive`

### 3. Confirm manifest scopes

Your `appsscript.json` should include the required advanced services and OAuth scopes.

Example:

~~~json
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
~~~

## Configuration

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
- `FETCH_REMOTE_IMAGES`
- `IGNORE_EML_ATTACHMENTS`

### Default config example

~~~javascript
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
      SAVE_INLINE_IMAGE_FILES: 'false',
      FETCH_REMOTE_IMAGES: 'true',
      IGNORE_EML_ATTACHMENTS: 'true'
    },
    true
  );

  Logger.log('Config saved to Script Properties.');
}
~~~

## First-time run

Run these functions in order:

~~~javascript
setupConfig()
verifySetup()
dryRunScan()
~~~

What they do:

- `setupConfig()` writes the initial Script Properties
- `verifySetup()` confirms labels, folders, services, and access
- `dryRunScan()` shows what would happen without writing files

## Real processing run

After the dry run looks correct, run:

~~~javascript
resetTestProcessedState()
setDryRunFalse()
runReceiptIngestion()
~~~

## Re-run test emails

If you want to test already-processed emails again:

~~~javascript
resetTestProcessedState()
dryRunScan()
setDryRunFalse()
runReceiptIngestion()
~~~

## Scheduling

To automate processing, install the weekly trigger:

~~~javascript
installWeeklyTrigger()
~~~

Current helper behavior:

- runs `runReceiptIngestion()`
- weekly
- Friday
- around 5 PM

## Running this in another person’s Gmail account

If this workflow is used in another person’s Gmail account, that user should authorize the script and create the installable trigger from their own account.

Recommended flow:

1. Open or copy the script into that user’s Google account
2. Enable the required Apps Script services
3. Authorize the script from that user’s account
4. Create the Gmail labels in that user’s Gmail
5. Create the Drive folder structure in that user’s Drive
6. Run:
   - `setupConfig()`
   - `verifySetup()`
   - `dryRunScan()`
7. Install the trigger with:
   - `installWeeklyTrigger()`

If the trigger is created by a different account, the workflow can run in the wrong Gmail or Drive context.

## Changing label names

If you want to use different Gmail labels, update both:

1. the labels in Gmail
2. the labels expected by the script

### Update the `LABELS` config

Example:

~~~javascript
LABELS: 'bob-expenses,sarah-expenses,mike-expenses'
~~~

### Update the label-to-folder mapping

Example:

~~~javascript
function mapLabelToPerson_(labelName) {
  if (labelName === 'bob-expenses') return 'bob';
  if (labelName === 'sarah-expenses') return 'sarah';
  if (labelName === 'mike-expenses') return 'mike';
  return 'unknown';
}
~~~

These values must match exactly:

- the Gmail label
- the value inside `LABELS`
- the value inside `mapLabelToPerson_()`

After changing them, run:

~~~javascript
setupConfig()
verifySetup()
dryRunScan()
~~~

## Date filtering

The workflow currently only processes emails on or after:

~~~javascript
MIN_EMAIL_DATE: '2026-01-01'
~~~

This is enforced in two ways:

1. an `after:` filter is added to the Gmail query
2. each message date is checked again in code before processing

To change the cutoff, update `MIN_EMAIL_DATE` and rerun:

~~~javascript
setupConfig()
dryRunScan()
~~~

## Receipt detection and naming

Vendor naming is based on best-effort extraction from sources such as:

- attachment names
- sender display names inside forwarded content
- known vendor patterns in the email body
- HTML title or heading content
- sender domain
- cleaned subject fallback

Order or reference numbers are extracted from:

- subject line
- plain text body
- HTML body
- attachment names

This is heuristic by design and can be extended as new receipt formats appear.

## How images are handled

The script tries to preserve images in the rendered PDF by:

1. embedding `cid:` inline images from MIME parts
2. fetching a limited number of remote `<img src="...">` assets
3. converting those assets into data URIs before PDF generation

To keep processing reliable, it also avoids or limits:

- tracking pixels
- beacon URLs
- some static map or marker URLs
- extremely long signed URLs
- noisy image-heavy content that adds little archival value

## Idempotency

The workflow avoids duplicate processing by:

- applying a processed Gmail label
- storing processed message IDs in Script Properties

Deleting files from Drive does **not** automatically make the original emails eligible again.

For testing, use:

~~~javascript
resetTestProcessedState()
~~~

## Key functions

### `setupConfig()`
Writes the base Script Properties used by the workflow.

### `verifySetup()`
Checks Gmail labels, Drive folders, manifest access, and required services.

### `dryRunScan()`
Simulates processing so you can validate naming, routing, and extraction before writing files.

### `resetTestProcessedState()`
Clears test processed-state tracking so messages can be tested again.

### `setDryRunFalse()`
Switches the workflow out of dry-run mode.

### `runReceiptIngestion()`
Processes matching emails, writes files to Drive, and marks messages as processed.

### `installWeeklyTrigger()`
Creates a recurring trigger so the script runs automatically.

## Limitations

This project is intentionally Apps Script-only.

That means:

- PDF generation is based on Apps Script blob conversion
- output fidelity is strong, but not identical to a browser screenshot renderer
- vendor extraction is heuristic, not guaranteed for every merchant format
- some remote or authenticated images may still fail to render
- full visual screenshot capture of the email body is not the goal of this workflow

The strongest archival output in this project is the combination of:

- rendered email-body PDF
- preserved original attachments
- structured folder organization
- cleaner searchable filenames

## Repo purpose

This repo exists to turn labeled Gmail receipt emails into a cleaner and more maintainable Google Drive archive for long-term receipt storage and retrieval.