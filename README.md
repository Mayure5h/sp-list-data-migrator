# SharePoint List Migrator

Browser-console scripts to export custom-column data from one or more SharePoint
lists on a source tenant, and import that data into matching lists on a target
tenant.

## Usage

### 1. Export (run on the source site)

1. Open the source SharePoint site in your browser.
2. Open DevTools (F12) > Console.
3. Open `sp-export.js`, edit the `LIST_NAMES` array with the exact list Titles
   you want to export, and paste the whole script into the console.
4. A JSON file (`sharepoint-export-<timestamp>.json`) downloads automatically.

### 2. Import (run on the target site)

1. Make sure the destination lists already exist on the target tenant, with the
   same custom column internal names as the source.
2. Open the target SharePoint site, open DevTools > Console.
3. Paste `sp-import.js` into the console and run it.
4. When prompted, select the JSON file from step 1.

## Known limitations

- Only custom columns are exported (`CanBeDeleted eq true`) — built-in system
  columns (Attachments, ContentType, Editor, etc.) are skipped.
- Attachments are not exported (SharePoint REST doesn't return file bytes for
  item queries).
- Lookup and Person/Group columns are exported/imported as raw Ids (or arrays
  of Ids for multi-value fields). These Ids are tenant-specific and will
  almost certainly not point to the same record on the target tenant — you'll
  need to build an old-Id → new-Id mapping and rewrite the exported JSON
  before importing, or handle those columns manually afterward.
- No throttling/retry logic; for very large lists you may need to add delays
  between requests to avoid SharePoint throttling (HTTP 429).

## Requirements

- Read access to the source lists.
- Contribute access to the target lists.
- Both scripts must be run directly in the browser console while logged into
  the respective SharePoint site (they rely on `_spPageContextInfo` and the
  browser's authenticated session — no credentials are stored or transmitted
  anywhere else).
