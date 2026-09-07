/**
 * SHAREPOINT LIST IMPORT SCRIPT
 * ------------------------------------------------------------
 * HOW TO USE:
 * 1. In the TARGET tenant, create the destination list(s) with the SAME internal
 *    column names as the source (the exported JSON tells you what these are).
 * 2. Navigate to the TARGET SharePoint site in your browser.
 * 3. Open the browser console (F12 > Console tab).
 * 4. Paste this whole script and press Enter.
 * 5. A file picker will pop up — select the JSON file downloaded by sp-export.js.
 * 6. The script will create one item per exported row in the matching list.
 *
 * NOTES / LIMITATIONS:
 * - The destination list must already exist, with the Title matching the key
 *   used in the export (e.g. "List1"), and must already have the same custom
 *   columns (internal names) created on it.
 * - Lookup and Person/Group columns are written using their raw Id (or array of Ids
 *   for multi-value fields). These IDs are almost certainly WRONG on the target
 *   tenant, since item/user IDs aren't consistent across tenants or even across
 *   re-created lists. You'll likely need to build an old-Id -> new-Id mapping for
 *   the target list/user list and rewrite those values in the JSON before importing,
 *   or strip those fields out and set them manually afterward.
 * - No throttling/retry logic is included — for very large lists (thousands of items)
 *   you may hit SharePoint's request throttling. Consider adding a delay between calls
 *   if that happens.
 * - Requires contribute permissions on the target lists.
 */
(async function () {
  const webUrl = _spPageContextInfo.webAbsoluteUrl;

  function pickFile() {
    return new Promise((resolve, reject) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".json,application/json";
      input.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return reject(new Error("No file selected"));
        const reader = new FileReader();
        reader.onload = (evt) => {
          try {
            resolve(JSON.parse(evt.target.result));
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = reject;
        reader.readAsText(file);
      };
      input.click();
    });
  }

  async function getDigest() {
    const res = await fetch(`${webUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose" },
      credentials: "same-origin",
    });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  async function getListItemEntityTypeName(listTitle) {
    const url =
      `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')` +
      `?$select=ListItemEntityTypeFullName`;
    const res = await fetch(url, {
      headers: { Accept: "application/json;odata=verbose" },
      credentials: "same-origin",
    });
    const data = await res.json();
    return data.d.ListItemEntityTypeFullName;
  }

  console.log("Select the JSON file exported by sp-export.js...");
  const importData = await pickFile();

  let digest = await getDigest();

  for (const listTitle of Object.keys(importData)) {
    const { fields, items } = importData[listTitle];
    console.log(`Importing into list: ${listTitle} (${items.length} items)`);

    let entityType;
    try {
      entityType = await getListItemEntityTypeName(listTitle);
    } catch (err) {
      console.error(
        `  List "${listTitle}" not found on target site. Create it first. Skipping.`
      );
      continue;
    }

    const LOOKUP_TYPES = ["Lookup", "LookupMulti", "User", "UserMulti"];

    let successCount = 0;
    for (const item of items) {
      const body = { __metadata: { type: entityType } };
      for (const f of fields) {
        const isLookup = LOOKUP_TYPES.includes(f.TypeAsString);
        const key = isLookup ? `${f.InternalName}Id` : f.InternalName;
        const val = item[key];
        if (val === undefined || val === null) continue;

        if (isLookup && Array.isArray(val)) {
          // Multi-value lookup/person: wrap as { results: [id, id, ...] }
          body[key] = { results: val };
        } else {
          body[key] = val;
        }
      }

      try {
        const res = await fetch(
          `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`,
          {
            method: "POST",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": digest,
            },
            credentials: "same-origin",
            body: JSON.stringify(body),
          }
        );

        if (!res.ok) {
          const errText = await res.text();
          console.error(`  Failed to import item into ${listTitle}:`, errText);
        } else {
          successCount++;
        }
      } catch (err) {
        console.error(`  Network error importing item into ${listTitle}:`, err);
      }
    }
    console.log(`  -> ${successCount}/${items.length} items imported successfully`);
  }

  console.log("Import complete.");
})();
