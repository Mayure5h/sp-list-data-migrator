/**
 * SHAREPOINT LIST EXPORT SCRIPT
 * ------------------------------------------------------------
 * HOW TO USE:
 * 1. Navigate to the SOURCE SharePoint site (e.g. https://tenant.sharepoint.com/sites/YourSite/SitePages/Home.aspx)
 * 2. Open the browser console (F12 > Console tab)
 * 3. Edit the LIST_NAMES array below with the exact Titles of the lists you want to export
 * 4. Paste this whole script into the console and press Enter
 * 5. A JSON file will automatically download containing all custom-column data
 *
 * NOTES / LIMITATIONS:
 * - Only exports "custom" columns (CanBeDeleted = true), i.e. columns you or someone
 *   added — not built-in system columns like Attachments, ContentType, Editor, etc.
 * - Lookup and Person/Group fields are exported as raw Id values (single number) or
 *   arrays of Ids (multi-value), stored under "<InternalName>Id" — matching how
 *   SharePoint REST expects them to be WRITTEN back on import.
 *   These IDs are tenant/list-specific and generally will NOT resolve to the same
 *   underlying record on a different tenant — you'll need to remap or handle those
 *   columns manually before/after import.
 * - Attachments are NOT included (SharePoint REST doesn't return file bytes in item queries).
 * - Requires read access to the lists on the source site.
 */
(async function () {
  // <<< EDIT THIS: exact list Titles as they appear in SharePoint >>>
  const LIST_NAMES = ["List1", "List2"];

  const webUrl = window._spPageContextInfo
    ? _spPageContextInfo.webAbsoluteUrl
    : window.location.origin;

  async function getJson(url) {
    const res = await fetch(url, {
      headers: { Accept: "application/json;odata=verbose" },
      credentials: "same-origin",
    });
    if (!res.ok) {
      throw new Error(`Request failed (${res.status}): ${url}`);
    }
    return res.json();
  }

  // Pull only custom (non-system) fields for a list
  async function getCustomFields(listTitle) {
    const url =
      `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields` +
      `?$filter=Hidden eq false and ReadOnlyField eq false and CanBeDeleted eq true`;
    const data = await getJson(url);
    return data.d.results.map((f) => ({
      InternalName: f.InternalName,
      Title: f.Title,
      TypeAsString: f.TypeAsString,
      Required: f.Required,
    }));
  }

  // Field types that need $expand rather than a plain $select, and that need
  // their value flattened down to a raw Id (or array of Ids) after fetching.
  const LOOKUP_TYPES = ["Lookup", "LookupMulti", "User", "UserMulti"];

  // Page through all items. Lookup/Person fields are selected via $expand
  // + "/Id" instead of their bare internal name — selecting a lookup field's
  // internal name directly can throw "$select ... must contain <field>" errors,
  // especially when the internal name itself happens to end in "Id".
  async function getAllItems(listTitle, fields) {
    const selectParts = [];
    const expandParts = [];

    for (const f of fields) {
      if (LOOKUP_TYPES.includes(f.TypeAsString)) {
        selectParts.push(`${f.InternalName}/Id`);
        expandParts.push(f.InternalName);
      } else {
        selectParts.push(f.InternalName);
      }
    }

    const selectQS = selectParts.join(",");
    const expandQS = [...new Set(expandParts)].join(",");

    let items = [];
    let url =
      `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
      `?$select=${selectQS}` +
      (expandQS ? `&$expand=${expandQS}` : "") +
      `&$top=2000`;

    while (url) {
      const data = await getJson(url);
      items = items.concat(data.d.results);
      url = data.d.__next || null;
    }

    // Flatten Lookup/User fields down to a raw Id (single) or array of Ids
    // (multi-value), stored under "<InternalName>Id" — the same property name
    // SharePoint REST expects when you WRITE a lookup/person value back.
    for (const item of items) {
      for (const f of fields) {
        if (!LOOKUP_TYPES.includes(f.TypeAsString)) continue;
        const raw = item[f.InternalName];
        let idVal = null;
        if (raw && Array.isArray(raw.results)) {
          idVal = raw.results.map((r) => r.Id);
        } else if (raw && typeof raw.Id !== "undefined") {
          idVal = raw.Id;
        }
        item[`${f.InternalName}Id`] = idVal;
        delete item[f.InternalName];
      }
    }

    return items;
  }

  const exportPayload = {};

  for (const listTitle of LIST_NAMES) {
    console.log(`Exporting list: ${listTitle} ...`);
    try {
      const fields = await getCustomFields(listTitle);
      const items = await getAllItems(listTitle, fields);
      exportPayload[listTitle] = { fields, items };
      console.log(`  -> ${items.length} items, ${fields.length} custom columns`);
    } catch (err) {
      console.error(`  Failed to export "${listTitle}":`, err);
    }
  }

  const blob = new Blob([JSON.stringify(exportPayload, null, 2)], {
    type: "application/json",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `sharepoint-export-${Date.now()}.json`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  console.log("Export complete. File downloaded.");
  console.log(exportPayload);
})();
