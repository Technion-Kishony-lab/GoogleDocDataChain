// Document.gs

function replacePlaceholders() {
  const cache = {};
  const runsByElement = new Map();
  collectRuns(DocumentApp.getActiveDocument().getBody(), runsByElement);
  runsByElement.forEach((runs, txt) => {
    runs.sort((a, b) => b.start - a.start);
    runs.forEach(r => replaceRun(txt, r, cache));
  });
}

function collectRuns(el, runsByElement) {
  if (el.getType() === DocumentApp.ElementType.TEXT) {
    const txt = el.asText(), text = txt.getText(), runs = [];
    for (let i = 0; i < text.length; ) {
      const url = txt.getLinkUrl(i);
      if (!url || getParam(url, LABEL_PARAM) !== LABEL_VALUE) { i++; continue; }
      let start = i, end = i;
      while (end + 1 < text.length && txt.getLinkUrl(end + 1) === url) end++;
      runs.push({ start, end, url });
      i = end + 1;
    }
    if (runs.length) runsByElement.set(txt, runs);
  } else if (el.getNumChildren) {
    for (let i = 0; i < el.getNumChildren(); i++) {
      collectRuns(el.getChild(i), runsByElement);
    }
  }
}

function replaceRun(txt, r, cache) {
  const field = getParam(r.url, 'field'),
        sheetId = extractSheetId(r.url),
        sheetName = getParam(r.url, 'sheet');
  if (!field || !sheetId || !sheetName) return;

  const cacheKey = `${sheetId}::${sheetName}`;

  if (!cache[cacheKey]) {
    let data = {}, sheetGid = null;
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      const sh = ss.getSheetByName(sheetName);
      if (sh) {
        const lr = sh.getLastRow();
        if (lr >= 2) {
          const vals = sh.getRange(2, 2, lr - 1, 2).getValues(); // B:C
          vals.forEach((row, i) => {
            const k = String(row[0] ?? '').trim();
            if (k) data[k] = { value: String(row[1] ?? ''), row: i + 2 };
          });
        }
        sheetGid = sh.getSheetId();
      }
    } catch (e) {
      // ignore: file missing / no access → data stays empty
    }
    cache[cacheKey] = { data, sheetGid };
  }

  const entry = cache[cacheKey].data[field];
  const rawVal = entry ? entry.value : '<NOT FOUND>';
  const val = formatScientificNotation(rawVal);
  const row = entry ? entry.row : 1;
  const col = columnToLetter(3);

  txt.deleteText(r.start, r.end);
  const insertedLength = insertScientificText(txt, val, r.start, r.start);
  const start = r.start, end2 = start + insertedLength - 1;

  const base = r.url.split('?')[0];
  const url =
    cache[cacheKey].sheetGid
      ? `${base}?gid=${cache[cacheKey].sheetGid}&range=${col}${row}&${LABEL_PARAM}=${LABEL_VALUE}&sheet=${encodeURIComponent(sheetName)}&field=${encodeURIComponent(field)}`
      : r.url; // keep original link if we couldn't resolve the tab

  txt.setLinkUrl(start, end2, url)
     .setForegroundColor(start, end2, getColor())
     .setUnderline(start, end2, false);
}
