// Code.gs

const LABEL_PARAM = 'dataChain';
const LABEL_VALUE = '1';
const RECENT_SHEETS_KEY = 'RECENT_SHEETS';
const TEXT_COLOR_KEY = 'TEXT_COLOR';

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Data‑Chain')
    .addItem('Insert value', 'showSidebar')
    .addItem('Update document', 'replacePlaceholders')
    .addSeparator()
    .addItem('Set color', 'setColor')
    .addItem('Clear recent sheets', 'clearRecentSheets')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

function setColor() {
  const ui = DocumentApp.getUi();
  const currentColor = getColor();
  const resp = ui.prompt(
    'Set color for inserted values',
    `Enter hex color code (Current: ${currentColor})`,
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;
  let c = resp.getResponseText().trim();
  if (!/^#?[0-9A-Fa-f]{6}$/.test(c)) {
    ui.alert('Invalid format. Must be 6‑digit hex, with or without “#”.');
    return;
  }
  if (!c.startsWith('#')) c = '#' + c;
  PropertiesService.getUserProperties().setProperty(TEXT_COLOR_KEY, c);
}

function getColor() {
  return PropertiesService.getUserProperties().getProperty(TEXT_COLOR_KEY) || '#000000';
}

function replacePlaceholders() {
  const cache = {};
  const runsByElement = new Map();
  collectRuns(DocumentApp.getActiveDocument().getBody(), runsByElement);
  runsByElement.forEach((runs, txt) => {
    runs.sort((a, b) => b.start - a.start);
    runs.forEach(r => replaceRun(txt, r, cache));
  });
}

function getRecentSheets() {
  const p = PropertiesService.getUserProperties().getProperty(RECENT_SHEETS_KEY);
  return p ? JSON.parse(p) : [];
}

function getRecentSheetNames() {
  const ids = getRecentSheets();
  return ids.map(id => {
    try {
      const name = DriveApp.getFileById(id).getName();
      return { id, name };
    } catch (e) {
      return { id, name: id };
    }
  });
}

function updateRecentSheets(id) {
  const prop = PropertiesService.getUserProperties();
  const arr = getRecentSheets().filter(x => x !== id);
  arr.unshift(id);
  prop.setProperty(RECENT_SHEETS_KEY, JSON.stringify(arr.slice(0, 5)));
}

function getTabs(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  return ss.getSheets().map(sh => sh.getName());
}

function getFields(sheetId, sheetName) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 2, lr - 1, 1)
           .getValues()
           .map(r => r[0].toString());
}

function formatScientificNotation(value) {
  const str = value.toString();
  const match = str.match(/^(-?\d*\.?\d+)[eE]([+-]?\d+)$/);
  if (!match) return str;
  
  const base = match[1];
  const exp = match[2];
  
  // If base is 1, just return 10^exp
  if (base === '1') {
    return `10^${exp}`;
  }
  
  return `${base}·10^${exp}`;
}

function insertScientificText(element, text, start, end) {
  // Check for format like "10^5" (no base)
  let match = text.match(/^10\^(.+)$/);
  if (match) {
    const exp = match[1];
    const fullText = `10${exp}`;
    element.insertText(start, fullText);
    
    // Make the exponent superscript
    const expStart = start + 2; // "10" = 2 chars
    const expEnd = expStart + exp.length - 1;
    element.setTextAlignment(expStart, expEnd, DocumentApp.TextAlignment.SUPERSCRIPT);
    
    return fullText.length;
  }
  
  // Check for format like "2.3·10^7" (with base)
  match = text.match(/^(.+)·10\^(.+)$/);
  if (!match) {
    element.insertText(start, text);
    return text.length;
  }
  
  const base = match[1];
  const exp = match[2];
  const fullText = `${base}·10${exp}`;
  
  element.insertText(start, fullText);
  
  // Make the exponent superscript
  const expStart = start + base.length + 3; // "·10" = 3 chars
  const expEnd = expStart + exp.length - 1;
  element.setTextAlignment(expStart, expEnd, DocumentApp.TextAlignment.SUPERSCRIPT);
  
  return fullText.length;
}

function insertFromSidebar(sheetId, sheetName, fieldName) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(sheetName);
  const lr = sh.getLastRow();
  if (lr < 2) {
    return { error: 'Sheet has no data.' };
  }
  // grab columns B (field) + C (value)
  const data = sh.getRange(2, 2, lr - 1, 2).getValues();
  let rowIndex = 1;
  const found = data.find((row, i) => {
    rowIndex = i + 2;            // actual sheet row
    return row[0].toString() === fieldName;
  });
  if (!found) {
    return { error: `No value found for “${fieldName}” in sheet.` };
  }
  const rawValue = found[1].toString();
  const value = formatScientificNotation(rawValue);

  const cursor = DocumentApp.getActiveDocument().getCursor();
  if (!cursor) return { error: 'Place the cursor first.' };

  const element = cursor.getElement();
  const offset = cursor.getOffset();
  
  const insertedLength = insertScientificText(element, value, offset, offset);

  // build link back to the exact cell (col C)
  const col = columnToLetter(3);  // C
  const sheetGid = sh.getSheetId();
  const base = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
  const url = base
    + `?gid=${sheetGid}`
    + `&range=${col}${rowIndex}`
    + `&${LABEL_PARAM}=${LABEL_VALUE}`
    + `&sheet=${encodeURIComponent(sheetName)}`
    + `&field=${encodeURIComponent(fieldName)}`;

  const txt = element.asText();
  txt.setLinkUrl(offset, offset + insertedLength - 1, url)
     .setForegroundColor(offset, offset + insertedLength - 1, getColor())
     .setUnderline(offset, offset + insertedLength - 1, false);

  updateRecentSheets(sheetId);
  return {};
}


function extractSheetId(url) {
  const m = url.match(/\/d\/([A-Za-z0-9_-]+)/);
  return m ? m[1] : null;
}

function getParam(url, name) {
  const m = url.match(new RegExp('[?&]' + name + '=([^&#]+)'));
  return m ? decodeURIComponent(m[1]) : null;
}

function columnToLetter(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

function clearRecentSheets() {
  PropertiesService.getUserProperties().deleteProperty(RECENT_SHEETS_KEY);
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





function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile("Sidebar");
  const html = template.evaluate().setTitle("Data Chain");
  DocumentApp.getUi().showSidebar(html);
}

function showHelp() {
  const template = HtmlService.createTemplateFromFile("Help");
  const html = template.evaluate().setTitle("Data Chain - Help");
  DocumentApp.getUi().showSidebar(html);
}
