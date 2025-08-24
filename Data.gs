// Data.gs

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
    return { error: `No value found for "${fieldName}" in sheet.` };
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

function clearRecentSheets() {
  PropertiesService.getUserProperties().deleteProperty(RECENT_SHEETS_KEY);
}
