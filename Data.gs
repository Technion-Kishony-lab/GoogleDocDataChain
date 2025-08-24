// Data.gs
// Spreadsheet data operations and management functions

/**
 * Retrieves the list of recently used spreadsheet IDs
 * @function getRecentSheets
 * @returns {Array<string>} Array of spreadsheet IDs
 * @description Gets the list of recently accessed spreadsheets from user properties
 */
function getRecentSheets() {
  try {
    const p = PropertiesService.getUserProperties().getProperty(RECENT_SHEETS_KEY);
    return p ? JSON.parse(p) : [];
  } catch (error) {
    console.error('Failed to get recent sheets from properties:', error.message);
    return [];
  }
}

/**
 * Retrieves names and IDs of recently used spreadsheets
 * @function getRecentSheetNames
 * @returns {Array<{id: string, name: string}>} Array of objects with sheet ID and name
 * @description Gets the list of recently accessed spreadsheets with their names
 * @throws {Error} When spreadsheet access fails
 */
function getRecentSheetNames() {
  try {
    const ids = getRecentSheets();
    if (!ids || ids.length === 0) return [];
    
    // Batch fetch file names to reduce API calls
    const fileInfos = ids.map(id => {
      try {
        const file = DriveApp.getFileById(id);
        return { id, name: file.getName() };
      } catch (e) {
        console.error(`Failed to get name for sheet ${id}:`, e.message);
        return { id, name: id };
      }
    });
    
    return fileInfos;
  } catch (error) {
    console.error('Failed to get recent sheet names:', error.message);
    return [];
  }
}

/**
 * Updates the list of recently used spreadsheets
 * @function updateRecentSheets
 * @param {string} id - The spreadsheet ID to add to recent list
 * @description Adds a spreadsheet ID to the beginning of the recent list (max 5 items)
 */
function updateRecentSheets(id) {
  try {
    if (!id || typeof id !== 'string') {
      console.error('Invalid sheet ID provided to updateRecentSheets');
      return;
    }
    
    const prop = PropertiesService.getUserProperties();
    const arr = getRecentSheets().filter(x => x !== id);
    arr.unshift(id);
    prop.setProperty(RECENT_SHEETS_KEY, JSON.stringify(arr.slice(0, 5)));
  } catch (error) {
    console.error('Failed to update recent sheets:', error.message);
  }
}

/**
 * Retrieves the list of sheet tabs from a spreadsheet
 * @function getTabs
 * @param {string} sheetId - The Google Sheets spreadsheet ID
 * @returns {Array<string>} Array of sheet tab names
 * @description Gets all sheet names from a Google Sheets spreadsheet
 * @throws {Error} When spreadsheet access fails
 */
function getTabs(sheetId) {
  try {
    if (!sheetId || typeof sheetId !== 'string') {
      throw new Error('Invalid sheet ID provided');
    }
    
    // Check cache first
    const cacheKey = `tabs_${sheetId}`;
    const cachedTabs = getCachedData(cacheKey);
    if (cachedTabs) {
      return cachedTabs;
    }
    
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      throw new Error('Could not open spreadsheet');
    }
    
    // Get all sheets at once to reduce API calls
    const sheets = ss.getSheets();
    const sheetNames = [];
    
    // Use a more efficient loop
    for (let i = 0; i < sheets.length; i++) {
      sheetNames.push(sheets[i].getName());
    }
    
    // Cache the result
    setCachedData(cacheKey, sheetNames);
    
    return sheetNames;
  } catch (error) {
    console.error(`Failed to get tabs for sheet ${sheetId}:`, error.message);
    throw new Error(`Failed to access spreadsheet: ${error.message}`);
  }
}

/**
 * Retrieves field names from a specific sheet tab
 * @function getFields
 * @param {string} sheetId - The Google Sheets spreadsheet ID
 * @param {string} sheetName - The name of the sheet tab
 * @returns {Array<string>} Array of field names from column B
 * @description Gets field names from the second column (B) of a sheet, starting from row 2
 * @throws {Error} When spreadsheet or sheet access fails
 */
function getFields(sheetId, sheetName) {
  try {
    if (!sheetId || typeof sheetId !== 'string') {
      throw new Error('Invalid sheet ID provided');
    }
    
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('Invalid sheet name provided');
    }
    
    // Check cache first
    const cacheKey = `fields_${sheetId}_${sheetName}`;
    const cachedFields = getCachedData(cacheKey);
    if (cachedFields) {
      return cachedFields;
    }
    
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      throw new Error('Could not open spreadsheet');
    }
    
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    const lr = sh.getLastRow();
    if (lr < 2) {
      return [];
    }
    
    // Get all values in one API call and process efficiently
    const values = sh.getRange(2, 2, lr - 1, 1).getValues();
    const fields = [];
    
    // Use a more efficient loop instead of map
    for (let i = 0; i < values.length; i++) {
      const value = values[i][0];
      if (value !== null && value !== undefined) {
        fields.push(value.toString());
      }
    }
    
    // Cache the result
    setCachedData(cacheKey, fields);
    
    return fields;
  } catch (error) {
    console.error(`Failed to get fields for sheet ${sheetId}, tab ${sheetName}:`, error.message);
    throw new Error(`Failed to get fields: ${error.message}`);
  }
}

/**
 * Inserts a value from a spreadsheet into the document at cursor position
 * @function insertFromSidebar
 * @param {string} sheetId - The Google Sheets spreadsheet ID
 * @param {string} sheetName - The name of the sheet tab
 * @param {string} fieldName - The field name to insert
 * @returns {Object} Result object with success/error status
 * @description Retrieves a value from a spreadsheet and inserts it as linked text in the document
 * @throws {Error} When spreadsheet access or document insertion fails
 */
function insertFromSidebar(sheetId, sheetName, fieldName) {
  try {
    // Validate inputs
    if (!sheetId || typeof sheetId !== 'string') {
      return { error: 'Invalid sheet ID provided.' };
    }
    
    if (!sheetName || typeof sheetName !== 'string') {
      return { error: 'Invalid sheet name provided.' };
    }
    
    if (!fieldName || typeof fieldName !== 'string') {
      return { error: 'Invalid field name provided.' };
    }

    // Open spreadsheet
    const ss = SpreadsheetApp.openById(sheetId);
    if (!ss) {
      return { error: 'Could not open spreadsheet. Check permissions and sheet ID.' };
    }

    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      return { error: `Sheet "${sheetName}" not found in spreadsheet.` };
    }

    const lr = sh.getLastRow();
    if (lr < 2) {
      return { error: 'Sheet has no data. Expected data starting from row 2.' };
    }

    // Get data from columns B (field) + C (value)
    const data = sh.getRange(2, 2, lr - 1, 2).getValues();
    let rowIndex = 1;
    const found = data.find((row, i) => {
      rowIndex = i + 2; // actual sheet row
      return row[0].toString() === fieldName;
    });

    if (!found) {
      return { error: `No value found for "${fieldName}" in sheet.` };
    }

    const rawValue = found[1].toString();
    const value = formatScientificNotation(rawValue);

    // Check cursor position
    const cursor = DocumentApp.getActiveDocument().getCursor();
    if (!cursor) {
      return { error: 'Place the cursor first.' };
    }

    const element = cursor.getElement();
    if (!element) {
      return { error: 'Could not get document element at cursor position.' };
    }

    const offset = cursor.getOffset();
    
    // Insert the text
    const insertedLength = insertScientificText(element, value, offset, offset);

    // Build link back to the exact cell (col C)
    const col = columnToLetter(3); // C
    const sheetGid = sh.getSheetId();
    const base = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
    const url = base
      + `?gid=${sheetGid}`
      + `&range=${col}${rowIndex}`
      + `&${LABEL_PARAM}=${LABEL_VALUE}`
      + `&sheet=${encodeURIComponent(sheetName)}`
      + `&field=${encodeURIComponent(fieldName)}`;

    // Apply formatting
    const txt = element.asText();
    txt.setLinkUrl(offset, offset + insertedLength - 1, url)
       .setForegroundColor(offset, offset + insertedLength - 1, getColor())
       .setUnderline(offset, offset + insertedLength - 1, false);

    updateRecentSheets(sheetId);
    return { success: true };
  } catch (error) {
    console.error('Error in insertFromSidebar:', error.message);
    return { error: `Insert failed: ${error.message}` };
  }
}

/**
 * Clears the list of recently used spreadsheets
 * @function clearRecentSheets
 * @description Removes all recently used spreadsheet IDs from user properties and clears cache
 * @fires clearAllCache
 * @fires DocumentApp.getUi().alert
 */
function clearRecentSheets() {
  try {
    PropertiesService.getUserProperties().deleteProperty(RECENT_SHEETS_KEY);
    // Also clear cache for better performance
    clearAllCache();
    const ui = DocumentApp.getUi();
    ui.alert('Recent sheets and cache cleared successfully!');
  } catch (error) {
    console.error('Failed to clear recent sheets:', error.message);
    const ui = DocumentApp.getUi();
    ui.alert('Failed to clear recent sheets: ' + error.message);
  }
}
