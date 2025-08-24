// Utils.gs
// Utility functions for data formatting and manipulation

/**
 * Formats numbers in scientific notation to display format
 * @function formatScientificNotation
 * @param {*} value - The value to format (number, string, etc.)
 * @returns {string} Formatted string representation
 * @description Converts scientific notation (e.g., 2.3e7) to display format (2.3·10^7)
 */
function formatScientificNotation(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }
    
    // Early return for non-scientific notation
    const str = value.toString();
    if (str.indexOf('e') === -1 && str.indexOf('E') === -1) {
      return str;
    }
    
    const match = str.match(/^(-?\d*\.?\d+)[eE]([+-]?\d+)$/);
    if (!match) return str;
    
    const base = match[1];
    const exp = match[2];
    
    // If base is 1, just return 10^exp
    if (base === '1') {
      return `10^${exp}`;
    }
    
    return `${base}·10^${exp}`;
  } catch (error) {
    console.error('Error formatting scientific notation:', error.message);
    return value ? value.toString() : '';
  }
}

/**
 * Inserts text with scientific notation formatting into a document element
 * @function insertScientificText
 * @param {DocumentApp.Element} element - The document element to insert text into
 * @param {string} text - The text to insert
 * @param {number} start - The starting position for insertion
 * @param {number} end - The ending position (unused, kept for compatibility)
 * @returns {number} The length of the inserted text
 * @description Inserts text and applies superscript formatting for scientific notation
 * @throws {Error} When element or text parameters are invalid
 */
function insertScientificText(element, text, start, end) {
  try {
    if (!element) {
      throw new Error('Element is required');
    }
    
    if (!text || typeof text !== 'string') {
      throw new Error('Text must be a non-empty string');
    }
    
    if (typeof start !== 'number' || start < 0) {
      throw new Error('Start position must be a non-negative number');
    }
    
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
  } catch (error) {
    console.error('Error inserting scientific text:', error.message);
    // Fallback to simple text insertion
    if (element && text) {
      element.insertText(start, text);
      return text.length;
    }
    return 0;
  }
}

/**
 * Extracts spreadsheet ID from a Google Sheets URL
 * @function extractSheetId
 * @param {string} url - The Google Sheets URL
 * @returns {string|null} The spreadsheet ID or null if not found
 * @description Parses a Google Sheets URL to extract the spreadsheet ID
 */
function extractSheetId(url) {
  try {
    if (!url || typeof url !== 'string') {
      return null;
    }
    
    const m = url.match(/\/d\/([A-Za-z0-9_-]+)/);
    return m ? m[1] : null;
  } catch (error) {
    console.error('Error extracting sheet ID from URL:', error.message);
    return null;
  }
}

/**
 * Extracts a parameter value from a URL query string
 * @function getParam
 * @param {string} url - The URL containing query parameters
 * @param {string} name - The parameter name to extract
 * @returns {string|null} The parameter value or null if not found
 * @description Parses URL query parameters using regex
 */
function getParam(url, name) {
  try {
    if (!url || typeof url !== 'string') {
      return null;
    }
    
    if (!name || typeof name !== 'string') {
      return null;
    }
    
    const m = url.match(new RegExp('[?&]' + name + '=([^&#]+)'));
    return m ? decodeURIComponent(m[1]) : null;
  } catch (error) {
    console.error(`Error getting parameter '${name}' from URL:`, error.message);
    return null;
  }
}

/**
 * Converts a column number to Excel-style column letter(s)
 * @function columnToLetter
 * @param {number} col - The column number (1-based)
 * @returns {string} The column letter(s) (A, B, C, ..., Z, AA, AB, etc.)
 * @description Converts numeric column indices to letter-based column references
 * @throws {Error} When column number is invalid
 */
function columnToLetter(col) {
  try {
    if (!col || typeof col !== 'number' || col < 1) {
      throw new Error('Column number must be a positive integer');
    }
    
    // Cache for common column numbers (A-Z, AA-ZZ)
    if (col <= 26) {
      return String.fromCharCode(64 + col);
    }
    if (col <= 702) { // AA-ZZ
      const first = Math.floor((col - 1) / 26);
      const second = col % 26 || 26;
      return String.fromCharCode(64 + first) + String.fromCharCode(64 + second);
    }
    
    // General case for larger columns
    let s = '';
    let column = col;
    while (column > 0) {
      const m = (column - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      column = Math.floor((column - 1) / 26);
    }
    return s;
  } catch (error) {
    console.error('Error converting column to letter:', error.message);
    return 'A';
  }
}
