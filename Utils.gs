// Utils.gs

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
