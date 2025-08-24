// Utils.gs

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
