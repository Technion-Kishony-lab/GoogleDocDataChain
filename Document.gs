// Document.gs

function replacePlaceholders() {
  const timer = startTimer('replacePlaceholders');
  
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      throw new Error('No active document found');
    }
    
    const body = doc.getBody();
    if (!body) {
      throw new Error('Document body not found');
    }
    
    const cache = {};
    const runsByElement = new Map();
    
    // Collect all runs first
    const collectTimer = startTimer('collectRuns');
    collectRuns(body, runsByElement);
    endTimer(collectTimer);
    
    if (runsByElement.size === 0) {
      console.log('No placeholders found to update');
      endTimer(timer);
      return;
    }
    
    let updatedCount = 0;
    const totalRuns = Array.from(runsByElement.values()).reduce((sum, runs) => sum + runs.length, 0);
    
    // Process runs in batches for better performance
    const batchSize = 10;
    let processedCount = 0;
    
    const processTimer = startTimer('processRuns');
    runsByElement.forEach((runs, txt) => {
      // Sort runs in descending order to avoid position shifts
      runs.sort((a, b) => b.start - a.start);
      
      // Process runs in batches
      for (let i = 0; i < runs.length; i += batchSize) {
        const batch = runs.slice(i, i + batchSize);
        
        batch.forEach(r => {
          try {
            replaceRun(txt, r, cache);
            updatedCount++;
          } catch (error) {
            console.error('Error replacing run:', error.message);
          }
        });
        
        processedCount += batch.length;
        
        // Log progress for large documents
        if (totalRuns > 50 && processedCount % 20 === 0) {
          console.log(`Progress: ${processedCount}/${totalRuns} runs processed`);
        }
      }
    });
    endTimer(processTimer);
    
    console.log(`Updated ${updatedCount} placeholders`);
    endTimer(timer);
  } catch (error) {
    console.error('Error in replacePlaceholders:', error.message);
    endTimer(timer);
    throw new Error(`Failed to update document: ${error.message}`);
  }
}

function collectRuns(el, runsByElement) {
  try {
    if (!el) {
      return;
    }
    
    if (el.getType() === DocumentApp.ElementType.TEXT) {
      const txt = el.asText();
      if (!txt) return;
      
      const text = txt.getText();
      const runs = [];
      
      for (let i = 0; i < text.length; ) {
        try {
          const url = txt.getLinkUrl(i);
          if (!url || getParam(url, LABEL_PARAM) !== LABEL_VALUE) { 
            i++; 
            continue; 
          }
          
          let start = i, end = i;
          while (end + 1 < text.length && txt.getLinkUrl(end + 1) === url) end++;
          runs.push({ start, end, url });
          i = end + 1;
        } catch (error) {
          console.error(`Error processing text at position ${i}:`, error.message);
          i++;
        }
      }
      
      if (runs.length) runsByElement.set(txt, runs);
    } else if (el.getNumChildren) {
      const numChildren = el.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        try {
          const child = el.getChild(i);
          if (child) {
            collectRuns(child, runsByElement);
          }
        } catch (error) {
          console.error(`Error processing child ${i}:`, error.message);
        }
      }
    }
  } catch (error) {
    console.error('Error in collectRuns:', error.message);
  }
}

function replaceRun(txt, r, cache) {
  try {
    if (!txt || !r || !r.url) {
      console.error('Invalid parameters for replaceRun');
      return;
    }
    
    const field = getParam(r.url, 'field');
    const sheetId = extractSheetId(r.url);
    const sheetName = getParam(r.url, 'sheet');
    
    if (!field || !sheetId || !sheetName) {
      console.error('Missing required parameters for replaceRun');
      return;
    }

    const cacheKey = `${sheetId}::${sheetName}`;

    // Use global cache for spreadsheet data
    if (!cache[cacheKey]) {
      let data = {}, sheetGid = null;
      try {
        const ss = SpreadsheetApp.openById(sheetId);
        if (ss) {
          const sh = ss.getSheetByName(sheetName);
          if (sh) {
            const lr = sh.getLastRow();
            if (lr >= 2) {
              // Get all data in one API call
              const vals = sh.getRange(2, 2, lr - 1, 2).getValues(); // B:C
              
              // Use efficient loop instead of forEach
              for (let i = 0; i < vals.length; i++) {
                const row = vals[i];
                const k = String(row[0] ?? '').trim();
                if (k) {
                  data[k] = { value: String(row[1] ?? ''), row: i + 2 };
                }
              }
            }
            sheetGid = sh.getSheetId();
          }
        }
      } catch (e) {
        console.error(`Failed to load data for ${cacheKey}:`, e.message);
        // data stays empty
      }
      cache[cacheKey] = { data, sheetGid };
    }

    const entry = cache[cacheKey].data[field];
    const rawVal = entry ? entry.value : '<NOT FOUND>';
    const val = formatScientificNotation(rawVal);
    const row = entry ? entry.row : 1;
    const col = columnToLetter(3);

    // Validate text range before deletion
    const textLength = txt.getText().length;
    if (r.start < 0 || r.end >= textLength || r.start > r.end) {
      console.error(`Invalid text range: start=${r.start}, end=${r.end}, textLength=${textLength}`);
      return;
    }

    // Perform all text operations in sequence to minimize API calls
    txt.deleteText(r.start, r.end);
    const insertedLength = insertScientificText(txt, val, r.start, r.start);
    const start = r.start, end2 = start + insertedLength - 1;

    const base = r.url.split('?')[0];
    const url =
      cache[cacheKey].sheetGid
        ? `${base}?gid=${cache[cacheKey].sheetGid}&range=${col}${row}&${LABEL_PARAM}=${LABEL_VALUE}&sheet=${encodeURIComponent(sheetName)}&field=${encodeURIComponent(field)}`
        : r.url; // keep original link if we couldn't resolve the tab

    // Apply all formatting in one chain to reduce API calls
    txt.setLinkUrl(start, end2, url)
       .setForegroundColor(start, end2, getColor())
       .setUnderline(start, end2, false);
  } catch (error) {
    console.error('Error in replaceRun:', error.message);
  }
}
