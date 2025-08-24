// Comprehensive test for Data Chain Google Apps Script add-on
const fs = require('fs');

// Enhanced mock objects
global.DocumentApp = {
  getUi: () => ({
    createMenu: (name) => {
      const menu = {
        addItem: (label, func) => menu,
        addSeparator: () => menu,
        addToUi: () => {}
      };
      return menu;
    }
  }),
  getActiveDocument: () => ({
    getBody: () => ({
      getParagraphs: () => [],
      getTables: () => []
    }),
    getCursor: () => ({
      getElement: () => ({
        insertText: (start, text) => {},
        asText: () => ({
          setLinkUrl: (start, end, url) => ({ setForegroundColor: () => ({ setUnderline: () => {} }) })
        }),
        setTextAlignment: (start, end, alignment) => {}
      }),
      getOffset: () => 0
    })
  }),
  TextAlignment: {
    SUPERSCRIPT: 'SUPERSCRIPT'
  }
};

global.HtmlService = {
  createHtmlOutputFromFile: (filename) => ({
    getContent: () => `<html>Mock content for ${filename}</html>`
  })
};

global.PropertiesService = {
  getUserProperties: () => ({
    getProperty: (key) => null,
    setProperty: (key, value) => {},
    deleteProperty: (key) => {}
  })
};

global.SpreadsheetApp = {
  openById: (id) => ({
    getSheets: () => ([
      { getName: () => 'Sheet1' },
      { getName: () => 'Sheet2' }
    ]),
    getSheetByName: (name) => ({
      getLastRow: () => 5,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => [
          ['Field1', 'Value1'],
          ['Field2', '2.3e7'],
          ['Field3', '1.5e-3']
        ]
      }),
      getSheetId: () => '12345'
    })
  })
};

global.DriveApp = {
  getFileById: (id) => ({
    getName: () => `Sheet_${id}`
  })
};

// Load modules
function loadModule(filename) {
  const content = fs.readFileSync(filename, 'utf8');
  const nodeContent = content
    .replace(/function\s+(\w+)/g, 'global.$1 = function')
    .replace(/const\s+/g, 'let ')
    .replace(/let\s+/g, 'let ');
  
  eval(nodeContent);
}

loadModule('Constants.gs');
loadModule('Utils.gs');
loadModule('Data.gs');
loadModule('Document.gs');
loadModule('UI.gs');
loadModule('Code.gs');

// Test functions
console.log('🧪 Testing Data Chain add-on behaviors...\n');

// Test 1: Scientific notation formatting
console.log('1. Testing scientific notation formatting:');
const testValues = [
  '2.3e7',
  '1.5e-3', 
  '1e5',
  'normal text',
  '2.3E7'
];

testValues.forEach(value => {
  const formatted = global.formatScientificNotation(value);
  console.log(`   ${value} → ${formatted}`);
});

// Test 2: URL parsing utilities
console.log('\n2. Testing URL parsing:');
const testUrls = [
  'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit',
  'https://docs.google.com/spreadsheets/d/abc123/edit?gid=0&range=A1'
];

testUrls.forEach(url => {
  const sheetId = global.extractSheetId(url);
  const gid = global.getParam(url, 'gid');
  console.log(`   Sheet ID: ${sheetId}, GID: ${gid}`);
});

// Test 3: Column to letter conversion
console.log('\n3. Testing column to letter conversion:');
[1, 2, 3, 26, 27, 28, 702, 703].forEach(col => {
  const letter = global.columnToLetter(col);
  console.log(`   Column ${col} → ${letter}`);
});

// Test 4: Recent sheets functionality
console.log('\n4. Testing recent sheets:');
try {
  const recent = global.getRecentSheets();
  console.log(`   Recent sheets: ${JSON.stringify(recent)}`);
  
  const names = global.getRecentSheetNames();
  console.log(`   Sheet names: ${JSON.stringify(names)}`);
} catch (error) {
  console.log(`   Error: ${error.message}`);
}

// Test 5: Sheet data retrieval
console.log('\n5. Testing sheet data retrieval:');
try {
  const tabs = global.getTabs('test-sheet-id');
  console.log(`   Tabs: ${JSON.stringify(tabs)}`);
  
  const fields = global.getFields('test-sheet-id', 'Sheet1');
  console.log(`   Fields: ${JSON.stringify(fields)}`);
} catch (error) {
  console.log(`   Error: ${error.message}`);
}

// Test 6: Core functions
console.log('\n6. Testing core functions:');
try {
  global.onOpen();
  console.log('   ✅ onOpen works');
  
  const includeResult = global.include('Sidebar.html');
  console.log('   ✅ include works');
  
  global.clearRecentSheets();
  console.log('   ✅ clearRecentSheets works');
} catch (error) {
  console.log(`   ❌ Error: ${error.message}`);
}

console.log('\n✅ All local behaviors tested!');
