// Test file for Data Chain Google Apps Script add-on
// Simulates Google Apps Script environment for testing

// Mock Google Apps Script objects
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
    })
  })
};

global.HtmlService = {
  createHtmlOutputFromFile: (filename) => ({
    getContent: () => `<html>Mock content for ${filename}</html>`
  })
};

// Import the modules
const fs = require('fs');

// Read and evaluate the GS files
function loadModule(filename) {
  const content = fs.readFileSync(filename, 'utf8');
  // Remove Google Apps Script specific code and convert to Node.js compatible
  const nodeContent = content
    .replace(/function\s+(\w+)/g, 'global.$1 = function')
    .replace(/const\s+/g, 'let ')
    .replace(/let\s+/g, 'let ');
  
  eval(nodeContent);
}

// Load all modules
loadModule('Code.gs');
loadModule('Constants.gs');
loadModule('Utils.gs');
loadModule('Data.gs');
loadModule('Document.gs');
loadModule('UI.gs');

// Test functions
console.log('Testing Data Chain add-on...');

// Test onOpen function
try {
  global.onOpen();
  console.log('✅ onOpen function works');
} catch (error) {
  console.log('❌ onOpen function failed:', error.message);
}

// Test include function
try {
  const result = global.include('Sidebar.html');
  console.log('✅ include function works:', result.substring(0, 50) + '...');
} catch (error) {
  console.log('❌ include function failed:', error.message);
}

console.log('Test completed!');
