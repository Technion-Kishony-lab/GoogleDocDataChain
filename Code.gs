// Code.gs
// Main entry point for Data Chain Google Docs add-on
// This file contains the essential entry code and delegates to modular functions

/**
 * Entry point - called when the document opens
 * Creates the Data-Chain menu in Google Docs
 * @function onOpen
 * @description Initializes the Data-Chain add-on by creating the menu in Google Docs
 * @fires DocumentApp.getUi().createMenu
 */
function onOpen() {
  try {
    const ui = DocumentApp.getUi();
    if (!ui) {
      console.error('Could not get UI instance');
      return;
    }
    
    ui.createMenu('Data‑Chain')
      .addItem('Insert value', 'showSidebar')
      .addItem('Update document', 'replacePlaceholders')
      .addSeparator()
      .addItem('Set color', 'setColor')
      .addItem('Clear recent sheets', 'clearRecentSheets')
      .addSeparator()
      .addItem('Help', 'showHelp')
      .addToUi();
  } catch (error) {
    console.error('Error in onOpen:', error.message);
  }
}

/**
 * Main function to replace all placeholders in the document
 * Delegates to Document.gs module
 * @function replacePlaceholders
 * @description Updates all linked values in the document with current spreadsheet data
 * @fires collectRuns
 * @fires replaceRun
 * @throws {Error} When document operations fail
 */
function replacePlaceholders() {
  try {
    const cache = {};
    const runsByElement = new Map();
    collectRuns(DocumentApp.getActiveDocument().getBody(), runsByElement);
    runsByElement.forEach((runs, txt) => {
      runs.sort((a, b) => b.start - a.start);
      runs.forEach(r => replaceRun(txt, r, cache));
    });
  } catch (error) {
    console.error('Error in replacePlaceholders:', error.message);
    const ui = DocumentApp.getUi();
    ui.alert(`Failed to update document: ${error.message}`);
  }
}

/**
 * Helper function to include HTML files
 * Used by UI.gs module
 * @function include
 * @param {string} filename - The name of the HTML file to include
 * @returns {string} The HTML content of the file
 * @description Loads and returns the content of an HTML file for use in templates
 * @throws {Error} When file cannot be loaded
 */
function include(filename) {
  try {
    if (!filename || typeof filename !== 'string') {
      throw new Error('Invalid filename provided');
    }
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error(`Failed to include file ${filename}:`, error.message);
    return `<!-- Error loading ${filename}: ${error.message} -->`;
  }
}
