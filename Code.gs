// Code.gs
// Main entry point for Data Chain Google Docs add-on
// This file contains the essential entry code and delegates to modular functions

/**
 * Entry point - called when the document opens
 * Creates the Data-Chain menu in Google Docs
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
