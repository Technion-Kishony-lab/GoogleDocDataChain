// Code.gs
// Main entry point for Data Chain Google Docs add-on
// This file contains the essential entry code and delegates to modular functions

/**
 * Entry point - called when the document opens
 * Creates the Data-Chain menu in Google Docs
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Data‑Chain')
    .addItem('Insert value', 'showSidebar')
    .addItem('Update document', 'replacePlaceholders')
    .addSeparator()
    .addItem('Set color', 'setColor')
    .addItem('Clear recent sheets', 'clearRecentSheets')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Main function to replace all placeholders in the document
 * Delegates to Document.gs module
 */
function replacePlaceholders() {
  const cache = {};
  const runsByElement = new Map();
  collectRuns(DocumentApp.getActiveDocument().getBody(), runsByElement);
  runsByElement.forEach((runs, txt) => {
    runs.sort((a, b) => b.start - a.start);
    runs.forEach(r => replaceRun(txt, r, cache));
  });
}

/**
 * Helper function to include HTML files
 * Used by UI.gs module
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
