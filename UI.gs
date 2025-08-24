// UI.gs
// User interface functions for the Data Chain add-on

/**
 * Creates the Data-Chain menu in Google Docs
 * @function onOpen
 * @description Entry point that initializes the add-on menu when document opens
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
      .addItem('Performance stats', 'showPerformanceStats')
      .addItem('Help', 'showHelp')
      .addToUi();
  } catch (error) {
    console.error('Error in onOpen:', error.message);
  }
}

/**
 * Opens a dialog to set the color for inserted values
 * @function setColor
 * @description Prompts user for hex color code and validates input
 * @fires DocumentApp.getUi().prompt
 * @fires DocumentApp.getUi().alert
 * @fires PropertiesService.getUserProperties().setProperty
 */
function setColor() {
  try {
    const ui = DocumentApp.getUi();
    const currentColor = getColor();
    const resp = ui.prompt(
      'Set color for inserted values',
      `Enter hex color code (Current: ${currentColor})`,
      ui.ButtonSet.OK_CANCEL
    );

    if (resp.getSelectedButton() !== ui.Button.OK) return;
    
    let c = resp.getResponseText().trim();
    if (!c) {
      ui.alert('Please enter a color code.');
      return;
    }
    
    if (!/^#?[0-9A-Fa-f]{6}$/.test(c)) {
      ui.alert('Invalid format. Must be 6‑digit hex, with or without "#".');
      return;
    }
    
    if (!c.startsWith('#')) c = '#' + c;
    
    PropertiesService.getUserProperties().setProperty(TEXT_COLOR_KEY, c);
    ui.alert('Color updated successfully!');
  } catch (error) {
    console.error('Error in setColor:', error.message);
    const ui = DocumentApp.getUi();
    ui.alert(`Failed to set color: ${error.message}`);
  }
}

/**
 * Retrieves the current text color setting
 * @function getColor
 * @returns {string} Hex color code (defaults to '#000000' if not set)
 * @description Gets the user's preferred color for inserted values from properties
 */
function getColor() {
  try {
    const color = PropertiesService.getUserProperties().getProperty(TEXT_COLOR_KEY);
    return color || '#000000';
  } catch (error) {
    console.error('Failed to get color from properties:', error.message);
    return '#000000';
  }
}

/**
 * Includes HTML file content for use in templates
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

/**
 * Shows the main Data Chain sidebar
 * @function showSidebar
 * @description Opens the sidebar interface for inserting values from spreadsheets
 * @fires HtmlService.createTemplateFromFile
 * @fires DocumentApp.getUi().showSidebar
 * @throws {Error} When sidebar template cannot be loaded
 */
function showSidebar() {
  try {
    const template = HtmlService.createTemplateFromFile("Sidebar");
    if (!template) {
      throw new Error('Could not load Sidebar template');
    }
    const html = template.evaluate().setTitle("Data Chain");
    DocumentApp.getUi().showSidebar(html);
  } catch (error) {
    console.error('Error showing sidebar:', error.message);
    const ui = DocumentApp.getUi();
    ui.alert(`Failed to open sidebar: ${error.message}`);
  }
}

/**
 * Shows the help documentation sidebar
 * @function showHelp
 * @description Opens the help sidebar with usage instructions
 * @fires HtmlService.createTemplateFromFile
 * @fires DocumentApp.getUi().showSidebar
 * @throws {Error} When help template cannot be loaded
 */
function showHelp() {
  try {
    const template = HtmlService.createTemplateFromFile("Help");
    if (!template) {
      throw new Error('Could not load Help template');
    }
    const html = template.evaluate().setTitle("Data Chain - Help");
    DocumentApp.getUi().showSidebar(html);
  } catch (error) {
    console.error('Error showing help:', error.message);
    const ui = DocumentApp.getUi();
    ui.alert(`Failed to open help: ${error.message}`);
  }
}

/**
 * Shows performance statistics in a dialog
 * @function showPerformanceStats
 * @description Displays performance metrics for various operations
 * @fires getPerformanceStats
 * @fires DocumentApp.getUi().alert
 * @throws {Error} When performance stats cannot be retrieved
 */
function showPerformanceStats() {
  try {
    const stats = getPerformanceStats();
    const ui = DocumentApp.getUi();
    
    if (stats.message) {
      ui.alert('Performance Stats', stats.message, ui.ButtonSet.OK);
      return;
    }
    
    let message = 'Performance Statistics:\n\n';
    Object.keys(stats).forEach(operation => {
      const stat = stats[operation];
      message += `${operation}:\n`;
      message += `  Count: ${stat.count}\n`;
      message += `  Avg Time: ${Math.round(stat.avgTime)}ms\n`;
      message += `  Min Time: ${stat.minTime}ms\n`;
      message += `  Max Time: ${stat.maxTime}ms\n\n`;
    });
    
    ui.alert('Performance Stats', message, ui.ButtonSet.OK);
  } catch (error) {
    console.error('Error showing performance stats:', error.message);
    const ui = DocumentApp.getUi();
    ui.alert('Failed to show performance stats: ' + error.message);
  }
}
