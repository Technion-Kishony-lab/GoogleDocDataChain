// UI.gs

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

function getColor() {
  try {
    const color = PropertiesService.getUserProperties().getProperty(TEXT_COLOR_KEY);
    return color || '#000000';
  } catch (error) {
    console.error('Failed to get color from properties:', error.message);
    return '#000000';
  }
}

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
