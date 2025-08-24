// UI.gs

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

function setColor() {
  const ui = DocumentApp.getUi();
  const currentColor = getColor();
  const resp = ui.prompt(
    'Set color for inserted values',
    `Enter hex color code (Current: ${currentColor})`,
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;
  let c = resp.getResponseText().trim();
  if (!/^#?[0-9A-Fa-f]{6}$/.test(c)) {
    ui.alert('Invalid format. Must be 6‑digit hex, with or without "#".');
    return;
  }
  if (!c.startsWith('#')) c = '#' + c;
  PropertiesService.getUserProperties().setProperty(TEXT_COLOR_KEY, c);
}

function getColor() {
  return PropertiesService.getUserProperties().getProperty(TEXT_COLOR_KEY) || '#000000';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile("Sidebar");
  const html = template.evaluate().setTitle("Data Chain");
  DocumentApp.getUi().showSidebar(html);
}

function showHelp() {
  const template = HtmlService.createTemplateFromFile("Help");
  const html = template.evaluate().setTitle("Data Chain - Help");
  DocumentApp.getUi().showSidebar(html);
}
