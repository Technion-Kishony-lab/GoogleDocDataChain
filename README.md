# Data Chain

A Google Apps Script add-on for Google Docs that creates dynamic links between documents and spreadsheets.

## What it does

- **Insert values** from Google Sheets directly into Google Docs
- **Auto-update** document placeholders when spreadsheet data changes
- **Scientific notation** formatting (e.g., `2.3·10^7`)
- **Clickable links** back to source spreadsheet cells
- **Recent sheets** tracking for quick access

## How it works

1. **Insert**: Select a spreadsheet, tab, and field to insert a linked value
2. **Update**: Click "Update Document" to refresh all placeholders with latest data
3. **Link**: Click any inserted value to jump to its source cell in the spreadsheet

## Setup

1. Open Google Docs
2. Go to Extensions > Apps Script
3. Copy `Code.gs` content to the script editor
4. Create HTML files for `Sidebar.html`, `Help.html`, and `logo.html`
5. Deploy as a Google Docs add-on

## Usage

- **Insert value**: Opens sidebar to select spreadsheet data
- **Update document**: Refreshes all linked values
- **Set color**: Customize text color for inserted values
- **Help**: Shows usage instructions
