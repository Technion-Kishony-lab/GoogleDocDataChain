# Data Chain

A Google Apps Script add-on for Google Docs that creates dynamic links between documents and spreadsheets.

## What it does

- **Insert values** from Google Sheets directly into Google Docs
- **Auto-update** document placeholders when spreadsheet data changes
- **Scientific notation** formatting (e.g., `2.3·10^7`)
- **Clickable links** back to source spreadsheet cells
- **Recent sheets** tracking for quick access
- **Performance monitoring** and caching for optimal speed
- **Comprehensive error handling** for robust operation

## Features

### Core Functionality
- **Dynamic Linking**: Insert spreadsheet values that update automatically
- **Scientific Notation**: Automatic formatting of scientific notation
- **Clickable Links**: Click any value to jump to its source cell
- **Recent Sheets**: Quick access to recently used spreadsheets

### Performance & Reliability
- **Caching System**: 5-minute cache for frequently accessed data
- **Batch Processing**: Efficient handling of large documents
- **Progress Tracking**: Real-time progress for long operations
- **Error Recovery**: Graceful handling of failures
- **Performance Monitoring**: Built-in timing and statistics

### User Experience
- **Modern UI**: Clean, responsive sidebar interface
- **Color Customization**: Set custom colors for inserted values
- **Help System**: Comprehensive documentation and guidance
- **Performance Stats**: Monitor operation performance

## How it works

1. **Insert**: Select a spreadsheet, tab, and field to insert a linked value
2. **Update**: Click "Update Document" to refresh all placeholders with latest data
3. **Link**: Click any inserted value to jump to its source cell in the spreadsheet

## Project Structure

```
doc_sheet_chain/
├── Code.gs              # Main entry point
├── Constants.gs         # Global constants
├── UI.gs               # User interface functions
├── Data.gs             # Spreadsheet operations
├── Document.gs         # Document manipulation
├── Utils.gs            # Utility functions
├── Cache.gs            # Caching system
├── Performance.gs      # Performance monitoring
├── Sidebar.html        # Main sidebar interface
├── Help.html           # Help documentation
├── logo.html           # Logo component
├── appsscript.json     # Project manifest
├── API_DOCUMENTATION.md # Comprehensive API docs
└── README.md           # This file
```

## Setup

1. Open Google Docs
2. Go to Extensions > Apps Script
3. Create all the `.gs` files with their respective content
4. Create HTML files for `Sidebar.html`, `Help.html`, and `logo.html`
5. Deploy as a Google Docs add-on

## Usage

### Menu Options
- **Insert value**: Opens sidebar to select spreadsheet data
- **Update document**: Refreshes all linked values
- **Set color**: Customize text color for inserted values
- **Clear recent sheets**: Clears recent sheets and cache
- **Performance stats**: View operation performance metrics
- **Help**: Shows usage instructions

### Spreadsheet Format
Your spreadsheet should follow this format:
```
| Description    | Field    | Value |
|----------------|----------|-------|
| Number of cats | num_cats | 123   |
| Number of dogs | num_dogs | 456   |
```

## Performance Features

- **Smart Caching**: Frequently accessed data is cached for 5 minutes
- **Batch Operations**: Large documents are processed in efficient batches
- **Progress Tracking**: Real-time progress updates for long operations
- **Performance Monitoring**: Built-in timing and statistics tracking

## Error Handling

- **Input Validation**: All inputs are validated before processing
- **Graceful Failures**: Operations continue even if individual items fail
- **User Feedback**: Clear error messages and success notifications
- **Debug Logging**: Comprehensive console logging for troubleshooting

## Documentation

- **API Documentation**: Complete function reference in `API_DOCUMENTATION.md`
- **JSDoc Comments**: All functions include detailed JSDoc documentation
- **Code Comments**: Comprehensive inline documentation
- **Help System**: Built-in help with usage examples

## Technical Details

- **Runtime**: Google Apps Script V8
- **Caching**: Script cache with automatic expiration
- **Error Logging**: Stackdriver integration
- **OAuth Scopes**: Minimal required permissions
- **Modular Architecture**: Clean separation of concerns

## Contributing

1. Follow the existing code style and documentation patterns
2. Add JSDoc comments for all new functions
3. Include error handling and performance monitoring
4. Update documentation as needed

## Version History

- **v1.0**: Initial release with basic functionality
- **v1.1**: Added error handling and performance improvements
- **v1.2**: Added caching system and modular architecture
- **v1.3**: Added performance monitoring and comprehensive documentation
