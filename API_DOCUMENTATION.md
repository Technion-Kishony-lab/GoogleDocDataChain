# Data Chain API Documentation

## Overview

Data Chain is a Google Apps Script add-on that creates dynamic links between Google Docs and Google Sheets. It allows users to insert values from spreadsheets into documents with automatic updates and scientific notation formatting.

## Module Structure

### Core Modules

#### `Code.gs` - Main Entry Point
- **onOpen()**: Initializes the add-on menu
- **replacePlaceholders()**: Updates all document placeholders
- **include(filename)**: Helper for HTML file inclusion

#### `Constants.gs` - Global Constants
- **LABEL_PARAM**: URL parameter for identifying Data Chain links
- **LABEL_VALUE**: Value for the label parameter
- **RECENT_SHEETS_KEY**: Property key for recent sheets
- **TEXT_COLOR_KEY**: Property key for text color preference

#### `UI.gs` - User Interface Functions
- **onOpen()**: Creates the Data-Chain menu
- **setColor()**: Sets text color for inserted values
- **getColor()**: Retrieves current text color setting
- **showSidebar()**: Opens the main sidebar interface
- **showHelp()**: Opens help documentation
- **showPerformanceStats()**: Displays performance metrics
- **include(filename)**: Includes HTML files in templates

#### `Data.gs` - Spreadsheet Operations
- **getRecentSheets()**: Gets list of recent spreadsheet IDs
- **getRecentSheetNames()**: Gets recent sheets with names
- **updateRecentSheets(id)**: Updates recent sheets list
- **getTabs(sheetId)**: Gets sheet tab names
- **getFields(sheetId, sheetName)**: Gets field names from sheet
- **insertFromSidebar(sheetId, sheetName, fieldName)**: Inserts value at cursor
- **clearRecentSheets()**: Clears recent sheets and cache

#### `Document.gs` - Document Manipulation
- **replacePlaceholders()**: Updates all placeholders in document
- **collectRuns(el, runsByElement)**: Collects linked text runs
- **replaceRun(txt, r, cache)**: Replaces single text run

#### `Utils.gs` - Utility Functions
- **formatScientificNotation(value)**: Formats scientific notation
- **insertScientificText(element, text, start, end)**: Inserts formatted text
- **extractSheetId(url)**: Extracts sheet ID from URL
- **getParam(url, name)**: Gets URL parameter value
- **columnToLetter(col)**: Converts column number to letter

#### `Cache.gs` - Caching System
- **getCachedData(key)**: Retrieves cached data
- **setCachedData(key, value)**: Stores data in cache
- **clearCachedData(key)**: Clears specific cache entry
- **clearAllCache()**: Clears all cache

#### `Performance.gs` - Performance Monitoring
- **startTimer(operation)**: Starts performance timer
- **endTimer(timer)**: Ends timer and logs result
- **getPerformanceStats()**: Gets performance statistics
- **clearPerformanceLog()**: Clears performance log
- **batchOperation(operations, batchSize)**: Processes operations in batches
- **debounce(func, wait)**: Creates debounced function

## Function Reference

### Core Functions

#### `onOpen()`
Initializes the Data-Chain add-on when a document opens.

**Returns:** `void`

**Example:**
```javascript
// Called automatically when document opens
onOpen();
```

#### `replacePlaceholders()`
Updates all linked placeholders in the document with current spreadsheet data.

**Returns:** `void`

**Throws:** `Error` when document operations fail

**Example:**
```javascript
// Update all placeholders in the document
replacePlaceholders();
```

#### `insertFromSidebar(sheetId, sheetName, fieldName)`
Inserts a value from a spreadsheet into the document at the cursor position.

**Parameters:**
- `sheetId` (string): Google Sheets spreadsheet ID
- `sheetName` (string): Name of the sheet tab
- `fieldName` (string): Field name to insert

**Returns:** `Object` with success/error status

**Example:**
```javascript
const result = insertFromSidebar('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms', 'Sheet1', 'population');
if (result.error) {
  console.error(result.error);
}
```

### Data Functions

#### `getTabs(sheetId)`
Retrieves the list of sheet tabs from a spreadsheet.

**Parameters:**
- `sheetId` (string): Google Sheets spreadsheet ID

**Returns:** `Array<string>` of sheet tab names

**Throws:** `Error` when spreadsheet access fails

**Example:**
```javascript
const tabs = getTabs('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms');
console.log('Available tabs:', tabs);
```

#### `getFields(sheetId, sheetName)`
Retrieves field names from a specific sheet tab.

**Parameters:**
- `sheetId` (string): Google Sheets spreadsheet ID
- `sheetName` (string): Name of the sheet tab

**Returns:** `Array<string>` of field names from column B

**Throws:** `Error` when spreadsheet or sheet access fails

**Example:**
```javascript
const fields = getFields('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms', 'Sheet1');
console.log('Available fields:', fields);
```

### Utility Functions

#### `formatScientificNotation(value)`
Formats numbers in scientific notation to display format.

**Parameters:**
- `value` (*): The value to format

**Returns:** `string` formatted representation

**Example:**
```javascript
const formatted = formatScientificNotation(2.3e7);
console.log(formatted); // "2.3·10^7"
```

#### `columnToLetter(col)`
Converts a column number to Excel-style column letter(s).

**Parameters:**
- `col` (number): The column number (1-based)

**Returns:** `string` column letter(s)

**Throws:** `Error` when column number is invalid

**Example:**
```javascript
const letter = columnToLetter(3);
console.log(letter); // "C"
```

### Cache Functions

#### `getCachedData(key)`
Retrieves data from script cache with automatic expiration check.

**Parameters:**
- `key` (string): The cache key

**Returns:** `*` The cached value or null if not found/expired

**Example:**
```javascript
const cachedTabs = getCachedData('tabs_1234567890');
if (cachedTabs) {
  console.log('Using cached tabs:', cachedTabs);
}
```

#### `setCachedData(key, value)`
Stores data in script cache with timestamp for expiration.

**Parameters:**
- `key` (string): The cache key
- `value` (*): The value to cache

**Returns:** `void`

**Example:**
```javascript
setCachedData('tabs_1234567890', ['Sheet1', 'Sheet2']);
```

### Performance Functions

#### `startTimer(operation)`
Creates a timer object to track operation performance.

**Parameters:**
- `operation` (string): The name of the operation being timed

**Returns:** `Object` Timer object with operation name and start time

**Example:**
```javascript
const timer = startTimer('getTabs');
// ... perform operation
const duration = endTimer(timer);
```

#### `getPerformanceStats()`
Calculates and returns performance metrics from the log.

**Returns:** `Object` Performance statistics for all tracked operations

**Example:**
```javascript
const stats = getPerformanceStats();
console.log('Performance stats:', stats);
```

## Error Handling

All functions include comprehensive error handling with:
- Input validation
- Try-catch blocks
- Detailed error messages
- Graceful fallbacks
- Console logging for debugging

## Performance Considerations

- **Caching**: 5-minute cache for frequently accessed data
- **Batch Processing**: Operations processed in batches of 10
- **Progress Tracking**: Progress logging for large documents
- **Performance Monitoring**: Built-in timing and statistics

## Security

The add-on requires the following OAuth scopes:
- `https://www.googleapis.com/auth/documents` - Document access
- `https://www.googleapis.com/auth/spreadsheets` - Spreadsheet access
- `https://www.googleapis.com/auth/drive` - Drive file access
- `https://www.googleapis.com/auth/script.external_request` - External requests

## Dependencies

- Google Apps Script V8 runtime
- Google Docs API
- Google Sheets API
- Google Drive API

## Browser Compatibility

- Google Chrome (recommended)
- Mozilla Firefox
- Microsoft Edge
- Safari (limited support)

## Version History

- **v1.0**: Initial release with basic functionality
- **v1.1**: Added error handling and performance improvements
- **v1.2**: Added caching system and modular architecture
- **v1.3**: Added performance monitoring and comprehensive documentation
