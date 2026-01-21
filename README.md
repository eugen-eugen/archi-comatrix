# Comatrix

Export Archi model elements to Excel spreadsheet.

## Description

Comatrix is an Archi JavaScript extension that reads all elements from an ArchiMate model and exports them to an Excel file (`comatrix.xlsx`) using the ExcelJS library.

## Features

- Extracts all elements from the current Archi model
- Exports to Excel format with the following columns:
  - ID
  - Name
  - Type
  - Documentation
- Auto-filter enabled for easy sorting
- Styled header row

## Installation

1. Navigate to the comatrix directory:
```bash
cd comatrix
```

2. Install dependencies:
```bash
npm install
```

## Usage

### Option 1: Bundled Version (Recommended)
1. Open your ArchiMate model in Archi
2. Run the script: `Scripts > comatrix-bundled.ajs` (located in `dist/` folder)
3. The Excel file `comatrix.xlsx` will be created in the same directory as your model
4. The file explorer will automatically open showing the output file

**No npm install required!** The bundled version includes all dependencies in a single 2.6 MB file.

### Option 2: Development Version
1. Install dependencies: `npm install`
2. Run the script: `Scripts > comatrix.ajs`

## Building

To create the bundled version:

```bash
npm run build
```

This generates `dist/comatrix-bundled.ajs` with ExcelJS and all dependencies included.

## Project Structure

```
comatrix/
├── package.json              # Project configuration and dependencies
├── comatrix.ajs             # Main Archi JavaScript file
├── src/
│   └── main/
│       └── excelGenerator.js # Excel generation functions
├── scripts/                  # Additional scripts
└── tests/                    # Test files
```

## Requirements

- Archi (https://www.archimatetool.com/)
- Node.js and npm (for dependency management)
- ExcelJS library (automatically installed via npm)

## Output Format

The generated Excel file contains one worksheet named "Elements" with:
- Header row with bold text and gray background
- Auto-filter enabled
- One row per element with ID, Name, Type, and Documentation

## Development

To add more features:
1. Modify [src/main/excelGenerator.js](src/main/excelGenerator.js) to add new columns or worksheets
2. Update [comatrix.ajs](comatrix.ajs) to call new functions

## License

MIT
