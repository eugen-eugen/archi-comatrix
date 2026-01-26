# Comatrix

Export Archi model relationships to Excel connectivity matrix with baseline comparison support.

## Description

Comatrix is an Archi JavaScript extension that analyzes ArchiMate model relationships and exports them to a styled Excel connectivity matrix (`comatrix.xlsx`). It shows which A-elements (application components) connect to which B-elements through specific interfaces (Schnittstellen), with support for baseline comparison to highlight changes between model versions.

## Features

- **Connectivity Matrix**: Generates a matrix showing connections between A-elements and B-elements via interfaces
- **Baseline Comparison**: Compare current model against a baseline to identify:
  - New elements/connections (green)
  - Changed connections (yellow/orange)
  - Removed elements/connections (red)
- **Domain Detection**: Automatically groups elements by domain using aggregation/composition relationships
- **Relationship Types**: Focuses on serving relationships and related ArchiMate connections
- **Visual Styling**: Color-coded cells with proper borders, fonts, and text rotation
- **Excel Features**: Frozen panes (first 4 columns and 2 rows), auto-fit column widths
- **Intern/Extern Classification**: Automatically detects whether connections are internal or external based on domains

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

**No npm install required!** The bundled version includes all dependencies in a single ~931 KiB file.

### Option 2: Development Version
For development, you can work directly with the source files in `src/main/` and rebuild using webpack.

### Baseline Comparison
To compare a current model against a baseline:
1. The script will prompt you to select a baseline model file
2. It analyzes both models and highlights differences:
   - **Green (RGB 155/187/89)**: New elements or connections
   - **Yellow/Orange (RGB 255/192/0 or 255/255/0)**: Changed connections
   - **Red (RGB 192/80/77)**: Removed elements or connections
3. Legend is displayed in cells A1-C1 for easy reference

## Building

To create the bundled version:

```bash
npm run build
```

This generates `dist/comatrix-bundled.ajs` with xlsx-js-style and all dependencies included.

## Project Structure

```
comatrix/
├── package.json              # Project configuration and dependencies
├── webpack.config.js         # Webpack bundling configuration
├── src/
│   └── main/
│       ├── comatrix.js       # Main entry point with matrix building logic
│       └── output2Excel.js   # Excel generation with styling and comparison
├── dist/
│   └── comatrix-bundled.ajs # Bundled script for Archi
├── tests/                    # Test files
└── .github/
    └── workflows/
        └── build-release.yml # Automated build to release branch
```

## Requirements

- Archi (https://www.archimatetool.com/) with jArchi plugin
- Node.js and npm (for building only, not required to run bundled version)
- xlsx-js-style library (automatically included in bundle)

## Output Format

The generated Excel file contains one worksheet named "Matrix" with:
- **Row 0 (Legend)**: Color-coded legend in A1-C1, domain names for B-elements
- **Row 1 (Headers)**: Column headers with vertical text for B-elements
- **Row 2+ (Data)**: Each row represents an A-element + Interface combination
  - Column A: Domäne (Domain)
  - Column B: Anwendungssystem (Application System / A-element)
  - Column C: Angebotene Schnittstelle (Offered Interface)
  - Column D: intern/extern classification
  - Columns E+: B-elements with "x" marking connections

### Color Coding
- **Green fill**: New elements or connections (not in baseline)
- **Yellow/Orange fill**: Changed connections or modified elements
- **Red fill**: Removed elements or connections (in baseline but not current)
- **Text colors**: White text on red background, black text on all others
- **Blue background (RGB 184/204/228)**: Standard B-element columns

## Development

The codebase is split into two main modules:

1. **comatrix.js**: Main orchestration
   - `buildComatrix()`: Extracts relationships from Archi model
   - `merge()`: Combines baseline and current matrices
   - `extractElements()`: Finds triggering relationships
   - `runComatrix()`: Main execution flow with file dialogs

2. **output2Excel.js**: Excel generation
   - Styling definitions (20+ style objects)
   - Baseline comparison logic
   - Cell coloring based on comparison results
   - Excel file writing using Java FileOutputStream

To add more features:
1. Modify relationship detection in `buildComatrix()` 
2. Enhance comparison logic in `output2Excel()` 
3. Add new styling or layout options

## License

MIT
