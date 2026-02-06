# Comatrix

Export Archi model relationships to Excel connectivity matrix with baseline comparison support, and generate application lists.

## Description

Comatrix is an Archi JavaScript extension that provides two main tools:

1. **Connectivity Matrix (comatrix)**: Analyzes ArchiMate model relationships and exports them to a styled Excel connectivity matrix (`comatrix.xlsx`). It shows which A-elements (application components) connect to which B-elements through specific interfaces (Schnittstellen), with support for baseline comparison to highlight changes between model versions.

2. **Application List (applist)**: Generates a list of all application components with specialization "Geschäftsanwendung" and their associated domains (`applist.xlsx`).

> **📖 New to these scripts?** See [doc/metamodel.md](doc/metamodel.md) for a comprehensive guide on how to structure your Archi model to work with these scripts, including which elements, relationships, specializations, and properties influence the outputs.

## Features

### Connectivity Matrix (comatrix)
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

### Application List (applist)
- **Application Listing**: Lists all application components with specializations:
  - "Geschäftsanwendung"
  - "Register"
  - "Querschnittsanwendung"
- **Domain Assignment**: Shows ALL domains (Domäne) for each application from grouping elements with specialization "Domäne"
  - Multiple domains are comma-separated
  - Handles nested groupings by traversing the entire hierarchy
- **Fachbereich Assignment**: Shows ALL Fachbereich for each application from grouping elements with specialization "Fachbereich"
  - Multiple Fachbereich are comma-separated
  - Handles nested groupings by traversing the entire hierarchy
- **Cycle Detection**: Detects circular grouping references and displays "cycle" instead of infinite loop
- **Type Classification**: Displays the specialization type for each application
- **Sorted Output**: Applications sorted by Fachbereich, then domain, then type, then name
- **Clean Formatting**: Styled Excel output with headers and borders

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

### Connectivity Matrix

#### Option 1: Bundled Version (Recommended)
1. Open your ArchiMate model in Archi
2. Run the script: `Scripts > comatrix-bundled.ajs` (located in `dist/` folder)
3. The Excel file `comatrix.xlsx` will be created in the same directory as your model
4. The file explorer will automatically open showing the output file

**No npm install required!** The bundled version includes all dependencies in a single ~932 KiB file.

#### Option 2: Development Version
For development, you can work directly with the source files in `src/main/` and rebuild using webpack.

#### Baseline Comparison
To compare a current model against a baseline:
1. Add a property "baseline" to your model with the name of the baseline model
2. Ensure both models are opened in Archi
3. Run the script - it will automatically detect and compare against the baseline
4. The Excel output highlights differences:
   - **Green (RGB 155/187/89)**: New elements or connections
   - **Yellow/Orange (RGB 255/192/0 or 255/255/0)**: Changed connections
   - **Red (RGB 192/80/77)**: Removed elements or connections
5. Legend is displayed in cells A1-C1 for easy reference

### Application List

1. Open your ArchiMate model in Archi
2. Run the script: `Scripts > applist-bundled.ajs` (located in `dist/` folder)
3. The Excel file `applist.xlsx` will be created in the same directory as your model
4. The file explorer will automatically open showing the output file

The application list includes:
- All application components with specializations: "Geschäftsanwendung", "Register", "Querschnittsanwendung"
- Their specialization type (Typ column)
- Their associated domains (Domäne) from ALL grouping elements with specialization "Domäne"
  - If an application is grouped by multiple domains (directly or through nested groupings), all are shown comma-separated
- Their associated Fachbereich from ALL grouping elements with specialization "Fachbereich"
  - If an application belongs to multiple Fachbereich (directly or through nested groupings), all are shown comma-separated
- Sorted by Fachbereich, then domain, then type, then application name
- Cycle detection: If circular grouping references are detected, "cycle" is displayed

## Building

To create the bundled versions:

```bash
npm run build
```

This generates both:
- `dist/comatrix-bundled.ajs` - Connectivity matrix script (~932 KiB)
- `dist/applist-bundled.ajs` - Application list script (~906 KiB)

Both include xlsx-js-style and all dependencies.

## Project Structure

```
comatrix/
├── package.json                  # Project configuration and dependencies
├── webpack.config.js             # Webpack bundling configuration
├── doc/
│   └── metamodel.md              # Archi model structure guide
├── src/
│   └── main/
│       ├── comatrix.js           # Main entry point - connectivity matrix
│       ├── applist.js            # Application list generator
│       ├── output2Excel.js       # Excel generation with styling
│       └── model.js              # Archi model interaction functions
├── dist/
│   ├── comatrix-bundled.ajs     # Bundled connectivity matrix script
│   └── applist-bundled.ajs      # Bundled application list script
├── tests/                        # Test files
└── .github/
    └── workflows/
        └── build-release.yml     # Automated build to release branch
```

## Requirements

- Archi (https://www.archimatetool.com/) with jArchi plugin
- Node.js and npm (for building only, not required to run bundled version)
- xlsx-js-style library (automatically included in bundle)

## Output Format

### Connectivity Matrix (comatrix.xlsx)

The generated Excel file contains one worksheet named "Matrix" with:
- **Row 0 (Legend)**: Color-coded legend in A1-C1, domain names for B-elements
- **Row 1 (Headers)**: Column headers with vertical text for B-elements
- **Row 2+ (Data)**: Each row represents an A-element + Interface combination
  - Column A: Domäne (Domain)
  - Column B: Anwendungssystem (Application System / A-element)
  - Column C: Angebotene Schnittstelle (Offered Interface)
  - Column D: intern/extern classification
  - Columns E+: B-elements with "x" marking connections

#### Color Coding
- **Green fill**: New elements or connections (not in baseline)
- **Yellow/Orange fill**: Changed connections or modified elements
- **Red fill**: Removed elements or connections (in baseline but not current)
- **Text colors**: White text on red background, black text on all others
- **Blue background (RGB 184/204/228)**: Standard B-element columns

### Application List (applist.xlsx)

The generated Excel file contains one worksheet named "Anwendungen" with:
- **Row 1 (Header)**: "Anwendung", "Typ", "Domäne", "Fachbereich" columns with gray background
- **Row 2+ (Data)**: All application components with specializations "Geschäftsanwendung", "Register", or "Querschnittsanwendung"
  - Column A: Anwendung (Application name)
  - Column B: Typ (Specialization: Geschäftsanwendung, Register, or Querschnittsanwendung)
  - Column C: Domäne (All domains from groupings with specialization "Domäne", comma-separated if multiple)
  - Column D: Fachbereich (All Fachbereich from groupings with specialization "Fachbereich", comma-separated if multiple)
  
**Special values:**
- Multiple domains/Fachbereich: Shown as comma-separated list (e.g., "Domain1, Domain2")
- Cycle detected: Shows "cycle" if circular grouping references exist
- No grouping found: Shows "(keine Domäne)" or "(kein Fachbereich)"

Applications are sorted by Fachbereich, then domain, then type, then alphabetically by name.

## Development

The codebase is organized into several modules:

1. **model.js**: Archi model interaction
   - `findAllGroupings()`: Helper function that traverses aggregation/composition relationships to find all groupings with a specific specialization
   - `findDomain()`: Finds ALL domains by traversing aggregation/composition relationships, looking for specialization "Domäne"
     - Returns comma-separated list of all found domains
     - Detects and reports cycles
     - Handles nested groupings
   - `findFachbereich()`: Finds ALL Fachbereich by traversing aggregation/composition relationships, looking for specialization "Fachbereich"
     - Returns comma-separated list of all found Fachbereich
     - Detects and reports cycles
     - Handles nested groupings
   - `extractElements()`: Finds triggering relationships starting with NST_*

2. **comatrix.js**: Connectivity matrix generator
   - `buildComatrix()`: Extracts relationships from Archi model
   - `merge()`: Combines baseline and current matrices
   - `sortElementsByDomain()`: Sorts elements by domain and name
   - `runComatrix()`: Main execution flow with baseline detection

3. **applist.js**: Application list generator
   - Queries application components with specializations: "Geschäftsanwendung", "Register", "Querschnittsanwendung"
   - Uses `findDomain()` and `findFachbereich()` to get ALL classifications (comma-separated)
   - Handles nested groupings and detects cycles
   - Generates Excel with 4 columns: Anwendung, Typ, Domäne (can be multi-value), Fachbereich (can be multi-value)
   - `generateAppListExcel()`: Creates styled Excel output with wider columns for multi-value cells

4. **output2Excel.js**: Excel generation utilities
   - Styling definitions with color constants
   - `getFontStyle()`: Unified font styling
   - `getTextColor()`: Text color based on background
   - Baseline comparison logic
   - Cell coloring based on comparison results
   - Excel file writing using Java FileOutputStream

To add more features:
1. Modify relationship detection in `buildComatrix()` or extend `extractElements()`
2. Enhance comparison logic in `output2Excel()` 
3. Add new specializations or filtering in `applist.js`
4. Extend domain detection logic in `model.js`

## Documentation

### Metamodel Guide

For a comprehensive understanding of how to structure your Archi model to work effectively with these scripts, see **[doc/metamodel.md](doc/metamodel.md)**.

This guide covers:
- Which Archi element types are analyzed (application components, groupings)
- Which relationship types matter (triggering, aggregation, composition)
- Required specializations (Domäne, Fachbereich, Geschäftsanwendung, Register, Querschnittsanwendung)
- Properties and naming conventions (NST_* prefix, Schnittstelle property)
- How model structure influences outputs
- Example structures and troubleshooting tips

**When to read this:**
- Setting up a new Archi model for these scripts
- Troubleshooting empty or unexpected outputs
- Understanding why certain elements appear/don't appear in the Excel files
- Planning your model's organization strategy

## License

MIT
