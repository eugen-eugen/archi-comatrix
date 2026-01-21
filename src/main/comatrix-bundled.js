/**
 * comatrix-bundled.js
 * Entry point for webpack bundling - Excel export using SheetJS with styles
 */

const XLSX = require("xlsx-js-style");

/**
 * Builds a connectivity matrix from an Archi model
 * @param {Object} currentModel - The Archi model to analyze
 * @returns {Object} Matrix data structure with aElements, bElementsMap, sortedAElements, sortedBElements, and relationships
 */
function buildComatrix(currentModel) {
  console.log("Step 1: Analyzing relationships and building matrix...");

  // Extract relationships from the model
  const relationships = extractElements(currentModel);
  console.log(`Found ${relationships.length} triggering relationships`);

  // Collect all unique A elements (targets) and B elements (sources)
  // Map structure: A element name -> {domain, schnittstellenMap}
  const aElements = new Map();
  const bElementsMap = new Map(); // Store B elements with their domains

  relationships.forEach((rel) => {
    const aName = rel.target.name;
    const bName = rel.source.name;
    const schnittstelle = rel.schnittstelle || "N/A";
    const aDomain = rel.targetDomain;
    const bDomain = rel.sourceDomain;

    if (!aElements.has(aName)) {
      aElements.set(aName, {
        domain: aDomain,
        schnittstellenMap: new Map(),
      });
    }

    const aData = aElements.get(aName);
    if (!aData.schnittstellenMap.has(schnittstelle)) {
      aData.schnittstellenMap.set(schnittstelle, new Set());
    }

    aData.schnittstellenMap.get(schnittstelle).add(bName);

    // Store B element with its domain
    if (!bElementsMap.has(bName)) {
      bElementsMap.set(bName, bDomain);
    }
  });

  console.log(
    `Found ${aElements.size} target elements (A) and ${bElementsMap.size} source elements (B)`,
  );
  console.log("Step 2: Sorting elements...");

  // Sort A elements by domain (empty domains last), then by element name
  const sortedAElements = Array.from(aElements.keys()).sort((a, b) => {
    const domainA = aElements.get(a).domain;
    const domainB = aElements.get(b).domain;

    // Empty domains go to the end
    if (domainA === "" && domainB !== "") return 1;
    if (domainA !== "" && domainB === "") return -1;

    // Compare domains
    if (domainA !== domainB) {
      return domainA.localeCompare(domainB);
    }

    // Same domain, sort by element name
    return a.localeCompare(b);
  });

  const sortedBElements = Array.from(bElementsMap.keys()).sort((a, b) => {
    const domainA = bElementsMap.get(a);
    const domainB = bElementsMap.get(b);

    // Empty domains go to the end
    if (domainA === "" && domainB !== "") return 1;
    if (domainA !== "" && domainB === "") return -1;

    // Compare domains
    if (domainA !== domainB) {
      return domainA.localeCompare(domainB);
    }

    // Same domain, sort by element name
    return a.localeCompare(b);
  });

  console.log("Matrix data structure built successfully.");

  return {
    aElements,
    bElementsMap,
    sortedAElements,
    sortedBElements,
    relationships,
  };
}

/**
 * Merges two connectivity matrices (baseline and current)
 * @param {Object} comatrixBase - Matrix data structure from buildComatrix(baselineModel)
 * @param {Object} comatrixCurrent - Matrix data structure from buildComatrix(currentModel)
 * @returns {Object} Merged matrix data structure containing all elements from both models
 */
function merge(comatrixBase, comatrixCurrent) {
  console.log("Merging baseline and current matrices...");

  // Merge A elements
  const aElements = new Map();
  
  // Copy baseline A elements
  comatrixBase.aElements.forEach((value, key) => {
    aElements.set(key, {
      domain: value.domain,
      schnittstellenMap: new Map(value.schnittstellenMap),
    });
  });

  // Merge current A elements
  comatrixCurrent.aElements.forEach((value, key) => {
    if (aElements.has(key)) {
      // Element exists in both - merge Schnittstellen
      const existingData = aElements.get(key);
      value.schnittstellenMap.forEach((bSet, schnittstelle) => {
        if (!existingData.schnittstellenMap.has(schnittstelle)) {
          existingData.schnittstellenMap.set(schnittstelle, new Set(bSet));
        } else {
          // Merge B elements for this Schnittstelle
          bSet.forEach((b) =>
            existingData.schnittstellenMap.get(schnittstelle).add(b),
          );
        }
      });
    } else {
      // New element from current
      aElements.set(key, {
        domain: value.domain,
        schnittstellenMap: new Map(value.schnittstellenMap),
      });
    }
  });

  // Merge B elements
  const bElementsMap = new Map();
  
  // Copy baseline B elements
  comatrixBase.bElementsMap.forEach((domain, name) => {
    bElementsMap.set(name, domain);
  });

  // Add current B elements
  comatrixCurrent.bElementsMap.forEach((domain, name) => {
    if (!bElementsMap.has(name)) {
      bElementsMap.set(name, domain);
    }
  });

  // Sort merged elements
  const sortedAElements = Array.from(aElements.keys()).sort((a, b) => {
    const domainA = aElements.get(a).domain;
    const domainB = aElements.get(b).domain;

    if (domainA === "" && domainB !== "") return 1;
    if (domainA !== "" && domainB === "") return -1;

    if (domainA !== domainB) {
      return domainA.localeCompare(domainB);
    }

    return a.localeCompare(b);
  });

  const sortedBElements = Array.from(bElementsMap.keys()).sort((a, b) => {
    const domainA = bElementsMap.get(a);
    const domainB = bElementsMap.get(b);

    if (domainA === "" && domainB !== "") return 1;
    if (domainA !== "" && domainB === "") return -1;

    if (domainA !== domainB) {
      return domainA.localeCompare(domainB);
    }

    return a.localeCompare(b);
  });

  // Merge relationships
  const relationships = [
    ...comatrixBase.relationships,
    ...comatrixCurrent.relationships,
  ];

  console.log(
    `Merged matrix: ${aElements.size} A-elements, ${bElementsMap.size} B-elements, ${relationships.length} relationships`,
  );

  return {
    aElements,
    bElementsMap,
    sortedAElements,
    sortedBElements,
    relationships,
  };
}

/**
 * Outputs a connectivity matrix to Excel with optional baseline comparison
 * @param {Object} comatrix - Matrix data structure from buildComatrix() or merge()
 * @param {String} outputPath - Path to save the Excel file
 * @param {Object} comatrixBase - Optional baseline matrix for comparison (from buildComatrix)
 */
function output2Excel(comatrix, outputPath, comatrixBase) {
  const { aElements, bElementsMap, sortedAElements, sortedBElements } =
    comatrix;

  // Build baseline sets if baseline matrix is provided
  let baselineSets = null;
  if (comatrixBase) {
    console.log("Building baseline sets for comparison...");

    baselineSets = {
      aElements: new Set(),
      bElements: new Set(),
    };

    comatrixBase.aElements.forEach((value, name) => {
      baselineSets.aElements.add(name);
    });

    comatrixBase.bElementsMap.forEach((domain, name) => {
      baselineSets.bElements.add(name);
    });

    console.log(
      `Baseline contains ${baselineSets.aElements.size} A-elements and ${baselineSets.bElements.size} B-elements`,
    );
    console.log("New elements will be highlighted with green fill.\n");
  }

  console.log("Step 3: Building matrix data...");

  // Build domain row for B elements (first row)
  const bDomainRow = [
    "",
    "",
    "",
    "",
    ...sortedBElements.map((bName) => bElementsMap.get(bName)),
  ];

  // Build header row (second row)
  const headerRow = [
    "Domäne",
    "Anwendungssystem",
    "Angebotene Schnittstelle",
    "intern/extern",
    ...sortedBElements,
  ];

  // Build data rows
  const dataRows = [];
  sortedAElements.forEach((aName) => {
    const aData = aElements.get(aName);
    const aDomain = aData.domain;
    const schnittstellenMap = aData.schnittstellenMap;
    const sortedSchnittstellen = Array.from(schnittstellenMap.keys()).sort();

    sortedSchnittstellen.forEach((schnittstelle) => {
      const targetSet = schnittstellenMap.get(schnittstelle);

      // Determine if intern or extern based on domains
      // Check if any B element in this row has a connection
      let hasConnectionWithDomain = false;
      let hasConnectionWithoutDomain = false;

      sortedBElements.forEach((bName) => {
        if (targetSet.has(bName)) {
          const bDomain = bElementsMap.get(bName);
          if (aDomain !== "" && bDomain !== "") {
            hasConnectionWithDomain = true;
          } else {
            hasConnectionWithoutDomain = true;
          }
        }
      });

      // Determine intern/extern: if at least one connection is extern, mark as extern
      const internExtern = hasConnectionWithoutDomain
        ? "extern"
        : hasConnectionWithDomain
          ? "intern"
          : "";

      const row = [aDomain, aName, schnittstelle, internExtern];

      // Add "x" for each B element if connection exists
      sortedBElements.forEach((bName) => {
        row.push(targetSet.has(bName) ? "x" : "");
      });

      dataRows.push(row);
    });
  });

  // Combine domain row, header, and data
  const data = [bDomainRow, headerRow, ...dataRows];

  console.log(
    `Step 4: Creating worksheet with ${dataRows.length} rows and ${sortedBElements.length + 2} columns...`,
  );

  // Create workbook and worksheet
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(data);

  // Calculate column widths based on content
  console.log("Step 4.5: Calculating auto-fit column widths...");
  const colWidths = [];

  for (let col = 0; col < headerRow.length; col++) {
    let maxWidth = 0;

    if (col < 4) {
      // For Domäne, Anwendungssystem, Angebotene Schnittstelle, and intern/extern columns, calculate based on content
      for (let row = 0; row < data.length; row++) {
        const cellValue = data[row][col];
        if (cellValue) {
          const cellLength = String(cellValue).length;
          maxWidth = Math.max(maxWidth, cellLength);
        }
      }
      // Add some padding and set minimum width
      const width = Math.max(10, maxWidth + 2);
      colWidths.push({ wch: width });
    } else {
      // For B-element columns (vertical text), set fixed width for 2 lines
      colWidths.push({ wch: 5 });
    }
  }

  worksheet["!cols"] = colWidths;

  // Style the rows
  console.log("Step 5: Applying styles and borders to all cells...");

  // Define border style
  const borderStyle = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
  };

  // Base style for all cells: Calibri 11 with borders
  const baseStyle = {
    font: { name: "Calibri", sz: 11 },
    border: borderStyle,
  };

  // Style for B-element domain row (row 0) columns D onwards - normal, RGB 217/217/217
  const domainStyleGray = {
    font: { name: "Calibri", sz: 11 },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "D9D9D9" } },
  };

  // Specific styles for A1, B1, C1
  const styleA1 = {
    font: { name: "Calibri", sz: 11 },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89
  };

  const styleB1 = {
    font: { name: "Calibri", sz: 11 },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "FFC000" } }, // RGB 255/192/0
  };

  const styleC1 = {
    font: { name: "Calibri", sz: 11 },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "C0504D" } }, // RGB 192/80/77
  };

  // Style for main header row (row 1) columns A-D - bold, RGB 217/217/217
  const headerStyleStandard = {
    font: { name: "Calibri", sz: 11, bold: true },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "D9D9D9" } },
  };

  // Style for main header row (row 1) B-element columns - normal, vertical, RGB 184/204/228
  const headerStyleVertical = {
    font: { name: "Calibri", sz: 11 },
    alignment: { horizontal: "center", textRotation: 90, wrapText: true },
    border: borderStyle,
    fill: { fgColor: { rgb: "B8CCE4" } }, // RGB 184/204/228
  };

  // Style for data cells columns A-D - Calibri 11 with borders, no fill
  const dataStyleLeftColumns = {
    font: { name: "Calibri", sz: 11 },
    border: borderStyle,
  };

  // Style for data cells columns A-D (bold for intern rows) - Calibri 11 with borders, no fill
  const dataStyleLeftColumnsBold = {
    font: { name: "Calibri", sz: 11, bold: true },
    border: borderStyle,
  };
  // Style for new A-elements (not in baseline) - green fill like A1
  const dataStyleNewA = {
    font: { name: "Calibri", sz: 11 },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89
  };

  const dataStyleNewABold = {
    font: { name: "Calibri", sz: 11, bold: true },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89
  };
  // Style for data cells columns E onwards - Calibri 11 with borders, RGB 184/204/228, centered
  const dataStyleRightColumns = {
    font: { name: "Calibri", sz: 11 },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "B8CCE4" } }, // RGB 184/204/228
  };

  // Style for "x" cells - bold, centered, RGB 184/204/228
  const dataStyleX = {
    font: { name: "Calibri", sz: 11, bold: true },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "B8CCE4" } }, // RGB 184/204/228
  };

  // Style for "x" cells in new A or B elements - bold, centered, green fill
  const dataStyleXNew = {
    font: { name: "Calibri", sz: 11, bold: true },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89 (green)
  };

  // Apply base style to all cells first
  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
      if (!worksheet[cellRef]) {
        worksheet[cellRef] = { v: "" };
      }
      // Data rows (row 2+)
      if (row >= 2) {
        const isIntern = data[row][3] === "intern"; // Check column D (intern/extern)
        const cellValue = data[row][col];
        const aElementName = data[row][1]; // Column B has A-element name (Anwendungssystem)
        const isNewA =
          baselineSets && !baselineSets.aElements.has(aElementName);

        if (col < 4) {
          // Columns A-D: green if new A-element, bold if intern
          if (isNewA) {
            worksheet[cellRef].s = isIntern
              ? { ...dataStyleNewABold }
              : { ...dataStyleNewA };
          } else {
            worksheet[cellRef].s = isIntern
              ? { ...dataStyleLeftColumnsBold }
              : { ...dataStyleLeftColumns };
          }
        } else {
          // Columns E+: check if B-element is new
          const bElementName = headerRow[col];
          const isNewB =
            baselineSets && !baselineSets.bElements.has(bElementName);

          if (cellValue === "x") {
            // "x" cell: green if A or B element is new, otherwise normal blue
            worksheet[cellRef].s =
              isNewA || isNewB ? { ...dataStyleXNew } : { ...dataStyleX };
          } else {
            // Empty cell: always use normal blue styling
            worksheet[cellRef].s = { ...dataStyleRightColumns };
          }
        }
      } else {
        worksheet[cellRef].s = { ...baseStyle };
      }
    }
  }

  // Apply specific styles to A1, B1, C1
  if (worksheet["A1"]) worksheet["A1"].s = styleA1;
  if (worksheet["B1"]) worksheet["B1"].s = styleB1;
  if (worksheet["C1"]) worksheet["C1"].s = styleC1;

  // Apply styles to B-element domain row (row 0) columns D onwards - gray
  for (let col = 3; col < bDomainRow.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
    if (worksheet[cellRef]) {
      worksheet[cellRef].s = domainStyleGray;
    }
  }

  // Apply styles to main header row (row 1) and domain row (row 0) for B-elements
  for (let col = 0; col < headerRow.length; col++) {
    const cellRef1 = XLSX.utils.encode_cell({ r: 1, c: col });
    const cellRef0 = XLSX.utils.encode_cell({ r: 0, c: col });

    if (col < 4) {
      // Columns 0-3: bold with gray
      if (worksheet[cellRef1]) {
        worksheet[cellRef1].s = headerStyleStandard;
      }
    } else {
      // Columns 4+ (B elements): check if new
      const bElementName = headerRow[col];
      const isNewB = baselineSets && !baselineSets.bElements.has(bElementName);

      if (isNewB) {
        // New B-element: green fill in rows 0 and 1
        const styleNewBDomain = {
          font: { name: "Calibri", sz: 11 },
          alignment: { horizontal: "center" },
          border: borderStyle,
          fill: { fgColor: { rgb: "9BBB59" } }, // Green like A1
        };
        const styleNewBHeader = {
          font: { name: "Calibri", sz: 11 },
          alignment: { horizontal: "center", textRotation: 90, wrapText: true },
          border: borderStyle,
          fill: { fgColor: { rgb: "9BBB59" } }, // Green like A1
        };
        if (worksheet[cellRef0]) worksheet[cellRef0].s = styleNewBDomain;
        if (worksheet[cellRef1]) worksheet[cellRef1].s = styleNewBHeader;
      } else {
        // Existing B-element: normal styling
        if (worksheet[cellRef1]) {
          worksheet[cellRef1].s = headerStyleVertical;
        }
      }
    }
  }

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, "Matrix");

  // Set freeze panes: freeze first 2 rows and first 4 columns (A-D)
  workbook.Workbook = { Views: [{ xSplit: 4, ySplit: 2 }] };

  console.log("Step 6: Generating Excel binary...");
  const excelBuffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsm",
    cellStyles: true,
  });

  console.log(`Step 7: Writing ${excelBuffer.length} bytes to file...`);
  const FileOutputStream = Java.type("java.io.FileOutputStream");
  const fos = new FileOutputStream(outputPath, false);
  const javaBytes = Java.to(Array.from(excelBuffer), "byte[]");
  fos.write(javaBytes);
  fos.close();

  console.log(`Step 8: Excel file created successfully: ${outputPath}`);
}

/**
 * Finds the domain for an element by traversing aggregation/composition relationships
 * @param {Object} element - The Archi element
 * @returns {String} The domain name or empty string if not found
 */
function findDomain(element) {
  // Check if this element itself has the "Domäne" property
  const domainProp = element.prop("Domäne");
  if (domainProp) {
    return element.name;
  }

  // Find incoming aggregation or composition relationships
  const incomingRels = $(element)
    .inRels("aggregation-relationship")
    .add($(element).inRels("composition-relationship"));

  // Traverse up the hierarchy
  const visited = new Set();
  visited.add(element.id);

  for (let i = 0; i < incomingRels.length; i++) {
    const rel = incomingRels[i];
    const parent = rel.source;

    if (visited.has(parent.id)) {
      continue; // Avoid circular references
    }

    // Recursively search in parent
    const domain = findDomain(parent);
    if (domain !== "") {
      return domain;
    }
  }

  return "";
}

/**
 * Extracts triggering relationship information from Archi model
 * @param {Object} model - Archi model object
 * @returns {Array} Array of relationship objects with source, target, and Schnittstelle property
 */
function extractElements(model) {
  const relationships = [];

  // Get all triggering relationships from the model where name starts with NST_
  $(model)
    .find("triggering-relationship")
    .each((relationship) => {
      if (relationship.name && relationship.name.startsWith("NST_")) {
        // Get the Schnittstelle property values, can be multiply with same key
        const schnittstellen = relationship.prop("Schnittstelle", true) || "";

        // Find domains for source and target
        const sourceDomain = findDomain(relationship.source);

        const targetDomain = findDomain(relationship.target);
        schnittstellen.forEach((schnittstelle) => {
          relationships.push({
            id: relationship.id,
            name: relationship.name,
            source: relationship.source,
            target: relationship.target,
            schnittstelle: schnittstelle,
            sourceDomain: sourceDomain,
            targetDomain: targetDomain,
          });
        });
      }
    });

  return relationships;
}

/**
 * Main execution function
 */
function runComatrix() {
  console.clear();
  console.show();
  console.log("=== Comatrix - Generate Connectivity Matrix ===\n");

  // Check if a model is selected
  if (!model) {
    console.log(
      "ERROR: No model is selected. Please open or create a model first.",
    );
    return;
  }

  console.log(`Selected model: ${model.name}`);
  console.log(`Model path: ${model.path || "(not saved)"}\n`);

  // List all loaded models
  const loadedModels = $.model.getLoadedModels();
  console.log(`Total loaded models: ${loadedModels.length}`);
  loadedModels.forEach((m, index) => {
    const marker = m === model ? "→" : " ";
    console.log(`${marker} ${index + 1}. ${m.name}`);
  });
  console.log("");

  // Check for baseline model
  const baselineProperty = model.prop("baseline");
  let baselineModel = null;
  let compareMode = false;

  if (!baselineProperty || baselineProperty.trim() === "") {
    console.log("ℹ Property 'baseline' is not set in the selected model.");
    console.log("Running in single model mode.\n");
  } else {
    console.log(`Looking for baseline model: "${baselineProperty}"`);

    // Search for baseline model among loaded models
    baselineModel = loadedModels.find((m) => m.name === baselineProperty);

    if (baselineModel) {
      console.log(`✓ Baseline model found: ${baselineModel.name}`);
      console.log(`  Baseline path: ${baselineModel.path || "(not saved)"}`);
      console.log("Running in COMPARE MODE.\n");
      compareMode = true;
    } else {
      console.log(`✗ Baseline model "${baselineProperty}" is not opened.`);
      console.log("Running in single model mode.\n");
    }
  }

  // Extract all triggering relationships from the selected model
  console.log("Extracting triggering relationships starting with NST_...");
  const elements = extractElements(model);
  console.log(`Found ${elements.length} triggering relationships`);

  // Print domain information for verification
  console.log("\n=== Domain Information ===");
  const uniqueElements = new Map();

  elements.forEach((rel) => {
    if (!uniqueElements.has(rel.source.name)) {
      uniqueElements.set(rel.source.name, rel.sourceDomain);
    }
    if (!uniqueElements.has(rel.target.name)) {
      uniqueElements.set(rel.target.name, rel.targetDomain);
    }
  });

  console.log(`Total unique elements: ${uniqueElements.size}\n`);

  // Define output path
  const outputDir = model.path
    ? model.path.substring(0, model.path.lastIndexOf("/"))
    : __DIR__;
  const outputPath = outputDir + "/comatrix.xlsm";

  console.log(`Output file: ${outputPath}\n`);

  // Generate Excel file
  console.log("Generating Excel file...");

  try {
    let comatrix;
    let comatrixBase = null;

    if (compareMode) {
      // Build matrices for both models
      console.log("Building matrix for baseline model...");
      comatrixBase = buildComatrix(baselineModel);
      console.log(`Baseline: ${comatrixBase.relationships.length} relationships\n`);

      console.log("Building matrix for current model...");
      const comatrixCurrent = buildComatrix(model);
      console.log(`Current: ${comatrixCurrent.relationships.length} relationships\n`);

      if (comatrixCurrent.relationships.length === 0) {
        console.log(
          "⚠ No NST_* triggering relationships found in the current model.",
        );
        return;
      }

      // Merge baseline and current
      comatrix = merge(comatrixBase, comatrixCurrent);
      console.log("");
    } else {
      // Single model mode - just build from current model
      comatrix = buildComatrix(model);

      if (comatrix.relationships.length === 0) {
        console.log(
          "⚠ No NST_* triggering relationships found in the selected model.",
        );
        return;
      }
    }

    // Output to Excel with optional baseline comparison
    output2Excel(comatrix, outputPath, comatrixBase);

    console.log("\n=== Export Complete ===");
    console.log(`Matrix file saved to: ${outputPath}`);
    console.log(
      `Total triggering relationships processed: ${comatrix.relationships.length}`,
    );

    // Open the file location in file browser
    try {
      java.awt.Desktop.getDesktop().open(new java.io.File(outputDir));
    } catch (e) {
      console.log("Could not open file browser automatically.");
    }
  } catch (error) {
    console.log(`\n✗ ERROR: Failed to create Excel file`);
    console.log(`Error message: ${error.message}`);
    console.log(`Error type: ${error.constructor.name}`);
    if (error.stack) {
      console.log(`Stack trace:\n${error.stack}`);
    }
  }
}

// Export functions for bundling
module.exports = {
  buildComatrix,
  merge,
  output2Excel,
  extractElements,
  runComatrix,
};

// Execute the main function when script is loaded
runComatrix();
