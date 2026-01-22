/**
 * comatrix-bundled.js
 * Entry point for webpack bundling - Excel export using SheetJS with styles
 */

const XLSX = require("xlsx-js-style");
const path = require("path-browserify");

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
      
      // Update domain: use current if baseline is empty or differs
      if (!existingData.domain || existingData.domain === "" || existingData.domain !== value.domain) {
        existingData.domain = value.domain;
      }
      
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

  // Add/update current B elements - prefer current domain if baseline is empty or differs
  comatrixCurrent.bElementsMap.forEach((domain, name) => {
    if (bElementsMap.has(name)) {
      const baselineDomain = bElementsMap.get(name);
      // Use current domain if baseline is empty or differs
      if (!baselineDomain || baselineDomain === "" || baselineDomain !== domain) {
        bElementsMap.set(name, domain);
      }
    } else {
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
 * @param {Object} comatrixCurrent - Optional current matrix for comparison (from buildComatrix)
 */
function output2Excel(comatrix, outputPath, comatrixBase, comatrixCurrent) {
  const { aElements, bElementsMap, sortedAElements, sortedBElements } =
    comatrix;

  // Build baseline and current sets if provided for comparison
  let baselineSets = null;
  let currentSets = null;
  let changedAElements = new Set(); // A-elements with changed connections
  let changedBElements = new Set(); // B-elements with changed connections
  let baselineRowCombinations = new Set(); // Set of "AElement|Schnittstelle" combinations in baseline
  let currentRowCombinations = new Set(); // Set of "AElement|Schnittstelle" combinations in current
  let changedRowCombinations = new Set(); // Row combinations with connection changes

  if (comatrixBase && comatrixCurrent) {
    console.log("Building baseline and current sets for comparison...");

    baselineSets = {
      aElements: new Set(),
      bElements: new Set(),
    };

    // Add only element names (not domains) for comparison
    comatrixBase.aElements.forEach((value, name) => {
      baselineSets.aElements.add(name);
    });

    comatrixBase.bElementsMap.forEach((domain, name) => {
      baselineSets.bElements.add(name);
    });

    currentSets = {
      aElements: new Set(),
      bElements: new Set(),
    };

    // Add only element names (not domains) for comparison
    comatrixCurrent.aElements.forEach((value, name) => {
      currentSets.aElements.add(name);
    });

    comatrixCurrent.bElementsMap.forEach((domain, name) => {
      currentSets.bElements.add(name);
    });

    // Build row combinations (A-element + Schnittstelle)
    comatrixBase.aElements.forEach((data, aName) => {
      data.schnittstellenMap.forEach((bSet, schnittstelle) => {
        baselineRowCombinations.add(`${aName}|${schnittstelle}`);
      });
    });

    comatrixCurrent.aElements.forEach((data, aName) => {
      data.schnittstellenMap.forEach((bSet, schnittstelle) => {
        currentRowCombinations.add(`${aName}|${schnittstelle}`);
      });
    });

    // Detect changed row combinations (A-element + Schnittstelle with different connections)
    console.log("Analyzing connection changes...");

    // Check combinations that exist in both baseline and current
    currentRowCombinations.forEach((combo) => {
      if (baselineRowCombinations.has(combo)) {
        const [aName, schnittstelle] = combo.split("|");

        const baseBSet = comatrixBase.aElements
          .get(aName)
          .schnittstellenMap.get(schnittstelle);
        const currentBSet = comatrixCurrent.aElements
          .get(aName)
          .schnittstellenMap.get(schnittstelle);

        // Check if the B-elements (connections) are different
        let hasChanges = false;

        baseBSet.forEach((bName) => {
          if (!currentBSet.has(bName)) {
            hasChanges = true;
          }
        });

        currentBSet.forEach((bName) => {
          if (!baseBSet.has(bName)) {
            hasChanges = true;
          }
        });

        if (hasChanges) {
          changedRowCombinations.add(combo);
        }
      }
    });

    // Check each B-element column for changes
    sortedBElements.forEach((bName) => {
      const inBaseline = baselineSets.bElements.has(bName);
      const inCurrent = currentSets.bElements.has(bName);

      if (inBaseline && inCurrent) {
        // B-element exists in both - check if any connections differ in this column
        let hasChanges = false;

        // Check all A-element + Schnittstelle combinations for this B-element
        comatrixBase.aElements.forEach((baseData, aName) => {
          baseData.schnittstellenMap.forEach((baseBSet, schnittstelle) => {
            const inBaselineRow = baseBSet.has(bName);

            // Check if this combination exists in current
            let inCurrentRow = false;
            if (comatrixCurrent.aElements.has(aName)) {
              const currentData = comatrixCurrent.aElements.get(aName);
              if (currentData.schnittstellenMap.has(schnittstelle)) {
                inCurrentRow = currentData.schnittstellenMap
                  .get(schnittstelle)
                  .has(bName);
              }
            }

            if (inBaselineRow !== inCurrentRow) {
              hasChanges = true;
            }
          });
        });

        // Also check combinations only in current
        comatrixCurrent.aElements.forEach((currentData, aName) => {
          currentData.schnittstellenMap.forEach(
            (currentBSet, schnittstelle) => {
              if (currentBSet.has(bName)) {
                // Check if this combination existed in baseline
                let inBaselineRow = false;
                if (comatrixBase.aElements.has(aName)) {
                  const baseData = comatrixBase.aElements.get(aName);
                  if (baseData.schnittstellenMap.has(schnittstelle)) {
                    inBaselineRow = baseData.schnittstellenMap
                      .get(schnittstelle)
                      .has(bName);
                  }
                }

                if (!inBaselineRow) {
                  hasChanges = true;
                }
              }
            },
          );
        });

        if (hasChanges) {
          changedBElements.add(bName);
        }
      }
    });

    console.log(
      `Baseline contains ${baselineSets.aElements.size} A-elements and ${baselineSets.bElements.size} B-elements`,
    );
    console.log(
      `Current contains ${currentSets.aElements.size} A-elements and ${currentSets.bElements.size} B-elements`,
    );
    console.log(
      `Elements with changed connections: ${changedAElements.size} A-elements, ${changedBElements.size} B-elements`,
    );
    console.log("New elements will be highlighted with green fill.");
    console.log("Removed elements will be highlighted with red fill.");
    console.log(
      "Elements with changed connections will be highlighted with orange fill.\n",
    );
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
    font: { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } },
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
    font: { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89
  };

  const dataStyleNewABold = {
    font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89
  };

  // Style for removed A-elements (in baseline but not in current) - red fill like C1
  const dataStyleRemovedA = {
    font: { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } },
    border: borderStyle,
    fill: { fgColor: { rgb: "C0504D" } }, // RGB 192/80/77
  };

  const dataStyleRemovedABold = {
    font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
    border: borderStyle,
    fill: { fgColor: { rgb: "C0504D" } }, // RGB 192/80/77
  };

  // Style for changed A-elements (in both, with connection changes) - orange fill like B1
  const dataStyleChangedA = {
    font: { name: "Calibri", sz: 11 },
    border: borderStyle,
    fill: { fgColor: { rgb: "FFC000" } }, // RGB 255/192/0
  };

  const dataStyleChangedABold = {
    font: { name: "Calibri", sz: 11, bold: true },
    border: borderStyle,
    fill: { fgColor: { rgb: "FFC000" } }, // RGB 255/192/0
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
    font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "9BBB59" } }, // RGB 155/187/89 (green)
  };

  // Style for "x" cells in removed A or B elements - bold, centered, red fill
  const dataStyleXRemoved = {
    font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: "C0504D" } }, // RGB 192/80/77 (red)
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
        const schnittstelle = data[row][2]; // Column C has Schnittstelle
        const rowCombo = `${aElementName}|${schnittstelle}`;

        // Check row combination status
        const isNewRow =
          baselineRowCombinations &&
          currentRowCombinations &&
          !baselineRowCombinations.has(rowCombo) &&
          currentRowCombinations.has(rowCombo);
        const isRemovedRow =
          baselineRowCombinations &&
          currentRowCombinations &&
          baselineRowCombinations.has(rowCombo) &&
          !currentRowCombinations.has(rowCombo);
        const isChangedRow = changedRowCombinations.has(rowCombo);

        if (col < 4) {
          // Columns A-D: same color for all cells in a row based on row combination status
          if (isNewRow) {
            // New row: all A-D cells green
            worksheet[cellRef].s = isIntern
              ? { ...dataStyleNewABold }
              : { ...dataStyleNewA };
          } else if (isChangedRow) {
            // Changed row: all A-D cells yellow
            worksheet[cellRef].s = isIntern
              ? { ...dataStyleChangedABold }
              : { ...dataStyleChangedA };
          } else if (isRemovedRow) {
            // Removed row: all A-D cells red with white text
            worksheet[cellRef].s = isIntern
              ? { ...dataStyleRemovedABold }
              : { ...dataStyleRemovedA };
          } else {
            // Normal row: normal coloring
            worksheet[cellRef].s = isIntern
              ? { ...dataStyleLeftColumnsBold }
              : { ...dataStyleLeftColumns };
          }
        } else {
          // Columns E+: check if B-element is new, removed, or changed
          const bElementName = headerRow[col];
          const isNewB =
            baselineSets &&
            currentSets &&
            !baselineSets.bElements.has(bElementName);
          const isRemovedB =
            baselineSets &&
            currentSets &&
            !currentSets.bElements.has(bElementName);
          const isChangedB = changedBElements.has(bElementName);

          if (cellValue === "x") {
            // Check if this specific connection is new or removed
            let connectionStatus = "existing";

            if (baselineSets && currentSets) {
              const inBaseline =
                comatrixBase.aElements.has(aElementName) &&
                comatrixBase.aElements
                  .get(aElementName)
                  .schnittstellenMap.has(schnittstelle) &&
                comatrixBase.aElements
                  .get(aElementName)
                  .schnittstellenMap.get(schnittstelle)
                  .has(bElementName);

              const inCurrent =
                comatrixCurrent.aElements.has(aElementName) &&
                comatrixCurrent.aElements
                  .get(aElementName)
                  .schnittstellenMap.has(schnittstelle) &&
                comatrixCurrent.aElements
                  .get(aElementName)
                  .schnittstellenMap.get(schnittstelle)
                  .has(bElementName);

              if (inCurrent && !inBaseline) {
                connectionStatus = "new";
              } else if (inBaseline && !inCurrent) {
                connectionStatus = "removed";
              }
            }

            // "x" cell: color based on connection status or element status
            if (connectionStatus === "new" || isNewRow || isNewB) {
              worksheet[cellRef].s = { ...dataStyleXNew };
            } else if (
              connectionStatus === "removed" ||
              isRemovedRow ||
              isRemovedB
            ) {
              worksheet[cellRef].s = { ...dataStyleXRemoved };
            } else {
              worksheet[cellRef].s = { ...dataStyleX };
            }
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
      // Columns 4+ (B elements): check if new, removed, or changed
      const bElementName = headerRow[col];
      const isNewB =
        baselineSets &&
        currentSets &&
        !baselineSets.bElements.has(bElementName);
      const isRemovedB =
        baselineSets && currentSets && !currentSets.bElements.has(bElementName);
      const isChangedB = changedBElements.has(bElementName);

      if (isNewB) {
        // New B-element: green fill in rows 0 and 1
        const styleNewBDomain = {
          font: { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } },
          alignment: { horizontal: "center" },
          border: borderStyle,
          fill: { fgColor: { rgb: "9BBB59" } }, // Green like A1
        };
        const styleNewBHeader = {
          font: { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } },
          alignment: { horizontal: "center", textRotation: 90, wrapText: true },
          border: borderStyle,
          fill: { fgColor: { rgb: "9BBB59" } }, // Green like A1
        };
        if (worksheet[cellRef0]) worksheet[cellRef0].s = styleNewBDomain;
        if (worksheet[cellRef1]) worksheet[cellRef1].s = styleNewBHeader;
      } else if (isRemovedB) {
        // Removed B-element: red fill in rows 0 and 1
        const styleRemovedBDomain = {
          font: { name: "Calibri", sz: 11 },
          alignment: { horizontal: "center" },
          border: borderStyle,
          fill: { fgColor: { rgb: "C0504D" } }, // Red like C1
        };
        const styleRemovedBHeader = {
          font: { name: "Calibri", sz: 11 },
          alignment: { horizontal: "center", textRotation: 90, wrapText: true },
          border: borderStyle,
          fill: { fgColor: { rgb: "C0504D" } }, // Red like C1
        };
        if (worksheet[cellRef0]) worksheet[cellRef0].s = styleRemovedBDomain;
        if (worksheet[cellRef1]) worksheet[cellRef1].s = styleRemovedBHeader;
      } else if (isChangedB) {
        // Changed B-element: yellow fill in rows 0 and 1
        const styleChangedBDomain = {
          font: { name: "Calibri", sz: 11 },
          alignment: { horizontal: "center" },
          border: borderStyle,
          fill: { fgColor: { rgb: "FFFF00" } }, // Yellow
        };
        const styleChangedBHeader = {
          font: { name: "Calibri", sz: 11 },
          alignment: { horizontal: "center", textRotation: 90, wrapText: true },
          border: borderStyle,
          fill: { fgColor: { rgb: "FFFF00" } }, // Yellow
        };
        if (worksheet[cellRef0]) worksheet[cellRef0].s = styleChangedBDomain;
        if (worksheet[cellRef1]) worksheet[cellRef1].s = styleChangedBHeader;
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
  // Normalize path to use forward slashes (path-browserify only handles POSIX paths)
  const normalizedPath = model.path ? model.path.replace(/\\/g, "/") : null;
  const outputDir = normalizedPath ? path.dirname(normalizedPath) : __DIR__;
  const outputPath = path.join(outputDir, "comatrix.xlsm");

  console.log(`Output file: ${outputPath}\n`);

  // Generate Excel file
  console.log("Generating Excel file...");

  try {
    let comatrix;
    let comatrixBase = null;
    let comatrixCurrent = null;

    if (compareMode) {
      // Build matrices for both models
      console.log("Building matrix for baseline model...");
      comatrixBase = buildComatrix(baselineModel);
      console.log(
        `Baseline: ${comatrixBase.relationships.length} relationships\n`,
      );

      console.log("Building matrix for current model...");
      comatrixCurrent = buildComatrix(model);
      console.log(
        `Current: ${comatrixCurrent.relationships.length} relationships\n`,
      );

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
    output2Excel(comatrix, outputPath, comatrixBase, comatrixCurrent);

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
