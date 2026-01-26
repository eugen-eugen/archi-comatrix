/**
 * comatrix.js
 * Entry point for webpack bundling - Excel export using SheetJS with styles
 */

const XLSX = require("xlsx-js-style");
const path = require("path");
const output2Excel = require("./output2Excel");
const { findDomain, extractElements } = require("./model");

/**
 * Sorts elements by domain (empty domains last), then by element name
 * @param {Map} elementsMap - Map with element names as keys and objects with domain property as values
 * @returns {Array} Sorted array of element names
 */
function sortElementsByDomain(elementsMap) {
  return Array.from(elementsMap.keys()).sort((name1, name2) => {
    const domain1 = elementsMap.get(name1).domain;
    const domain2 = elementsMap.get(name2).domain;

    // Empty domains go to the end
    if (domain1 === "" && domain2 !== "") return 1;
    if (domain1 !== "" && domain2 === "") return -1;

    // Compare domains
    if (domain1 !== domain2) {
      return domain1.localeCompare(domain2);
    }

    // Same domain, sort by element name
    return name1.localeCompare(name2);
  });
}

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
      bElementsMap.set(bName, { domain: bDomain });
    }
  });

  console.log(
    `Found ${aElements.size} target elements (A) and ${bElementsMap.size} source elements (B)`,
  );
  console.log("Step 2: Sorting elements...");

  // Sort A elements by domain (empty domains last), then by element name
  const sortedAElements = sortElementsByDomain(aElements);

  // Sort B elements by domain (empty domains last), then by element name
  const sortedBElements = sortElementsByDomain(bElementsMap);

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
      if (
        !existingData.domain ||
        existingData.domain === "" ||
        existingData.domain !== value.domain
      ) {
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
  comatrixBase.bElementsMap.forEach((data, name) => {
    bElementsMap.set(name, { domain: data.domain });
  });

  // Add/update current B elements - prefer current domain if baseline is empty or differs
  comatrixCurrent.bElementsMap.forEach((data, name) => {
    if (bElementsMap.has(name)) {
      const baselineData = bElementsMap.get(name);
      // Use current domain if baseline is empty or differs
      if (
        !baselineData.domain ||
        baselineData.domain === "" ||
        baselineData.domain !== data.domain
      ) {
        bElementsMap.set(name, { domain: data.domain });
      }
    } else {
      bElementsMap.set(name, { domain: data.domain });
    }
  });

  // Sort merged elements
  const sortedAElements = sortElementsByDomain(aElements);

  const sortedBElements = sortElementsByDomain(bElementsMap);

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

// Execute the main function when script is loaded
runComatrix();
