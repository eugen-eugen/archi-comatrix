/**
 * metacomatrix.js
 * Entry point: builds connectivity matrices routed by the meta-model relationship
 * property Report:Model:Comatrix, with optional baseline diff. Bundled to
 * metacomatrix-bundled.ajs.
 */

const path = require("path");
const { getParameter } = require("./params");
const { buildComatrices, toMatrix, diffMatrix } = require("./metamodelMatrix");
const { PROP_BASELINE, PROP_COMATRIX } = require("./archi-properties");
const { hasText } = require("./text");
const metaOutput2Excel = require("./metaOutput2Excel");

/**
 * Resolves the baseline model: --baselineModel path parameter, else the model
 * property "baseline" naming a currently loaded model.
 * @returns {Object|null} The baseline model or null
 */
function resolveBaselineModel() {
  const baselineModelPath = getParameter("baselineModel");
  if (baselineModelPath) {
    console.log(`Loading baseline model from parameter: "${baselineModelPath}"`);
    try {
      const loaded = $.model.load(baselineModelPath);
      if (loaded) {
        console.log(`✓ Baseline model loaded: ${loaded.name}`);
      }
      return loaded || null;
    } catch (error) {
      console.log(`✗ Error loading baseline model: ${error.message}`);
      return null;
    }
  }

  const baselineProperty = model.prop(PROP_BASELINE);
  if (hasText(baselineProperty)) {
    console.log(`Looking for baseline model: "${baselineProperty}"`);
    const baseline = $.model.getLoadedModels().find((m) => m.name === baselineProperty);
    if (baseline) {
      console.log(`✓ Baseline model found: ${baseline.name}`);
    } else {
      console.log(`✗ Baseline model "${baselineProperty}" is not opened.`);
    }
    return baseline || null;
  }

  return null;
}

/**
 * Main execution.
 */
function runMetaComatrix() {
  console.clear();
  console.show();
  console.log("=== Meta-Comatrix - Meta-model driven Connectivity Matrix ===\n");

  if (!model) {
    console.log("ERROR: No model is selected. Please open or select a model first.");
    return;
  }
  console.log(`Selected model: ${model.name}`);

  const baselineModel = resolveBaselineModel();
  const compareMode = !!baselineModel;
  console.log(compareMode ? "\nRunning in COMPARE MODE.\n" : "\nRunning in single model mode.\n");

  console.log(`Building comatrices routed by ${PROP_COMATRIX}...`);
  const currentComatrices = buildComatrices(model);
  console.log(`Found ${currentComatrices.size} comatrix/comatrices in the current model.`);

  if (currentComatrices.size === 0) {
    console.log(`⚠ No routed relationships found. Tag meta-model relationships with the ${PROP_COMATRIX} property.`);
    return;
  }

  const baselineComatrices = compareMode ? buildComatrices(baselineModel) : new Map();
  if (compareMode) {
    console.log(`Found ${baselineComatrices.size} comatrix/comatrices in the baseline model.`);
  }

  const sheets = [];
  currentComatrices.forEach((comatrix, comatrixId) => {
    const currentMatrix = toMatrix(comatrix);
    console.log(
      `  Comatrix "${comatrix.name}" (${comatrix.kind}): ${currentMatrix.rowNames.length} source(s) x ${currentMatrix.colNames.length} target(s), ${currentMatrix.edges.size} relationship(s).`,
    );

    if (compareMode && baselineComatrices.has(comatrixId)) {
      const baselineMatrix = toMatrix(baselineComatrices.get(comatrixId));
      sheets.push(diffMatrix(currentMatrix, baselineMatrix));
      console.log(`    Diffed against baseline comatrix "${baselineMatrix.name}".`);
    } else {
      sheets.push(currentMatrix);
    }
  });

  const normalizedPath = model.path ? model.path.replace(/\\/g, "/") : null;
  const outputDir = normalizedPath ? path.dirname(normalizedPath) : __DIR__;
  const outputPath = path.join(outputDir, "comatrix-metamodel.xlsx");

  console.log(`\nWriting ${sheets.length} worksheet(s) to: ${outputPath}`);
  try {
    metaOutput2Excel(sheets, outputPath);
    console.log("=== Export Complete ===");
    console.log(`Matrix file saved to: ${outputPath}`);

    try {
      java.awt.Desktop.getDesktop().open(new java.io.File(outputDir));
    } catch (openError) {
      console.log("Could not open file browser automatically.");
    }
  } catch (error) {
    console.log(`\n✗ ERROR: Failed to create Excel file`);
    console.log(`Error message: ${error.message}`);
    if (error.stack) {
      console.log(`Stack trace:\n${error.stack}`);
    }
  }
}

runMetaComatrix();
