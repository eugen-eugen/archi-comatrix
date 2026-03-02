/**
 * tgf.js
 * Exports Archi model to Trivial Graph Format (TGF)
 * TGF Format:
 *   - Nodes: node_id node_label
 *   - Separator: #
 *   - Edges: source_id target_id edge_label
 */

const path = require("path");

/**
 * Generates TGF file from Archi model
 * @param {Object} model - The Archi model to export
 * @param {String} outputPath - Path where TGF file should be saved
 */
function generateTGF(model, outputPath) {
  console.log("Generating TGF file...");

  const lines = [];
  const nodeIds = new Set();

  // Step 1: Collect all elements (nodes) from the model
  const elements = $(model).find("element").not("relationship");

  console.log(`Found ${elements.length} elements`);

  // Add nodes section
  elements.each((element) => {
    const id = element.id;
    const name = element.name || "(unnamed)";
    const type = element.type;

    nodeIds.add(id);
    // Format: id label (include type for clarity)
    lines.push(`${id} ${name} [${type}]`);
  });

  // Step 2: Add separator
  lines.push("#");

  // Step 3: Collect all relationships (edges)
  const relationships = $(model).find("relationship");

  console.log(`Found ${relationships.length} relationships`);

  let edgeCount = 0;
  relationships.each((rel) => {
    const sourceId = rel.source.id;
    const targetId = rel.target.id;
    const label = rel.name || rel.type;

    // Only include edges where both source and target are in our node set
    if (nodeIds.has(sourceId) && nodeIds.has(targetId)) {
      lines.push(`${sourceId} ${targetId} ${label}`);
      edgeCount++;
    }
  });

  console.log(`Exported ${edgeCount} edges`);

  // Step 4: Write to file
  const content = lines.join("\n") + "\n";

  try {
    const FileWriter = Java.type("java.io.FileWriter");
    const writer = new FileWriter(outputPath);
    writer.write(content);
    writer.close();

    console.log(`✓ TGF file written successfully`);
    return true;
  } catch (error) {
    console.log(`✗ ERROR: Failed to write file: ${error.message}`);
    return false;
  }
}

/**
 * Main function to execute TGF export
 */
function runTGF() {
  try {
    console.log("=== TGF Export ===\n");

    // Get the currently selected model
    const model = selection.filter("archimate-model").first();

    if (!model) {
      console.log("⚠ Please select an ArchiMate model in the model tree before running this script.");
      return;
    }

    console.log(`Selected model: ${model.name}`);

    // Define output path
    const normalizedPath = model.path ? model.path.replace(/\\/g, "/") : null;
    const outputDir = normalizedPath ? path.dirname(normalizedPath) : __DIR__;
    const outputPath = path.join(outputDir, "graph.tgf");

    console.log(`Output file: ${outputPath}\n`);

    // Generate TGF file
    const success = generateTGF(model, outputPath);

    if (success) {
      console.log("\n=== Export Complete ===");
      console.log(`TGF file saved to: ${outputPath}`);

      // Open the file location in file browser
      try {
        java.awt.Desktop.getDesktop().open(new java.io.File(outputDir));
      } catch (e) {
        console.log("Could not open file browser automatically.");
      }
    }
  } catch (error) {
    console.log(`\n✗ ERROR: Failed to create TGF file`);
    console.log(`Error message: ${error.message}`);
    console.log(`Error type: ${error.constructor.name}`);
    if (error.stack) {
      console.log(`Stack trace:\n${error.stack}`);
    }
  }
}

// Execute the main function when script is loaded
runTGF();
