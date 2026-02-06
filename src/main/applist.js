/**
 * applist.js
 * Generates Excel list of business applications with their domains and Fachbereich
 * Supports multiple domains/fachbereich per application (comma-separated)
 */

const XLSX = require("xlsx-js-style");
const path = require("path");
const { findDomain, findFachbereich } = require("./model");

// Color constants matching comatrix style
const COLOR_HEADER_GRAY = "D9D9D9";
const COLOR_TEXT_BLACK = "000000";

/**
 * Gets font style for Excel cells
 * @param {Boolean} bold - Whether text should be bold
 * @returns {Object} Font style object
 */
function getFontStyle(bold = false) {
  const style = {
    name: "Calibri",
    sz: 11,
    color: { rgb: COLOR_TEXT_BLACK },
  };
  if (bold) {
    style.bold = true;
  }
  return style;
}

/**
 * Generates Excel file with list of business applications and their classifications
 * @param {Array} applications - Array of {name, typ, domain, fachbereich} objects
 *                               domain and fachbereich can be comma-separated lists or "cycle"
 * @param {String} outputPath - Path where Excel file should be saved
 */
function generateAppListExcel(applications, outputPath) {
  console.log("Creating Excel workbook...");

  // Create workbook and worksheet
  const workbook = XLSX.utils.book_new();
  const worksheetData = [];

  // Define border style
  const borderStyle = {
    top: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
    bottom: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
    left: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
    right: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
  };

  // Header row
  worksheetData.push(["Anwendung", "Typ", "Domäne", "Fachbereich"]);

  // Data rows
  applications.forEach((app) => {
    worksheetData.push([app.name, app.typ, app.domain, app.fachbereich]);
  });

  // Create worksheet from data
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

  // Apply styles
  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  // Header style
  const headerStyle = {
    font: getFontStyle(true),
    alignment: { horizontal: "center" },
    border: borderStyle,
    fill: { fgColor: { rgb: COLOR_HEADER_GRAY } },
  };

  // Data style
  const dataStyle = {
    font: getFontStyle(),
    border: borderStyle,
  };

  // Apply header style to first row
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
    if (worksheet[cellRef]) {
      worksheet[cellRef].s = headerStyle;
    }
  }

  // Apply data style to all data rows
  for (let row = 1; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
      if (worksheet[cellRef]) {
        worksheet[cellRef].s = dataStyle;
      }
    }
  }

  // Set column widths (wider to accommodate multiple comma-separated values)
  worksheet["!cols"] = [
    { wch: 40 }, // Anwendung column
    { wch: 25 }, // Typ column
    { wch: 40 }, // Domäne column (wider for multiple domains)
    { wch: 40 }, // Fachbereich column (wider for multiple Fachbereich)
  ];

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, "Anwendungen");

  console.log("Writing Excel file...");

  // Write to file using Java FileOutputStream for Archi/GraalVM compatibility
  const FileOutputStream = Java.type("java.io.FileOutputStream");
  const output = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  const fos = new FileOutputStream(outputPath);
  fos.write(output);
  fos.close();

  console.log(`✓ Excel file created: ${outputPath}`);
}

/**
 * Main execution function
 */
function runAppList() {
  console.clear();
  console.show();
  console.log("=== AppList - Generate Application List ===\n");

  // Check if a model is selected
  if (!model) {
    console.log(
      "ERROR: No model is selected. Please open or create a model first.",
    );
    return;
  }

  console.log(`Selected model: ${model.name}`);
  console.log(`Model path: ${model.path || "(not saved)"}\n`);

  try {
    // Find all application components with specific specializations
    console.log(
      'Searching for application components with specializations: "Geschäftsanwendung", "Register", "Querschnittsanwendung"...',
    );

    const allAppComponents = $("application-component");
    console.log(
      `Total application components found: ${allAppComponents.length}`,
    );

    const applications = [];
    const allowedSpecializations = [
      "Geschäftsanwendung",
      "Register",
      "Querschnittsanwendung",
    ];

    allAppComponents.each((appComponent) => {
      const spec = appComponent.specialization;
      if (allowedSpecializations.includes(spec)) {
        const domain = findDomain(appComponent);
        const fachbereich = findFachbereich(appComponent);
        applications.push({
          name: appComponent.name,
          typ: spec,
          domain: domain || "(keine Domäne)",
          fachbereich: fachbereich || "(kein Fachbereich)",
        });
      }
    });

    console.log(`Found ${applications.length} applications\n`);

    if (applications.length === 0) {
      console.log(
        "⚠ No application components with the specified specializations found.",
      );
      return;
    }

    // Sort by fachbereich, then domain, then typ, then name
    applications.sort((a, b) => {
      // Empty Fachbereich go to the end
      if (
        a.fachbereich === "(kein Fachbereich)" &&
        b.fachbereich !== "(kein Fachbereich)"
      )
        return 1;
      if (
        a.fachbereich !== "(kein Fachbereich)" &&
        b.fachbereich === "(kein Fachbereich)"
      )
        return -1;

      // Compare Fachbereich
      if (a.fachbereich !== b.fachbereich) {
        return a.fachbereich.localeCompare(b.fachbereich);
      }

      // Empty domains go to the end
      if (a.domain === "(keine Domäne)" && b.domain !== "(keine Domäne)")
        return 1;
      if (a.domain !== "(keine Domäne)" && b.domain === "(keine Domäne)")
        return -1;

      // Compare domains
      if (a.domain !== b.domain) {
        return a.domain.localeCompare(b.domain);
      }

      // Compare typ
      if (a.typ !== b.typ) {
        return a.typ.localeCompare(b.typ);
      }

      // Same fachbereich, domain, and typ - sort by name
      return a.name.localeCompare(b.name);
    });

    // Print summary
    console.log("=== Applications by Fachbereich and Domain ===");
    let currentFachbereich = null;
    let currentDomain = null;
    applications.forEach((app) => {
      if (app.fachbereich !== currentFachbereich) {
        currentFachbereich = app.fachbereich;
        console.log(`\n${currentFachbereich}:`);
        currentDomain = null; // Reset domain when fachbereich changes
      }
      if (app.domain !== currentDomain) {
        currentDomain = app.domain;
        console.log(`  ${currentDomain}:`);
      }
      console.log(`    - ${app.name} (${app.typ})`);
    });
    console.log("");

    // Define output path
    const normalizedPath = model.path ? model.path.replace(/\\/g, "/") : null;
    const outputDir = normalizedPath ? path.dirname(normalizedPath) : __DIR__;
    const outputPath = path.join(outputDir, "applist.xlsx");

    console.log(`Output file: ${outputPath}\n`);

    // Generate Excel file
    generateAppListExcel(applications, outputPath);

    console.log("\n=== Export Complete ===");
    console.log(`Application list saved to: ${outputPath}`);
    console.log(`Total applications: ${applications.length}`);

    // Count by type
    const countByType = {};
    applications.forEach((app) => {
      countByType[app.typ] = (countByType[app.typ] || 0) + 1;
    });
    console.log(
      `  - Geschäftsanwendung: ${countByType["Geschäftsanwendung"] || 0}`,
    );
    console.log(`  - Register: ${countByType["Register"] || 0}`);
    console.log(
      `  - Querschnittsanwendung: ${countByType["Querschnittsanwendung"] || 0}`,
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
runAppList();
