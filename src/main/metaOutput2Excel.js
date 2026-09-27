/**
 * metaOutput2Excel.js
 * Writes meta-model connectivity matrices to a single workbook, one worksheet per
 * meta-model. Square element x element layout, Domäne grouping, optional diff
 * coloring (green = added, red = removed).
 */

const XLSX = require("xlsx-js-style");

const COLOR_ADDED = "9BBB59";
const COLOR_CHANGED = "FFC000";
const COLOR_REMOVED = "C0504D";
const COLOR_HEADER_GRAY = "D9D9D9";
const COLOR_COLUMN_BLUE = "B8CCE4";
const COLOR_SEPARATOR = "808080";
const COLOR_TEXT_BLACK = "000000";
const COLOR_TEXT_WHITE = "FFFFFF";

const BORDER = {
  top: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
  bottom: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
  left: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
  right: { style: "thin", color: { rgb: COLOR_TEXT_BLACK } },
};

/**
 * Builds a display sequence with a separator between Domäne groups.
 * @param {Array<String>} names - Sorted element names
 * @param {Map<String,String>} nameDomain - name -> domain
 * @returns {Array<Object>} Items { name } or { sep: true }
 */
function withSeparators(names, nameDomain) {
  const seq = [];
  let prevDomain = null;

  names.forEach((name, index) => {
    const domain = nameDomain.get(name) || "";
    if (index > 0 && domain !== prevDomain) {
      seq.push({ sep: true });
    }
    seq.push({ name: name });
    prevDomain = domain;
  });

  return seq;
}

/**
 * Creates a styled string cell.
 */
function cell(value, style) {
  return { v: value, t: "s", s: style };
}

/**
 * Resolves the fill/text style for a data cell given its diff state.
 */
function markStyle(state) {
  const style = {
    font: { name: "Calibri", sz: 11, color: { rgb: COLOR_TEXT_BLACK } },
    alignment: { horizontal: "center" },
    border: BORDER,
  };

  if (state === "added") {
    style.fill = { fgColor: { rgb: COLOR_ADDED } };
  } else if (state === "removed") {
    style.fill = { fgColor: { rgb: COLOR_REMOVED } };
    style.font.color = { rgb: COLOR_TEXT_WHITE };
  }
  return style;
}

/**
 * Resolves the header style for a row/col element name given its diff state.
 */
function headerNameStyle(state, rotate) {
  const style = {
    font: { name: "Calibri", sz: 11, color: { rgb: COLOR_TEXT_BLACK }, bold: true },
    alignment: rotate ? { horizontal: "center", textRotation: 90, wrapText: true } : { horizontal: "left" },
    border: BORDER,
    fill: { fgColor: { rgb: COLOR_COLUMN_BLUE } },
  };

  if (state === "added") {
    style.fill = { fgColor: { rgb: COLOR_ADDED } };
  } else if (state === "changed") {
    style.fill = { fgColor: { rgb: COLOR_CHANGED } };
  } else if (state === "removed") {
    style.fill = { fgColor: { rgb: COLOR_REMOVED } };
    style.font.color = { rgb: COLOR_TEXT_WHITE };
  }
  return style;
}

const SEP_STYLE = {
  border: BORDER,
  fill: { fgColor: { rgb: COLOR_SEPARATOR } },
};

const GRAY_STYLE = {
  font: { name: "Calibri", sz: 11, color: { rgb: COLOR_TEXT_BLACK }, bold: true },
  alignment: { horizontal: "center" },
  border: BORDER,
  fill: { fgColor: { rgb: COLOR_HEADER_GRAY } },
};

/**
 * Colored legend cell style (green = added, orange = changed, red = removed).
 */
function legendStyle(color) {
  return {
    font: {
      name: "Calibri",
      sz: 11,
      color: { rgb: color === COLOR_REMOVED ? COLOR_TEXT_WHITE : COLOR_TEXT_BLACK },
      bold: true,
    },
    alignment: { horizontal: "center" },
    border: BORDER,
    fill: { fgColor: { rgb: color } },
  };
}

/**
 * Builds a worksheet object for one matrix/diff sheet.
 */
function buildWorksheet(sheet) {
  const isDiff = sheet.diff === true;
  const domainOf = (name) => sheet.nameDomain.get(name) || "";
  const rowStateOf = (name) => (isDiff ? sheet.rowState.get(name) : "same");
  const colStateOf = (name) => (isDiff ? sheet.colState.get(name) : "same");

  const edgeStateOf = (source, target) => {
    const key = source + "\u0000" + target;
    if (isDiff) {
      return sheet.edgeState.has(key) ? sheet.edgeState.get(key) : null;
    }
    return sheet.edges.has(key) ? "same" : null;
  };

  const colSeq = withSeparators(sheet.colNames, sheet.nameDomain);
  const rowSeq = withSeparators(sheet.rowNames, sheet.nameDomain);
  const totalCols = 2 + colSeq.length;
  const grid = [];

  /*  Legend row (diff sheets only), colored like the classic comatrix.  */
  if (isDiff) {
    const legendRow = [];
    legendRow.push(cell("Grün = Hinzugefügt", legendStyle(COLOR_ADDED)));
    legendRow.push(cell("Orange = geändert", legendStyle(COLOR_CHANGED)));
    legendRow.push(cell("Rot = gelöscht", legendStyle(COLOR_REMOVED)));
    for (let i = 3; i < totalCols; i++) {
      legendRow.push(cell("", GRAY_STYLE));
    }
    grid.push(legendRow);
  }

  /*  Domain row: target element domains.  */
  const domainRow = [];
  domainRow.push(cell("", GRAY_STYLE));
  domainRow.push(cell("", GRAY_STYLE));
  colSeq.forEach((item) => {
    domainRow.push(item.sep ? cell("", SEP_STYLE) : cell(domainOf(item.name), GRAY_STYLE));
  });
  grid.push(domainRow);

  /*  Row 1: header with target element names.  */
  const headerRow = [];
  headerRow.push(cell("Domäne", GRAY_STYLE));
  headerRow.push(cell("Quelle \\ Ziel", GRAY_STYLE));
  colSeq.forEach((item) => {
    headerRow.push(item.sep ? cell("", SEP_STYLE) : cell(item.name, headerNameStyle(colStateOf(item.name), true)));
  });
  grid.push(headerRow);

  /*  Data rows.  */
  rowSeq.forEach((rowItem) => {
    const row = [];
    if (rowItem.sep) {
      const width = 2 + colSeq.length;
      for (let i = 0; i < width; i++) {
        row.push(cell("", SEP_STYLE));
      }
      grid.push(row);
      return;
    }

    const rowName = rowItem.name;
    row.push(cell(domainOf(rowName), GRAY_STYLE));
    row.push(cell(rowName, headerNameStyle(rowStateOf(rowName), false)));

    colSeq.forEach((colItem) => {
      if (colItem.sep) {
        row.push(cell("", SEP_STYLE));
        return;
      }
      const state = edgeStateOf(rowName, colItem.name);
      row.push(state ? cell("x", markStyle(state)) : cell("", markStyle(null)));
    });
    grid.push(row);
  });

  /*  Assemble worksheet from the cell grid.  */
  const ws = {};
  const numRows = grid.length;
  const numCols = 2 + colSeq.length;
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const ref = XLSX.utils.encode_cell({ r: r, c: c });
      ws[ref] = grid[r][c];
    }
  }
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: numRows - 1, c: numCols - 1 } });

  const cols = [{ wch: 20 }, { wch: 32 }];
  colSeq.forEach((item) => {
    cols.push(item.sep ? { wch: 2 } : { wch: 4 });
  });
  ws["!cols"] = cols;

  const ySplit = isDiff ? 3 : 2;
  ws["!freeze"] = {
    xSplit: 2,
    ySplit: ySplit,
    topLeftCell: XLSX.utils.encode_cell({ r: ySplit, c: 2 }),
    activePane: "bottomRight",
    state: "frozen",
  };

  return ws;
}

/**
 * Sanitizes and de-duplicates an Excel worksheet name (<= 31 chars, no []:*?/\).
 */
function sheetNameFor(rawName, used) {
  let name = (rawName || "Metamodell").replace(/[\[\]\:\*\?\/\\]/g, " ").trim();
  if (name.length === 0) {
    name = "Metamodell";
  }
  if (name.length > 31) {
    name = name.substring(0, 31);
  }

  let candidate = name;
  let counter = 2;
  while (used.has(candidate)) {
    const suffix = " (" + counter + ")";
    candidate = name.substring(0, 31 - suffix.length) + suffix;
    counter++;
  }
  used.add(candidate);
  return candidate;
}

/**
 * Writes all matrix/diff sheets to a single workbook file.
 * @param {Array<Object>} sheets - Matrix or diff sheet models
 * @param {String} outputPath - Destination .xlsm path
 */
function metaOutput2Excel(sheets, outputPath) {
  const workbook = XLSX.utils.book_new();
  const used = new Set();

  sheets.forEach((sheet) => {
    const ws = buildWorksheet(sheet);
    XLSX.utils.book_append_sheet(workbook, ws, sheetNameFor(sheet.name, used));
  });

  const excelBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsm", cellStyles: true });

  const FileOutputStream = Java.type("java.io.FileOutputStream");
  const fos = new FileOutputStream(outputPath, false);
  const javaBytes = Java.to(Array.from(excelBuffer), "byte[]");
  fos.write(javaBytes);
  fos.close();
}

module.exports = metaOutput2Excel;
