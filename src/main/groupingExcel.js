/**
 * groupingExcel.js
 * Excel row and column grouping functionality
 */

// Constants
const COLUMN_E = 4; // First column for B-elements (columns A-D are fixed)

/**
 * Generic function to group vectors by a key function
 * @param {Array} vectors - Array of vectors (rows or columns) to group
 * @param {Function} keyFunc - Function to extract grouping key from a vector
 * @returns {Array} Array of group objects with vectors
 */
function groupVectorsByKey(vectors, keyFunc) {
  if (vectors.length === 0) return [];

  const groups = [];
  let currentGroupKey = null;
  let currentGroupVectors = [];

  vectors.forEach((vector) => {
    const groupKey = keyFunc(vector);

    if (groupKey !== currentGroupKey) {
      // Save previous group
      if (currentGroupVectors.length > 0) {
        groups.push({
          key: currentGroupKey,
          vectors: currentGroupVectors,
        });
      }
      // Start new group
      currentGroupKey = groupKey;
      currentGroupVectors = [];
    }

    currentGroupVectors.push(vector);
  });

  // Don't forget the last group
  if (currentGroupVectors.length > 0) {
    groups.push({
      key: currentGroupKey,
      vectors: currentGroupVectors,
    });
  }

  return groups;
}

/**
 * Helper function to insert group separator rows
 * @param {Array} dataRows - Array of data rows
 * @returns {Array} Data rows with separator rows inserted
 */
function insertGroupSeparators(dataRows) {
  if (dataRows.length === 0) return [];

  // Group rows by Domain|Application
  const groups = groupVectorsByKey(dataRows, (row) => `${row[0]}|${row[1]}`);

  // Build result with separator rows containing aggregated "x" marks
  const result = [];
  groups.forEach((group) => {
    if (group.vectors.length === 1) {
      // Single row group: no separator needed, just add the row
      result.push(group.vectors[0]);
    } else {
      // Multi-row group: add separator followed by all rows
      const firstRow = group.vectors[0];
      const numColumns = firstRow.length;
      const separatorRow = [firstRow[0], firstRow[1], "", ""];

      // For each B-element column, check if any row in group has "x"
      for (let col = COLUMN_E; col < numColumns; col++) {
        let hasX = false;
        for (let row of group.vectors) {
          if (row[col] === "x") {
            hasX = true;
            break;
          }
        }
        separatorRow.push(hasX ? "x" : "");
      }

      // Add separator row
      result.push(separatorRow);

      // Add all group member rows
      group.vectors.forEach((row) => result.push(row));
    }
  });

  return result;
}

/**
 * Helper function to check if a row is a separator row
 * @param {Array} row - Data row
 * @returns {boolean} True if row is a separator
 */
function isSeparatorRow(row) {
  // Separator rows have empty Schnittstelle (column 2)
  return row[2] === "";
}

/**
 * Helper function to apply Excel row grouping for rows with same Domain and Application
 * @param {Object} worksheet - XLSX worksheet object
 * @param {Array} dataRows - Array of data rows with separators
 */
function applyRowGrouping(worksheet, dataRows) {
  // Initialize !rows array if not exists
  if (!worksheet["!rows"]) {
    worksheet["!rows"] = [];
  }

  // Track groups: separator rows mark the start of groups
  const groups = buildGroups(
    dataRows,
    isSeparatorRow,
    (vector) => ({ domain: vector[0], application: vector[1] }),
    2, // Data rows start at row 2 (0-indexed)
  );

  // Apply grouping levels to worksheet (only to non-separator rows)
  groups.forEach((group) => {
    for (let rowIdx = group.startIndex; rowIdx <= group.endIndex; rowIdx++) {
      if (!worksheet["!rows"][rowIdx]) {
        worksheet["!rows"][rowIdx] = {};
      }
      worksheet["!rows"][rowIdx].level = 1; // Set outline level 1
      worksheet["!rows"][rowIdx].hidden = true; // Initially collapsed
    }
  });

  console.log(`Applied grouping to ${groups.length} groups of rows (initially collapsed)`);
}

/**
 * Generic function to build groups from vectors (rows or columns)
 * @param {Array} dataVectors - Array of data vectors (rows or columns) to process
 * @param {Function} isSeparatorFunc - Function to check if a vector is a separator
 * @param {Function} groupingKeyFunc - Function to extract grouping key from a vector
 * @param {number} indexOffset - Offset for Excel index calculation (where first vector appears in Excel)
 * @returns {Array} Array of group objects
 */
function buildGroups(dataVectors, isSeparatorFunc, groupingKeyFunc, indexOffset = 0) {
  const groups = [];
  let currentGroup = null;

  dataVectors.forEach((vector, index) => {
    const excelIndex = index + indexOffset;

    if (isSeparatorFunc(vector)) {
      // Save previous group if it exists and has more than one element
      if (currentGroup && currentGroup.endIndex > currentGroup.startIndex) {
        groups.push(currentGroup);
      }
      // Start new group (separator is NOT part of the group)
      const groupKey = groupingKeyFunc(vector);
      currentGroup = {
        ...groupKey,
        separatorIndex: excelIndex,
        startIndex: excelIndex + 1, // Group starts after separator
        endIndex: excelIndex, // Will be updated as we find group members
      };
    } else if (currentGroup) {
      // Check if this vector belongs to the current group
      const vectorKey = groupingKeyFunc(vector);
      const keysMatch = Object.keys(vectorKey).every((k) => vectorKey[k] === currentGroup[k]);

      if (keysMatch) {
        // This is a group member
        currentGroup.endIndex = excelIndex;
      } else {
        // Different key - close current group and reset
        if (currentGroup.endIndex > currentGroup.startIndex) {
          groups.push(currentGroup);
        }
        currentGroup = null; // Single element, no group
      }
    }
    // else: single element with no current group - no action needed
  });

  // Don't forget the last group
  if (currentGroup && currentGroup.endIndex > currentGroup.startIndex) {
    groups.push(currentGroup);
  }
  return groups;
}

/**
 * Helper function to insert column separators
 * @param {Array} data - Matrix data with domain row, header row, and data rows
 * @returns {Array} Data with column separators inserted
 */
function insertColumnSeparators(data) {
  if (data.length === 0 || data[0].length <= COLUMN_E) return data;

  // Create array of column indices starting from COLUMN_E
  const columnIndices = [];
  for (let col = COLUMN_E; col < data[0].length; col++) {
    columnIndices.push(col);
  }

  // Group column indices by domain (from row 0)
  const columnGroups = groupVectorsByKey(columnIndices, (colIdx) => data[0][colIdx]);

  // Build new data structure with separator columns
  const result = [];
  for (let row of data) {
    const newRow = [...row.slice(0, COLUMN_E)];

    columnGroups.forEach((group) => {
      if (group.vectors.length === 1) {
        // Single column group: no separator needed
        newRow.push(row[group.vectors[0]]);
      } else {
        // Multi-column group: add separator followed by all columns
        // Separator column shows aggregated data
        let separatorValue = "";

        if (row === data[0]) {
          // Row 0 (domain row): show domain name
          separatorValue = group.key;
        } else if (row === data[1]) {
          // Row 1 (header row): empty
          separatorValue = "";
        } else {
          // Data rows: aggregate "x" marks
          let hasX = false;
          for (let colIdx of group.vectors) {
            if (row[colIdx] === "x") {
              hasX = true;
              break;
            }
          }
          separatorValue = hasX ? "x" : "";
        }

        newRow.push(separatorValue);

        // Add all group member columns
        group.vectors.forEach((colIdx) => newRow.push(row[colIdx]));
      }
    });

    result.push(newRow);
  }

  return result;
}

/**
 * Helper function to check if a column is a separator column
 * @param {Array} data - Matrix data
 * @param {number} col - Column index
 * @returns {boolean} True if column is a separator
 */
function isSeparatorColumn(data, col) {
  // Separator columns have domain name in row 0 and empty in row 1
  // and the next column also has a value in row 1 (it's not the end)
  if (col < COLUMN_E || col >= data[1].length - 1) return false;

  const row0Value = data[0][col];
  const row1Value = data[1][col];
  const row1NextValue = data[1][col + 1];

  // Separator if row 1 is empty, row 0 has domain, and next column in row 1 has a value
  return row1Value === "" && row0Value !== "" && row1NextValue !== "";
}

/**
 * Helper function to apply column grouping
 * @param {Object} worksheet - XLSX worksheet object
 * @param {Array} data - Matrix data with separators
 */
function applyColumnGrouping(worksheet, data) {
  // Initialize !cols array if not exists
  if (!worksheet["!cols"]) {
    worksheet["!cols"] = [];
  }

  // Transpose columns into column vectors
  const columnVectors = [];
  for (let col = COLUMN_E; col < data[0].length; col++) {
    columnVectors.push(data.map((row) => row[col]));
  }

  // Use buildGroups to find column groups
  const columnGroups = buildGroups(
    columnVectors,
    (colVector) => colVector[1] === "" && colVector[0] !== "",
    (colVector) => ({ domain: colVector[0] }),
    COLUMN_E,
  );

  // Apply grouping levels to worksheet (only to non-separator columns)
  columnGroups.forEach((group) => {
    for (let colIdx = group.startIndex; colIdx <= group.endIndex; colIdx++) {
      if (!worksheet["!cols"][colIdx]) {
        worksheet["!cols"][colIdx] = {};
      }
      worksheet["!cols"][colIdx].level = 1; // Set outline level 1
      worksheet["!cols"][colIdx].hidden = true; // Initially collapsed
    }
  });

  console.log(`Applied grouping to ${columnGroups.length} groups of columns (initially collapsed)`);
}

module.exports = {
  insertGroupSeparators,
  isSeparatorRow,
  applyRowGrouping,
  insertColumnSeparators,
  isSeparatorColumn,
  applyColumnGrouping,
};
