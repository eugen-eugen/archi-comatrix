/**
 * model.js
 * Functions for extracting and analyzing Archi model elements
 */

/**
 * Helper function to find all groupings with a specific specialization
 * @param {Object} element - The Archi element
 * @param {String} targetSpecialization - The specialization to look for (e.g., "Domäne", "Fachbereich")
 * @param {Set} visited - Set of already visited element IDs to detect cycles
 * @param {Object} currentPath - Set of elements in current traversal path for cycle detection
 * @returns {Object} Object with {found: Array<String>, hasCycle: Boolean}
 */
function findAllGroupings(element, targetSpecialization, visited, currentPath) {
  const result = { found: [], hasCycle: false };

  // Check if we've encountered a cycle in the current path
  if (currentPath.has(element.id)) {
    result.hasCycle = true;
    return result;
  }

  // Check if this element itself has the target specialization
  if (element.specialization === targetSpecialization) {
    result.found.push(element.name);
  }

  // Mark as visited
  visited.add(element.id);
  currentPath.add(element.id);

  // Find incoming aggregation or composition relationships
  const incomingRels = $(element)
    .inRels("aggregation-relationship")
    .add($(element).inRels("composition-relationship"));

  // Traverse up the hierarchy
  for (let i = 0; i < incomingRels.length; i++) {
    const rel = incomingRels[i];
    const parent = rel.source;

    // Skip if already visited (but not in current path - that would be a cycle)
    if (visited.has(parent.id) && !currentPath.has(parent.id)) {
      continue;
    }

    // Recursively search in parent
    const parentResult = findAllGroupings(
      parent,
      targetSpecialization,
      visited,
      currentPath,
    );

    if (parentResult.hasCycle) {
      result.hasCycle = true;
    }

    // Add found groupings from parent
    parentResult.found.forEach((name) => {
      if (!result.found.includes(name)) {
        result.found.push(name);
      }
    });
  }

  // Remove from current path when backtracking
  currentPath.delete(element.id);

  return result;
}

/**
 * Finds all domains for an element by traversing aggregation/composition relationships
 * @param {Object} element - The Archi element
 * @returns {String} Comma-separated domain names, "cycle" if cycle detected, or empty string if none found
 */
function findDomain(element) {
  const visited = new Set();
  const currentPath = new Set();
  const result = findAllGroupings(element, "Domäne", visited, currentPath);

  if (result.hasCycle) {
    return "cycle";
  }

  if (result.found.length === 0) {
    return "";
  }

  return result.found.join(", ");
}

/**
 * Finds all Fachbereich for an element by traversing aggregation/composition relationships
 * @param {Object} element - The Archi element
 * @returns {String} Comma-separated Fachbereich names, "cycle" if cycle detected, or empty string if none found
 */
function findFachbereich(element) {
  const visited = new Set();
  const currentPath = new Set();
  const result = findAllGroupings(element, "Fachbereich", visited, currentPath);

  if (result.hasCycle) {
    return "cycle";
  }

  if (result.found.length === 0) {
    return "";
  }

  return result.found.join(", ");
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

module.exports = {
  findDomain,
  findFachbereich,
  extractElements,
};
