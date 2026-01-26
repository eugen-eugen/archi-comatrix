/**
 * model.js
 * Functions for extracting and analyzing Archi model elements
 */

/**
 * Finds the domain for an element by traversing aggregation/composition relationships
 * @param {Object} element - The Archi element
 * @returns {String} The domain name or empty string if not found
 */
function findDomain(element) {
  // Check if this element's specialization is "Domäne"
  const specialization = element.specialization;
  if (specialization === "Domäne") {
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

module.exports = {
  findDomain,
  extractElements,
};
