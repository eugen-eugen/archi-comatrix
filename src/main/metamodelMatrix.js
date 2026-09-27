/**
 * metamodelMatrix.js
 * Meta-model driven connectivity matrix: gathers meta-models across a model's
 * views and builds a source x target adjacency matrix per comatrix (rows = source
 * elements, cols = target elements), plus a diff between two such matrices
 * (baseline vs current).
 */

const {
  extractMetaModel,
  findMetaModelMatches,
  findMetaModelRelationshipMatches,
  isRelationshipAllowed,
} = require("@eugen-eugen/archi-metamodel");
const { findDomain } = require("./model");
const { PROP_COMATRIX } = require("./archi-properties");

/*  Field separator for directed-edge keys "source<sep>target".  */
const EDGE_SEP = "\u0000";

/**
 * Sorts element names by domain (empty domain last), then by name.
 * @param {Array<String>} names - Element names
 * @param {Map<String,String>} nameDomain - name -> domain
 * @returns {Array<String>} Sorted names
 */
function sortByDomain(names, nameDomain) {
  return names.slice().sort((name1, name2) => {
    const domain1 = nameDomain.get(name1) || "";
    const domain2 = nameDomain.get(name2) || "";

    if ((domain1 === "") !== (domain2 === "")) {
      return domain1 === "" ? 1 : -1;
    }
    if (domain1 !== domain2) {
      return domain1.localeCompare(domain2);
    }
    return name1.localeCompare(name2);
  });
}

/**
 * Gathers every distinct meta-model reachable from a model's views.
 * Identity of a meta-model is the id of its metaView (embedded: the view itself;
 * linked: the referenced meta view). Subject concepts are the real elements shown
 * on the contributing views that correspond to a meta-model element type.
 * @param {Object} currentModel - The Archi model to scan
 * @returns {Map<String,Object>} metaViewId -> entry { id, name, mm, views, concepts }
 */
function gatherMetamodels(currentModel) {
  const metamodels = new Map();

  $(currentModel)
    .find("view")
    .each((view) => {
      const mm = extractMetaModel(view);
      if (!mm) {
        return;
      }

      const metaViewId = mm.metaView.id;
      let entry = metamodels.get(metaViewId);
      if (!entry) {
        entry = {
          id: metaViewId,
          name: mm.metaView.name,
          mm: mm,
          views: [],
          concepts: new Map(),
        };
        metamodels.set(metaViewId, entry);
      }
      entry.views.push(view);

      /*  Collect real elements shown on this contributing view.  */
      $(view)
        .find("element")
        .each((visual) => {
          const concept = visual.concept;
          if (!concept) {
            return;
          }

          /*  Drop the meta-model definition elements themselves.  */
          if (entry.mm.elementIds[concept.id]) {
            return;
          }

          /*  Keep only elements corresponding to a meta-model element type.  */
          if (findMetaModelMatches(concept, entry.mm).length === 0) {
            return;
          }
          entry.concepts.set(concept.id, concept);
        });
    });

  return metamodels;
}

/**
 * Reads the Report:Model:Comatrix routing of a meta-model relationship.
 * @param {Object} mmRelationship - A meta-model relationship concept
 * @returns {Object|null} { id, name, kind } or null when the property key is absent
 */
function comatrixRouteOf(mmRelationship, mm) {
  const keys = mmRelationship.prop();
  if (!keys || keys.indexOf(PROP_COMATRIX) < 0) {
    console.log(`✗ Meta-model relationship "${mmRelationship.name}" does not have the ${PROP_COMATRIX} property.`);
    return null;
  }

  const raw = mmRelationship.prop(PROP_COMATRIX);
  const value = (raw === undefined || raw === null ? "" : String(raw)).trim();
  if (value === "") {
    return { id: "mm:" + mm.metaView.id, name: mm.metaView.name, kind: "default" };
  }
  return { id: value, name: value, kind: "named" };
}

/**
 * Builds all comatrices for a model by routing each allowed direct relationship to
 * the comatrix (or comatrices) named on its matching meta-model relationship(s).
 * - key absent on the meta-model relationship  -> the relationship is ignored,
 * - key present but empty                       -> default comatrix (metaView id + name),
 * - key present with a value                    -> named comatrix (value = id and name),
 *   merged globally across meta-models.
 * Rows/cols contain only elements incident to a routed relationship.
 * @param {Object} currentModel - The Archi model
 * @returns {Map<String,Object>} comatrixId -> { id, name, kind, rowNames, colNames, edges, nameDomain }
 */
function buildComatrices(currentModel) {
  const metamodels = gatherMetamodels(currentModel);
  const comatrices = new Map();

  const ensure = (route) => {
    let comatrix = comatrices.get(route.id);
    if (!comatrix) {
      comatrix = {
        id: route.id,
        name: route.name,
        kind: route.kind,
        rowNames: new Set(),
        colNames: new Set(),
        edges: new Set(),
        nameDomain: new Map(),
      };
      comatrices.set(route.id, comatrix);
    }
    return comatrix;
  };

  metamodels.forEach((entry) => {
    const mm = entry.mm;
    console.log(`Processing meta-model: "${mm.metaView.name}"`);

    entry.concepts.forEach((concept) => {
      $(concept)
        .outRels()
        .each((rel) => {
          const target = rel.target;
          if (!entry.concepts.has(target.id)) {
            return;
          }
          if (!isRelationshipAllowed(rel.source, rel.target, rel, mm)) {
            return;
          }

          const matches = findMetaModelRelationshipMatches(rel, mm);
          if (!matches || matches.length === 0) {
            return;
          }

          const sourceName = rel.source.name;
          const targetName = target.name;

          matches.forEach((mmRel) => {
            const route = comatrixRouteOf(mmRel, mm);
            if (!route) {
              return;
            }

            const comatrix = ensure(route);
            comatrix.edges.add(sourceName + EDGE_SEP + targetName);
            comatrix.rowNames.add(sourceName);
            comatrix.colNames.add(targetName);
            if (!comatrix.nameDomain.has(sourceName)) {
              comatrix.nameDomain.set(sourceName, findDomain(rel.source));
            }
            if (!comatrix.nameDomain.has(targetName)) {
              comatrix.nameDomain.set(targetName, findDomain(target));
            }
          });
        });
    });
  });

  return comatrices;
}

/**
 * Finalizes a comatrix into a matrix sheet with Domäne-sorted rows (sources) and
 * columns (targets).
 * @param {Object} comatrix - A value from buildComatrices()
 * @returns {Object} Matrix { id, name, rowNames, colNames, nameDomain, edges }
 */
function toMatrix(comatrix) {
  const rowNames = sortByDomain(Array.from(comatrix.rowNames), comatrix.nameDomain);
  const colNames = sortByDomain(Array.from(comatrix.colNames), comatrix.nameDomain);
  return {
    id: comatrix.id,
    name: comatrix.name,
    rowNames,
    colNames,
    nameDomain: comatrix.nameDomain,
    edges: comatrix.edges,
  };
}

/**
 * Diffs a current matrix against a baseline matrix (same comatrix id).
 * Rows (sources) and columns (targets) are diffed on their own axes.
 * @param {Object} current - Matrix from toMatrix() (current model)
 * @param {Object} baseline - Matrix from toMatrix() (baseline model)
 * @returns {Object} Diff sheet { id, name, rowNames, colNames, nameDomain, rowState, colState, edgeState, diff }
 */
function diffMatrix(current, baseline) {
  const axisState = (currentAxis, baselineAxis) => {
    const cur = new Set(currentAxis);
    const base = new Set(baselineAxis);
    const all = new Set([...currentAxis, ...baselineAxis]);
    const state = new Map();
    all.forEach((name) => {
      const inCurrent = cur.has(name);
      const inBaseline = base.has(name);
      state.set(name, inCurrent && inBaseline ? "same" : inCurrent ? "added" : "removed");
    });
    return { all, state };
  };

  const rows = axisState(current.rowNames, baseline.rowNames);
  const cols = axisState(current.colNames, baseline.colNames);

  const nameDomain = new Map();
  new Set([...rows.all, ...cols.all]).forEach((name) => {
    nameDomain.set(name, current.nameDomain.has(name) ? current.nameDomain.get(name) : baseline.nameDomain.get(name));
  });

  const edgeState = new Map();
  const allEdges = new Set([...current.edges, ...baseline.edges]);
  allEdges.forEach((edge) => {
    const inCurrent = current.edges.has(edge);
    const inBaseline = baseline.edges.has(edge);
    edgeState.set(edge, inCurrent && inBaseline ? "same" : inCurrent ? "added" : "removed");
  });

  /*  Source/target elements present in both models but with a changed relationship.  */
  const changedSources = new Set();
  const changedTargets = new Set();
  edgeState.forEach((state, edge) => {
    if (state === "same") {
      return;
    }
    const parts = edge.split(EDGE_SEP);
    changedSources.add(parts[0]);
    changedTargets.add(parts[1]);
  });
  rows.state.forEach((state, name) => {
    if (state === "same" && changedSources.has(name)) {
      rows.state.set(name, "changed");
    }
  });
  cols.state.forEach((state, name) => {
    if (state === "same" && changedTargets.has(name)) {
      cols.state.set(name, "changed");
    }
  });

  const rowNames = sortByDomain(Array.from(rows.all), nameDomain);
  const colNames = sortByDomain(Array.from(cols.all), nameDomain);

  return {
    id: current.id,
    name: current.name,
    rowNames,
    colNames,
    nameDomain,
    rowState: rows.state,
    colState: cols.state,
    edgeState,
    diff: true,
  };
}

module.exports = {
  EDGE_SEP,
  gatherMetamodels,
  buildComatrices,
  toMatrix,
  diffMatrix,
};
