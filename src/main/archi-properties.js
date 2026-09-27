/**
 * archi-properties.js
 * Central definitions of the Archi property keys and specialization names that
 * drive connectivity-matrix (comatrix) generation. Shared by both the classic
 * comatrix scripts and the meta-model driven metacomatrix.
 */

/*
 *  Custom property, set on a triggering relationship, whose value(s) name the
 *  interface ("Schnittstelle") the relationship represents. Multiple values are
 *  allowed; each becomes a separate row group in the classic comatrix.
 */
const PROP_SCHNITTSTELLE = "Schnittstelle";

/*
 *  Model-level property naming the baseline model to compare against. Used to
 *  enable COMPARE MODE when no --baselineModel parameter is supplied.
 */
const PROP_BASELINE = "baseline";

/*
 *  Property on a meta-model relationship that routes it to a named comatrix.
 *  Key absent  -> the relationship is ignored,
 *  value empty -> default comatrix (meta-model identity),
 *  value set   -> named comatrix (value = identity and worksheet name).
 */
const PROP_COMATRIX = "Report:Model:Comatrix";

/*
 *  Specialization of a grouping that represents a domain ("Domäne"). An element's
 *  domain is derived by traversing aggregation/composition up to such a grouping.
 */
const SPECIALIZATION_DOMAENE = "Domäne";

/*
 *  Specialization of a grouping that represents a "Fachbereich" (business area).
 *  Derived the same way as the domain.
 */
const SPECIALIZATION_FACHBEREICH = "Fachbereich";

module.exports = {
  PROP_SCHNITTSTELLE,
  PROP_BASELINE,
  PROP_COMATRIX,
  SPECIALIZATION_DOMAENE,
  SPECIALIZATION_FACHBEREICH,
};
