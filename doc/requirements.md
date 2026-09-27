# Meta-Comatrix — Requirements

Metamodel-driven connectivity matrix for Archi, implemented as a new script
(`metacomatrix`) alongside the existing `comatrix`. It replaces the hard-coded
"triggering relationships with a `Schnittstelle` property" approach with a flexible
approach driven by **meta-models**.

Consumes the `@eugen-eugen/archi-metamodel` package (bundled at build time).

## 1. Scope

- New, separate entry point `src/main/metacomatrix.js` → `dist/metacomatrix-bundled.ajs`.
- The existing `comatrix` / `applist` / `tgf` scripts are left unchanged.

## 2. Meta-model discovery

- Iterate over **all views** of the model and call `extractMetaModel(view)` from
  the metamodel package.
- Views for which `extractMetaModel` returns `null` are skipped.
- A meta-model is defined either:
  - **embedded** — a visual group with property `meta` on a view, or
  - **linked** — a view reference on a view pointing to a meta view (a view with
    property `meta`).
- **Identity** of a meta-model is the **UUID of its `metaView`**:
  - embedded → the view holding the meta group (the view itself),
  - linked → the referenced meta view.
- Views are grouped by this identity. `extractMetaModel` returns at most one
  meta-model per view (embedded group beats linked reference); this is acceptable.

## 3. Comatrix partitioning — `Report:Model:Comatrix`

Comatrices are not one-per-meta-model. Instead, each **allowed direct relationship**
is routed to a comatrix by the property `Report:Model:Comatrix` on the **matching
meta-model relationship**:

- **key absent** on the meta-model relationship → the relationship is **ignored**
  (does not appear in any comatrix).
- **key present but empty** → routed to the **default** comatrix of its meta-model
  (identity `mm:<metaView UUID>`, name = meta view name), i.e. today's behavior.
- **key present with a value** → routed to a **named** comatrix; the value string is
  **both the identity and the worksheet name**.
- Named comatrices with the **same value merge across meta-models** (union of the
  involved relationships/elements). Default comatrices stay per meta-model.
- A relationship whose matched meta-model relationship set carries several values is
  routed to **every** matching comatrix.
- Consequently one meta-model can yield several comatrices: one default (empty-value
  edges) plus one per distinct value.

## 4. Subject elements (rows/cols of a comatrix)

Candidate elements are the elements **shown on the meta-model's contributing views**:

- **embedded** meta-model → elements shown on that one view.
- **linked** meta-model → union of elements across **all** views that reference it.
- The meta view's own elements are **not** candidates.

Filtering / keying:

- Keep only elements that **correspond to a meta-model element type**
  (`findMetaModelMatches` non-empty).
- Exclude the **meta-model definition elements themselves** (`mm.elementIds`).
- **Dedupe by concept**; **key rows/cols by element name**.
- A comatrix's **rows are the source** elements and its **columns the target**
  elements of the relationships routed to it; an element not used on a given side
  does not appear on that axis.

## 5. Matrix shape

- **Rows = source elements, columns = target elements** (asymmetric): an element is
  a **row** only if it is the source of a routed relationship, and a **column** only
  if it is the target of one.
- A cell is marked with `x` when a **direct** model relationship goes from the row
  (source) element to the column (target) element, the meta-model allows it
  (`isRelationshipAllowed`), and the relationship is **routed to this comatrix** (§3).
  - "Direct" = an actual model relationship between the two concepts (no derived /
    transitive relationships).
- Association / non-directional relationships are marked in their model
  source→target direction.

## 6. Domäne grouping

- Rows and columns are sorted by **Domäne** (found by traversing
  aggregation/composition to a grouping with specialization `Domäne`), empty domains
  last, then by name.
- A separator row/column is inserted between Domäne groups.

## 7. Diff mode (baseline comparison)

- **Baseline selection** reuses the existing mechanism:
  1. `--baselineModel <path>` parameter, else
  2. model property `baseline` naming a currently loaded model.
- **Optional**: without a baseline, produce plain matrices (no diff coloring).
- **Comatrix matching across models**: by comatrix **identity** — the value string
  for named comatrices, `mm:<metaView UUID>` for the default. The diff is applied
  only when the same identity exists in **both** baseline and current models.
- **What is diffed**:
  - **Relationships (cells)**: added / removed / unchanged.
  - **Elements**: **rows (sources)** and **columns (targets)** are diffed on their
    own axes — added / removed / changed (union of both models per axis). An element
    present in **both** models on that axis but whose routed relationships differ is
    marked **changed**; `added`/`removed` take precedence over `changed`.
- **Colors** (reusing the existing comatrix palette):
  - **Green** = added (only in current),
  - **Orange** = changed (element in both, but its relationships changed),
  - **Red** = removed (only in baseline),
  - neutral = unchanged.
- A colored **legend row** (green / orange / red) is shown at the top of each diff
  worksheet.
- **Comatrix present in only one model**:
  - current-only comatrices are emitted (as plain matrices),
  - baseline-only comatrices are **skipped**.

## 8. Output

- A **single workbook** `comatrix-metamodel.xlsx`, written to the model's directory.
- **One worksheet per comatrix**, named after the comatrix (the property value, or
  the meta view name for the default), sanitized to ≤ 31 chars and made unique.
- Header row/column frozen; cell mark = `x`; diff legend included on diff sheets.

## 9. Non-goals (for now)

- No baseline is mandatory (single-model mode is supported).
- No unit tests (the code depends on the jArchi runtime globals `$`, `model`, `Java`).
