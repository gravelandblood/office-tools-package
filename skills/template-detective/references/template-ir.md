# Template IR

The Template IR is the contract between deterministic DOCX extraction, LLM rule induction, rendering, and validation.

## Top-Level Shape

```json
{
  "version": 1,
  "mode": "generalize",
  "sources": [],
  "formatAtoms": [],
  "structure": [],
  "dataSchema": {},
  "rules": [],
  "conflicts": [],
  "validation": []
}
```

## Sources

Each source records the sample document and generated profile:

- `id`: stable source ID.
- `docxPath`: input DOCX path.
- `profilePath`: normalized profile JSON path.
- `role`: `example`, `holdout`, or `baseline`.
- `fingerprint`: package/profile fingerprint.

## Format Atoms

Format atoms are deduplicated formatting objects:

- `id`: stable ID such as `fmt.paragraph.title`.
- `kind`: `paragraph`, `run`, `table`, `row`, `cell`, `section`, `numbering`, or `style`.
- `fingerprint`: deterministic hash.
- `properties`: normalized OOXML property subset.
- `evidence`: locations where the atom appears.
- `stability`: `stable`, `variant`, `conditional`, or `unknown`.

The atom must keep enough properties to render, compare, and explain a decision. Do not reduce it to a human label like "title style".

## Structure Nodes

Structure nodes represent document order:

- `paragraph`: text runs, paragraph atom, role candidates.
- `table`: table atom, rows, cells, possible loop candidates.
- `section`: page setup and header/footer references.
- `headerFooter`: header/footer content.
- `media`: image or drawing references.

Every node should carry:

- `id`
- `kind`
- `sourceRefs`
- `formatAtomRefs`
- `textPreview`
- `children`

## Data Schema

The data schema describes runtime inputs:

- `fields`: scalar slots such as `company.name`.
- `arrays`: loop inputs such as `shareholders[]`.
- `objects`: grouped fields.
- `required`: required field paths.
- `types`: `string`, `number`, `date`, `boolean`, `richText`, or `unknown`.
- `examples`: values observed in source documents.

Prefer semantic names only when evidence supports them. Otherwise use neutral names such as `field.001` and record a naming suggestion separately.

## Rules

Rules are the reusable template logic:

- `staticText`: literal text remains fixed.
- `slot`: dynamic scalar text.
- `loop`: repeated paragraphs/table rows/cells driven by array data.
- `conditional`: optional block driven by a condition.
- `format`: formatting atom applies to a role or condition.
- `normalization`: generated IDs or nonsemantic differences to ignore during compare.

Every rule must include:

- `id`
- `kind`
- `target`
- `expression` or `value`
- `confidence`: 0 to 1.
- `evidence`
- `alternatives`
- `validationStatus`: `unverified`, `passed`, `failed`, or `partial`.

## Conflicts

Conflicts explain why inference is not fully deterministic:

- `sameRoleDifferentFormat`: same inferred role has different format atoms.
- `sameLabelDifferentMeaning`: repeated label seems to map to different fields.
- `contentLengthRisk`: observed slot text length is too narrow for generalization.
- `optionalBlockAmbiguity`: block appears/disappears without enough input evidence.
- `tableShapeAmbiguity`: repeated rows/columns cannot be confidently mapped to arrays.
- `directFormattingDrift`: direct formatting differs from named style or repeated atom.

Each conflict must include:

- `severity`: `info`, `low`, `medium`, `high`.
- `classification`: `conditionalRule`, `likelyAnomaly`, or `unresolvedConflict`.
- `evidence`
- `recommendedAction`

## Validation

Validation records replay results:

- `sourceId`
- `renderedPath`
- `textEqual`
- `profileEqual`
- `acceptedDifferences`
- `failures`

Validation failures should feed back into rules or conflicts; they should not be ignored.

## Office-Native Compile Target

Treat the IR as a detection and planning layer, not the final template runtime. Prefer compiling accepted rules back into Office-native structures:

- `slot` -> Word content control (`w:sdt`) with `w:tag`, `w:alias`, and optional custom XML binding.
- `loop` -> repeating section content control when a stable table row or block range is known.
- `staticText` -> unchanged WordprocessingML.
- `conditional` -> native content-control range plus sidecar condition.
- `conflict` -> sidecar gap/report only; do not patch unresolved conflicts.

Static rules should be emitted as preserve items, not actual patches. Actual patches must only target structures that need wrapping or metadata.

Before writing a DOCX, every patch must have an exact Office range:

- label/value slots need the value run or text-node range, not the full label paragraph.
- paragraph slots need a whole-paragraph or exact run range.
- loops need validated table-row or block boundaries.
- conditionals need a native content-control range plus sidecar condition.

`template compile-office` may apply scalar slot patches when the dynamic value is isolated in a single safe run. It may apply table loop patches when the baseline table is uniquely located and body rows can be wrapped without changing table properties. It must skip cross-run values, ambiguous loops, conditionals, and conflicts until the IR carries a validated exact range.

Skipped compile targets should carry scored candidates with `kind`, `index`, `confidence`, evidence reasons, text preview, fingerprint, and range. Human or LLM review can feed selected candidates back through `--accept-candidates`; compiled results must mark those as overrides.

Custom behavior should default to an OfficeCLI/OOXML patcher:

- Use OOXML patching for content controls, custom XML parts, data binding, deletion, cloning, and deterministic rendering.
- Use COM/JSAPI only for operations that require a live Office/WPS application, such as field updates, active selection, pagination-sensitive verification, save-as/export, or user-facing add-in flows.
- Do not use UIA for template rendering; reserve UIA for product UI features such as PDF conversion.
