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
