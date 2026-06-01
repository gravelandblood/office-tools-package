# Template Detective Workflow

## Inputs

- One or more `.docx` output examples.
- Optional structured inputs that produced those examples.
- Optional user notes about which fields should be dynamic.

## Phase 1: Evidence Extraction

For every DOCX sample:

```powershell
node packages/cli/bin/office-tools.js template profile sample.docx --out sample.profile.json
```

Keep the generated profile as the source of truth for formatting. It contains text parts, paragraphs, runs, tables, sections, styles, numbering, fingerprints, and conflict signals.

For one or more samples, create the first Template IR:

```powershell
node packages/cli/bin/office-tools.js template analyze sample-a.docx sample-b.docx --out-ir template-ir.json
```

This IR contains deterministic format atoms, structure nodes, slot/loop candidates, and conflicts. Treat it as a starting point for LLM reasoning, not as the final generalized template.

Use `template infer-format` only to build a preservation baseline:

```powershell
node packages/cli/bin/office-tools.js template infer-format sample.docx `
  --out-template baseline.docx `
  --out-data baseline.data.json `
  --out-profile baseline.profile.json
```

## Phase 2: Normalization

Normalize profiles into the Template IR, or inspect the IR produced by `template analyze`:

- Convert repeated `pPr`, `rPr`, table, row, and cell properties into reusable format atoms.
- Preserve original evidence locations for every atom and rule candidate.
- Keep paragraph/table order stable.
- Normalize obvious noise such as volatile generated IDs separately from semantic properties.

Do not erase inconsistencies. If two visually similar paragraphs use different direct formatting, record both variants and defer the decision to rule induction.

Current analyzer behavior: multi-sample alignment uses anchor similarity from labels, table headers, text, formatting, and nearby position. Inspect `ir.alignment` before trusting slot/loop rules. A high unmatched count means the samples may be different report types or need better section-level alignment.

## Phase 3: Rule Induction

Infer rules in this order:

1. Structural anchors: headings, labels, table headers, section boundaries, repeated blocks.
2. Static text: stable literals that should stay in the template.
3. Slots: changing text surrounded by stable anchors.
4. Loops: repeated rows or paragraph groups with stable internal shape.
5. Conditionals: optional sections or rows that correlate with input fields.
6. Formatting rules: style atoms for each role, with exceptions and confidence.

LLM reasoning should operate on compact IR summaries plus selected evidence, not raw XML. Ask the model to produce candidate rules with evidence IDs, confidence, and alternative explanations.

## Phase 4: Candidate Rendering

Build a candidate template and data mapping, then replay the examples. Rendering must preserve:

- Text content for replayed examples.
- Paragraph, run, table, and section counts unless the rule explicitly changes structure.
- Formatting fingerprints for unchanged regions.
- Accepted formatting-rule differences only where the inferred rule predicts them.

## Phase 5: Validation

Run:

```powershell
node packages/cli/bin/office-tools.js template compare-format original.docx rendered.docx
```

Treat failures as rule evidence:

- Text mismatch means slot/static text mapping is wrong.
- Profile mismatch in stable regions means formatting preservation is wrong.
- Mismatch only in predicted dynamic regions may be acceptable if explained by the rule.

## Conflict Handling

Classify every conflict:

- `stableRule`: evidence agrees; render/compare passes.
- `conditionalRule`: evidence differs but a data condition explains it.
- `likelyAnomaly`: one sample deviates from a stronger repeated pattern.
- `unresolvedConflict`: evidence is insufficient or contradictory.

Never collapse an `unresolvedConflict` into a silent default. Show the conflict with evidence and the practical risk.

## Acceptance Bar

A template inference result is useful only when it includes:

- A candidate DOCX template or a clear reason it cannot be generated yet.
- A structured data schema.
- Rule/conflict report with evidence.
- Replay validation results.
- Known generalization risks for content length, optional blocks, table row counts, and inconsistent formatting.
