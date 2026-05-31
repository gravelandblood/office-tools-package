---
name: template-detective
description: Infer reusable Word/DOCX templates from one or more example documents, especially when Codex must separate stable formatting from dynamic content, detect inconsistent formatting, induce slots/loops/conditionals with LLM reasoning, render candidates, and validate that regenerated DOCX output matches the examples.
---

# Template Detective

## Overview

Use this skill as a heavy workflow, not as a one-shot converter. The goal is to infer a reusable DOCX template whose rendered output can round-trip against examples while exposing uncertain or contradictory patterns.

## Workflow

1. Extract deterministic evidence first.
   - Run `office-tools template profile <sample.docx> --out <profile.json>` for every sample.
   - Run `office-tools template analyze <sample-a.docx> <sample-b.docx> --out-ir <ir.json>` to build the first deterministic IR, rule candidates, and conflict report.
   - Run `office-tools template infer-format` only when a byte-preserving placeholder baseline is needed.
   - Do not ask an LLM to reason over raw OOXML unless the normalized profile is insufficient.

2. Normalize before reasoning.
   - Load `references/template-ir.md` and inspect the `template analyze` IR before doing any custom mapping.
   - Collapse repeated formatting atoms by fingerprint.
   - Keep evidence pointers: part name, paragraph/table/run index, source text preview, and fingerprint.

3. Induce rules with explicit confidence.
   - Static text: same text and same role across examples.
   - Slots: text that changes while surrounding structure and formatting remain stable.
   - Loops: repeated table rows or repeated paragraph groups with aligned structure.
   - Conditionals: blocks present in some examples and absent in others.
   - Conflicts: same inferred role with incompatible formatting, unstable labels, missing rows, or contradictory sample evidence.

4. Render and compare.
   - Generate candidate data from each example.
   - Render with the candidate template.
   - Compare original vs rendered using `office-tools template compare-format`.
   - A rule is not accepted as stable until it survives render/compare on at least one held-out or replayed example.

5. Report uncertainty instead of hiding it.
   - Do not silently pick one style when examples disagree.
   - Emit `stableRule`, `conditionalRule`, `likelyAnomaly`, or `unresolvedConflict`.
   - Include the smallest evidence set that explains the decision.

## Product Modes

- `preserve`: preserve a single source document with fine-grained text placeholders. This is the current low-level CLI baseline.
- `generalize`: infer semantic slots, loops, conditionals, and formatting rules from one or more examples.
- `audit`: report inconsistencies and generalization risks without generating a template.

## References

- `references/workflow.md`: detailed phase-by-phase process and acceptance gates.
- `references/template-ir.md`: template IR fields, rule types, conflict model, and validation expectations.
- `schemas/template-ir.schema.json`: machine-checkable IR schema for generated plans.
