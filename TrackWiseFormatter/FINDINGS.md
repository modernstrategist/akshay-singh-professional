# Pilot Feasibility Findings — TrackWise Formatting

**Question:** Can a legacy IVC SOP be automatically reformatted into the new TrackWise
TWD format with high fidelity? **Answer: Yes — demonstrated end-to-end on GEN-027.**

## Verified result (GEN-027 → SOP-DDR-GEN-000003)

| Item | Outcome |
|---|---|
| Legacy → Full Document Number | `GEN-027` → `SOP-DDR-GEN-000003` (lookup) ✅ |
| Title | `HANDLING OF MANUFACTURING REWORKS` → `Handling Of Manufacturing Reworks` ✅ |
| Revision number | `5` → `006` (incremented) ✅ |
| Department | `General` ✅ |
| Confidentiality statement | Present in page header on every page ✅ |
| Eight sections | All present, numbered `1.`–`8.`, content mapped ✅ |
| Responsibilities | Mapped from legacy singular `RESPONSIBILITY` ✅ |
| Procedure | Two `PROCEDURE` / `PROCEDURE: (continued)` blocks merged (50 paragraphs) ✅ |
| Forms | Absent in source → `N/A` + exception ✅ |
| Revision History | 4-line rolling table (rev 3, 4, 5 + new `006` `(see header)`) ✅ |
| Source document | Untouched ✅ |

## Exception Report produced

```
[Warning] MissingSection (Forms): Source has no Forms/Attachments section; set to N/A.
[Info]    ImageCarryover (Procedure): Embedded exhibits/flowchart images are not carried into the draft in Phase 1; flagged for manual placement (Phase 3).
[Info]    Numbering (Procedure): Procedure sub-step numbering normalization (e.g. I.A.1 -> 1.1.1) is deferred to Phase 3.
```

## Feasibility conclusion

- The hardest OpenXML transformations — header table rebuild, per-page confidentiality,
  8-section mapping with N/A, and the 4-line rolling revision history — are **all proven**.
- Building against a **copy of the approved template** gives authentic formatting and
  guarantees the regulated source text is never rewritten.
- Remaining work (NTR, numbering normalization, images, AI-assisted mapping, UI,
  Azure orchestration) is **additive and scoped to Phases 2–3**, not a feasibility risk.

**Recommendation:** proceed to Phase 2. The use case is feasible on the chosen
.NET + OpenXML + Azure stack.
