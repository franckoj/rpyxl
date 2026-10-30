# Compatibility and performance remediation plan

## Scope and evidence

This plan responds to a user-reported comparison of **rpyxl 0.1.1 against openpyxl on Windows with Python 3.14**. The reported probes have not been rerun for this document. Findings below describe that report, not verified behavior of the current source tree. The repository's `pyproject.toml` declares version 0.1.1 and lists Python classifiers through 3.13; Python 3.14 compatibility must be tested explicitly rather than inferred from packaging.

Implementation is being kept on `fix/text-fidelity`; no release is authorized by this plan. At the user's request, runtime testing is deferred until the implementation phase is complete. New regression tests are prepared but have not been run against a rebuilt native extension. No milestone below is considered validated or complete yet.

## Current implementation status — unverified

| Area                        | Implemented source changes                                                                                                               | Remaining boundaries                                                                                                                  |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| Text fidelity               | Whitespace preservation, XML newline normalization, CR escaping                                                                          | Empty/absent cells and broader OOXML escape sequences still need acceptance coverage                                                  |
| Temporal values and formats | Date/datetime/time ISO conversion, fractional seconds, millisecond serial rounding, 1900/1904 epoch reads, number-format persistence     | Elapsed durations remain numeric; locale-specific built-in format codes and write-only custom formats are limited                     |
| File-like I/O               | Binary stream loading/saving, caller ownership and input-position preservation, output replacement/truncation, read-only buffer lifetime | Output requires seek/write/truncate; inputs are buffered, not zero-copy streaming                                                     |
| Cell styles                 | Font/fill/border/alignment/protection models and style-table overlay                                                                     | Full openpyxl style API and all advanced theme/style references are not covered                                                       |
| Worksheet metadata          | Merges, dimensions, freeze panes, filters, visibility, tab color, protection, page setup                                                 | Full merged-border propagation is not implemented                                                                                     |
| Worksheet edits             | Sparse insert/delete rows/columns and move_range with metadata movement                                                                  | Formula translation and edits involving reference-bearing features are explicitly rejected                                            |
| Workbook features           | Properties, defined names, comments, hyperlinks, validations, conditional formatting, tables, basic bar/column charts                    | No formula calculation, threaded comments, new-image API, or general editable chart support; unknown XML is best-effort, not lossless |
| GIL handling                | Detached native parsing/saving, single-cell reads, and row-channel waits                                                                 | Thread speedup is not verified; Python metadata work still has costs                                                                  |
| Regression preparation      | Text matrix, rich-feature cross-library comparisons, stream lifetime/truncation and functional threading tests                           | Tests have not been executed in this implementation phase                                                                             |

### Architecture and performance caveats

- Cell values stay in the Rust engine. Python wrappers hold sparse rich-feature metadata; this is not an openpyxl backend and adds no runtime dependency on openpyxl.
- `workbook.py` and `worksheet.py` become packages to support their public feature submodule import paths. `Native*` exports are available for callers intentionally using unwrapped native types.
- New value-only and cell-style-only workbooks use one native Rust save pass. Style components are snapshotted and serialized once per unique value, transferred as interned records, and combined with number formats in the Rust writer. Loaded original metadata and other rich features still require a conservative overlay, but cell styles are not rebuilt there. The new native writer path needs a rebuilt wheel and full validation; no performance improvement is claimed from static checks.
- Metadata worksheet parsing discards value/formula payloads during parsing, but input bytes, retained metadata, and dependency parts still consume memory. Read-only value iteration avoids loading rich metadata; first metadata access loads it for all sheets with a warning.
- Unknown package parts are retained only on a best-effort basis. Reference-sensitive operations are rejected where correct rewriting is not implemented. Do not advertise full drop-in compatibility before acceptance testing.
- No testing or publication results are implied by successful syntax/format checks.

| Reported observation                                                             | Planning implication                                                                                                                                |
| -------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| 21 of 44 compatibility probes passed                                             | Preserve the original probe inventory and outcomes as the baseline; passing a subset is not full drop-in compatibility.                             |
| Leading whitespace stripped; LF becomes CRLF on read                             | Prioritize exact text fidelity before expanding APIs.                                                                                               |
| Date/time and number-format probes failed                                        | Treat Python values, Excel serials, workbook epoch, and formatting as a combined compatibility contract.                                            |
| `Cell.font`, `Cell.fill`, and related style APIs unavailable                     | Separate API support, serialization, and preservation of existing styles.                                                                           |
| Merge, dimensions, freeze panes, filter, and editing support missing             | Add worksheet metadata before structural mutations, then test their interactions.                                                                   |
| `BytesIO` read/write failed                                                      | Add file-like I/O after scalar fidelity, with explicit stream ownership and positioning behavior.                                                   |
| Charts, comments, tables, validation, names, protection, and metadata absent     | Deliver feature-specific preservation and editing increments, not a single broad compatibility claim.                                               |
| Interoperability results reported as `396/196` cell mismatches; features `1/10`  | Retain these figures verbatim until the two mismatch populations and all denominators are identified. Do not combine them or infer a mismatch rate. |
| Approximately 7.3× speedup on supported workloads; 15 versus 16 comparable cases | Retain the result as reported and workload-limited; audit the eligible case set and aggregation before republishing it.                             |
| Thread result approximately 1× at 100% CPU                                       | No demonstrated concurrency benefit; CPU percentage alone does not establish GIL release or parallel execution.                                     |
| Repeated-run RSS stable                                                          | Encouraging within the measured run, not proof of leak freedom, bounded streaming memory, or all-workload stability.                                |

## Shared reproduction and release gates

Every milestone below includes reproduction, permanent regressions, and acceptance. A failing probe becoming a passing local test is necessary but not sufficient: retain artifacts and verify cross-library behavior.

- [ ] Recover the original 44 probes, input workbooks, raw interoperability results, feature checks, and performance logs. If unavailable, label replacement probes as reconstructed rather than exact reproductions.
- [ ] Record Windows build, CPU/core count, RAM, Python version/build and architecture, rpyxl wheel or source revision/build profile, openpyxl version, dependency versions, commands, and fixture hashes. Reproduce on Windows/Python 3.14 first; track other supported platforms and Python versions separately.
- [ ] Assign stable case IDs and distinguish pass, fail, unsupported, skipped, and harness error. Preserve exception types/messages and expected versus actual values/types. Unsupported or skipped cases must not count as passes.
- [ ] Store minimal deterministic fixtures and machine-readable outcomes alongside regression tests. Use `tests/compatibility`, `tests/integration`, and `tests/unit` according to test scope; new scoped regressions are prepared, while the complete original harness still needs to be recovered.

### Required writer-reader matrix

For each applicable data or feature fixture, test all four directions:

| Writer   | Reader   | Purpose                              |
| -------- | -------- | ------------------------------------ |
| openpyxl | openpyxl | Reference/control and fixture sanity |
| openpyxl | rpyxl    | Import compatibility                 |
| rpyxl    | openpyxl | Export interoperability              |
| rpyxl    | rpyxl    | Internal round-trip consistency      |

Run each direction with normal writing and `write_only=True` writing where the API supports the feature, and normal reading plus `read_only=True` reading where inspection is supported. Include both file paths and `BytesIO` after M2. Streaming modes are not assumed to expose editing or every metadata API: record mode-specific limitations explicitly, and inspect their serialized output through a normal reader or ZIP/XML checks where necessary. Do not silently omit a failing mode or call an unsupported feature supported.

For preservation work, additionally test **openpyxl fixture → rpyxl load → unrelated edit/save → openpyxl inspection**, plus a no-edit save. Validate values, types, feature semantics, package parts, content types, and relationships as relevant; byte-for-byte XLSX equality is not required. Reference-library limitations or normalization must be documented rather than mistaken for an rpyxl pass or failure.

Each milestone closes only when its scoped regression matrix passes, existing passing probes remain green, and the original failed case IDs have explicit outcomes. Publish counts with denominators and reasons for exclusions. A milestone may establish partial support, but must not imply complete openpyxl parity.

## Ordered milestones

### M0 / P0 — Exact text fidelity

- [ ] **IN PROGRESS — first implementation and regression tests added; native build validation blocked.** Worksheet readers now disable XML text trimming; readers normalize literal XML CR/CRLF before unescaping; writers escape explicit carriage returns and declare whitespace preservation. Rebuild and run tests before marking these fixes verified. Remaining M0 cases include empty/absent cells and broader OOXML escaping.

**Reproduce:** Use separate cells containing leading/trailing spaces, whitespace-only strings, tabs, LF, CRLF, isolated CR, multiline text, empty strings versus absent cells, Unicode, XML-sensitive characters, and literal escape-like text. Include both shared-string and inline-string fixtures and inspect actual ZIP/XML payloads to locate whether corruption occurs during writing, parsing, or Python conversion.

**Implementation scope:** Preserve text through serialization and parsing, including `xml:space` and OOXML character escaping as required. Distinguish XML newline normalization from the public cell-value contract; avoid blanket trimming or newline replacement. Cover both normal and streaming paths.

**Regressions:** Add parameterized exact-string comparisons across the shared matrix, with first differing code point in failure output. Include OOXML-encoded CR/LF cases and reference controls so XML parser behavior is not confused with application-level conversion.

**Acceptance:** Leading whitespace remains intact, LF does not become CRLF, and all scoped strings retain their expected code points across applicable directions/modes. Empty-string and absent-cell behavior matches the documented reference contract. Close only after the native extension is rebuilt and all scoped regressions actually pass.

### M1 / P1 — Date/time values and number formats

- [ ] Implement and validate date/time conversion and number-format round trips after M0.

**Reproduce:** Isolate `date`, `datetime`, `time`, and `timedelta`; fractional days/seconds; date-formatted numeric cells; plain numeric cells; and built-in/custom formats. Include 1900 and 1904 epochs, Excel's fictitious 1900 leap day, pre-epoch/negative serial cases, blank cells, and precision boundaries. Record the reference behavior for each edge rather than assuming every value is supported. Include timezone-aware inputs and require explicit, reference-aligned handling or rejection.

**Implementation scope:** Carry epoch and style/format identifiers through loading, cell conversion, and writing. Preserve format codes separately from values. Define conversion precision and unsupported-input behavior; never silently stringify temporal values or classify all formatted numbers as dates.

**Regressions:** Apply the matrix to temporal values and numeric controls, asserting Python type, value, epoch, and format code after reload. Cover `data_only` separately where formula caches affect a case; do not require formula calculation from either library.

**Acceptance:** Scoped temporal types and formats match the reference contract in normal and streaming modes, using only explicit reference-derived precision tolerances. Ambiguous Excel edge cases have documented, tested behavior. Previously passing numeric/text cases remain unchanged.

### M2 / P1 — `BytesIO` and file-like I/O

- [ ] Add file-like reading and writing without regressing path-based I/O.

**Reproduce:** Load an openpyxl-created workbook from `BytesIO`, save an rpyxl workbook to a fresh `BytesIO`, explicitly rewind it, and reload using each reader. Repeat normal and streaming workflows. Record behavior for nonzero initial positions, closed streams, and corrupt or non-ZIP input against the reference.

**Implementation scope:** Define a binary file-like contract, including required read/write/seek capabilities, error behavior, and stream ownership. Do not close caller-owned streams unexpectedly. Explicitly reject unsupported non-seekable streams; do not imply support from `BytesIO` alone. Preserve deferred-read resources for the lifetime of a read-only workbook and release internal resources on close.

**Regressions:** Extend M0/M1 fixtures to paths and streams; check caller stream usability after save/close, repeated operations, malformed input, and read-only iteration lifetime. Test the declared stream-position contract rather than assuming automatic rewinding.

**Acceptance:** All applicable writer-reader and mode combinations load/save valid XLSX bytes, maintain scalar fidelity, and obey the documented position/ownership contract. Streaming memory behavior is measured separately; accepting `BytesIO` must not imply constant-memory I/O.

### M3 / P1 — Styles and basic metadata preservation

- [ ] Preserve existing styles/basic workbook and worksheet metadata, then expose a bounded style API.

**Reproduce:** Build fixtures with font, fill, border, alignment, number format, and cell protection settings; shared styles; theme/indexed colors; multiple sheets; titles/order/visibility; and core document properties. First demonstrate no-edit and unrelated-cell-edit losses separately from missing attributes such as `Cell.font` and `Cell.fill`.

**Implementation scope:** Retain style tables, theme references, metadata, and their relationships without corrupting identifiers. Implement the scoped public style attributes and copy/assignment semantics incrementally, keeping number formats consistent with M1. Define which unmodeled parts can be safely preserved; never claim opaque pass-through is safe when relationships or references may change.

**Regressions:** Assert semantic style fields rather than raw style IDs. Check styled blank cells, style reuse, unrelated edits, and serialization through the shared matrix. For streaming modes, test style access/creation only where supported and inspect output using normal readers.

**Acceptance:** Scoped existing styles/properties survive no-edit and unrelated-edit saves; supported style assignments are readable by openpyxl. Unsupported access or unsafe preservation has an explicit, tested outcome, not silent loss. Publish a per-field preservation/read/write/edit support table as future documentation work.

### M4 / P2 — Worksheet metadata and structural edits

- [ ] Add merge ranges, dimensions, freeze panes, and filters before implementing structural edits.
- [ ] Add row/column insertion/deletion and range moves in separate, reviewable increments.

**Reproduce:** Isolate merged ranges and anchor cells; row heights/column widths and hidden flags; worksheet used-range bounds; freeze panes; auto-filter ranges; and each mutation API. Add combinations of those features with styles, formulas, sparse cells, and multiple sheets.

**Implementation scope:** Implement metadata load/save and public APIs first, then mutations with an explicit reference-update policy. Check openpyxl behavior for formulas, merged ranges, names, and feature references; do not promise automatic translation or dependency repair beyond tested semantics. Reject unsupported mutations rather than partially modifying a workbook.

**Regressions:** Test API state before save and semantic state after both readers reload. Include boundary insertions/deletions, overlapping or invalid merges, moved cells, and unaffected-sheet checks. Apply the shared matrix to serialized metadata; editing is a normal-mode gate, while read-only inspection and write-only metadata timing restrictions are recorded separately.

**Acceptance:** Each metadata feature and edit operation has its own passing gate. Scoped reference behavior, including any non-updated references, is explicit. No unrelated text, values, styles, or relationships are lost, and invalid operations leave workbook state consistent.

### M5 / P2 — Advanced feature preservation, then targeted APIs

- [ ] Preserve and expose comments.
- [ ] Preserve and expose tables, including names and range references.
- [ ] Preserve and expose data validation.
- [ ] Preserve and expose workbook- and sheet-scoped defined names.
- [ ] Preserve and expose worksheet/workbook protection settings.
- [ ] Preserve and expose remaining document/workbook metadata in a named scope.
- [ ] Preserve charts and drawing relationships; add chart creation/editing only in separately scoped increments.

**Reproduce:** Recover the reported ten feature probes and identify the single reported pass. Create one minimal fixture per feature before combined workbooks. Include relationship-backed parts, comment authors, table IDs/names, validation formulas/ranges, scoped names, protection settings, and chart references to edited sheets/ranges. Enumerate metadata fields instead of treating “metadata” as a single undefined feature.

**Implementation scope:** Start each feature with load/no-edit-save and unrelated-edit preservation, then read APIs, then creation/editing. Track package parts, content types, relationship targets, identifiers, and reference dependencies. Protection settings are compatibility features, not a claim of encryption or strong security. If a feature cannot safely survive save, provide a visible diagnostic or reject the operation according to an explicit policy; do not silently discard it.

**Regressions:** Use the shared matrix and preservation pipeline for each feature, then combinations with M3/M4 edits. Inspect package integrity as well as reference-reader semantics. Record streaming limitations per feature. Do not equate XML part presence alone with usable chart/table/comment support.

**Acceptance:** Each checkbox closes independently only after its scoped matrix passes. The feature score is recomputed from the same identified inventory; new or unsupported features are listed separately. Claims distinguish preservation, reading, creation, and editing, with no blanket parity assertion.

### M6 / P3 — Concurrency evidence and safe execution

- [ ] Establish a concurrency baseline, then optimize only demonstrated bottlenecks.

**Reproduce:** Repeat the reported thread workload with serial execution and 1/2/4 workers, bounded by available cores, using independent workbooks and unique outputs. Separate read-heavy, write-heavy, CPU-heavy, and storage-bound tasks. Record wall time, process CPU time, per-core utilization, throughput, correctness, and storage/cache conditions; define how CPU percentage is normalized.

**Implementation scope:** Profile before changing synchronization or GIL handling. Release the GIL only around Rust work that does not access Python objects, with explicit ownership and synchronization. Do not infer shared-workbook thread safety from independent-workbook tests. Use process-based execution as a separately labeled comparison, not evidence of thread scalability.

**Regressions:** Compare concurrent outputs with serial results and the applicable cross-library/mode matrix. Test repeated independent operations, propagated exceptions, resource cleanup, and race-sensitive paths. Define a bounded stress workload; unsupported shared-instance access must remain explicitly unsupported.

**Acceptance:** Report speedup/efficiency with variability and workload scope. A flat approximately 1× result remains “no demonstrated benefit,” not a concurrency success. Publish a concurrency claim only after repeatable improvement and correctness are demonstrated; otherwise document the limitation without blocking compatibility fixes.

### M7 / P3 — Benchmark and memory methodology audit

- [ ] Reconcile the 15-versus-16 comparable-case discrepancy and the reported approximately 7.3× aggregate.
- [ ] Rebuild correctness-gated timing and memory reporting after compatibility fixes.

**Reproduce:** Recover the original case list, raw samples, eligibility decisions, timings, aggregation formula, and mismatch outputs. Identify both populations behind `396/196` and the denominator behind `1/10`. If raw evidence is unavailable, leave those figures unresolved and publish a new, separately labeled baseline.

**Implementation scope:** Review existing benchmark entry points in `benchmarks/` and `docs/benchmark_methodology.md` before future changes. Define comparable work by dataset, cells actually processed, operation boundaries, mode, values/types, features, output validity, and storage conditions. Separate unsupported, failed, and correctness-mismatched cases from eligible performance pairs. Never reward dropped features or skipped work with a speedup claim.

**Regressions and measurement:**

- Use fixed seeds, versioned fixtures, pinned environment/build details, warmups, repeated samples, and randomized library order. Report sample counts and dispersion alongside medians.
- Separate load, full iteration, time-to-first-row, save/finalization, and end-to-end workflows. Fully consume lazy readers and include writer finalization for comparable tasks.
- Report paired per-case ratios and an explicitly defined aggregate, such as the geometric mean over the same eligible paired cases. Show exclusions and denominators; do not mix mode/dataset populations without labeling them.
- Measure peak RSS in isolated processes, distinguish baseline from peak/delta, and report units and sampling method. Run bounded repeated-use tests with explicit close/cleanup and multiple dataset sizes; separate Python allocator measurements from total process RSS.
- Keep concurrency measurements separate from single-workload speedups. Re-run correctness checks across the shared matrix before accepting timed cases, including `BytesIO` only as a separately labeled I/O workload.

**Acceptance:** Every aggregate can be recalculated from published samples and an identical paired case list; the 15/16 discrepancy is resolved or explicitly remains unresolved. The reported approximately 7.3× result is not generalized beyond its validated workload. Repeated RSS stability is reported with run length and workload bounds, not as proof of no leaks or constant memory. Reconcile historical and current counts without rewriting the original evidence.

## Delivery sequence and completion policy

Proceed **M0 → M1 → M2 → M3 → M4 → M5 → M6 → M7** for implementation priority. Evidence capture and benchmark-case auditing can begin immediately, but performance optimization must not displace fidelity fixes. Split large milestones by feature rather than delaying all regressions until the end.

For each implementation increment, retain a minimal failing reproduction, add regressions, make the focused change, run targeted tests, then run the broader compatibility suite. Record actual commands, environment, outcomes, remaining failures, and support boundaries before checking a task off. Future documentation updates should correct broad drop-in, streaming-memory, and concurrency claims only to the extent supported by those results.

**Current status:** M0 is in progress in the parent work. No milestone is accepted by this document; all remaining work is unchecked. No implementation or validation results are asserted here.
