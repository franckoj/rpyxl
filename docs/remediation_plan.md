# Compatibility and performance remediation plan

## Scope and evidence

This plan responds to a user-reported comparison of **rpyxl 0.1.1 against openpyxl on Windows with Python 3.14**. The original reported findings are retained below as the baseline; subsequent scoped validation is recorded separately. The repository still declares version 0.1.1, so the artifact's source commit distinguishes branch builds from the published release. Python 3.14 compatibility is assessed through explicit test runs, not inferred from packaging classifiers.

Implementation is being kept on `fix/text-fidelity`; no release is authorized by this plan. Runtime validation has started: the native-style wheel at `43271c2` passed 203 local Python tests and the original harness's four interoperability directions (zero cell mismatches, 10/10 feature checks). Its API result was 43/44, with the cell-reference-like table name `T1` rejected. The subsequent native-read/passthrough implementation at `6899d7e` passes 99 Rust core/OOXML tests and 346 installed-wheel Python tests on each of Windows Python 3.11 and 3.14. Its SHA-256-verified wheel also passes all 346 tests locally on Python 3.14. Release gates remain open rather than treating scoped test success as complete API parity.

## Current implementation status — scoped testing in progress

| Area                        | Implemented source changes                                                                                                               | Remaining boundaries                                                                                                                  |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| Text fidelity               | Whitespace preservation, XML newline normalization, CR escaping                                                                          | Empty/absent cells and broader OOXML escape sequences still need acceptance coverage                                                  |
| Temporal values and formats | Date/datetime/time ISO conversion, fractional seconds, millisecond serial rounding, 1900/1904 epoch reads, number-format persistence     | Elapsed durations remain numeric; locale-specific built-in format codes and write-only custom formats are limited                     |
| File-like I/O               | Binary stream loading/saving, caller ownership and input-position preservation, output replacement/truncation, read-only buffer lifetime | Output requires seek/write/truncate; inputs are buffered, not zero-copy streaming                                                     |
| Cell styles                 | Interned copy-on-write style models and native style-table persistence                                                                   | Full openpyxl style API and all advanced theme/style references are not covered                                                       |
| Worksheet metadata          | Merges, dimensions, freeze panes, filters, visibility, tab color, protection, page setup                                                 | Full merged-border propagation is not implemented                                                                                     |
| Worksheet edits             | Sparse insert/delete rows/columns and move_range with metadata movement                                                                  | Formula translation and edits involving reference-bearing features are explicitly rejected                                            |
| Workbook features           | Properties, defined names, comments, hyperlinks, validations, conditional formatting, tables, basic bar/column charts                    | No formula calculation, threaded comments, new-image API, or general editable chart support; unknown XML is best-effort, not lossless |
| GIL handling                | Detached native parsing/saving, single-cell reads, and row-channel waits                                                                 | Thread speedup is not verified; Python metadata work still has costs                                                                  |
| Regression preparation      | Text matrix, rich-feature cross-library comparisons, stream lifetime/truncation and functional threading tests                           | 112 Rust tests and 450 Python tests per CI version passed at `5883f7f`; one older-native-only test skipped; broader acceptance open   |

### Architecture and performance caveats

- Cell values stay in the Rust engine. Python wrappers hold sparse rich-feature metadata; this is not an openpyxl backend and adds no runtime dependency on openpyxl.
- `workbook.py` and `worksheet.py` become packages to support their public feature submodule import paths. `Native*` exports are available for callers intentionally using unwrapped native types.
- New value-only and cell-style-only workbooks use one native Rust save pass. Style components are snapshotted and serialized once per unique value, transferred as interned records, and combined with number formats in the Rust writer. The native writer is verified in the `43271c2` wheel. The latest branch adds native style reads and an eligible loaded-workbook preservation path, validated in the `6899d7e` wheel. No controlled before/after speedup claim is established for this phase.
- Normal cell styles share immutable workbook-local records. Cell-bound views detach on nested or alias edits; standalone source models remain mutable and use lazy revision tracking only after pool assignment. XML-backed and unsupported models keep the conservative copy path. Cached defaults avoid rebuilding a style for each unstyled cell. Latest source loads original cell XFs natively, decodes Python models only on access, and tracks sparse component overrides, including explicit clears. Unknown style XML retains the conservative mutable-copy path.
- Metadata worksheet parsing discards value/formula payloads during parsing, but input bytes, retained metadata, and dependency parts still consume memory. Read-only value iteration avoids loading rich metadata; first metadata access loads it for all sheets with a warning.
- Unknown package parts are retained only on a best-effort basis. Reference-sensitive operations are rejected where correct rewriting is not implemented. Do not advertise full drop-in compatibility before acceptance testing.
- No testing or publication results are implied by successful syntax/format checks.

### Native loaded-workbook optimization — scoped wheel validation complete

- Editable loads retain a bounded compressed source snapshot in Rust. Normal style access no longer triggers a Python metadata scan; each accessed source style decodes once per workbook. Read-only style/metadata access still uses the existing deferred Python path.
- Eligible value/style-only saves write a single ZIP output: unchanged members are raw-copied, worksheet framing and row attributes are retained, and native cells/styles are written once. Original style IDs, named styles, conditional-formatting DXFs, themes, relationships, and XF-level provenance remain anchored to the source stylesheet. Explicit `s="0"` is distinguished from an absent style attribute.
- Rich metadata or unproven structural edits use explicit generated-native output followed by the conservative Python overlay. Strictly proven plain workbooks can retain native package preservation during row/column edits and range moves. The original stylesheet base remains available even when package passthrough is ineligible. Both paths stage file output before replacing an existing destination.
- This is not a lossless arbitrary-OOXML editor. Saves explicitly reject detected unsupported cell payloads (for example shared/array-formula attributes, inline rich-text runs, and cell metadata) instead of silently regenerating them incorrectly. Loading/reading remains available. Unsupported rich shared-string details and broader formula semantics still need acceptance coverage.
- XML template parsing has depth/node limits and linear-size structural keys. Resolved style components have a 64 MiB aggregate allocation budget; highly repetitive pathological style files may be rejected rather than exhaust memory.
- Passthrough still buffers worksheet XML and generated output strings, and retaining compressed input has a memory cost. Do not describe this as constant-memory or zero-copy processing, or infer lower CPU usage without measuring total CPU seconds.
- Validation at `6899d7e`: **99 Rust core/OOXML tests passed** in GitHub's Windows build; **346 installed-wheel Python tests passed on each of Python 3.11 and 3.14**, with no skips. The verified wheel also passed **346 tests locally on Python 3.14**. The local MSVC linker/SDK remains unavailable, so native compilation uses CI. Workflow: https://github.com/franckoj/rpyxl/actions/runs/41416206292. No release is authorized.
- CI caught an implicit-General storage expectation and an explicit-default style bug. Native style assignments now retain `s="0"` for populated and blank cells without making untouched implicit styles explicit. Regressions also cover lazy passthrough followed by rich-metadata fallback and sheet-local name scope updates.
- The original harness's quick all-section run against the preceding `a1e52dc` wheel recorded **43/44 API probes**, **zero cell mismatches and 10/10 features in all four interoperability directions**, and measurements for all **20 performance scenarios**. All sections and reports completed, but the outer command later timed out, so a clean process exit was not confirmed. The remaining API difference is rejection of the cell-reference-like table name `T1`.
- Quick-mode timing uses one repetition and incremental sampled RSS, not absolute peak memory. It is not comparable to the original large benchmark. Styled writes and row insertion/deletion remain performance targets: profiling attributes substantial costs to Python style-capability scans and to XML tree copying/double serialization in the structural-edit overlay. Native row moves themselves were small in those profiles. Do not claim universal wall-time, CPU, or memory superiority. Ignored evidence is under `target/local-test-results/native-read-a1e52dc-quick/` and `target/local-test-results/profile-a1e52dc-targeted/`.

### Targeted style and structural performance — `5883f7f`

- Reuse dataclass layouts and a conservative exact-cell-shape check for native-only eligibility; reuse the decision within one save, but recompute after metadata promotion and on later saves. Unchanged tracked leaf styles reuse tokens and snapshots without recursive freezing. Nested edits, aliases, opaque XML, and subclasses retain their established conservative handling.
- Metadata `dump()` no longer deep-copies the full XML tree or serializes twice. It restores temporarily replaced attribute dictionaries in `finally`, then repairs namespace-compatible `mc:Ignorable` on the serialized output without changing global namespace registrations.
- A native, atomic, all-sheet proof enables structural package preservation only for a narrow plain-workbook subset. It pins source/framing identity, rejects unsupported references and markup, checks live content again, and discards obsolete original row frames after moves. Unproven sources and rich edits retain the Python fallback. Normal native cell attribute loading now decodes entity-escaped coordinates, types, and style IDs consistently with the proof.
- Validation: **112 Rust tests passed**, and **450 Python tests passed with one skip** on both Windows Python 3.11 and 3.14 in CI. The skip intentionally targets native wheels lacking the new private API. The SHA-256-verified wheel also passed the same local Python suite and is installed in `.venv-test`. Workflow: https://github.com/franckoj/rpyxl/actions/runs/41437474050.
- Targeted benchmark: Windows/Python 3.14.3; compare old wheel `6899d7e`, new wheel `5883f7f`, and openpyxl 3.1.5 in fresh processes, rotating order across three repetitions. Each affected case uses **20,000 rows × 20 columns**. The following are same-session medians, not comparisons against historical absolute timings.

| Workload                                | Old rpyxl wall | New rpyxl wall | openpyxl wall | Old → new sampled incremental RSS |
| --------------------------------------- | -------------: | -------------: | ------------: | --------------------------------: |
| Styled write                            |        18.49 s |        13.63 s |       17.45 s |                 304.2 → 300.5 MiB |
| Insert/delete rows, including load/save |        16.36 s |         5.55 s |       11.75 s |                 722.7 → 152.4 MiB |

- Styled write used about **26% less wall time and 27% less CPU time** than the old wheel, but its memory use remains above openpyxl (194.5 MiB in this run). Row editing used about **66% less wall time, 65% less CPU time, and 79% less sampled memory growth** than the old wheel. Streaming openpyxl checks compared **400,000 cells per case per rpyxl wheel** with zero mismatches; nine sampled style checks per styled output also matched.
- The append control warrants follow-up: a separate nine-pair alternating-order investigation measured median total wall time about **11% higher at 10,000×20 and 13% higher at 50,000×20** for the new wheel, with a roughly 19% save-stage increase. Individual timings varied and direct-native profiles did not establish a specific source-level cause. Do not dismiss the signal or claim all workloads improved. Read iteration was roughly unchanged in the original three-repetition control.
- A compact native per-cell style-ID redesign remains future work; this increment does not eliminate the Python style sidecar or establish styled-memory parity. The complete 20-scenario large matrix has not been repeated after this increment. Timing/profile/output-validation evidence is retained in ignored `target/local-test-results/optimization-5883f7f/`.

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
