# Future backend evaluation plan

**Status: post-0.2.0 experiment on `perf/rust-backends`; released defaults remain unchanged.**

The first opt-in writer adapter and dev-only Calamine probe have initial Windows correctness/performance evidence. See [the experiment scope and commands](backend_experiment.md) and [the measured report](reports/rust_backend_experiment_2026-11-02.md). The standard-mode crate writer lost to native in all completed cases, so the default stays native. Calamine streaming is promising but not yet a same-contract read-only comparison or security-qualified production reader. Runtime constant-memory writing is deferred because the pinned writer's spill I/O can panic. Additional multicore changes remain explicitly deferred.

Release the current work separately from this evaluation. This document does not change default backend selection; preliminary measured results are linked above and do not establish the adoption gates below. The 0.2.0 baseline remains experimental and incompletely compatible with openpyxl. Earlier statements that publication was not authorized in the remediation history describe those earlier validation stages, not this future implementation plan.

## Baseline and priorities

- Cell values currently live in the Rust engine; Python wrappers add sparse rich-feature metadata and a style sidecar. Preserve this single authoritative value model rather than adding another complete cell store.
- The current loaded-workbook backend retains package parts and style provenance, with native saves for eligible cases and a conservative overlay for rich/unproven edits. Keep it for existing-workbook edits, including unrelated-edit and no-edit preservation workflows. It is not a lossless arbitrary-OOXML editor.
- The released 0.2.0 baseline uses `rust_xlsxwriter = "0.100"` only as a development dependency. The experiment pins `=0.100.0` for an **explicitly selected new-workbook adapter**, leaving the default writer unchanged. Standard mode is evaluated first; constant-memory/write-only adoption has an additional fallible-I/O gate. These APIs cannot edit existing packages.
- Calamine is **not a current runtime dependency**; the experiment pins `=0.38.0` only for dev-only reader tests and native benchmarks. Evaluate it separately for Rust reading rather than coupling reader replacement to the writer experiment. Historical documentation mentioning a Calamine runtime reader is not evidence of the present dependency graph.
- All adapters must run in-process in Rust through the existing PyO3 boundary. No Python subprocess backend, hidden openpyxl writer, or Python-to-Rust-to-Python serialization loop.

The initial decision is selective adoption, not a wholesale backend rewrite. A reader success does not require adopting a writer, and a write-only success does not justify replacing normal or loaded-workbook saves. There are no 10–100× speedup promises.

## Proposed routing and ownership

| Workload                             | Candidate                                 | Default during evaluation                          |
| ------------------------------------ | ----------------------------------------- | -------------------------------------------------- |
| New workbook, normal mode            | Opt-in `rust_xlsxwriter` adapter          | Current native writer                              |
| New workbook, append-only/write-only | Separately opt-in constant-memory adapter | Current write-only writer                          |
| Existing workbook load/edit/save     | Current preservation backend              | Unchanged; never route to a new-workbook generator |
| Value/style/formula reading          | Independent Calamine prototype            | Current native reader                              |

Before opening output, build a conservative capability decision covering workbook origin, mode, styles and metadata. A candidate must support the requested semantics or decline explicitly. Unknown styles, subclasses, opaque XML and unsupported features must not silently disappear. Test routing after mutation and metadata promotion, not just at workbook creation.

For normal writing, borrow/iterate the current Rust cells directly where possible. Do not clone the full workbook into a second Rust worksheet model; verify whether the candidate's internal storage makes this unavoidable. If it does, account for both stores in peak RSS and reject normal-mode adoption when the memory gate fails. Cache and compile each unique format once per workbook using the existing interned style identity, with explicit invalidation for mutable styles. Do not construct formats per cell or traverse Python styles inside the Rust cell loop.

For constant-memory writing, stream rows directly through a bounded buffer rather than first collecting them in the existing full cell store. Characterize the candidate's ordering requirements, temporary-file use, supported features and save lifecycle. Memory may still depend on unique strings/formats, ZIP state and metadata: do not describe arbitrary inputs as strictly O(1). Existing random-access semantics must not silently become append-only.

Stage output before destination replacement, including file-like output where feasible. Decide fallback **before irreversible streaming writes**. If a candidate fails after consuming a non-replayable row source, return an actionable error and leave the destination intact; do not claim transparent fallback without a tested replay strategy.

## Independent Calamine read experiment

Evaluate a pinned, compatible Calamine version in an isolated proof of concept. [Calamine 0.38.0's `XlsxCellReader` documentation](https://docs.rs/calamine/0.38.0/calamine/struct.XlsxCellReader.html) exposes streaming values with style IDs, a style palette, and formula/formula-metadata methods. This is more capable than a value-only reader, but is not a package-preservation contract:

- `next_cell_with_style_id()` returns style ID 0 both for implicit default styles and for explicit default styles. `next_style_id()` skips style ID 0. Neither result alone proves retention of the original distinction between an absent style attribute and explicit `s="0"`.
- Test cached values versus formulas, shared-formula metadata, array formulas, date epochs, exact text, blank styled cells, style inheritance and unknown XML. Semantic formula expansion is not original XML provenance or formula calculation.
- The `next_*` methods advance the same XML stream. Choose a suitable single-pass record API or measure any required extra passes; do not mix methods and assume independent iteration.
- Do not use materialized ranges plus a second complete RPyXL cell store for a purported streaming path. Populate the authoritative store once for editable loads, or expose a bounded read-only iterator with correct borrowed-data lifetimes.
- Retain current ZIP/XML resource limits and error behavior, or document and test a stricter compatible policy. Parser speed cannot compensate for weakened DTD/entity, decompression or allocation defenses.

If style or formula provenance is insufficient, restrict the candidate to explicitly supported read-only/value workflows or reject it. Keep the existing parser/preservation path for editable loaded workbooks unless an independent preservation gate proves equivalence. Adding raw XML side parsing to recover missing provenance must be included in end-to-end time and memory.

## Staged proof of concept

1. **Freeze the baseline and contract.** Record the merged release revision, wheel hashes, dependencies and actual combined test counts. Keep `5883f7f` results as historical pre-merge evidence, not merged validation. Inventory supported modes/features, errors and unsupported cases with stable fixture IDs. Predeclare benchmark settings and gates below.
2. **Build an isolated new-workbook writer prototype.** Use the current development dependency first, without changing default runtime routing. Translate scalar values and unique styles, save paths and `BytesIO`, then validate output through independent readers and XML inspection. Measure pure Rust and Python-wrapper costs separately.
3. **Prototype constant-memory writing separately.** Verify row ordering, row-source consumption, temporary-file cleanup, cancellation/failure, repeated saves and unsupported operations. Instrument memory growth and avoid duplicate cell stores. Do not infer streaming behavior from the crate feature name alone.
4. **Run the independent reader prototype.** Test Calamine against the same read fixtures and safety corpus, including provenance probes. It may ship later, in a narrower mode, or not at all.
5. **Introduce opt-in runtime routing only after the scoped gates pass.** Select and pin compatible versions/features in a separate implementation change. Keep the current backend as the default and explicitly test decline/fallback and forced-current selection. Retain reproducible artifacts and publish supported/unsupported scope.
6. **Adopt only winning paths.** Enable a candidate by default only for its proven workload subset after portable-wheel validation. Leave other paths on the current backend. Remove an experimental adapter if it adds complexity without a measured end-to-end benefit.

## Matched benchmark protocol

Run both **pure Rust native** benchmarks and **RPyXL end-to-end** benchmarks through the installed release wheel. A fast crate microbenchmark is not an RPyXL speedup. Include Python wrapper overhead, style conversion, capability scans, buffering, temporary files and final ZIP close in end-to-end totals. Split build/append, format compilation, save/load and first-row latency as diagnostics, not substitutes for total time.

Compare the current backend, candidate adapter, and appropriate external controls (openpyxl and direct Rust crate usage). Match normal with normal, write-only with write-only and read-only with read-only. Use identical deterministic data, shape, style cardinality, formulas, string duplication, output destination and compression algorithm/level. If a backend cannot match compression or feature semantics, label it a separate configuration and do not aggregate it as a like-for-like speedup. Record output sizes and ZIP methods as checks.

At minimum include:

- Dense numeric, mixed text/boolean/date/formula, sparse and repeated/high-cardinality string inputs.
- Unstyled, reused-style and high-cardinality-style cases; blank styled cells and explicit default styles.
- 20,000 × 20 styled writing and load/insert/delete/save to retain the targeted baseline, plus 10,000 × 20 and 50,000 × 20 append controls. Existing-edit cases are routing/preservation controls, not candidates for a new-workbook writer.
- Normal and streaming reads with time to first row and full consumption, paths and `BytesIO`, small fixtures and a row-count scale sweep large enough to expose memory growth.

Use fresh processes and release builds on the same machine/session; rotate paired execution order. After a documented warmup, collect at least seven measured pairs per case, retaining raw per-run data, medians, spread and a paired 95% confidence interval. Record OS, CPU/RAM, Python/Rust/toolchain versions, source/wheel hashes and all dependency versions. Repeat qualifying results on a second supported platform before default adoption.

Measure wall time, total process CPU seconds, absolute peak RSS, baseline RSS, sampled incremental RSS, output bytes and temporary-disk peak. State sampling frequency and missed-peak limitations. For constant-memory qualification, measure equal-style-cardinality workloads at N, 2N and 4N rows, separately reporting unique-format/string growth. Historical incremental RSS cannot be relabeled as absolute peak memory.

## Proposed measurable exit gates

These are acceptance criteria for future work, not results already achieved. Freeze any workload-specific exception before measuring; do not move thresholds afterward to make a candidate pass.

| Gate                        | Required evidence before default adoption of a scoped path                                                                                                                                                                                                                                                                                                             |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Correctness                 | Zero value/type mismatches on all applicable deterministic fixtures; all applicable feature assertions pass in the four writer/reader directions from the remediation plan. Unsupported/skipped cases are listed, never counted as passes. Existing passing regressions remain green.                                                                                  |
| Style/provenance            | Zero mismatches in the fixture inventory for font, fill, border, alignment, protection, number formats, theme/indexed colors, blank cells, explicit/default styles and nested/alias mutations. Inspect relevant XML/relationships as well as rendered API values. Unsupported provenance excludes that route.                                                          |
| Safety and routing          | All existing malformed-file/resource-limit tests and new candidate failure tests pass; no silent payload loss, no changed loaded-workbook route, and no destination corruption. Test capability decline, metadata promotion, I/O failure and forced-current rollback.                                                                                                  |
| Performance benefit         | For a declared target subset, at least 10% lower median end-to-end wall time or 20% lower absolute peak RSS versus current, with the paired confidence interval supporting an improvement in that metric. Other required control cases must not regress by more than 5% in wall time, CPU or peak RSS without an explicitly accepted, documented opt-in-only tradeoff. |
| Constant-memory claim       | At fixed unique styles/strings and metadata, absolute peak RSS at 4N rows increases by at most 20% over N, and instrumentation confirms no row-count-sized duplicate store. Report temporary-disk growth and any limits; vary style/string cardinality separately.                                                                                                     |
| Compatibility and packaging | Build and install portable wheels across the project's supported Linux, macOS and Windows targets/architectures and Python matrix; smoke-test paths, streams and failures from installed wheels. Build with the agreed minimum supported Rust version (MSRV) and stable Rust, with dependency/features/license review completed.                                       |
| Reproducibility             | Retain fixture hashes, raw results, commands, build/dependency versions and validation artifacts. A second supported platform confirms qualifying improvements; a maintainer reviews the scope and rollback path before a default change.                                                                                                                              |

The known append regression (approximately 11–13% in the recorded paired investigation) stays open. It must not be hidden in aggregate speedups or attributed to a backend without evidence. Styled memory around 300 MiB versus openpyxl's 195 MiB in the recorded incremental-RSS run is another explicit control, not proof that any candidate will fix it. Historical 43/44 API probes include deliberate rejection of table name `T1`; that difference is not full drop-in compatibility.

## Dependency and wheel constraints

The existing `"0.100"` development requirement does not establish suitability as a runtime dependency or a compatible MSRV. Before promotion, verify the exact selected `rust_xlsxwriter` version, constant-memory feature names/restrictions, license, transitive dependencies, ZIP/compression settings and minimum Rust requirements. Do the same independently for Calamine; version 0.38 documentation is an API reference, not an instruction to select it regardless of compatibility.

Agree and document the project's MSRV rather than inferring it from Rust 2021 edition. Test the resolved lockfile and selected features with that compiler and the current stable toolchain. Check PyO3/ABI3 compatibility, wheel size/build time, Linux portability, macOS architectures and Windows linking/SDK requirements. No runtime compiler, external executable or Python subprocess may be required by an installed wheel. A dependency requiring an unsupported compiler or wheel target blocks promotion until explicitly resolved.

## Rollback and subsequent iterations

Keep a tested explicit current-backend override and conservative capability routing. A failure in fidelity, safety, wheel portability or a performance gate means **no default promotion**. If a defaulted subset later regresses, disable only that route, retain the current backend and add a permanent reproducer. A partially successful prototype may remain opt-in with clear limits; otherwise remove it rather than maintain two general-purpose implementations.

If the existing backend wins, retain it. Only after the proof of concept and profiles identify the dominant costs, consider native coordinate parsing, format-cache improvements, borrowed reads, fewer PyO3 crossings, row batching, compact native per-cell style IDs or reduced XML buffering. Evaluate each against the same fidelity/safety controls and matched end-to-end benchmarks; do not bundle speculative optimizations into this release or assume a backend swap is necessary.

## References

- [0.2.0 release notes](../CHANGELOG.md)
- [Current remediation evidence, preservation limits and writer-reader matrix](remediation_plan.md)
- [Benchmark methodology](benchmark_methodology.md)
- [Current writer crate manifest](../crates/rpyxl-ooxml/Cargo.toml)
- [Calamine 0.38.0 streaming cell API](https://docs.rs/calamine/0.38.0/calamine/struct.XlsxCellReader.html)
