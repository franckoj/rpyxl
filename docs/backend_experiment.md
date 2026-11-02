# Rust backend experiment

Active development branch: `perf/rust-backends`, based on released `v0.2.0` (`2e8de70`). The fully merged `fix/text-fidelity` branch has been retired locally and remotely. This experiment does not publish a release or change the default backend.

## Scope

- New editable workbooks can explicitly select the in-process Rust `rust_xlsxwriter` adapter.
- Existing-workbook reads, edits, preservation, structural safeguards and default writing remain on the released native engine.
- Calamine is a dev-only native read benchmark/probe, not a runtime reader replacement.
- Additional multicore work is deferred until backend correctness and single-process measurements are understood. Existing native parallelism is not expanded here.

```python
from rpyxl import Workbook
from rpyxl.styles import Font

book = Workbook(backend="rust_xlsxwriter")
sheet = book.active
sheet.append(["Amount", 12.5])
sheet["A1"].font = Font(name="Arial", size=12, bold=True, color="FF123456")
book.save("candidate.xlsx")
book.close()
```

The default remains `Workbook(backend="native")`. Loading any workbook retains the native backend. Selecting an unknown backend, supplying a loaded/native workbook to the candidate constructor, or requesting candidate write-only mode fails explicitly. There is no hidden native fallback during a forced candidate save.

## Current candidate constraints

This is a conservative prototype, not an interchangeable writer. The complete Rust-side subset and caveats are in [`RUST_XLSXWRITER.md`](../crates/rpyxl-ooxml/RUST_XLSXWRITER.md).

- Only new editable 1900-epoch workbooks and supported scalar values/styles/formulas are eligible.
- Python rich metadata, comments, hyperlinks, merges, dimensions and other overlay-dependent features are rejected rather than dropped. Opaque/custom style models are also rejected.
- Explicit fonts need representable names/sizes. Explicit RGB colors must be eight-digit `FF`-prefixed ARGB; ordinary six-digit rpyxl colors are padded with `00` and therefore rejected by this prototype. Theme/tint colors are not supported.
- Empty strings, literal OOXML escape tokens, unsafe formula rewrites, unrepresentable integers, non-finite numbers and attribute control whitespace have explicit guards.
- Uncomputed formula caches must remain absent. CR cell text and missing formula caches require a post-serialization fixup; its archive work is part of measured time and temporary-disk use.
- Runtime serialization uses the crate's standard in-memory worksheets, so it introduces a second cell store. No memory or speed benefit is assumed.
- Constant-memory runtime mode is deferred: the pinned crate uses panicking temporary-file writes in its row-spill path. Under this workspace's release panic-abort setting that is not an acceptable Python I/O error boundary. An upstream fallible-I/O fix or separately reviewed patch is required first.
- Validation and serialization are staged before modifying destinations. Failures during the final destination copy are not fully atomic. Caller-owned binary streams retain ownership and use the established replacement/truncation conventions.

## Correctness and wheel validation

The branch-wheel workflow now follows `perf/rust-backends`, builds the real native ABI, runs formatting/Clippy and Rust regressions, tests the installed wheel, and uploads the native reader comparison executable. Local source-only checks against the released wheel do not validate the new adapter.

Use the full existing suite plus `tests/unit/test_crate_backend.py` and the Rust adapter/probe tests. Candidate tests require the real method and may not silently skip or substitute the current writer. Repeated saves, source mutation, copy-on-write, unsupported input, streams and unchanged destinations must remain covered.

## Writer benchmark

```sh
python benchmarks/compare_backends.py --backends native rust_xlsxwriter openpyxl --repetitions 7
```

The driver uses fresh processes, rotated order, deterministic identical data (including identical eight-digit style colors), at least seven measured repetitions, raw results, package/native provenance, CPU seconds, absolute and incremental RSS, output size, stage timings and semantic output validation. It creates a separate session directory rather than overwriting old evidence. Save-time fixups and final close are included.

These are explicitly **library-default compression comparisons**, not strict compression-level-matched results: the native writer, candidate and openpyxl can use different deflaters/default settings. An observed speed difference cannot automatically be assigned to cell serialization. Compare file sizes and compression methods too. Normal and write-only modes remain separate; no candidate write-only result is fabricated.

No default promotion is planned from a single machine. Record fidelity and memory costs even if the candidate is faster; keep the native backend if it wins.

## Native reader benchmark

```sh
cargo run --release -p rpyxl-ooxml --example bench_backend_read -- native input.xlsx
cargo run --release -p rpyxl-ooxml --example bench_backend_read -- calamine input.xlsx
```

The Calamine arm uses its cell-reader stream without constructing another RPyXL store. The native arm uses the current editable materialization path. This architectural difference is intentional and must be labeled: it is **not** a same-API Python reader speedup. Compare cached-value scalar checksums/counts only for the declared common schema. Dates, errors, formula provenance, CR escape handling and other gaps must be reported rather than normalized away to create a false pass.

Both arms perform the existing ZIP safety precheck; timing reports distinguish that phase and include it in the overall total. This does not establish XML-security equivalence or qualify Calamine for production routing.

## Next decision

Collect correctness and performance evidence first. Adopt only a proven workload subset; preserve an explicit native override. Evaluate streaming after its I/O failure behavior is resolved, and consider multicore optimization only after these backend decisions.
