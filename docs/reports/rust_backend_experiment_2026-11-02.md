# Rust backend experiment — 2026-11-02

## Decision

**Keep the current native writer as the default.** The opt-in `rust_xlsxwriter` integration is slower and uses more memory on every completed writer workload in this experiment. Reusing a mature crate is useful for correctness and maintenance, but does not automatically make this integration faster.

**Continue evaluating Calamine for a narrowly scoped values-only streaming reader.** Its probe is promising, but compares a cell stream with native full-workbook materialization. It is not a compatible replacement for the current editable reader, and it has not passed the required security, provenance or cross-platform gates.

No default routing, published release, `main` commit, or additional multicore implementation is part of this experiment. Work stays on `perf/rust-backends`; the fully merged `fix/text-fidelity` branch was retired locally and remotely.

## Build and correctness

- Release baseline: `v0.2.0`, `2e8de70bd705aa9f12e4ffba048a00f7e246dc93`.
- Candidate wheel/executable source: `88ffb13d37b8bf67273b2a4480bb7bad79f2b31c`.
- [Candidate CI](https://github.com/franckoj/rpyxl/actions/runs/41479509946): Rust formatting, Clippy with warnings denied, Rust core/OOXML regressions, reader executable build, and installed-wheel Python 3.11/3.14 suites passed. Each Python job ran **603 passed, 1 skipped**.
- Local exact-wheel regression suite: **603 passed, 1 skipped**. After adding benchmark-gate unit tests: **660 passed, 1 skipped**. These 57 additional cases test the benchmark harness, not new library APIs.
- Fixed the first CI compile failure by removing ambiguous `AsRef` calls on quick-xml byte-text values.
- Fixed getter-only implicit default styles being incorrectly submitted as explicit unsupported theme fonts. Explicit unsupported style assignments still fail before changing destinations. The default/native style compilation path is unchanged.
- Corrected a rejection test to use the actually unsupported centered-vertical/indented alignment, and added accepted top/bottom/default-vertical round-trip tests.
- Verified GitHub artifact digests and the wheel hash in `build-info.json` before using the binaries.

Wheel SHA-256:

```text
28b8fa526c64ef928133b83692861d012f26a173e04adb5a5374705364118b6b
```

Native extension SHA-256:

```text
7c1b576d713053b74669baef994a3a110f119443600ede6ff3ebd259814c9b7c
```

The candidate retains distribution version `0.2.0`; the hashes and ABI distinguish it from the published release. Do not use the distribution version alone as experimental provenance.

## Method

Windows AMD64, Python 3.14.3, 14 logical CPUs, approximately 31.45 GiB RAM; openpyxl 3.1.5. Runtime writer crate `rust_xlsxwriter =0.100.0`; dev-only reader crate `calamine =0.38.0`.

Writer measurements use identical deterministic inputs, fresh processes, seven rotated repetitions per backend, and median create + generate/append/style + save + final-close time. Import/provenance work and independent output validation are outside that measured interval. Save includes format conversion, adapter staging, formula-cache/CR fixups when applicable, and final output copying. Every output is fully validated by openpyxl, authored-style assertions, formula-cache checks and ZIP integrity checks before a case is declared comparable.

These are **library-default compression comparisons**, not compression-level-matched serializer benchmarks. RSS is sampled: absolute process RSS includes interpreter/import costs, while incremental RSS subtracts the measured baseline. No discarded warmups, confidence intervals, temporary-disk high-water measurement, second-platform confirmation, write-only candidate, or 10–100× performance claim is supplied by these results. This is initial experiment evidence, not satisfaction of all adoption gates in the future plan.

### Completed exact-wheel run: 2,000 × 20

All **105/105 runs** succeeded and all five cases passed semantic comparison. Main cases contain 40,000 cells; the separate plain append control contains 1,000 × 20 cells. All modes here are editable, not write-only.

| Scenario | Native wall s | Crate wall s | openpyxl wall s | Crate/native time | Native CPU s | Crate CPU s |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Plain | 0.088 | 0.131 | 0.671 | 1.49× | 0.078 | 0.125 |
| Styled | 1.193 | 1.346 | 1.691 | 1.13× | 1.156 | 1.281 |
| Formulas | 0.085 | 0.385 | 0.440 | 4.52× | 0.063 | 0.375 |
| Dates | 0.285 | 0.337 | 0.729 | 1.18× | 0.281 | 0.313 |
| Append control | 0.046 | 0.065 | 0.239 | 1.41× | 0.031 | 0.047 |

Higher crate/native time means the candidate is slower. CPU is total process CPU during the same measured interval, not CPU utilization percentage.

| Scenario | Native absolute RSS MiB | Crate absolute RSS MiB | openpyxl absolute RSS MiB | Native incremental MiB | Crate incremental MiB |
| --- | ---: | ---: | ---: | ---: | ---: |
| Plain | 49.4 | 56.3 | 48.6 | 13.6 | 20.7 |
| Styled | 65.6 | 72.1 | 53.6 | 29.8 | 36.4 |
| Formulas | 48.6 | 56.0 | 48.3 | 12.9 | 20.4 |
| Dates | 52.0 | 58.6 | 53.3 | 16.3 | 22.9 |
| Append control | 42.8 | 47.1 | 43.0 | 7.1 | 11.5 |

Native is faster than openpyxl here, but the small-case absolute-memory baseline is often similar or worse. Styled memory remains a problem. Do not generalize older lower-memory claims to every workload or to these absolute-RSS figures.

### Larger completed cases: 20,000 × 20

Seven repetitions per backend completed for plain and styled output, with semantic validation passing. This preliminary session used the `42b0ffb` native wheel plus the two local Python fixes later committed at `88ffb13`; it is **not** an exact final-wheel run. Both native and candidate arms used the same Python files and native extension within that session.

| Scenario | Native wall s | Crate wall s | openpyxl wall s | Native absolute RSS MiB | Crate absolute RSS MiB | openpyxl absolute RSS MiB |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Plain | 0.696 | 1.027 | 4.450 | 162.1 | 225.5 | 174.4 |
| Styled | 13.441 | 14.115 | 19.055 | 331.4 | 379.7 | 218.5 |

The full larger invocation **did not complete**: an openpyxl formula child exceeded its 300-second timeout and the parent reached its 1,000-second scheduling deadline. The precise cause of that stall was not isolated; do not attribute it to openpyxl computation, OneDrive, or this library without evidence. Incomplete formula repetitions and unstarted date/control cases are excluded from conclusions. Original logs/results are retained, not relabeled as a successful full run.

### Reader probe: 20,000 × 20

Seven alternating fresh processes per backend; one identical plain scalar workbook containing 400,000 nonempty cells. All **14/14 runs** passed the common scalar gate with matching sheet order, nonempty bounds/counts and checksum `9ea4a00de39be602`; no excluded cells. Integral float/int representations are normalized by the executable checksum; raw type histograms remain in the evidence.

| Metric, median | Native full workbook | Calamine stream |
| --- | ---: | ---: |
| Internal total including ZIP precheck and drop | 0.706 s | 0.239 s |
| External process lifetime | 1.462 s | 1.026 s |
| Externally sampled absolute RSS | 113.0 MiB | 9.1 MiB |
| Last sampled CPU (lower bound, not final CPU) | 0.703 s | 0.266 s |

The approximately **2.96× internal-time advantage** and **92% lower sampled RSS** are for **different storage/API contracts**. Native builds the current editable workbook and then iterates it; Calamine consumes a cell stream without a `Range` or a second RPyXL store. Existing ZIP prevalidation is included for both, but this does not establish equivalent XML security. This fixture does not qualify dates, errors, formula/style provenance, CR escape handling, or package preservation. A noncryptographic scalar checksum is a comparison gate, not a complete fidelity proof.

Executable SHA-256: `06603f9691ac71dee02bfb5448c7754f905b3fca1c4c1852aaf38ca25a41d581`.
Input SHA-256: `aef2a81db9819086f203517fd6a65780392ce0bf57393eac0ea3165803fe95fc`.

## Interpretation and next work

1. **Do not promote the crate writer.** The standard-mode adapter constructs another cell store after the existing native model, compiles/preflights formats, stages output and sometimes rewrites the archive. The duplicate-store cost is architectural. Formula fixup is a likely contributor to the measured formula gap, not a fully profiled attribution. Styled wrapper work is shared by both writer arms and remains a dominant optimization target.
2. **Optimize the current single-core flow first.** Profile the styled build/assignment and save phases; consider a native bulk style API, compact per-cell style IDs, fewer Python/PyO3 crossings, and reuse of compiled formats across saves with correct invalidation. Change and benchmark one bottleneck at a time. Do not add a second authoritative value model.
3. **Keep streaming writer work gated on fallible I/O.** Upstream row spilling uses panicking I/O, unsafe under this workspace's release `panic=abort`. A reviewed upstream fix/patch is required before testing a direct row-stream adapter. It must consume rows without first storing the whole workbook.
4. **Evaluate a scoped Calamine read-only path next.** First match the current read-only contract and security limits; test epochs, temporal values, blanks, exact text/CR, formula caches, style IDs, corrupt input and caller-stream ownership. Benchmark full consumption and first-row latency against native read-only—not just the editable path measured here. Keep editable loaded-workbook preservation native.
5. **Portable qualification before adoption.** Repeat qualifying same-contract tests/benchmarks on a second platform, include scaling/confidence intervals and temporary-disk measurements, and meet the predeclared gates. Additional multicore work comes later, after single-core bottlenecks are clear.

## Reproduction and retained evidence

Commands run from the repository root with the verified candidate wheel extracted at `dist/installed-backend-88ffb13` and dependencies in `.venv-backends`:

```sh
PYTHONPATH=dist/installed-backend-88ffb13 .venv-backends/Scripts/python.exe -m pytest tests/unit tests/compatibility tests/integration -q

PYTHONPATH=dist/installed-backend-88ffb13 .venv-backends/Scripts/python.exe benchmarks/compare_backends.py --backends native rust_xlsxwriter openpyxl --scenarios plain styled formulas dates --rows 2000 --cols 20 --repetitions 7 --timeout 60 --parent-timeout 500 --output-dir dist/backend-results

.venv-backends/Scripts/python.exe benchmarks/compare_read_backends.py --executable dist/rpyxl-reader-comparison-windows-x64-88ffb13d37b8bf67273b2a4480bb7bad79f2b31c/bench_backend_read.exe --input dist/backend-results/20261102T161503Z-ca95b8de/plain-r01-o1-native.xlsx --repetitions 7 --timeout 120 --parent-timeout 600 --sample-interval 0.005 --output-dir dist/reader-results
```

The reader command requires the existing fixture; generate the deterministic plain case with `compare_backends.py --keep-workbooks` on another machine and record its actual hash rather than assuming compressed workbook bytes are deterministic.

Tracked compact summaries and individual measurements: [JSON evidence](rust_backend_results_2026-11-02.json). Full local raw reports/logs/provenance remain under:

- `dist/backend-results/20261102T164448Z-efa8f40c/` — completed exact-wheel run.
- `dist/backend-results/20261102T161503Z-ca95b8de/` — incomplete larger invocation, including completed plain/styled cases and retained reader input.
- `dist/reader-results/20261102T165152Z-6e8b3b3d/` — completed reader probe.

These ignored directories are local evidence, not distributed package contents. The original `.venv-test` was restored to the published PyPI `0.2.0` wheel after a Windows uninstall permission failure; use the separate candidate environment/extraction for this experiment.
