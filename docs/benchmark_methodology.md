# Benchmark Methodology

## Overview

The RPyXL benchmark suite measures performance and memory characteristics against `openpyxl` across equivalent workloads and environments.

## Methodology Rules

1. **Isolation**: Tests execute in controlled environments with explicit garbage collection between trials.
2. **Warmup**: Two warmup rounds precede measured repetitions to stabilize JIT/OS disk caches.
3. **Statistical Rigor**: At least 7 repetitions are measured for every workload. Metrics report Median, Mean, Min, Max, and Standard Deviation.
4. **Reproducible Datasets**: Deterministic random seeds (`Random(42)`) ensure logically identical datasets across runs.
5. **Resource Measurement**:
   - `psutil` monitors process resident set size (RSS) memory before, during, and after execution.
   - `time.perf_counter()` captures high-precision wall-clock time.
   - `time.process_time()` measures CPU execution time (user + system).
6. **No Fabricated Data**: All reported values derive directly from real measurements.

## Metrics & Formulas

- **Speedup**:
  $$\text{speedup} = \frac{\text{openpyxl median time}}{\text{rpyxl median time}}$$
- **Memory Reduction**:
  $$\text{reduction (\%)} = \left(1 - \frac{\text{rpyxl peak RSS}}{\text{openpyxl peak RSS}}\right) \times 100$$
- **Throughput**:
  $$\text{throughput} = \frac{\text{total units}}{\text{median wall time}}$$
