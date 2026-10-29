# RPyXL Comprehensive Benchmark Report (Phases 1, 2, and 3)

Generated at: 2026-10-20 12:27:10 UTC

**Environment**: Linux (aarch64), 8 Logical Cores, 3916.36 MB RAM

## 1. Phase 3: High-Performance XLSX Writer & CPU Utilization Benchmarks

| Operation | openpyxl (Wall) | RPyXL (Wall) | Speedup | openpyxl CPU Time | RPyXL CPU Time | openpyxl RSS | RPyXL RSS | RPyXL Throughput |
|---|---|---|---|---|---|---|---|---|
| Standard Save (10,000 rows x 10 cols) | 0.4137s | 0.0422s | **9.80x** | 0.416s (106%) | 0.042s (100%) | 137.9 MB | 147.7 MB | 236,988 rows/s |
| Write-Only Streaming (10,000 rows x 10 cols) | 0.4449s | 0.0567s | **7.85x** | 0.447s (126%) | 0.057s (270%) | 137.3 MB | 137.3 MB | 176,472 rows/s |
| Write-Only Streaming (50,000 rows x 10 cols) | 2.2154s | 0.2801s | **7.91x** | 2.225s (123%) | 0.280s (324%) | 137.3 MB | 137.3 MB | 178,520 rows/s |
| Roundtrip Write + Read (5,000 rows x 10 cols) | 0.5367s | 0.0595s | **9.02x** | 0.538s (105%) | 0.060s (96%) | 146.2 MB | 145.2 MB | 84,052 rows/s |

### Phase 3 Performance Graphs

#### 1. Writing Execution Time (Wall-Clock)
![Writing Time](phase3_writing_time.png)

#### 2. Speedup Factor over openpyxl
![Writing Speedup](phase3_speedup.png)

#### 3. Peak RSS Memory Utilization
![Writing Memory](phase3_memory.png)

#### 4. CPU Consumption and Utilization Percentage
![Writing CPU](phase3_cpu_utilization.png)

## 2. Phase 2: High-Performance XLSX Reader Benchmarks

| Operation | openpyxl | RPyXL | Speedup | openpyxl RSS | RPyXL RSS | RPyXL Throughput |
|---|---|---|---|---|---|---|
| Full XLSX Read (Standard) | 0.5753s | 0.0623s | **9.24x** | 162.9 MB | 126.5 MB | 160,609 rows/s |
| Full XLSX Read (read_only=True) | 0.3288s | 0.0687s | **4.79x** | 109.9 MB | 126.7 MB | 145,617 rows/s |
| Time to First Row (read_only=True) | 2.68 ms | 0.96 ms | **2.78x** | 117.6 MB | 124.6 MB | Instantaneous |
| Time to First Row (Standard vs RO) | 534.83 ms | 0.96 ms | **555.9x** | 164.5 MB | 124.6 MB | Instantaneous |
| Calamine Full Read Baseline | 0.5753s (openpyxl std) | 0.0448s (calamine) | **12.85x** | 162.9 MB | 116.6 MB | 223,270 rows/s |

### Phase 2 Performance Graphs

#### 1. Total Reading Execution Time (10,000 rows)
![Reading Time](phase2_reading_time.png)

#### 2. Speedup Factor over openpyxl
![Reading Speedup](phase2_speedup.png)

#### 3. Peak RSS Memory Utilization
![Reading Memory](phase2_memory.png)

## 3. Phase 1: In-Memory Operations Benchmarks

| Operation | Dataset | openpyxl (s) | RPyXL (s) | Speedup | openpyxl RSS | RPyXL RSS | Mem Reduction | RPyXL Throughput |
|---|---|---|---|---|---|---|---|---|
| Workbook Creation | 100 instances | 0.0254s | 0.0s | **729.11x** | 0.63 MB | 0.05 MB | 92.1% | 2,874,306.6 workbooks/s |
| Cell Access (Write) | 100x100 (10,000 cells) | 0.0029s | 0.0045s | **0.65x** | 77.93 MB | 76.57 MB | 1.7% | 2,217,622.9 cells/s |
| Row Appending | 5,000 rows (50,000 cells) | 0.0523s | 0.0071s | **7.39x** | 97.76 MB | 80.49 MB | 17.7% | 705,882.3 rows/s |
| In-Memory Row Iteration | 5,000 rows (50,000 cells) | 0.016s | 0.0042s | **3.81x** | 93.82 MB | 94.79 MB | -1.0% | 11,932,698.6 cells/s |
