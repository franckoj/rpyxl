# RPyXL Phase 5 Multi-Scale Benchmark Report

Generated at: 2026-10-20 13:08:05 UTC

**Environment**: Linux (aarch64), 8 Logical Cores, 3916.36 MB RAM

## 1. Multi-Scale Execution Summary

| Scale | Cells | Std Save openpyxl | Std Save RPyXL | Save Speedup | Write-Only openpyxl | Write-Only RPyXL | WO Speedup | Read openpyxl | Read RPyXL | Read Speedup | RPyXL Peak RAM | openpyxl Peak RAM |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Small (1k cells) | 1,000 | 0.0093s | 0.0020s | **4.71x** | 0.0130s | 0.0022s | **5.89x** | 0.0102s | 0.0024s | **4.18x** | 75.8 MB | 75.7 MB |
| Medium (10k cells) | 10,000 | 0.0512s | 0.0101s | **5.05x** | 0.0493s | 0.0071s | **6.97x** | 0.0428s | 0.0113s | **3.80x** | 80.1 MB | 79.7 MB |
| Large (100k cells) | 100,000 | 0.5115s | 0.0719s | **7.12x** | 0.4204s | 0.0648s | **6.49x** | 0.3389s | 0.0687s | **4.93x** | 117.9 MB | 107.8 MB |
| Stress (500k cells) | 500,000 | 2.7976s | 0.3447s | **8.12x** | 2.1202s | 0.2961s | **7.16x** | 1.6987s | 0.3312s | **5.13x** | 218.1 MB | 213.8 MB |

## 2. Detailed Throughput & Performance

| Scale | Standard Save Throughput | Write-Only Throughput | Read-Only Streaming Throughput | RPyXL CPU Utilization |
|---|---|---|---|---|
| Small (1k cells) | **505,082 cells/s** | **452,753 cells/s** | **409,948 cells/s** | 100% |
| Medium (10k cells) | **987,167 cells/s** | **1,414,669 cells/s** | **888,284 cells/s** | 96% |
| Large (100k cells) | **1,391,000 cells/s** | **1,542,970 cells/s** | **1,455,411 cells/s** | 67% |
| Stress (500k cells) | **1,450,658 cells/s** | **1,688,650 cells/s** | **1,509,715 cells/s** | 137% |

## 3. Visual Performance Comparisons

### 3.1 Standard In-Memory Save Time Across Scales
![Standard Save Time](phase5_scale_writing_time.png)

### 3.2 Streaming Write-Only Time Across Scales
![Write-Only Time](phase5_scale_write_only_time.png)

### 3.3 Reading Time Across Scales
![Reading Time](phase5_scale_reading_time.png)

### 3.4 Peak Memory Footprint (RAM)
![Memory Scaling](phase5_scale_memory.png)

