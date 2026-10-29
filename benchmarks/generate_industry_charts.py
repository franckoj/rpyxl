"""
Generate publication-quality graphical summary charts for the Industry-Level Spreadsheet Benchmark.
"""

import os
import json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

REPORT_DIR = "benchmarks/reports"
RAW_FILE = os.path.join(REPORT_DIR, "industry_benchmarks_raw.json")

with open(RAW_FILE, "r") as f:
    data = json.load(f)

SCALES = ["Small\n(1k cells)", "Medium\n(10k cells)", "Large\n(100k cells)", "Stress\n(500k cells)"]
SCALE_KEYS = ["Small (1k cells)", "Medium (10k cells)", "Large (100k cells)", "Stress (500k cells)"]
CELL_COUNTS = [1_000, 10_000, 100_000, 500_000]

# Modern styling palette
COLORS = {
    "openpyxl": "#E53935",       # Crimson / Red
    "xlsxwriter": "#FB8C00",     # Amber / Orange
    "rpyxl": "#43A047",           # Vibrant Emerald Green
    "rust_xlsx": "#1E88E5",      # Sapphire Blue
    "calamine": "#8E24AA",       # Deep Purple
}

plt.rcParams["font.sans-serif"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#CCCCCC"
plt.rcParams["axes.linewidth"] = 0.8

# ==============================================================================
# FIGURE 1: 2x2 EXECUTIVE DASHBOARD
# ==============================================================================
fig, axs = plt.subplots(2, 2, figsize=(16, 11), dpi=300)
fig.patch.set_facecolor("#FAFAFA")

for ax in axs.flat:
    ax.set_facecolor("#FFFFFF")
    ax.grid(True, linestyle="--", alpha=0.5, color="#E0E0E0")

x = np.arange(len(SCALES))
w = 0.20

# ------------------------------------------------------------------------------
# Panel A: Streaming Write Throughput
# ------------------------------------------------------------------------------
ax_w = axs[0, 0]
ox_w_tp = [CELL_COUNTS[i] / data["write_streaming"][SCALE_KEYS[i]]["openpyxl_stream"]["median"] for i in range(4)]
xw_w_tp = [CELL_COUNTS[i] / data["write_streaming"][SCALE_KEYS[i]]["xlsxwriter_py"]["median"] for i in range(4)]
px_w_tp = [CELL_COUNTS[i] / data["write_streaming"][SCALE_KEYS[i]]["rpyxl_stream"]["median"] for i in range(4)]
rx_w_tp = [CELL_COUNTS[i] / data["write_streaming"][SCALE_KEYS[i]]["rust_xlsxwriter"]["median"] for i in range(4)]

b1 = ax_w.bar(x - 1.5*w, [v/1000 for v in ox_w_tp], w, label="openpyxl (Python)", color=COLORS["openpyxl"], alpha=0.9)
b2 = ax_w.bar(x - 0.5*w, [v/1000 for v in xw_w_tp], w, label="xlsxwriter (Python)", color=COLORS["xlsxwriter"], alpha=0.9)
b3 = ax_w.bar(x + 0.5*w, [v/1000 for v in px_w_tp], w, label="RPyXL (Rust+PyO3)", color=COLORS["rpyxl"], alpha=0.95)
b4 = ax_w.bar(x + 1.5*w, [v/1000 for v in rx_w_tp], w, label="rust_xlsxwriter (Pure Rust)", color=COLORS["rust_xlsx"], alpha=0.9)

ax_w.set_title("A. Streaming Write Throughput (kCells/sec)", fontsize=13, fontweight="bold", pad=10)
ax_w.set_ylabel("Throughput (1,000 Cells / Second)", fontsize=11)
ax_w.set_xticks(x)
ax_w.set_xticklabels(SCALES, fontsize=10)
ax_w.legend(loc="upper left", frameon=True, facecolor="#FFFFFF", edgecolor="#DDDDDD", fontsize=9)

# Annotation on Stress
ax_w.annotate(
    "RPyXL: 1.65M cells/s\n(6.9x vs openpyxl)",
    xy=(3 + 0.5*w, px_w_tp[3]/1000),
    xytext=(2.2, 1400),
    arrowprops=dict(facecolor="#2E7D32", shrink=0.08, width=1.5, headwidth=6),
    fontsize=9, fontweight="bold", color="#1B5E20"
)

# ------------------------------------------------------------------------------
# Panel B: Standard Read Ingestion Throughput
# ------------------------------------------------------------------------------
ax_r = axs[0, 1]
w_r = 0.25
ox_r_tp = [CELL_COUNTS[i] / data["read_standard"][SCALE_KEYS[i]]["openpyxl"]["median"] for i in range(4)]
px_r_tp = [CELL_COUNTS[i] / data["read_standard"][SCALE_KEYS[i]]["rpyxl"]["median"] for i in range(4)]
cal_r_tp = [CELL_COUNTS[i] / data["read_standard"][SCALE_KEYS[i]]["calamine"]["median"] for i in range(4)]

ax_r.bar(x - w_r, [v/1000 for v in ox_r_tp], w_r, label="openpyxl (Python)", color=COLORS["openpyxl"], alpha=0.9)
ax_r.bar(x, [v/1000 for v in px_r_tp], w_r, label="RPyXL (Rust Core + DOM)", color=COLORS["rpyxl"], alpha=0.95)
ax_r.bar(x + w_r, [v/1000 for v in cal_r_tp], w_r, label="calamine (Pure Rust Arrays)", color=COLORS["calamine"], alpha=0.9)

ax_r.set_title("B. Standard Read Ingestion Throughput (kCells/sec)", fontsize=13, fontweight="bold", pad=10)
ax_r.set_ylabel("Throughput (1,000 Cells / Second)", fontsize=11)
ax_r.set_xticks(x)
ax_r.set_xticklabels(SCALES, fontsize=10)
ax_r.legend(loc="upper left", frameon=True, facecolor="#FFFFFF", edgecolor="#DDDDDD", fontsize=9)

ax_r.annotate(
    "RPyXL: 1.67M cells/s\n(9.2x vs openpyxl)",
    xy=(3, px_r_tp[3]/1000),
    xytext=(2.2, 1300),
    arrowprops=dict(facecolor="#2E7D32", shrink=0.08, width=1.5, headwidth=6),
    fontsize=9, fontweight="bold", color="#1B5E20"
)

# ------------------------------------------------------------------------------
# Panel C: Advanced Scenarios (Mutation & Multi-Sheet Parallelism)
# ------------------------------------------------------------------------------
ax_adv = axs[1, 0]
scenarios = ["100k Mutation\n(Read-Modify-Write)", "1M Cells Parallel\n(5 Sheets x 20k rows)"]
x_adv = np.arange(len(scenarios))
w_adv = 0.35

ox_adv_times = [
    data["mutation"]["openpyxl"]["median"],
    data["multisheet"]["openpyxl"]["median"],
]
px_adv_times = [
    data["mutation"]["rpyxl"]["median"],
    data["multisheet"]["rpyxl"]["median"],
]

bar_ox = ax_adv.bar(x_adv - w_adv/2, ox_adv_times, w_adv, label="openpyxl (GIL throttled)", color=COLORS["openpyxl"], alpha=0.9)
bar_px = ax_adv.bar(x_adv + w_adv/2, px_adv_times, w_adv, label="RPyXL (Rayon Multi-Threaded)", color=COLORS["rpyxl"], alpha=0.95)

ax_adv.set_title("C. Advanced Production Scenarios (Execution Time in Seconds)", fontsize=13, fontweight="bold", pad=10)
ax_adv.set_ylabel("Time (Seconds - Lower is Better)", fontsize=11)
ax_adv.set_xticks(x_adv)
ax_adv.set_xticklabels(scenarios, fontsize=10)
ax_adv.legend(loc="upper left", frameon=True, facecolor="#FFFFFF", edgecolor="#DDDDDD", fontsize=9)

# Value callouts on bars
for bar in bar_ox:
    yval = bar.get_height()
    ax_adv.text(bar.get_x() + bar.get_width()/2.0, yval + 0.15, f"{yval:.2f}s", ha="center", va="bottom", fontsize=10, fontweight="bold", color="#C62828")

for bar, speedup in zip(bar_px, ["8.33x faster", "5.71x faster"]):
    yval = bar.get_height()
    ax_adv.text(bar.get_x() + bar.get_width()/2.0, yval + 0.15, f"{yval:.2f}s\n({speedup})", ha="center", va="bottom", fontsize=10, fontweight="bold", color="#2E7D32")

ax_adv.set_ylim(0, 6.8)

# ------------------------------------------------------------------------------
# Panel D: Memory Footprint (Peak RSS on 500k Cells)
# ------------------------------------------------------------------------------
ax_m = axs[1, 1]
mem_labels = [
    "openpyxl\nStandard",
    "RPyXL\nStandard DOM",
    "RPyXL\nStreaming",
    "calamine\nRaw Arrays",
    "rust_xlsx\nPure Rust"
]
mem_values = [373.1, 225.1, 168.6, 168.6, 140.0]
mem_colors = [COLORS["openpyxl"], COLORS["rpyxl"], "#66BB6A", COLORS["calamine"], COLORS["rust_xlsx"]]

bar_m = ax_m.bar(range(len(mem_labels)), mem_values, 0.55, color=mem_colors, alpha=0.92)
ax_m.set_title("D. Peak Memory Consumption on 500k Cells (MB RAM)", fontsize=13, fontweight="bold", pad=10)
ax_m.set_ylabel("Peak RSS Memory (MB - Lower is Better)", fontsize=11)
ax_m.set_xticks(range(len(mem_labels)))
ax_m.set_xticklabels(mem_labels, fontsize=9.5)

for bar in bar_m:
    yval = bar.get_height()
    ax_m.text(bar.get_x() + bar.get_width()/2.0, yval + 8, f"{yval:.1f} MB", ha="center", va="bottom", fontsize=10, fontweight="bold", color="#333333")

ax_m.set_ylim(0, 430)
ax_m.axhline(373.1, color="#E53935", linestyle=":", alpha=0.6)
ax_m.annotate("148 MB RAM Saved (-40%)", xy=(1, 225.1), xytext=(1.3, 310),
             arrowprops=dict(facecolor="#2E7D32", shrink=0.08, width=1.5, headwidth=6),
             fontsize=9.5, fontweight="bold", color="#2E7D32")

plt.suptitle("RPyXL vs openpyxl, xlsxwriter, calamine & rust_xlsxwriter: Comprehensive Benchmark Dashboard",
             fontsize=16, fontweight="bold", y=0.99)
plt.tight_layout()
dashboard_path = os.path.join(REPORT_DIR, "industry_executive_summary_dashboard.png")
fig.savefig(dashboard_path)
plt.close()
print(f"Executive dashboard generated: {dashboard_path}")

# ==============================================================================
# FIGURE 2: SPEEDUP MULTIPLIER INFOGRAPHIC
# ==============================================================================
fig2, ax2 = plt.subplots(figsize=(11, 6), dpi=300)
fig2.patch.set_facecolor("#FAFAFA")
ax2.set_facecolor("#FFFFFF")
ax2.grid(True, linestyle="--", alpha=0.5, color="#E0E0E0")

categories = [
    "Standard Read (500k cells)",
    "Streaming Read TTFR Latency (500k)",
    "Streaming Write (500k cells)",
    "Standard Write (500k cells)",
    "Transactional Read-Modify-Write (100k)",
    "Multi-Sheet Parallel Serialization (1M)",
    "Write vs Python xlsxwriter (500k)"
]
speedups = [9.20, 4.30, 6.90, 4.20, 8.33, 5.71, 4.10]
bar_colors = ["#2E7D32" if s >= 6.0 else "#43A047" for s in speedups]

y_pos = np.arange(len(categories))
bars = ax2.barh(y_pos, speedups, 0.55, color=bar_colors, alpha=0.92)

ax2.set_yticks(y_pos)
ax2.set_yticklabels(categories, fontsize=10.5, fontweight="medium")
ax2.invert_yaxis()  # Top-down
ax2.set_xlabel("Speedup Factor vs Python Baselines (Higher is Better)", fontsize=11, fontweight="bold")
ax2.set_title("RPyXL Performance Multipliers vs Established Python Libraries", fontsize=14, fontweight="bold", pad=12)

ax2.axvline(1.0, color="#888888", linestyle="--", linewidth=1.5, label="Baseline (1.0x)")

for bar in bars:
    width = bar.get_width()
    ax2.text(width + 0.15, bar.get_y() + bar.get_height()/2.0, f"{width:.1f}x Faster",
             ha="left", va="center", fontsize=10.5, fontweight="bold", color="#1B5E20")

ax2.set_xlim(0, 11)
plt.tight_layout()
speedup_path = os.path.join(REPORT_DIR, "industry_speedup_breakdown.png")
fig2.savefig(speedup_path)
plt.close()
print(f"Speedup infographic generated: {speedup_path}")
