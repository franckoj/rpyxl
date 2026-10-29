# LinkedIn Launch Announcement: RPyXL v0.1.0

---

🚀 Excited to announce **RPyXL (v0.1.0)**: A high-performance, Rust-powered, 100% drop-in alternative to `openpyxl`!

If you've ever dealt with massive `.xlsx` spreadsheets in Python data pipelines, financial models, or ETL jobs, you know the pain: high memory usage, GIL bottlenecks, and scripts grinding to a halt for minutes on large files.

We built **RPyXL** to solve this once and for all by pairing bare-metal **Rust performance** with the beloved, ubiquitous **`openpyxl` API**.

---

### 💡 The 1-Line Migration

You don’t have to rewrite your codebase or learn a proprietary syntax:

```python
# Before
import openpyxl

# After (Everything else stays identical!)
import rpyxl as openpyxl

wb = openpyxl.load_workbook("large_financials.xlsx")
ws = wb.active
ws["A1"] = "High Speed Excel"
wb.save("output.xlsx")
```

---

### 📊 Benchmark Highlights: RPyXL vs openpyxl, xlsxwriter, calamine & rust_xlsxwriter

We tested 5 spreadsheet engines across 5 enterprise scenarios on a 500,000-cell realistic multi-type schema:

⚡ **Standard Ingestion**: **9.2× faster** than `openpyxl` (1.67M cells/sec vs 181k cells/sec), running within ~23% of pure `calamine` while providing a full mutable DOM.
⚡ **Streaming Export**: **6.9× faster** than `openpyxl` and **4.1× faster** than Python `xlsxwriter` (1.65M cells/sec), rivaling native `rust_xlsxwriter` (2.13M cells/sec).
⚡ **Transactional Mutation (Read-Modify-Write)**: **8.33× faster** (0.11s vs 0.95s on 100k cells).
⚡ **Multi-Sheet Parallelism**: **5.71× faster** (Over 1,000,000 cells across 5 sheets serialized in **0.97s** using multi-core Rayon threads, breaking through the Python GIL!).
🧠 **Memory Savings**: Uses **148 MB less RAM** on 500k cells (a 40% reduction).
🛡️ **Security Hardened**: Built-in protection against decompression ZIP bombs, XXE injection, and Billion Laughs XML attacks.

---

### 🗺️ What’s Next: The Roadmap to 50× Performance

A 5× to 10× speedup is just Phase 1. Here is what we're building next to unlock 20× to 50×+ performance (>10 Million cells/second):

1️⃣ **Assembly-Optimized & Parallel Deflate (`libdeflater`)**: Multi-threaded parallel chunk compression to cut `.save()` times by another 60%.
2️⃣ **Zero-Copy Arrow FFI (PyArrow / Polars Interop)**: Direct memory passing from DataFrames to Excel via the Arrow C Data Interface without Python object overhead.
3️⃣ **SIMD Vectorized Parsing (`simdutf8` / AVX-512 / ARM NEON)**: Vectorized XML scanning and SIMD numeric float/integer deserialization.
4️⃣ **Python 3.13 No-GIL Support**: Full free-threaded scaling across 16+ CPU cores.
5️⃣ **Native Excel Binary (.xlsb) Engine**: Sub-50ms instant serialization for enterprise massive tables.

---

📦 **PyPI / Wheels**: Universal ABI3 wheels (`cp39-abi3`) available for Linux, macOS (Intel & Apple Silicon), and Windows across Python 3.9 through 3.13.

⭐ Check out the open-source repo, benchmark reports, and migration guide:
👉 **GitHub**: https://github.com/franckoj/pyxl

I’d love to hear your thoughts, feedback, and edge-case benchmark requests! What’s your biggest bottleneck when working with spreadsheets in Python today?

#Python #Rust #OpenSource #DataEngineering #Performance #DataScience #Excel #SoftwareEngineering #PyO3
