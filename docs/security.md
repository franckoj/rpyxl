# RPyXL Security Architecture & Threat Model

This document outlines RPyXL's defensive architecture against untrusted spreadsheets, malicious file formats, and resource exhaustion vectors.

---

## 1. Threat Vectors & Defenses

Spreadsheet processing engines are frequent targets of malicious payloads because Excel files are complex container formats (`.xlsx` is a ZIP package containing XML streams and binary assets). RPyXL implements multi-layered protections in its Rust core.

### 1.1 ZIP Bomb & Decompression Attacks

#### Threat
A malicious actor crafts a highly compressed file (e.g. 10 KB on disk that expands into 100 GB in memory upon decompression), leading to denial-of-service (DoS) via RAM or disk space exhaustion.

#### RPyXL Defenses
Before decompressing any file entries, RPyXL's `validate_zip_archive_security` inspects the ZIP central directory headers:
1. **Maximum Uncompressed Entry Limit**:
   - Limit: `512 MB` per entry (`MAX_UNCOMPRESSED_FILE_SIZE`).
   - Any XML stream declaring an uncompressed size > 512 MB is rejected immediately.
2. **Maximum Total Decompressed Archive Size**:
   - Limit: `1,024 MB` (`MAX_TOTAL_UNCOMPRESSED_SIZE`).
   - The cumulative uncompressed size of all entries across the archive cannot exceed 1 GB.
3. **Decompression Compression Ratio Limit**:
   - Ratio: `100.0:1` (`MAX_COMPRESSION_RATIO`) for streams >= 5 MB uncompressed.
   - If an entry's compression ratio exceeds 100:1, it is identified as an anomalous compression bomb and rejected with `InvalidFileException("ZIP bomb detected: compression ratio exceeds maximum safety ratio")`.
4. **Archive Entry Count Limit**:
   - Limit: `10,000` entries (`MAX_ZIP_ENTRIES`).
   - Protects against ZIP archives with millions of tiny entries designed to exhaust inode tables or directory iterators.

---

### 1.2 XML Entity Expansion (Billion Laughs & XXE)

#### Threat
XML parsers can be tricked into expanding recursive entities (`<!ENTITY lol "lol"><!ENTITY lol1 "&lol;&lol;">...`) or fetching external network resources (`<!ENTITY xxe SYSTEM "http://malicious.com/secret">`).

#### RPyXL Defenses
1. **DTD / DOCTYPE Strict Rejection**:
   - Standard ECMA-376 OOXML parts strictly forbid DTD / `<!DOCTYPE>` declarations.
   - RPyXL's `quick-xml` reader explicitly checks for `Event::DocType(_)`. Encountering any DOCTYPE declaration immediately halts parsing with `XmlSecurityViolation("DTD / DOCTYPE declaration is forbidden in XLSX XML streams")`.
   - External entity resolution (XXE) is completely eliminated because DTD processing is disabled.
2. **Memory Bounded Strings**:
   - Shared strings tables are capped at `10,000,000` entries to prevent unbounded heap allocation from malformed `<sst>` tags.

---

### 1.3 Memory Safety & Dangling Pointer Defenses

#### Threat
C/C++ Excel engines frequently suffer from use-after-free, buffer overflows, or out-of-bounds pointer writes when traversing nested structures.

#### RPyXL Defenses
1. **100% Memory-Safe Rust**:
   - RPyXL's workbook model, cell store, and XML serializers are written in safe Rust.
   - Workbooks, worksheets, and cells maintain thread-safe, reference-counted ownership (`Arc<RwLock<...>>`). If a worksheet or cell reference is retained in Python while the workbook is closed, memory remains valid and isolated without dangling pointers.
2. **Excel Grid Boundary Enforcement**:
   - Row indices are strictly validated against `1 <= row <= 1,048,576`.
   - Column indices are strictly validated against `1 <= col <= 16,384`.
   - Zero-based or out-of-bounds coordinates raise `ValueError` before memory operations occur.

---

### 1.4 Path Traversal & Sheet Injection Defenses

#### Threat
Malicious sheet names or relationship targets attempting path traversal (`../../etc/passwd` or control characters).

#### RPyXL Defenses
1. **Sheet Title Whitelisting**:
   - Rejects forbidden Excel characters (`\ / ? * : [ ]`).
   - Enforces max 31 character length.
   - Prevents empty names and single-quote wrapping attacks.
2. **ZIP Path Normalization**:
   - Internal paths resolved from `_rels/.rels` are sanitized and normalized (`normalize_zip_path`) to prevent directory traversal outside the ZIP root.

---

## 2. Configuration & Overrides

For trusted enterprise environments processing ultra-large scientific datasets, limits can be configured via environment variables or parameters in future releases (default limits protect standard deployments).
