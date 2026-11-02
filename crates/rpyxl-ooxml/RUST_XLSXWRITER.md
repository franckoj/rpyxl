# Opt-in rust_xlsxwriter backend

The default OOXML writer and runtime reader are unchanged. The explicitly selected
backend is available as:

```rust,ignore
rpyxl_ooxml::crate_writer::save_workbook_rust_xlsxwriter(path, &workbook)?;
rpyxl_ooxml::crate_writer::write_workbook_rust_xlsxwriter(writer, &workbook)?;
```

The signatures use `P: AsRef<Path>` and `W: Write + Seek`, respectively, and return
`RpyxlResult<()>`. The writer need not be `Send`. Pass `&mut stream` to retain the
caller's stream; writing starts at its current position, without truncation.

## Supported input

Only **new, editable, normal, 1900-epoch core workbooks** are accepted. Multiple
sheets, sparse cells, sheet names/order, and the active sheet are supported. Empty
workbooks with no worksheets and the backend-reserved sheet name `History` are
rejected. A normal newly created workbook with its one empty worksheet works.

Values:

- Nonempty literal strings, including `=` prefixes, URLs, Unicode, spaces, tabs,
  LF, CR, and CRLF. No formula or hyperlink auto-detection. Excel's 32,767-character
  UTF-16-code-unit limit applies. Invalid XML characters and literal `_xHHHH_` escape tokens are
  rejected rather than changed by the backend/native reader combination.
- Signed integers only when both the conversion to `f64` **and the backend's
  decimal rendering** round-trip to the same integer. This accepts all integers
  through +/-2^53 and some larger exactly representable values, but rejects
  2^53+1, i64 extrema, and values such as 2^60 whose shortest decimal spelling
  would be read as a different integer by the native reader.
- Finite floating-point numbers and booleans. Numeric int/float classification
  on read follows the existing reader, not the original Rust enum variant.
- ISO dates, naive datetimes, and standalone times accepted by the existing
  `temporal_to_excel_serial` helper. The native `effective_number_format` helper
  supplies temporal defaults and replaces incompatible non-temporal formats.
  Native-reader temporal spelling/precision is unchanged (space-separated
  datetimes and millisecond resolution). No timezone-aware values/leap seconds.
- Empty cells, including styled blanks (unstyled blanks need not materialize).
- Ordinary scalar formulas with arithmetic/references, quoted sheet names and
  strings, and the legacy functions listed below. A missing cache stays missing
  (`CellValue::Empty` in data-only core reads, `None` at the Python boundary), not
  zero or an empty string. Cached numbers, booleans, temporal values, and
  unambiguous nonempty text are supported. Nested/empty caches are rejected.
  Text caches that would be inferred as numeric, boolean (`TRUE`/`FALSE`), or one
  of the crate's recognized Excel errors are rejected to avoid changing type.

Allowed function names (case-insensitive):

```text
SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK IF IFERROR AND OR NOT
ROUND ROUNDUP ROUNDDOWN ABS INT MOD SUMIF SUMIFS COUNTIF COUNTIFS
VLOOKUP HLOOKUP INDEX MATCH CONCATENATE LEFT RIGHT MID LEN TRIM UPPER LOWER
TEXT VALUE DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND TODAY NOW
ISBLANK ISNUMBER ISTEXT ISERROR TRUE FALSE
```

Other functions are deliberately rejected because the crate can rewrite modern
functions or silently create dynamic-array metadata. This is not a formula
syntax validator or evaluator. Quoted sheet-name tokens and string literals
containing `(` are conservatively rejected, including `='FILTER('!A1`: upstream
can recognize and rewrite function substrings inside single-quoted sheet names.
Ordinary quoted names such as `'Sales Q1'` remain supported.
Formulas are limited to 8,192 UTF-16 code units; array
braces, external/structured references, `@`, spill/error `#` syntax, and line
breaks are outside the initial subset, including those characters inside formula
string literals.

## RichStyle subset

All five XML fragments are first checked by the native namespace/XML-security
validator, then by a closed schema. Standard SpreadsheetML namespace prefixes
are accepted, as is inheritance from the stylesheet's default namespace.
At most 64,000 compiled style combinations (including the default) are accepted
before the backend's final deduplication. Unknown attributes, children, duplicate children, foreign namespaces, extension
payloads, and non-whitespace element text are errors, never silently discarded.
Formatting is represented semantically, not by preserving source fragment bytes,
comments, XML attribute order, or explicit spellings of default values.

- **Font:** explicit nonempty name and size (>0, <=409), bold, italic, strike,
  single/double/accounting underline, super/subscript, optional family 0..5,
  charset 0..255, and opaque RGB color. Missing family/color/scheme do not inherit
  the crate's implicit themed font settings. Only absent/`none` scheme is
  supported. `outline`, `shadow`, `condense`, and `extend` must be false/absent.
  Missing font name/size is rejected because the crate always emits both.
  Font names containing CR, LF, or tab are rejected, including character-reference
  spellings in the input XML, rather than emitting attribute values that change
  under XML whitespace normalization.
- **Fill:** empty/none pattern without colors, or solid pattern with an opaque
  foreground RGB and no explicit background. Other patterns and gradients are
  outside the initial subset. The crate emits the normal indexed-64 default
  solid-fill background; it is not a user-supplied indexed color.
- **Border:** ordinary left/right/top/bottom lines, all 13 modeled line styles,
  optional opaque RGB. Other sides may only be empty. Outline must be true/
  absent; diagonal directions must be false/absent. A color on a side without a
  visible line is rejected because the crate would discard it. The crate emits
  the default automatic color when a visible side has no explicit color.
- **Alignment:** all standard horizontal/vertical modes, rotation 0..180 or 255,
  wrap, shrink, indent 0..250, reading order 0..2. `relativeIndent` must be zero
  and `justifyLastLine` false/absent. Conflicts that cause the crate to drop
  shrink or change horizontal alignment for indentation are rejected.
- **Protection:** locked (default true), hidden (default false).
- **Number format:** nonempty format code without CR, LF, or tab, with the existing
  temporal-format rules. Supplied codes are validated even if the temporal helper
  would replace them. Format codes are not interpreted as loaded XF IDs.

OOXML booleans accept `0`, `1`, `false`, and `true`. Boolean font elements without
`val` mean true; underline without `val` means single. Optional/default metadata
is accepted without inventing a Python `None` string. Modeled integer-valued
style fields also accept integral decimal spellings such as `2.0`.

**Colors must be 8 hexadecimal digits beginning with `FF`.** Six-digit RGB,
`00`-alpha ARGB (including rpyxl's six-digit-to-`00` padding), other alpha values,
theme/indexed/automatic source colors, and tint are rejected. The adapter never
silently changes an explicit ARGB alpha byte to the crate's forced `FF`.

## Rejected provenance and metadata

Workbook `read_only`, `data_only`, 1904 epoch, package templates, source style XML,
source styles, worksheet XML/structural provenance, source XF IDs, unsupported
cell payload flags, comments, hyperlinks, and merged ranges are errors.
`source_xf_explicit = true` alone is allowed: new style assignments and explicit
blank/default writes legitimately set it. It does not bypass any workbook,
worksheet, or actual source-XF-ID provenance guard. In particular, loading a file produced by this backend
does **not** make the loaded workbook eligible for this backend again. Python's
additional non-style metadata checks belong to the binding layer.

## Memory, destination safety, and performance

Full validation and unique style compilation precede backend serialization.
Styles use an Arc-pointer cache for expected O(1) repeated lookup. Component XML
is hashed/compared only on the first encounter of each Arc, when content-equal
styles are deduplicated; a previously unseen style is still fully validated.

The runtime adapter uses standard `add_worksheet()` **in-memory** sheets and
shared strings. It creates a **second cell store**, and the library can also hold
serialized XML in memory. No memory win or default-writer promotion is claimed;
benchmarks must measure peak memory as well as elapsed time. No new thread pools
or multicore work were introduced.

**Constant-memory support is deferred.** In rust_xlsxwriter 0.100.0, row spilling
uses `expect`/`unwrap` on real temporary-file I/O. A disk-full failure can therefore
abort Python under the workspace's panic-abort policy; `map_err` cannot catch it.
The runtime adapter does not use constant/low-memory sheets, and the dependency's
`constant_memory` feature is not enabled here. Reconsider it only after fallible
spill I/O is verified or patched upstream, with an isolated benchmark first.

The completed ZIP is serialized to a temporary file before opening a filesystem
destination or touching a caller stream. Validation and serialization failures
therefore leave existing paths/streams intact and do not create absent paths.
The final copy is **not** atomic against destination I/O failures such as disk
full, and writing to a stream does not remove any pre-existing trailing bytes.

When CR text or uncached formulas occur, a streaming ZIP fixup converts generated
CR escapes to XML character references and removes the string type from empty
formula caches. It streams worksheet/shared-string XML, buffers at most one cell
or text event, raw-copies unrelated parts, and never changes the reader. This adds
disk/compression work for these workbooks;
CR-free books without uncached formulas skip the pass entirely. Temporary disk
capacity is required. No speedup is claimed without measurements on a working
Rust toolchain; benchmark both backends with the same eligible input and include
this staging/fixup cost.

`rust_xlsxwriter` is pinned to `=0.100.0` as a normal dependency, without
`constant_memory`. `calamine =0.38.0` is a dev-dependency solely for the separate reader
experiment; it is not a runtime reader replacement.
