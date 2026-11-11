#!/usr/bin/env python3
"""
Full test suite: openpyxl vs rpyxl (Rust port of openpyxl).

Sections (run all by default, or pick with --only):
  compat        ~45 API/feature probes: does each call work and round-trip?
  interop       correctness matrix: write with lib A, read with lib B, compare
                values (int/float/str/bool/date/datetime/formula/unicode/long
                text) and features (font, fill, number format, merged cells,
                column width, row height, freeze panes, auto filter, sheets)
  perf          ~20 performance scenarios (write, read, modify, styled, ...)
                -> wall time, CPU time, CPU %, peak memory, output file size
  scale         how time/memory grow with row count (linear? worse?)
  threads       4 loads serial vs 4 loads in threads (does Rust release the GIL?)
  concurrency   1/2/4 parallel processes (throughput scaling, total memory)
  leak          repeated load/save in one process (RSS growth)

Every measurement runs in a fresh subprocess so memory is not polluted.
The first library in --libs is the baseline for speedup/ratio columns.

Install:    pip install openpyxl rpyxl psutil
Default:    python full_test_openpyxl_vs_rpyxl.py
Full:       python full_test_openpyxl_vs_rpyxl.py --profile full
API only:   python full_test_openpyxl_vs_rpyxl.py --only compat
"""

import argparse
import csv
import datetime as dt
import gc
import importlib
import io
import json
import math
import os
import platform
import statistics
import subprocess
import sys
import tempfile
import threading
import time
from decimal import Decimal

MB = 1024 * 1024
ALL_SECTIONS = ["compat", "interop", "perf", "scale", "threads", "concurrency", "leak"]


# ==========================================================================
# helpers
# ==========================================================================
def sub(lib, name):
    """Import lib.<name> (e.g. 'styles', 'worksheet.table')."""
    try:
        return importlib.import_module(f"{lib.__name__}.{name}")
    except ImportError:
        return getattr(lib, name)


def _iter_rows(ws):
    try:
        for row in ws.iter_rows(values_only=True):
            yield row
    except TypeError:
        for row in ws.iter_rows():
            yield [c.value for c in row]


def norm(v):
    """Normalise a cell value into something JSON-able and comparable."""
    if isinstance(v, dt.datetime):
        return "dt:" + v.isoformat()
    if isinstance(v, dt.date):
        return "dt:" + dt.datetime.combine(v, dt.time()).isoformat()
    if isinstance(v, dt.time):
        return "tm:" + v.isoformat()
    if isinstance(v, Decimal):
        return float(v)
    return v


def same(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return type(a) is type(b) and a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-12)
    return a == b


def md_table(headers, rows):
    out = ["| " + " | ".join(headers) + " |", "|" + "|".join("---" for _ in headers) + "|"]
    for r in rows:
        out.append("| " + " | ".join(str(x) for x in r) + " |")
    return "\n".join(out)


# ==========================================================================
# perf data builders
# ==========================================================================
def make_row(r, cols):
    row = []
    for c in range(cols):
        k = c % 4
        if k == 0:
            row.append(r * cols + c)
        elif k == 1:
            row.append(r * 0.5 + c / 3)
        elif k == 2:
            row.append(f"text_{r}_{c}")
        else:
            row.append((r + c) % 2 == 0)
    return row


BASE_DT = dt.datetime(2024, 1, 1, 8, 0, 0)


def make_rich_row(r, cols):
    row = []
    for c in range(cols):
        k = c % 5
        if k == 0:
            row.append(r)
        elif k == 1:
            row.append(r * 0.25)
        elif k == 2:
            row.append(f"item_{r}_{c}")
        elif k == 3:
            row.append(BASE_DT + dt.timedelta(days=r % 1000))
        else:
            row.append(f"=A{r + 1}*2")
    return row


# ==========================================================================
# perf scenarios (library-agnostic, openpyxl-style API only)
# ==========================================================================
def sc_write_append(lib, rows, cols, ip, op):
    wb = lib.Workbook()
    ws = wb.active
    for r in range(rows):
        ws.append(make_row(r, cols))
    wb.save(op)


def sc_write_cell_assign(lib, rows, cols, ip, op):
    wb = lib.Workbook()
    ws = wb.active
    for r in range(rows):
        for c, v in enumerate(make_row(r, cols)):
            ws.cell(row=r + 1, column=c + 1, value=v)
    wb.save(op)


def sc_write_only_mode(lib, rows, cols, ip, op):
    wb = lib.Workbook(write_only=True)
    ws = wb.create_sheet()
    for r in range(rows):
        ws.append(make_row(r, cols))
    wb.save(op)


def sc_write_styled(lib, rows, cols, ip, op):
    st = sub(lib, "styles")
    font = st.Font(bold=True, color="FF0000")
    fill = st.PatternFill("solid", fgColor="FFFF00")
    border = st.Border(left=st.Side(style="thin"))
    align = st.Alignment(horizontal="center")
    wb = lib.Workbook()
    ws = wb.active
    for r in range(rows):
        for c, v in enumerate(make_row(r, cols)):
            cell = ws.cell(row=r + 1, column=c + 1, value=v)
            cell.font = font
            cell.fill = fill
            cell.border = border
            cell.alignment = align
            cell.number_format = "0.00"
    wb.save(op)


def sc_write_formulas(lib, rows, cols, ip, op):
    wb = lib.Workbook()
    ws = wb.active
    for r in range(rows):
        ws.append([r, r * 2, f"=SUM(A{r + 1}:B{r + 1})", f"=C{r + 1}*2"])
    wb.save(op)


def sc_write_dates(lib, rows, cols, ip, op):
    wb = lib.Workbook()
    ws = wb.active
    for r in range(rows):
        d = BASE_DT + dt.timedelta(days=r % 2000)
        ws.append([d, d.date(), d.time(), r])
    wb.save(op)


def sc_write_multisheet(lib, rows, cols, ip, op):
    wb = lib.Workbook()
    n = 10
    for s in range(n):
        ws = wb.active if s == 0 else wb.create_sheet(f"S{s}")
        for r in range(rows // n):
            ws.append(make_row(r, cols))
    wb.save(op)


def sc_write_bytesio(lib, rows, cols, ip, op):
    wb = lib.Workbook()
    ws = wb.active
    for r in range(rows):
        ws.append(make_row(r, cols))
    bio = io.BytesIO()
    wb.save(bio)
    with open(op, "wb") as f:
        f.write(bio.getvalue())


def sc_read_iter(lib, rows, cols, ip, op):
    ws = lib.load_workbook(ip).active
    n = sum(1 for row in _iter_rows(ws) for v in row if v is not None)
    assert n > 0


def sc_read_cells(lib, rows, cols, ip, op):
    ws = lib.load_workbook(ip).active
    max_r, max_c = ws.max_row, ws.max_column
    n = 0
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            if ws.cell(row=r, column=c).value is not None:
                n += 1
    assert n > 0


def sc_read_only_mode(lib, rows, cols, ip, op):
    wb = lib.load_workbook(ip, read_only=True)
    ws = wb.active
    n = sum(1 for row in _iter_rows(ws) for v in row if v is not None)
    wb.close()
    assert n > 0


def sc_read_data_only(lib, rows, cols, ip, op):
    ws = lib.load_workbook(ip, data_only=True).active
    n = sum(1 for _ in _iter_rows(ws))
    assert n > 0


def sc_read_iter_cols(lib, rows, cols, ip, op):
    ws = lib.load_workbook(ip).active
    n = 0
    for col in ws.iter_cols(values_only=True):
        n += len(col)
    assert n > 0


def sc_load_only(lib, rows, cols, ip, op):
    ws = lib.load_workbook(ip).active
    assert ws.max_row > 0


def sc_read_styled(lib, rows, cols, ip, op):
    ws = lib.load_workbook(ip).active
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            _ = cell.number_format
            _ = cell.font.bold
            _ = cell.is_date
            n += 1
    assert n > 0


def sc_read_multisheet(lib, rows, cols, ip, op):
    wb = lib.load_workbook(ip)
    n = 0
    for name in wb.sheetnames:
        for row in _iter_rows(wb[name]):
            n += len(row)
    assert n > 0


def sc_read_bytesio(lib, rows, cols, ip, op):
    with open(ip, "rb") as f:
        bio = io.BytesIO(f.read())
    ws = lib.load_workbook(bio).active
    assert sum(1 for _ in _iter_rows(ws)) > 0


def sc_modify(lib, rows, cols, ip, op):
    wb = lib.load_workbook(ip)
    ws = wb.active
    for r in range(1, min(ws.max_row, rows) + 1, 2):
        ws.cell(row=r, column=1).value = r * 10
        ws.cell(row=r, column=2).value = f"edited_{r}"
    wb.save(op)


def sc_roundtrip(lib, rows, cols, ip, op):
    wb = lib.load_workbook(ip)
    wb.save(op)


def sc_insert_delete_rows(lib, rows, cols, ip, op):
    wb = lib.load_workbook(ip)
    ws = wb.active
    ws.insert_rows(2, 100)
    ws.delete_rows(5, 100)
    wb.save(op)


def sc_t_serial(lib, rows, cols, ip, op):
    for _ in range(4):
        sc_read_iter(lib, rows, cols, ip, op)


def sc_t_threaded(lib, rows, cols, ip, op):
    from concurrent.futures import ThreadPoolExecutor

    with ThreadPoolExecutor(4) as ex:
        list(ex.map(lambda _: sc_read_iter(lib, rows, cols, ip, op), range(4)))


SCEN = {
    "write_append": (None, 1.0, sc_write_append),
    "write_cell_assign": (None, 0.5, sc_write_cell_assign),
    "write_only_mode": (None, 1.0, sc_write_only_mode),
    "write_styled": (None, 0.2, sc_write_styled),
    "write_formulas": (None, 0.5, sc_write_formulas),
    "write_dates": (None, 0.5, sc_write_dates),
    "write_multisheet": (None, 1.0, sc_write_multisheet),
    "write_bytesio": (None, 0.5, sc_write_bytesio),
    "read_iter": ("plain", 1.0, sc_read_iter),
    "read_cells": ("plain", 1.0, sc_read_cells),
    "read_only_mode": ("plain", 1.0, sc_read_only_mode),
    "read_data_only": ("rich", 0.2, sc_read_data_only),
    "read_iter_cols": ("plain", 0.5, sc_read_iter_cols),
    "load_only": ("plain", 1.0, sc_load_only),
    "read_styled": ("rich", 0.2, sc_read_styled),
    "read_multisheet": ("multi", 1.0, sc_read_multisheet),
    "read_bytesio": ("plain", 0.5, sc_read_bytesio),
    "modify": ("plain", 1.0, sc_modify),
    "roundtrip": ("plain", 1.0, sc_roundtrip),
    "insert_delete_rows": ("plain", 0.2, sc_insert_delete_rows),
    "t_serial": ("plain", 0.25, sc_t_serial),
    "t_threaded": ("plain", 0.25, sc_t_threaded),
}
PERF_SCENARIOS = [k for k in SCEN if not k.startswith("t_")]


# ==========================================================================
# input file generators
# ==========================================================================
def gen_plain(path, rows, cols):
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(rows):
        ws.append(make_row(r, cols))
    wb.save(path)


def gen_rich(path, rows, cols):
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    bold = Font(bold=True)
    for r in range(rows):
        ws.append(make_rich_row(r, cols))
        ws.cell(row=r + 1, column=1).font = bold
        ws.cell(row=r + 1, column=2).number_format = "0.00"
        ws.cell(row=r + 1, column=4).number_format = "yyyy-mm-dd hh:mm"
    wb.save(path)


def gen_multi(path, rows, cols):
    import openpyxl

    wb = openpyxl.Workbook()
    for s in range(5):
        ws = wb.active if s == 0 else wb.create_sheet(f"S{s}")
        for r in range(rows // 5):
            ws.append(make_row(r, cols))
    wb.save(path)


GEN = {"plain": gen_plain, "rich": gen_rich, "multi": gen_multi}


# ==========================================================================
# compat probes
# ==========================================================================
def _rt(lib, wb):
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb.save(p)
        return lib.load_workbook(p)
    finally:
        try:
            os.remove(p)
        except OSError:
            pass


def build_probes():
    P = []

    def probe(name):
        def deco(fn):
            P.append((name, fn))
            return fn

        return deco

    @probe("Workbook(), wb.active, ws.title")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.title = "X"
        assert _rt(lib, wb).active.title == "X"

    @probe("ws.append(row)")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, "a", 2.5])
        ws.append([3, "b", 4.5])
        assert [c.value for c in _rt(lib, wb).active[2]] == [3, "b", 4.5]

    @probe("ws['A1'] = v ; ws['A1'].value")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = "hi"
        ws["B2"].value = 7
        w2 = _rt(lib, wb).active
        assert w2["A1"].value == "hi" and w2["B2"].value == 7

    @probe("ws.cell(row, column, value=)")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.cell(row=3, column=2, value="z")
        assert _rt(lib, wb).active.cell(row=3, column=2).value == "z"

    @probe("ws['A1:C3'] range slice")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        for r in range(1, 4):
            ws.append([r, r, r])
        rows = ws["A1:C3"]
        assert len(rows) == 3 and len(rows[0]) == 3

    @probe("iter_rows(min/max row/col)")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        for r in range(5):
            ws.append([r] * 5)
        rows = list(ws.iter_rows(min_row=2, max_row=3, min_col=2, max_col=4))
        assert len(rows) == 2 and len(rows[0]) == 3

    @probe("iter_rows(values_only=True)")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, 2])
        ws.append([3, 4])
        assert list(ws.iter_rows(values_only=True)) == [(1, 2), (3, 4)]

    @probe("ws.iter_cols()")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, 2])
        ws.append([3, 4])
        cols = list(ws.iter_cols(values_only=True))
        assert [tuple(c) for c in cols] == [(1, 3), (2, 4)]

    @probe("ws.rows / ws.columns")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, 2])
        ws.append([3, 4])
        assert len(list(ws.rows)) == 2 and len(list(ws.columns)) == 2

    @probe("max_row / max_column / dimensions")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, 2, 3])
        ws.append([1, 2, 3])
        assert ws.max_row == 2 and ws.max_column == 3
        assert str(ws.dimensions) == "A1:C2"

    @probe("create_sheet / wb[name] / sheetnames / remove")
    def _(lib):
        wb = lib.Workbook()
        wb.create_sheet("B")
        wb.create_sheet("C", 0)
        assert wb.sheetnames[0] == "C" and "B" in wb.sheetnames
        wb.remove(wb["B"])
        assert "B" not in wb.sheetnames

    @probe("copy_worksheet")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 5
        c = wb.copy_worksheet(ws)
        assert c["A1"].value == 5

    @probe("number_format round-trip")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 1.2345
        ws["A1"].number_format = "0.00%"
        assert _rt(lib, wb).active["A1"].number_format == "0.00%"

    @probe("Font / PatternFill / Border / Alignment round-trip")
    def _(lib):
        st = sub(lib, "styles")
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = "x"
        c = ws["A1"]
        c.font = st.Font(bold=True, italic=True, size=14, color="FF0000")
        c.fill = st.PatternFill("solid", fgColor="00FF00")
        c.border = st.Border(left=st.Side(style="thin"))
        c.alignment = st.Alignment(horizontal="center", wrap_text=True)
        c2 = _rt(lib, wb).active["A1"]
        assert c2.font.bold and c2.font.italic and c2.font.size == 14
        assert c2.alignment.horizontal == "center" and c2.border.left.style == "thin"

    @probe("copy(cell.font) style copy")
    def _(lib):
        from copy import copy

        st = sub(lib, "styles")
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"].font = st.Font(bold=True)
        ws["B1"].font = copy(ws["A1"].font)
        assert ws["B1"].font.bold

    @probe("merge_cells / unmerge_cells")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.merge_cells("A1:B2")
        assert "A1:B2" in [str(r) for r in _rt(lib, wb).active.merged_cells.ranges]
        ws.unmerge_cells("A1:B2")
        assert len(ws.merged_cells.ranges) == 0

    @probe("column_dimensions width / row_dimensions height / hidden")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws.column_dimensions["A"].width = 33
        ws.row_dimensions[1].height = 44
        ws.column_dimensions["B"].hidden = True
        w2 = _rt(lib, wb).active
        assert w2.column_dimensions["A"].width == 33
        assert w2.row_dimensions[1].height == 44

    @probe("freeze_panes")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws.freeze_panes = "B2"
        assert _rt(lib, wb).active.freeze_panes == "B2"

    @probe("auto_filter.ref")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, 2])
        ws.append([3, 4])
        ws.auto_filter.ref = "A1:B2"
        assert _rt(lib, wb).active.auto_filter.ref == "A1:B2"

    @probe("insert_rows / delete_rows")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        for i in range(1, 6):
            ws.append([i])
        ws.insert_rows(2, 2)
        assert ws["A4"].value == 2
        ws.delete_rows(2, 2)
        assert ws["A2"].value == 2

    @probe("insert_cols / delete_cols")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.append([1, 2, 3])
        ws.insert_cols(2)
        assert ws["C1"].value == 2
        ws.delete_cols(2)
        assert ws["B1"].value == 2

    @probe("move_range")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 9
        ws.move_range("A1:A1", rows=2, cols=1)
        assert ws["B3"].value == 9

    @probe("formula stored as text")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 2
        ws["A2"] = "=A1*3"
        assert _rt(lib, wb).active["A2"].value == "=A1*3"

    @probe("load_workbook(data_only=True)")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 2
        ws["A2"] = "=A1*3"
        w2 = _rt(lib, wb)
        fd, p = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        w2.save(p)
        assert lib.load_workbook(p, data_only=True).active["A1"].value == 2
        os.remove(p)

    @probe("date / datetime / time values")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        d = dt.datetime(2024, 5, 6, 7, 8, 9)
        ws["A1"] = d
        ws["A2"] = dt.date(2024, 5, 6)
        ws["A3"] = dt.time(7, 8, 9)
        w2 = _rt(lib, wb).active
        assert w2["A1"].value == d
        assert w2["A2"].value in (dt.datetime(2024, 5, 6), dt.date(2024, 5, 6))
        assert w2["A3"].value == dt.time(7, 8, 9)

    @probe("Decimal value")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = Decimal("1.25")
        assert abs(float(_rt(lib, wb).active["A1"].value) - 1.25) < 1e-9

    @probe("save / load via BytesIO")
    def _(lib):
        wb = lib.Workbook()
        wb.active["A1"] = "q"
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        assert lib.load_workbook(bio).active["A1"].value == "q"

    @probe("load_workbook(read_only=True)")
    def _(lib):
        wb = lib.Workbook()
        wb.active.append([1, 2])
        fd, p = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        wb.save(p)
        w2 = lib.load_workbook(p, read_only=True)
        assert list(_iter_rows(w2.active))[0][0] == 1
        w2.close()
        os.remove(p)

    @probe("Workbook(write_only=True)")
    def _(lib):
        wb = lib.Workbook(write_only=True)
        ws = wb.create_sheet()
        ws.append([1, 2])
        ws.append([3, 4])
        assert list(_iter_rows(_rt(lib, wb).active)) == [(1, 2), (3, 4)]

    @probe("defined names")
    def _(lib):
        dn_mod = sub(lib, "workbook.defined_name")
        wb = lib.Workbook()
        dn = dn_mod.DefinedName("myname", attr_text="Sheet!$A$1")
        try:
            wb.defined_names["myname"] = dn
        except TypeError:
            wb.defined_names.append(dn)
        _rt(lib, wb)

    @probe("data validation")
    def _(lib):
        dv_mod = sub(lib, "worksheet.datavalidation")
        wb = lib.Workbook()
        ws = wb.active
        dv = dv_mod.DataValidation(type="list", formula1='"a,b,c"')
        ws.add_data_validation(dv)
        dv.add("A1")
        _rt(lib, wb)

    @probe("conditional formatting")
    def _(lib):
        rule = sub(lib, "formatting.rule")
        st = sub(lib, "styles")
        wb = lib.Workbook()
        ws = wb.active
        ws.conditional_formatting.add(
            "A1:A10",
            rule.CellIsRule(
                operator="greaterThan",
                formula=["5"],
                fill=st.PatternFill("solid", bgColor="FF0000"),
            ),
        )
        _rt(lib, wb)

    @probe("hyperlink")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = "link"
        ws["A1"].hyperlink = "https://example.com"
        h = _rt(lib, wb).active["A1"].hyperlink
        assert h is not None and "example.com" in str(getattr(h, "target", None) or h)

    @probe("cell comment")
    def _(lib):
        cm = sub(lib, "comments")
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws["A1"].comment = cm.Comment("hello", "me")
        assert _rt(lib, wb).active["A1"].comment.text == "hello"

    @probe("worksheet table")
    def _(lib):
        tb = sub(lib, "worksheet.table")
        wb = lib.Workbook()
        ws = wb.active
        ws.append(["a", "b"])
        ws.append([1, 2])
        ws.append([3, 4])
        ws.add_table(tb.Table(displayName="SalesTable", ref="A1:B3"))
        _rt(lib, wb)

    @probe("chart (BarChart)")
    def _(lib):
        ch = sub(lib, "chart")
        wb = lib.Workbook()
        ws = wb.active
        for i in range(1, 6):
            ws.append([i, i * 2])
        chart = ch.BarChart()
        chart.add_data(ch.Reference(ws, min_col=2, min_row=1, max_row=5))
        ws.add_chart(chart, "D2")
        _rt(lib, wb)

    @probe("tabColor / sheet_state hidden")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        wb.create_sheet("B")
        ws.sheet_properties.tabColor = "FF0000"
        wb["B"].sheet_state = "hidden"
        assert _rt(lib, wb)["B"].sheet_state == "hidden"

    @probe("wb.properties (creator/title)")
    def _(lib):
        wb = lib.Workbook()
        wb.properties.creator = "tester"
        wb.properties.title = "T"
        assert _rt(lib, wb).properties.creator == "tester"

    @probe("utils: get_column_letter / column_index_from_string")
    def _(lib):
        u = sub(lib, "utils")
        assert u.get_column_letter(28) == "AB"
        assert u.column_index_from_string("AB") == 28

    @probe("cell attrs: coordinate/column_letter/data_type/is_date")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["C5"] = dt.datetime(2024, 1, 1)
        c = ws["C5"]
        assert c.coordinate == "C5" and c.column_letter == "C" and c.row == 5
        assert c.is_date and c.data_type == "d"

    @probe("sheet protection")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws.protection.sheet = True
        assert _rt(lib, wb).active.protection.sheet

    @probe("page_setup (orientation)")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 1
        ws.page_setup.orientation = "landscape"
        assert _rt(lib, wb).active.page_setup.orientation == "landscape"

    @probe("wb.close()")
    def _(lib):
        wb = lib.Workbook()
        wb.close()

    @probe("large int / long string / unicode")
    def _(lib):
        wb = lib.Workbook()
        ws = wb.active
        ws["A1"] = 2**40
        ws["A2"] = "é日本😀" * 50
        ws["A3"] = "x" * 30000
        w2 = _rt(lib, wb).active
        assert w2["A1"].value == 2**40 and w2["A2"].value == "é日本😀" * 50
        assert len(w2["A3"].value) == 30000

    return P


def run_probes(libname):
    lib = importlib.import_module(libname)
    results = {}
    for name, fn in build_probes():
        try:
            fn(lib)
            results[name] = "OK"
        except BaseException as e:
            results[name] = f"FAIL: {type(e).__name__}: {str(e)[:70]}"
    print("RESULT_JSON:" + json.dumps(results))
    sys.exit(0)


# ==========================================================================
# interop (correctness) workers
# ==========================================================================
N_INTEROP = 200


def interop_row(r):
    return [
        r,
        r * 1.5 + 0.25,
        f"héllo 日本語 😀 <&> {r}\nline2",
        r % 2 == 0,
        dt.datetime(2024, 1, 1, 12, 30, 15) + dt.timedelta(days=r),
        None,
        f"=A{r + 1}+1",
        ("x" * 3000) if r % 50 == 0 else f"  pad {r}",
        (2**40) + r,
        1e-9 * (r + 1),
        dt.date(2024, 1, 1) + dt.timedelta(days=r),
    ]


EXPECTED_FEATURES = {
    "font_bold": True,
    "font_color": "FF0000",
    "fill_color": "FFFF00",
    "number_format": "0.000",
    "merged": ["L1:M2"],
    "col_width": 30,
    "row_height": 40,
    "freeze_panes": "B2",
    "auto_filter": "A1:K20",
    "sheetnames": ["Data", "Second"],
}


def write_interop(libname, path):
    lib = importlib.import_module(libname)
    st = sub(lib, "styles")
    wb = lib.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(N_INTEROP):
        ws.append(interop_row(r))
    ws2 = wb.create_sheet("Second")
    ws2.append(["second", 1, 2.5])
    ws2.append(["wörld", True, None])
    errs = {}

    def attempt(name, fn):
        try:
            fn()
        except BaseException as e:
            errs[name] = f"{type(e).__name__}: {str(e)[:60]}"

    attempt("font", lambda: setattr(ws["A1"], "font", st.Font(bold=True, color="FF0000")))
    attempt("fill", lambda: setattr(ws["B1"], "fill", st.PatternFill("solid", fgColor="FFFF00")))
    attempt("numfmt", lambda: setattr(ws["C1"], "number_format", "0.000"))
    attempt("merge", lambda: ws.merge_cells("L1:M2"))
    attempt("colwidth", lambda: setattr(ws.column_dimensions["A"], "width", 30))
    attempt("rowheight", lambda: setattr(ws.row_dimensions[1], "height", 40))
    attempt("freeze", lambda: setattr(ws, "freeze_panes", "B2"))
    attempt("filter", lambda: setattr(ws.auto_filter, "ref", "A1:K20"))
    wb.save(path)
    print("RESULT_JSON:" + json.dumps({"write_errors": errs}))
    sys.exit(0)


def read_dump(libname, path):
    lib = importlib.import_module(libname)
    wb = lib.load_workbook(path)
    out = {"sheets": {}, "features": {}}
    for name in wb.sheetnames:
        out["sheets"][name] = [[norm(c) for c in row] for row in _iter_rows(wb[name])]
    ws = wb["Data"]
    feat = out["features"]

    def chk(name, fn):
        try:
            feat[name] = fn()
        except BaseException as e:
            feat[name] = f"ERR {type(e).__name__}"

    chk("font_bold", lambda: bool(ws["A1"].font.bold))
    chk("font_color", lambda: str(ws["A1"].font.color.rgb)[-6:])
    chk("fill_color", lambda: str(ws["B1"].fill.fgColor.rgb)[-6:])
    chk("number_format", lambda: ws["C1"].number_format)
    chk("merged", lambda: [str(x) for x in ws.merged_cells.ranges])
    chk("col_width", lambda: ws.column_dimensions["A"].width)
    chk("row_height", lambda: ws.row_dimensions[1].height)
    chk("freeze_panes", lambda: ws.freeze_panes)
    chk("auto_filter", lambda: ws.auto_filter.ref)
    chk("sheetnames", lambda: list(wb.sheetnames))
    print("RESULT_JSON:" + json.dumps(out))
    sys.exit(0)


# ==========================================================================
# perf worker
# ==========================================================================
class PeakRSS(threading.Thread):
    def __init__(self, proc, interval=0.005):
        super().__init__(daemon=True)
        self.proc, self.interval = proc, interval
        self.peak = proc.memory_info().rss
        self._stop_evt = threading.Event()

    def run(self):
        while not self._stop_evt.is_set():
            try:
                self.peak = max(self.peak, self.proc.memory_info().rss)
            except Exception:
                pass
            time.sleep(self.interval)

    def stop(self):
        self._stop_evt.set()
        self.join()
        self.peak = max(self.peak, self.proc.memory_info().rss)


def perf_worker(libname, scenario, rows, cols, ip, op):
    import psutil

    lib = importlib.import_module(libname)
    fn = SCEN[scenario][2]
    proc = psutil.Process(os.getpid())
    gc.collect()
    base = proc.memory_info().rss
    sampler = PeakRSS(proc)
    sampler.start()
    c0 = proc.cpu_times()
    t0 = time.perf_counter()
    fn(lib, rows, cols, ip, op)
    t1 = time.perf_counter()
    c1 = proc.cpu_times()
    sampler.stop()
    wall = t1 - t0
    cpu = (c1.user - c0.user) + (c1.system - c0.system)
    out_mb = os.path.getsize(op) / MB if os.path.exists(op) else 0.0
    print(
        "RESULT_JSON:"
        + json.dumps(
            {
                "wall_s": wall,
                "cpu_s": cpu,
                "cpu_pct": cpu / wall * 100 if wall > 0 else 0.0,
                "peak_mb": max(sampler.peak - base, 0) / MB,
                "out_mb": out_mb,
            }
        )
    )
    sys.exit(0)


def leak_worker(libname, kind, rows, cols, ip, op, iters):
    import psutil

    lib = importlib.import_module(libname)
    proc = psutil.Process(os.getpid())
    series = []
    for _ in range(iters):
        if kind == "read":
            wb = lib.load_workbook(ip)
            _ = sum(1 for _r in _iter_rows(wb.active))
            del wb
        else:
            wb = lib.Workbook()
            ws = wb.active
            for r in range(rows):
                ws.append(make_row(r, cols))
            wb.save(op)
            del wb
        gc.collect()
        series.append(proc.memory_info().rss / MB)
    print("RESULT_JSON:" + json.dumps({"rss_mb": series}))
    sys.exit(0)


def compare_interop(dump):
    """Returns (mismatch_count, examples, failed_features)."""
    exp_rows = [[norm(v) for v in interop_row(r)] for r in range(N_INTEROP)]
    got = dump["sheets"].get("Data", [])
    mism, ex = 0, []
    for r, erow in enumerate(exp_rows):
        grow = got[r] if r < len(got) else []
        for c, ev in enumerate(erow):
            gv = grow[c] if c < len(grow) else None
            if not same(ev, gv):
                mism += 1
                if len(ex) < 4:
                    ex.append(f"r{r + 1}c{c + 1}: expected {str(ev)[:25]!r} got {str(gv)[:25]!r}")
    sec = dump["sheets"].get("Second", [])
    exp2 = [["second", 1, 2.5], ["wörld", True, None]]
    for r, erow in enumerate(exp2):
        grow = sec[r] if r < len(sec) else []
        for c, ev in enumerate(erow):
            gv = grow[c] if c < len(grow) else None
            if not same(ev, gv):
                mism += 1
                if len(ex) < 4:
                    ex.append(f"Second r{r + 1}c{c + 1}: expected {ev!r} got {gv!r}")
    bad = []
    for k, ev in EXPECTED_FEATURES.items():
        gv = dump["features"].get(k)
        if (
            isinstance(ev, (int, float))
            and not isinstance(ev, bool)
            and isinstance(gv, (int, float))
        ):
            ok = math.isclose(ev, gv, rel_tol=1e-6)
        else:
            ok = ev == gv
        if not ok:
            bad.append(f"{k} (expected {ev!r}, got {gv!r})")
    return mism, ex, bad


# ==========================================================================
# orchestrator
# ==========================================================================
class Runner:
    def __init__(self, args, tmp):
        self.a, self.tmp = args, tmp
        self.inputs = {}
        self.lines = []
        self.data = {}
        self.facts = {}

    def emit(self, s=""):
        print(s, flush=True)
        self.lines.append(s)

    def cmd(self, mode, lib, **kw):
        c = [sys.executable, os.path.abspath(__file__), "--mode", mode, "--lib", lib]
        for k, v in kw.items():
            c += [f"--{k.replace('_', '-')}", str(v)]
        return c

    def call(self, cmd):
        try:
            p = subprocess.run(cmd, capture_output=True, text=True, timeout=self.a.timeout)
        except subprocess.TimeoutExpired:
            raise RuntimeError(f"timeout > {self.a.timeout}s")
        if p.returncode != 0:
            msg = (p.stderr.strip() or p.stdout.strip()).splitlines()
            raise RuntimeError(msg[-1][:120] if msg else f"exit {p.returncode}")
        for line in p.stdout.splitlines():
            if line.startswith("RESULT_JSON:"):
                return json.loads(line[len("RESULT_JSON:") :])
        raise RuntimeError("no result line")

    def input_for(self, kind, rows, cols):
        key = (kind, rows, cols)
        if key not in self.inputs:
            path = os.path.join(self.tmp, f"in_{kind}_{rows}x{cols}.xlsx")
            print(f"  [generating {kind} input {rows:,}x{cols} with openpyxl] ", end="", flush=True)
            t = time.time()
            GEN[kind](path, rows, cols)
            print(f"{time.time() - t:.1f}s, {os.path.getsize(path) / MB:.1f} MB", flush=True)
            self.inputs[key] = path
        return self.inputs[key]

    def perf_once(self, lib, scenario, rows, cols, ip, tag=""):
        op = os.path.join(self.tmp, f"out_{lib}_{scenario}{tag}.xlsx")
        return self.call(
            self.cmd(
                "perf", lib, scenario=scenario, rows=rows, cols=cols, in_path=ip or "", out_path=op
            )
        )

    def perf_median(self, lib, scenario, rows, cols, ip, repeat):
        runs = []
        for i in range(repeat):
            print(f"  {scenario:<20} {lib:<9} [{i + 1}/{repeat}]", end="\r", flush=True)
            try:
                runs.append(self.perf_once(lib, scenario, rows, cols, ip))
            except RuntimeError as e:
                print(" " * 60, end="\r")
                return {"error": str(e)}
        print(" " * 60, end="\r")
        return {k: statistics.median(r[k] for r in runs) for k in runs[0]}

    def sec_compat(self):
        self.emit("\n## 1. API / feature compatibility\n")
        res = {}
        for lib in self.a.libs:
            try:
                res[lib] = self.call(self.cmd("probe", lib))
            except RuntimeError as e:
                res[lib] = {"__crash__": f"worker crashed: {e}"}
        names = [n for n, _ in build_probes()]
        rows = []
        for n in names:
            cells = []
            for lib in self.a.libs:
                v = res[lib].get(n, res[lib].get("__crash__", "n/a"))
                cells.append("OK" if v == "OK" else v.replace("|", "/"))
            rows.append([n] + cells)
        ok = {lib: sum(1 for n in names if res[lib].get(n) == "OK") for lib in self.a.libs}
        rows.append(["**TOTAL passing**"] + [f"**{ok[l]}/{len(names)}**" for l in self.a.libs])
        self.emit(md_table(["probe"] + self.a.libs, rows))
        self.data["compat"] = res
        self.facts["compat"] = {
            "total": len(names),
            "ok": ok,
            "failed": {
                l: [n for n in names if res[l].get(n, res[l].get("__crash__")) != "OK"]
                for l in self.a.libs
            },
        }

    def sec_interop(self):
        self.emit("\n## 2. Correctness / interoperability (writer x reader)\n")
        self.emit(
            "Each writer creates the same workbook; each reader loads it and every "
            "cell/feature is compared with the expected value.\n"
        )
        files, werr = {}, {}
        for w in self.a.libs:
            p = os.path.join(self.tmp, f"interop_{w}.xlsx")
            try:
                werr[w] = self.call(self.cmd("write_interop", w, out_path=p))["write_errors"]
                files[w] = p
            except RuntimeError as e:
                werr[w] = {"__crash__": str(e)}
        matrix, details = [], []
        for w in self.a.libs:
            row = [f"written by **{w}**"]
            for r in self.a.libs:
                if w not in files:
                    row.append("writer failed")
                    self.facts.setdefault("interop", []).append((w, r, None, None))
                    continue
                try:
                    d = self.call(self.cmd("read_dump", r, in_path=files[w]))
                    mism, ex, bad = compare_interop(d)
                    nfeat = len(EXPECTED_FEATURES)
                    self.facts.setdefault("interop", []).append((w, r, mism, len(bad)))
                    status = "PASS" if mism == 0 and not bad else "ISSUES"
                    row.append(
                        f"{status}: {mism} cell mismatches, features {nfeat - len(bad)}/{nfeat}"
                    )
                    if mism or bad:
                        details.append(f"- **{w} -> {r}**: " + "; ".join(ex + bad))
                except RuntimeError as e:
                    row.append(f"reader crashed: {str(e)[:50]}")
                    self.facts.setdefault("interop", []).append((w, r, None, None))
            matrix.append(row)
        self.emit(md_table([""] + [f"read by {l}" for l in self.a.libs], matrix))
        for w, e in werr.items():
            if e:
                self.emit(f"\nWriter `{w}` could not apply: {e}")
        if details:
            self.emit("\nDetails:\n" + "\n".join(details))
        self.data["interop"] = {"matrix": matrix, "details": details}

    def eff_rows(self, scale):
        return max(10, int(self.a.rows * scale))

    def sec_perf(self):
        self.emit("\n## 3. Performance matrix\n")
        self.emit(
            f"Base size {self.a.rows:,} rows x {self.a.cols} cols; some scenarios use a "
            f"fraction of the rows (shown). Median of {self.a.repeat} run(s). "
            f"Baseline = `{self.a.libs[0]}`.\n"
        )
        base = self.a.libs[0]
        table, raw = [], {}
        for sc in self.a.scenario_list:
            kind, scale, _ = SCEN[sc]
            rows = self.eff_rows(scale)
            ip = self.input_for(kind, rows, self.a.cols) if kind else None
            res = {
                lib: self.perf_median(lib, sc, rows, self.a.cols, ip, self.a.repeat)
                for lib in self.a.libs
            }
            raw[sc] = {"rows": rows, **res}
            b = res[base]
            for lib in self.a.libs:
                r = res[lib]
                if "error" in r:
                    table.append(
                        [sc, f"{rows:,}", lib, "FAILED", "", "", "", "", "", r["error"][:50]]
                    )
                    continue
                sp = mr = ""
                if lib != base and "error" not in b:
                    sp = f"{b['wall_s'] / r['wall_s']:.2f}x" if r["wall_s"] else ""
                    mr = f"{r['peak_mb'] / b['peak_mb']:.2f}x" if b["peak_mb"] else ""
                table.append(
                    [
                        sc,
                        f"{rows:,}",
                        lib,
                        f"{r['wall_s']:.3f}",
                        f"{r['cpu_s']:.3f}",
                        f"{r['cpu_pct']:.0f}",
                        f"{r['peak_mb']:.1f}",
                        f"{r['out_mb']:.2f}" if r["out_mb"] else "",
                        sp,
                        mr,
                    ]
                )
        self.emit(
            md_table(
                [
                    "scenario",
                    "rows",
                    "lib",
                    "wall s",
                    "cpu s",
                    "cpu %",
                    "peak MB",
                    "out MB",
                    "speedup",
                    "mem ratio",
                ],
                table,
            )
        )
        self.emit(
            "\nspeedup = baseline wall / lib wall (higher is faster); "
            "mem ratio = lib peak / baseline peak (lower is better)."
        )
        if len(self.a.libs) > 1:
            o = self.a.libs[1]
            h = []
            for sc, d in raw.items():
                if "error" in d[o] and "error" not in d[base]:
                    h.append(
                        [
                            sc,
                            f"{d[base]['wall_s']:.2f} -> FAILED",
                            "not supported",
                            "",
                            "",
                            "",
                            "",
                            d[o]["error"][:40],
                        ]
                    )
                    continue
                if "error" in d[base] or "error" in d[o]:
                    continue
                bw, ow = d[base], d[o]
                h.append(
                    [
                        sc,
                        f"{bw['wall_s']:.2f} -> {ow['wall_s']:.2f}",
                        f"{bw['wall_s'] / ow['wall_s']:.1f}x" if ow["wall_s"] else "",
                        f"{bw['cpu_s']:.2f} -> {ow['cpu_s']:.2f}",
                        f"{ow['cpu_s'] / bw['cpu_s']:.2f}x" if bw["cpu_s"] else "",
                        f"{bw['cpu_pct']:.0f}% -> {ow['cpu_pct']:.0f}%",
                        f"{bw['peak_mb']:.0f} -> {ow['peak_mb']:.0f}",
                        f"{ow['peak_mb'] / bw['peak_mb']:.2f}x" if bw["peak_mb"] else "",
                    ]
                )
            self.emit(f"\n**Head-to-head: speed, CPU and memory ({base} -> {o})**\n")
            self.emit(
                md_table(
                    [
                        "scenario",
                        "time s",
                        "speedup",
                        "CPU s",
                        "CPU ratio",
                        "CPU utilisation",
                        "peak MB",
                        "mem ratio",
                    ],
                    h,
                )
            )
            self.emit(
                "\nCPU ratio below 1.0x = less CPU work than the baseline. CPU utilisation "
                "is CPU time / wall time (about 100% = one core busy; above 100% = multi-threaded)."
            )
        self.data["perf"] = raw
        self._summarize(raw)
        os.makedirs(self.a.outdir, exist_ok=True)
        with open(os.path.join(self.a.outdir, "perf.csv"), "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(
                ["scenario", "rows", "lib", "wall_s", "cpu_s", "cpu_pct", "peak_mb", "out_mb"]
            )
            for sc, d in raw.items():
                for lib in self.a.libs:
                    r = d[lib]
                    if "error" not in r:
                        w.writerow(
                            [sc, d["rows"], lib]
                            + [
                                f"{r[k]:.4f}"
                                for k in ("wall_s", "cpu_s", "cpu_pct", "peak_mb", "out_mb")
                            ]
                        )

    def _summarize(self, raw):
        if len(self.a.libs) < 2:
            return
        base, other = self.a.libs[0], self.a.libs[1]
        sps, mrs = [], []
        for sc, d in raw.items():
            if "error" in d[base] or "error" in d[other]:
                continue
            if d[other]["wall_s"] > 0:
                sps.append(d[base]["wall_s"] / d[other]["wall_s"])
            if d[base]["peak_mb"] > 0 and d[other]["peak_mb"] > 0:
                mrs.append(d[other]["peak_mb"] / d[base]["peak_mb"])
        if sps:
            gm = math.exp(sum(math.log(x) for x in sps) / len(sps))
            self.emit(
                f"\n**Summary:** `{other}` vs `{base}` over {len(sps)} comparable scenarios: "
                f"geometric-mean speedup **{gm:.2f}x** (min {min(sps):.2f}x, max {max(sps):.2f}x); "
                f"median memory ratio **{statistics.median(mrs):.2f}x**."
            )
        failed = [sc for sc, d in raw.items() if any("error" in d[l] for l in self.a.libs)]
        if failed:
            self.emit(f"Scenarios with a failure: {', '.join(failed)}")

    def sec_scale(self):
        self.emit("\n## 4. Scaling with size\n")
        rows_out = []
        reps = min(self.a.repeat, 2)
        for n in self.a.sizes:
            ip = self.input_for("plain", n, self.a.cols)
            for lib in self.a.libs:
                w = self.perf_median(lib, "write_append", n, self.a.cols, None, reps)
                r = self.perf_median(lib, "read_iter", n, self.a.cols, ip, reps)
                cells = n * self.a.cols

                def f(x, k, fmt):
                    return "FAILED" if "error" in x else fmt.format(x[k])

                rows_out.append(
                    [
                        f"{n:,}",
                        lib,
                        f(w, "wall_s", "{:.2f}"),
                        "" if "error" in w else f"{w['wall_s'] / cells * 1e6:.2f}",
                        f(w, "peak_mb", "{:.0f}"),
                        f(r, "wall_s", "{:.2f}"),
                        "" if "error" in r else f"{r['wall_s'] / cells * 1e6:.2f}",
                        f(r, "peak_mb", "{:.0f}"),
                    ]
                )
        self.emit(
            md_table(
                [
                    "rows",
                    "lib",
                    "write s",
                    "write us/cell",
                    "write MB",
                    "read s",
                    "read us/cell",
                    "read MB",
                ],
                rows_out,
            )
        )
        self.emit("\nFlat us/cell across sizes = linear scaling; rising = super-linear.")
        self.data["scale"] = rows_out

    def sec_threads(self):
        self.emit("\n## 5. Threading (GIL release check)\n")
        rows = self.eff_rows(0.25)
        ip = self.input_for("plain", rows, self.a.cols)
        out = []
        for lib in self.a.libs:
            s = self.perf_median(lib, "t_serial", rows, self.a.cols, ip, self.a.repeat)
            t = self.perf_median(lib, "t_threaded", rows, self.a.cols, ip, self.a.repeat)
            if "error" in s or "error" in t:
                out.append([lib, "FAILED", "", "", "", (s.get("error") or t.get("error"))[:40]])
                continue
            self.facts.setdefault("threads", {})[lib] = (s["wall_s"] / t["wall_s"], t["cpu_pct"])
            out.append(
                [
                    lib,
                    f"{s['wall_s']:.2f}",
                    f"{t['wall_s']:.2f}",
                    f"{s['wall_s'] / t['wall_s']:.2f}x",
                    f"{t['cpu_pct']:.0f}",
                    f"{t['peak_mb']:.0f}",
                ]
            )
        self.emit(
            md_table(
                [
                    "lib",
                    "4 loads serial s",
                    "4 loads in 4 threads s",
                    "thread speedup",
                    "cpu % (threaded)",
                    "peak MB (threaded)",
                ],
                out,
            )
        )
        self.emit(
            "\nThread speedup near 4x and cpu % near 400 means the library releases the "
            "GIL while parsing; about 1x means loads are serialised."
        )
        self.data["threads"] = out

    def sec_concurrency(self):
        self.emit("\n## 6. Parallel processes (worker scaling)\n")
        rows = self.eff_rows(0.25)
        ip = self.input_for("plain", rows, self.a.cols)
        out = []
        for lib in self.a.libs:
            one = None
            for k in self.a.procs:
                cmds = []
                for i in range(k):
                    op = os.path.join(self.tmp, f"par_{lib}_{k}_{i}.xlsx")
                    cmds.append(
                        self.cmd(
                            "perf",
                            lib,
                            scenario="read_iter",
                            rows=rows,
                            cols=self.a.cols,
                            in_path=ip,
                            out_path=op,
                        )
                    )
                t0 = time.perf_counter()
                procs = [
                    subprocess.Popen(c, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                    for c in cmds
                ]
                results, failed = [], False
                for p in procs:
                    try:
                        so, se = p.communicate(timeout=self.a.timeout)
                    except subprocess.TimeoutExpired:
                        p.kill()
                        failed = True
                        continue
                    line = [l for l in so.splitlines() if l.startswith("RESULT_JSON:")]
                    if p.returncode != 0 or not line:
                        failed = True
                    else:
                        results.append(json.loads(line[0][12:]))
                total = time.perf_counter() - t0
                if failed or not results:
                    out.append([lib, k, "FAILED", "", "", ""])
                    continue
                avg = statistics.mean(r["wall_s"] for r in results)
                mem = sum(r["peak_mb"] for r in results)
                if k == 1:
                    one = total
                eff = f"{one / total * 100:.0f}%" if one else ""
                self.facts.setdefault("conc", {}).setdefault(lib, {})[k] = (
                    (one / total * 100) if one else None
                )
                out.append([lib, k, f"{total:.2f}", f"{avg:.2f}", f"{mem:.0f}", eff])
        self.emit(
            md_table(
                [
                    "lib",
                    "processes",
                    "total wall s",
                    "avg per-process s",
                    "sum peak MB",
                    "efficiency",
                ],
                out,
            )
        )
        self.emit(
            f"\nMachine has {os.cpu_count()} logical CPUs. Efficiency 100% = perfect "
            "scaling; lower means CPU/memory-bandwidth contention."
        )
        self.data["concurrency"] = out

    def sec_leak(self):
        self.emit("\n## 7. Memory leak check (repeated operations in one process)\n")
        rows = self.eff_rows(0.1)
        ip = self.input_for("plain", rows, self.a.cols)
        out = []
        for lib in self.a.libs:
            for kind in ("read", "write"):
                op = os.path.join(self.tmp, f"leak_{lib}_{kind}.xlsx")
                try:
                    s = self.call(
                        self.cmd(
                            "leak",
                            lib,
                            scenario=kind,
                            rows=rows,
                            cols=self.a.cols,
                            in_path=ip,
                            out_path=op,
                            iters=self.a.leak_iters,
                        )
                    )["rss_mb"]
                except RuntimeError as e:
                    out.append([lib, kind, "FAILED", "", "", str(e)[:40]])
                    continue
                base = s[1] if len(s) > 1 else s[0]
                growth = s[-1] - base
                verdict = "possible leak" if growth > max(10, 0.10 * base) else "stable"
                self.facts.setdefault("leak", []).append((lib, kind, verdict, growth))
                out.append([lib, kind, f"{s[0]:.0f}", f"{s[-1]:.0f}", f"{growth:+.0f}", verdict])
        self.emit(
            md_table(
                [
                    "lib",
                    "op",
                    "RSS after iter 1 (MB)",
                    "RSS after last (MB)",
                    "growth since iter 2 (MB)",
                    "verdict",
                ],
                out,
            )
        )
        self.data["leak"] = out

    def summary(self):
        L = self.a.libs
        self.emit("\n## Summary of findings\n")
        if len(L) < 2:
            self.emit("Only one library was tested, so there is nothing to compare.")
            return
        base, other = L[0], L[1]
        f = self.facts
        pts, blockers = [], []

        def gm(xs):
            return math.exp(sum(math.log(x) for x in xs) / len(xs)) if xs else float("nan")

        # compatibility
        if "compat" in f:
            c = f["compat"]
            bad = c["failed"][other]
            line = (
                f"**API compatibility:** `{other}` passes {c['ok'][other]}/{c['total']} probes "
                f"(`{base}`: {c['ok'][base]}/{c['total']})."
            )
            if bad:
                line += " Not working: " + "; ".join(bad[:10])
                line += f" (+{len(bad) - 10} more)." if len(bad) > 10 else "."
                blockers.append(f"{len(bad)} unsupported/broken API probes")
            pts.append(line)

        # correctness
        if "interop" in f:
            iss = [(w, r, m, b) for w, r, m, b in f["interop"] if m is None or m or b]
            if not iss:
                pts.append(
                    "**Correctness:** every writer/reader combination round-trips all "
                    "tested values and features correctly."
                )
            else:
                desc = ", ".join(
                    f"{w}->{r} ("
                    + ("failed" if m is None else f"{m} cell mismatches, {b} feature diffs")
                    + ")"
                    for w, r, m, b in iss
                )
                pts.append(f"**Correctness issues:** {desc}. See section 2 for details.")
                blockers.append("data/feature mismatches in round-trip tests")

        # performance
        raw = self.data.get("perf")
        speedup_gm = None
        if raw:
            ratios, mems, bad_sc = {}, {}, []
            cpu_b = cpu_o = 0.0
            for sc, d in raw.items():
                if "error" in d[other]:
                    bad_sc.append(sc)
                    continue
                if "error" in d[base]:
                    continue
                if d[other]["wall_s"] > 0:
                    ratios[sc] = d[base]["wall_s"] / d[other]["wall_s"]
                if d[base]["peak_mb"] > 0 and d[other]["peak_mb"] > 0:
                    mems[sc] = d[other]["peak_mb"] / d[base]["peak_mb"]
                cpu_b += d[base]["cpu_s"]
                cpu_o += d[other]["cpu_s"]
            if ratios:
                speedup_gm = gm(list(ratios.values()))
                best = max(ratios, key=ratios.get)
                worst = min(ratios, key=ratios.get)
                groups = {"write": [], "read/load": [], "modify/save": []}
                for sc, v in ratios.items():
                    if sc.startswith("write"):
                        groups["write"].append(v)
                    elif sc.startswith("read") or sc == "load_only":
                        groups["read/load"].append(v)
                    else:
                        groups["modify/save"].append(v)
                gtxt = ", ".join(f"{k} {gm(v):.1f}x" for k, v in groups.items() if v)
                pts.append(
                    f"**Speed:** `{other}` is {speedup_gm:.2f}x faster than `{base}` overall "
                    f"(geometric mean of {len(ratios)} scenarios). By group: {gtxt}. "
                    f"Best: {best} {ratios[best]:.1f}x; weakest: {worst} {ratios[worst]:.2f}x."
                )
                slower = [f"{k} ({v:.2f}x)" for k, v in ratios.items() if v < 0.9]
                if slower:
                    pts.append(
                        "**Clearly slower than baseline (below 0.9x):** "
                        + ", ".join(slower)
                        + ". Differences within about 10% are normal run-to-run noise."
                    )
            if mems:
                mb = min(mems, key=mems.get)
                mw = max(mems, key=mems.get)
                pts.append(
                    f"**Memory:** median peak-memory ratio {statistics.median(mems.values()):.2f}x "
                    f"of `{base}` (lower is better). Best: {mb} {mems[mb]:.2f}x; "
                    f"worst: {mw} {mems[mw]:.2f}x."
                )
            if cpu_b > 0:
                pts.append(
                    f"**CPU:** total CPU time {cpu_o:.0f}s vs {cpu_b:.0f}s "
                    f"({(1 - cpu_o / cpu_b) * 100:.0f}% less CPU work)."
                )
            if bad_sc:
                pts.append(
                    "**Scenarios that failed for " + other + ":** " + ", ".join(bad_sc) + "."
                )
                plur = "s" if len(bad_sc) != 1 else ""
                blockers.append(f"{len(bad_sc)} perf scenario{plur} failed")

            # always-tested operations
            sub_lines = []
            for sc in ALWAYS_SCENARIOS:
                if sc not in raw:
                    continue
                b, o, label = raw[sc][base], raw[sc][other], ALWAYS_LABELS[sc]
                if "error" in o:
                    extra = f"; `{base}` took {b['wall_s']:.1f}s" if "error" not in b else ""
                    sub_lines.append(
                        f"  - {label}: **not supported by `{other}`** ({o['error'][:70]}){extra}"
                    )
                elif "error" in b:
                    sub_lines.append(
                        f"  - {label}: failed for `{base}` ({b['error'][:60]}); "
                        f"`{other}` took {o['wall_s']:.2f}s"
                    )
                else:
                    sp = b["wall_s"] / o["wall_s"] if o["wall_s"] else float("nan")
                    mr = o["peak_mb"] / b["peak_mb"] if b["peak_mb"] else float("nan")
                    sub_lines.append(
                        f"  - {label}: {b['wall_s']:.2f}s -> {o['wall_s']:.2f}s (**{sp:.1f}x**), "
                        f"CPU {b['cpu_s']:.1f}s -> {o['cpu_s']:.1f}s ({o['cpu_pct']:.0f}% utilisation), "
                        f"peak memory {b['peak_mb']:.0f} -> {o['peak_mb']:.0f} MB ({mr:.2f}x)"
                    )
            if sub_lines:
                pts.insert(
                    0,
                    "**Always-tested operations** (append write, styled write, styled read, "
                    "insert/delete rows):\n" + "\n".join(sub_lines),
                )

        # scaling
        rows_scale = self.data.get("scale")
        if rows_scale:
            o_scale = [r for r in rows_scale if r[1] == other and r[3] and r[6]]
            if len(o_scale) >= 2:
                w0, w1 = float(o_scale[0][3]), float(o_scale[-1][3])
                r0, r1 = float(o_scale[0][6]), float(o_scale[-1][6])
                lin = (w1 / w0 < 1.5 and r1 / r0 < 1.5) if w0 and r0 else None
                if lin is not None:
                    pts.append(
                        "**Scaling:** "
                        + (
                            f"`{other}` scales roughly linearly with file size ({o_scale[0][0]} to {o_scale[-1][0]} rows)."
                            if lin
                            else f"`{other}` cost per cell changes noticeably with size "
                            f"(write {w0:.1f}->{w1:.1f} us/cell, read {r0:.1f}->{r1:.1f} us/cell); check section 4."
                        )
                    )

        # threads
        if "threads" in f and other in f["threads"]:
            sp, cp = f["threads"][other]
            if sp >= 2.5:
                verdict = "releases the GIL, so threads give real parallel speedup"
            elif sp >= 1.3:
                verdict = "is only partly parallel under threads"
            else:
                verdict = "shows no thread parallelism (loads are effectively serialised), so use processes instead"
            pts.append(
                f"**Threads:** 4 threaded loads ran {sp:.2f}x faster than serial "
                f"(CPU {cp:.0f}%); `{other}` {verdict}."
            )

        # concurrency
        if "conc" in f and other in f["conc"]:
            ks = {k: v for k, v in f["conc"][other].items() if v is not None}
            if ks:
                kmax = max(ks)
                pts.append(
                    f"**Parallel processes:** at {kmax} processes `{other}` kept {ks[kmax]:.0f}% "
                    f"efficiency on this {os.cpu_count()}-CPU machine "
                    + (
                        "(scales well)."
                        if ks[kmax] >= 80
                        else "(limited by available CPUs or memory bandwidth; size workers accordingly)."
                    )
                )

        # leak
        if "leak" in f:
            lk = [x for x in f["leak"] if x[0] == other]
            bad = [x for x in lk if x[2] != "stable"]
            if lk and not bad:
                pts.append(
                    f"**Memory leaks:** none detected for `{other}` over repeated load/save."
                )
            elif bad:
                pts.append(
                    "**Possible memory leak:** "
                    + ", ".join(f"{x[1]} (+{x[3]:.0f} MB)" for x in bad)
                    + f" for `{other}`."
                )
                blockers.append("possible memory growth over repeated operations")

        for pnt in pts:
            self.emit(f"- {pnt}")

        # overall verdict
        self.emit("\n**Overall verdict:**")
        if blockers:
            msg = f"`{other}` is not a full drop-in for `{base}` yet: " + "; ".join(blockers) + ". "
            if speedup_gm and speedup_gm >= 1.5:
                msg += (
                    f"It is clearly faster ({speedup_gm:.1f}x), so it is worth using for the code "
                    "paths whose features all passed, after validating output on real files."
                )
            self.emit(msg)
        elif speedup_gm is None:
            self.emit("Run the perf section for a speed verdict; no compatibility blockers found.")
        elif speedup_gm >= 1.5:
            self.emit(
                f"`{other}` looks like a drop-in replacement for the tested features, with a "
                f"{speedup_gm:.1f}x speedup. Validate against your real workbooks (styles, "
                "formulas, large files) before switching production."
            )
        elif speedup_gm >= 1.1:
            self.emit(
                f"`{other}` is compatible but only modestly faster ({speedup_gm:.2f}x); "
                "the gain may not justify switching."
            )
        else:
            self.emit(f"No meaningful speed benefit from `{other}` ({speedup_gm:.2f}x).")
        self.emit(
            "\n_Caveat: results come from synthetic data on this machine; run again on your "
            "real files and target hardware before deciding._"
        )


# ==========================================================================
# configuration: profiles, scenario groups, config file
# ==========================================================================
CORE_SCENARIOS = [
    "write_append",
    "write_only_mode",
    "write_styled",
    "read_iter",
    "read_only_mode",
    "read_cells",
    "load_only",
    "modify",
    "roundtrip",
]
ALWAYS_SCENARIOS = ["write_append", "write_styled", "read_styled", "insert_delete_rows"]
ALWAYS_LABELS = {
    "write_append": "Append write",
    "write_styled": "Styled write",
    "read_styled": "Styled read",
    "insert_delete_rows": "Insert/delete rows",
}
GROUPS = {
    "all": PERF_SCENARIOS,
    "core": CORE_SCENARIOS,
    "write": [x for x in PERF_SCENARIOS if x.startswith("write")],
    "read": [x for x in PERF_SCENARIOS if x.startswith("read") or x == "load_only"],
    "edit": ["modify", "roundtrip", "insert_delete_rows"],
}
PROFILES = {
    "quick": dict(
        sections=["compat", "perf"],
        scenarios=["write_append", "read_iter", "read_only_mode", "modify"],
        rows=5000,
        cols=10,
        repeat=1,
        sizes=[1000, 5000],
        procs=[1, 2, 4],
        leak_iters=4,
    ),
    "focus": dict(
        sections=["compat", "interop", "perf"],
        scenarios=CORE_SCENARIOS,
        rows=50_000,
        cols=20,
        repeat=2,
        sizes=[1000, 10_000, 50_000],
        procs=[1, 2, 4],
        leak_iters=6,
    ),
    "full": dict(
        sections=ALL_SECTIONS,
        scenarios=PERF_SCENARIOS,
        rows=50_000,
        cols=20,
        repeat=3,
        sizes=[1000, 10_000, 100_000],
        procs=[1, 2, 4],
        leak_iters=8,
    ),
}
CONFIG_KEYS = [
    "profile",
    "libs",
    "sections",
    "scenarios",
    "rows",
    "cols",
    "repeat",
    "sizes",
    "procs",
    "leak_iters",
    "timeout",
    "outdir",
]


def resolve_scenarios(tokens):
    out = []
    for t in tokens:
        if t in GROUPS:
            names = GROUPS[t]
        elif t in SCEN and not t.startswith("t_"):
            names = [t]
        else:
            sys.exit(f"Unknown scenario/group '{t}'. Use --list to see the options.")
        for n in names:
            if n not in out:
                out.append(n)
    return out


def print_list():
    print("Profiles (--profile):")
    for k, v in PROFILES.items():
        print(
            f"  {k:<6} sections={','.join(v['sections'])}  scenarios={len(v['scenarios'])}  "
            f"rows={v['rows']:,} cols={v['cols']} repeat={v['repeat']}"
        )
    print("\nSections (--only / --also / --skip):\n  " + ", ".join(ALL_SECTIONS))
    print("\nScenario groups (--scenarios):")
    for k, v in GROUPS.items():
        print(f"  {k:<7} {', '.join(v)}")
    print("\nScenarios and their row scale:")
    for k in PERF_SCENARIOS:
        print(f"  {k:<20} rows x {SCEN[k][1]}  input={SCEN[k][0] or 'none (writes)'}")


def build_config(args):
    profile = "quick" if args.quick else (args.profile or "focus")
    cfg = {k: (list(v) if isinstance(v, list) else v) for k, v in PROFILES[profile].items()}
    cfg.update(libs=["openpyxl", "rpyxl"], timeout=900, outdir="bench_results", profile=profile)
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            file_cfg = json.load(f)
        if "only" in file_cfg:
            file_cfg["sections"] = file_cfg.pop("only")
        if "profile" in file_cfg and not (args.profile or args.quick):
            profile = file_cfg["profile"]
            cfg.update(
                {k: (list(v) if isinstance(v, list) else v) for k, v in PROFILES[profile].items()}
            )
        unknown = set(file_cfg) - set(CONFIG_KEYS)
        if unknown:
            sys.exit(f"Unknown keys in config file: {sorted(unknown)}")
        cfg.update(file_cfg)
        cfg["profile"] = profile
    for key in (
        "libs",
        "rows",
        "cols",
        "repeat",
        "sizes",
        "procs",
        "leak_iters",
        "timeout",
        "outdir",
    ):
        v = getattr(args, key)
        if v is not None:
            cfg[key] = v
    if args.only:
        cfg["sections"] = args.only
    if args.also:
        cfg["sections"] = cfg["sections"] + [x for x in args.also if x not in cfg["sections"]]
    if args.skip:
        cfg["sections"] = [x for x in cfg["sections"] if x not in args.skip]
    if args.scenarios:
        cfg["scenarios"] = args.scenarios
    scen = resolve_scenarios(cfg["scenarios"])
    drop = set(resolve_scenarios(args.skip_scenarios)) if args.skip_scenarios else set()
    for x in ALWAYS_SCENARIOS:
        if x not in scen:
            scen.append(x)
    scen = [x for x in scen if x not in drop]
    for x in cfg["sections"]:
        if x not in ALL_SECTIONS:
            sys.exit(f"Unknown section '{x}'. Choose from {ALL_SECTIONS}")
    cfg["sections"] = [x for x in ALL_SECTIONS if x in cfg["sections"]]
    cfg["scenarios"] = scen
    return cfg


def run_backend_comparison(args, parser):
    import hashlib
    from pathlib import Path
    import uuid

    if args._backend_incompatible_options:
        parser.error(
            "--backend-comparison has a fixed validated matrix; use profile, rows, "
            "cols, repeat, timeout and outdir, not legacy section/scenario options"
        )
    if not args.reader_executable:
        parser.error("--backend-comparison requires --reader-executable from the branch CI build")
    if args.repeat < 7:
        parser.error("Backend comparison requires --repeat 7 or greater")
    if not 1 <= args.rows <= 1048576 or not 1 <= args.cols <= 16384 or args.timeout <= 0:
        parser.error("Rows/columns must fit Excel limits and timeout must be positive")

    root = Path(__file__).resolve().parent
    repo = root / "rpyxl"
    writer_driver = repo / "benchmarks" / "compare_backends.py"
    reader_driver = repo / "benchmarks" / "compare_read_backends.py"
    executable = Path(args.reader_executable).resolve()
    contract_tests = repo / "tests" / "unit" / "test_crate_backend.py"
    for path in (writer_driver, reader_driver, executable, contract_tests):
        if not path.is_file():
            parser.error(f"Required backend test/build file is missing: {path}")

    session_id = (
        dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:8]
    )
    artifacts = repo / "dist" / "main-backends" / session_id
    report_dir = Path(args.outdir).resolve() / ("backend-comparison-" + session_id)
    fixture = artifacts / "reader-input.xlsx"
    commands = [
        [sys.executable, "-m", "pytest", str(contract_tests), "-q"],
        [
            sys.executable,
            str(writer_driver),
            "--backends",
            "native",
            "rust_xlsxwriter",
            "openpyxl",
            "--scenarios",
            "plain",
            "styled",
            "formulas",
            "dates",
            "--rows",
            str(args.rows),
            "--cols",
            str(args.cols),
            "--repetitions",
            str(args.repeat),
            "--timeout",
            str(args.timeout),
            "--parent-timeout",
            str(args.timeout),
            "--output-dir",
            str(artifacts / "writer"),
        ],
        [
            sys.executable,
            str(reader_driver),
            "--executable",
            str(executable),
            "--input",
            str(fixture),
            "--repetitions",
            str(args.repeat),
            "--timeout",
            str(args.timeout),
            "--parent-timeout",
            str(args.timeout),
            "--output-dir",
            str(artifacts / "reader"),
        ],
    ]
    print(
        f"Explicit crate comparison: {args.rows:,} x {args.cols}, {args.repeat} repetitions",
        flush=True,
    )
    print(
        "Writers: native / rust_xlsxwriter / openpyxl; subset correctness gates required.",
        flush=True,
    )
    print(
        "Readers: native full materialization / Calamine stream; NOT equivalent APIs or security.",
        flush=True,
    )
    print(
        "Legacy compatibility/editing probes are not rerouted to unsupported crate operations.",
        flush=True,
    )
    print(f"Each stage has a {args.timeout}s bound; output: {report_dir}", flush=True)
    if args.dry_run:
        for command in commands:
            print(subprocess.list2cmdline(command))
        return 0

    lib = importlib.import_module("rpyxl")
    core = importlib.import_module("rpyxl._core")
    if not callable(getattr(lib.NativeWorkbook, "_save_rust_xlsxwriter", None)):
        parser.error(
            "Installed rpyxl lacks the candidate ABI. Use the rebuilt branch wheel; no native fallback."
        )
    artifacts.mkdir(parents=True, exist_ok=False)
    report_dir.mkdir(parents=True, exist_ok=False)

    def digest(path):
        return hashlib.sha256(Path(path).read_bytes()).hexdigest()

    evidence = {
        "scope": "Experimental subset, not production backend adoption",
        "python": sys.version,
        "rpyxl_package": lib.__file__,
        "native_extension": core.__file__,
        "native_extension_sha256": digest(core.__file__),
        "reader_executable_sha256": digest(executable),
        "main_py_sha256": digest(__file__),
        "rows": args.rows,
        "cols": args.cols,
        "repetitions": args.repeat,
        "commands": commands,
        "stages": [],
        "passed": False,
    }
    report = [
        "# Explicit Rust crate backend comparison",
        "",
        "Opt-in writer and dev-only reader probe; native defaults unchanged.",
        "Reader streaming versus full materialization is not a same-API speedup.",
        "",
    ]
    try:
        for index, (label, command) in enumerate(
            zip(("Writer contracts", "Writer performance", "Reader probe"), commands)
        ):
            if index == 2:
                print(
                    "Generating identical scalar reader input using main.py gen_plain (outside timing)",
                    flush=True,
                )
                gen_plain(str(fixture), args.rows, args.cols)
                evidence["reader_input_sha256"] = digest(fixture)
            log = report_dir / f"stage-{index + 1}.log"
            print(f"\n{label}: {subprocess.list2cmdline(command)}", flush=True)
            with log.open("w", encoding="utf-8") as output:
                try:
                    completed = subprocess.run(
                        command,
                        cwd=root,
                        stdout=output,
                        stderr=subprocess.STDOUT,
                        timeout=args.timeout + 30,
                    )
                    returncode = completed.returncode
                except subprocess.TimeoutExpired:
                    returncode = 124
            print(log.read_text(encoding="utf-8", errors="replace"), flush=True)
            evidence["stages"].append({"name": label, "returncode": returncode, "log": str(log)})
            if returncode:
                report.append(f"**FAILED: {label}** (exit {returncode}); remaining stages not run.")
                return returncode
            report.append(f"- {label}: passed; log `{log.name}`")

        writer = next((artifacts / "writer").glob("*/summary.json"))
        reader = next((artifacts / "reader").glob("*/summary.json"))
        evidence["writer"] = json.loads(writer.read_text(encoding="utf-8"))
        evidence["reader"] = json.loads(reader.read_text(encoding="utf-8"))
        evidence["passed"] = (
            evidence["writer"]["complete"]
            and all(case["semantic_comparable"] for case in evidence["writer"]["cases"].values())
            and evidence["reader"]["passed"]
        )
        report += [
            "",
            "## Writer matrix (median end-to-end seconds)",
            "",
            md_table(
                ["case", "native", "rust_xlsxwriter", "openpyxl"],
                [
                    [
                        name,
                        *(
                            f"{case['variants'][backend]['end_to_end_wall_seconds']['median']:.4f}"
                            for backend in ("native", "rust_xlsxwriter", "openpyxl")
                        ),
                    ]
                    for name, case in evidence["writer"]["cases"].items()
                ],
            ),
            "",
            "Library-default compression; not compression-level matched.",
            "",
            "## Reader probe",
            "",
            md_table(
                ["backend", "internal total s", "absolute sampled RSS MiB"],
                [
                    [
                        backend,
                        f"{stats['timing_ns']['overall_total']['median'] / 1e9:.4f}",
                        f"{stats['sampled_peak_rss_bytes']['median'] / MB:.1f}",
                    ]
                    for backend, stats in evidence["reader"]["statistics"].items()
                ],
            ),
            "",
            "Common scalar checksums must match; excluded types are not passes.",
            "Native full workbook versus Calamine stream: different storage/API contracts.",
        ]
        return 0 if evidence["passed"] else 1
    finally:
        (report_dir / "results.json").write_text(
            json.dumps(evidence, indent=2) + "\n", encoding="utf-8"
        )
        (report_dir / "report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
        print(f"Saved backend report and logs: {report_dir}", flush=True)


def main():
    ap = argparse.ArgumentParser(
        description="openpyxl vs rpyxl test suite (configurable)",
        epilog="Examples:\n"
        "  python %(prog)s                      # default 'focus' profile\n"
        "  python %(prog)s --profile quick      # 1-2 min smoke test\n"
        "  python %(prog)s --profile full       # everything (slow)\n"
        "  python %(prog)s --only compat        # API checks only (fast)\n"
        "  python %(prog)s --also threads leak  # focus profile + extra sections\n"
        "  python %(prog)s --scenarios write read_iter modify --repeat 1\n"
        "  python %(prog)s --config my.json     # settings from a JSON file\n"
        "  python %(prog)s --dry-run            # show the plan without running",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument(
        "--profile", choices=list(PROFILES), help="preset: quick | focus (default) | full"
    )
    ap.add_argument("--quick", action="store_true", help="alias for --profile quick")
    ap.add_argument("--config", help="JSON file with any of: " + ", ".join(CONFIG_KEYS))
    ap.add_argument("--save-config", help="write the effective settings to this JSON file")
    ap.add_argument(
        "--list", action="store_true", help="list profiles, sections and scenarios, then exit"
    )
    ap.add_argument(
        "--dry-run", action="store_true", help="print the plan and exit without running"
    )
    ap.add_argument("--libs", nargs="+", help="first one is the baseline (default: openpyxl rpyxl)")
    ap.add_argument("--only", nargs="+", choices=ALL_SECTIONS, help="run exactly these sections")
    ap.add_argument("--also", nargs="+", choices=ALL_SECTIONS, help="add sections to the profile")
    ap.add_argument(
        "--skip", nargs="+", choices=ALL_SECTIONS, help="remove sections from the profile"
    )
    ap.add_argument(
        "--scenarios", nargs="+", help="perf scenarios and/or groups: " + ", ".join(GROUPS)
    )
    ap.add_argument(
        "--skip-scenarios",
        nargs="+",
        help="perf scenarios/groups to leave out (also removes the always-run ones)",
    )
    ap.add_argument("--rows", type=int)
    ap.add_argument("--cols", type=int)
    ap.add_argument("--repeat", type=int, help="runs per scenario (median reported)")
    ap.add_argument("--sizes", nargs="+", type=int, help="row counts for the scale section")
    ap.add_argument(
        "--procs", nargs="+", type=int, help="process counts for the concurrency section"
    )
    ap.add_argument("--leak-iters", type=int)
    ap.add_argument("--timeout", type=int, help="seconds per subprocess (default 900)")
    ap.add_argument("--outdir")
    ap.add_argument(
        "--backend-comparison",
        action="store_true",
        help="explicit rust_xlsxwriter contracts/matrix and Calamine reader probe; no native fallback",
    )
    ap.add_argument(
        "--reader-executable", help="path to pinned branch CI bench_backend_read executable"
    )
    # internal worker args
    ap.add_argument("--mode", default="main", help=argparse.SUPPRESS)
    ap.add_argument("--lib", help=argparse.SUPPRESS)
    ap.add_argument("--scenario", help=argparse.SUPPRESS)
    ap.add_argument("--in-path", default="", help=argparse.SUPPRESS)
    ap.add_argument("--out-path", default="", help=argparse.SUPPRESS)
    ap.add_argument("--iters", type=int, default=5, help=argparse.SUPPRESS)
    args = ap.parse_args()

    # ---- worker dispatch ----
    if args.mode == "probe":
        run_probes(args.lib)
    if args.mode == "write_interop":
        write_interop(args.lib, args.out_path)
    if args.mode == "read_dump":
        read_dump(args.lib, args.in_path)
    if args.mode == "perf":
        perf_worker(args.lib, args.scenario, args.rows, args.cols, args.in_path, args.out_path)
    if args.mode == "leak":
        leak_worker(
            args.lib, args.scenario, args.rows, args.cols, args.in_path, args.out_path, args.iters
        )

    # ---- orchestrator ----
    if args.list:
        return print_list()
    args._backend_incompatible_options = any(
        (
            args.only,
            args.also,
            args.skip,
            args.scenarios,
            args.skip_scenarios,
            args.libs,
            args.config,
            args.save_config,
        )
    )
    args._repeat_explicit = args.repeat is not None
    cfg = build_config(args)
    for k, v in cfg.items():
        setattr(args, k, v)
    args.scenario_list = cfg["scenarios"]
    if args.backend_comparison:
        if not args._repeat_explicit and args.repeat < 7:
            args.repeat = 7
        raise SystemExit(run_backend_comparison(args, ap))
    if args.reader_executable:
        ap.error("--reader-executable requires --backend-comparison")

    plan = (
        f"profile={cfg['profile']} | sections={','.join(cfg['sections'])} | "
        f"perf scenarios={len(cfg['scenarios'])} | rows={cfg['rows']:,} cols={cfg['cols']} "
        f"repeat={cfg['repeat']} | libs={','.join(cfg['libs'])}"
    )
    if args.save_config:
        with open(args.save_config, "w", encoding="utf-8") as f:
            json.dump({k: cfg[k] for k in CONFIG_KEYS}, f, indent=2)
        print(f"Saved settings to {args.save_config}")
    if args.dry_run:
        print(plan)
        if "perf" in cfg["sections"]:
            print("Perf scenarios (effective rows):")
            for sc in cfg["scenarios"]:
                print(f"  {sc:<20} {max(10, int(cfg['rows'] * SCEN[sc][1])):>9,} rows")
        print("(dry run: nothing executed)")
        return

    try:
        import psutil
    except ImportError:
        sys.exit("psutil is required: pip install psutil")
    versions = {}
    for lib in args.libs:
        try:
            versions[lib] = getattr(importlib.import_module(lib), "__version__", "unknown")
        except ImportError:
            sys.exit(f"Cannot import '{lib}'. pip install {lib}")

    with tempfile.TemporaryDirectory() as tmp:
        R = Runner(args, tmp)
        R.emit("# openpyxl vs rpyxl: test report\n")
        R.emit(
            f"- Python {platform.python_version()} on {platform.system()} {platform.machine()}, "
            f"{os.cpu_count()} logical CPUs, {psutil.virtual_memory().total / 1024**3:.1f} GB RAM"
        )
        R.emit("- Libraries: " + ", ".join(f"{k} {v}" for k, v in versions.items()))
        R.emit(f"- Settings: {plan}")
        t_all = time.time()
        timings = []
        for sec in ALL_SECTIONS:
            if sec in cfg["sections"]:
                t0 = time.time()
                getattr(R, f"sec_{sec}")()
                timings.append([sec, f"{(time.time() - t0) / 60:.1f} min"])
        R.summary()
        R.emit("\n**Time spent per section** (use --skip / --only / --scenarios to trim):\n")
        R.emit(md_table(["section", "time"], timings))
        R.emit(f"\n_Total run time: {(time.time() - t_all) / 60:.1f} min_")

        os.makedirs(args.outdir, exist_ok=True)
        with open(os.path.join(args.outdir, "report.md"), "w", encoding="utf-8") as f:
            f.write("\n".join(R.lines))
        with open(os.path.join(args.outdir, "results.json"), "w", encoding="utf-8") as f:
            json.dump(R.data, f, indent=2, default=str)
        print(f"\nSaved report.md, results.json, perf.csv in {os.path.abspath(args.outdir)}")


if __name__ == "__main__":
    main()
