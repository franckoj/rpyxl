"""Phase 7 Security Hardening and Malicious Archive Defense Tests."""

import io
import os
import tempfile
import zipfile
import pytest

import rpyxl
from rpyxl.utils.exceptions import (
    InvalidFileException,
    SheetTitleException,
    OpenpyxlException,
)


def create_zip_archive(entries: dict[str, bytes]) -> str:
    """Helper to create a temporary ZIP file with given entries."""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)
    tmp.close()
    return tmp.name


def test_corrupted_truncated_zip():
    """Verify truncated or corrupted ZIP files raise InvalidFileException."""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(b"PK\x03\x04random corrupted data not a real zip archive")
    tmp.close()

    try:
        with pytest.raises((InvalidFileException, ValueError)):
            rpyxl.load_workbook(tmp.name)
    finally:
        os.unlink(tmp.name)


def test_missing_workbook_xml():
    """Verify archive missing workbook.xml raises InvalidFileException."""
    path = create_zip_archive({
        "[Content_Types].xml": b"<Types></Types>",
        "random.txt": b"Hello world",
    })

    try:
        with pytest.raises((InvalidFileException, ValueError)):
            rpyxl.load_workbook(path)
    finally:
        os.unlink(path)


def test_zip_bomb_high_compression_ratio_rejected():
    """Verify ZIP bomb with extreme compression ratio (>100:1) on large uncompressed stream is rejected."""
    # 10 MB of zeros compresses down to a few kilobytes (ratio ~1000:1)
    large_zeros = b"\x00" * (10 * 1024 * 1024)

    path = create_zip_archive({
        "_rels/.rels": b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": b'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/worksheets/sheet1.xml": large_zeros,
    })

    try:
        with pytest.raises((InvalidFileException, ValueError)) as excinfo:
            rpyxl.load_workbook(path)
        assert "ZIP bomb" in str(excinfo.value) or "compression ratio" in str(excinfo.value)
    finally:
        os.unlink(path)


def test_xml_doctype_xxe_entity_injection_rejected():
    """Verify XML files containing DOCTYPE/DTD entity injections are rejected."""
    malicious_shared_strings = (
        b'<?xml version="1.0" encoding="UTF-8"?>\n'
        b'<!DOCTYPE sst [\n'
        b'  <!ENTITY xxe SYSTEM "file:///etc/passwd">\n'
        b']>\n'
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        b'  <si><t>&xxe;</t></si>\n'
        b'</sst>'
    )

    path = create_zip_archive({
        "_rels/.rels": b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/_rels/workbook.xml.rels": b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        "xl/workbook.xml": b'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/worksheets/sheet1.xml": b'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>',
        "xl/sharedStrings.xml": malicious_shared_strings,
    })

    try:
        with pytest.raises((InvalidFileException, ValueError)) as excinfo:
            rpyxl.load_workbook(path)
        assert "security violation" in str(excinfo.value) or "DOCTYPE" in str(excinfo.value)
    finally:
        os.unlink(path)


def test_excel_bounds_protection():
    """Verify row and column boundaries strictly conform to Excel limits."""
    wb = rpyxl.Workbook()
    ws = wb.active

    # Row 0 and Col 0 are invalid
    with pytest.raises(ValueError):
        ws.cell(row=0, column=1)

    with pytest.raises(ValueError):
        ws.cell(row=1, column=0)

    # Upper Excel boundary limits: Max row 1,048,576, Max col 16,384
    ws.cell(row=1_048_576, column=16_384, value="MaxValid")
    assert ws.cell(row=1_048_576, column=16_384).value == "MaxValid"

    with pytest.raises(ValueError):
        ws.cell(row=1_048_577, column=1)

    with pytest.raises(ValueError):
        ws.cell(row=1, column=16_385)


def test_sheet_title_character_and_length_security():
    """Verify sheet titles cannot contain forbidden control/path chars or exceed 31 chars."""
    wb = rpyxl.Workbook()
    ws = wb.active

    # Forbidden characters: \ / ? * : [ ]
    for bad_char in ["\\", "/", "?", "*", ":", "[", "]"]:
        with pytest.raises((SheetTitleException, ValueError)):
            ws.title = f"Sheet{bad_char}Bad"

    # Max length: 31 characters
    valid_title = "A" * 31
    ws.title = valid_title
    assert ws.title == valid_title

    with pytest.raises((SheetTitleException, ValueError)):
        ws.title = "A" * 32
