"""Integration test validating Phase 1 objective flow."""

from rpyxl import Workbook


def test_phase1_master_flow():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"

    ws["A1"] = "Invoice ID"
    ws["B1"] = "Amount"
    ws.append(["INV001", 1500.50])

    assert ws["A1"].value == "Invoice ID"
    assert ws["B1"].value == "Amount"
    assert ws["A2"].value == "INV001"
    assert ws["B2"].value == 1500.50

    rows = list(ws.iter_rows(values_only=True))
    assert rows == [
        ("Invoice ID", "Amount"),
        ("INV001", 1500.50),
    ]

    wb.close()
