"""Pytest configuration and common fixtures."""

import openpyxl
import pytest
import rpyxl


@pytest.fixture
def rpyxl_wb():
    wb = rpyxl.Workbook()
    yield wb
    wb.close()


@pytest.fixture
def openpyxl_wb():
    wb = openpyxl.Workbook()
    yield wb
    wb.close()
