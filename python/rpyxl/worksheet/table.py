"""Compatibility import path for worksheet tables."""
from .._models import Table, TableColumn, Tables, TableStyleInfo

TableList = Tables

__all__ = ['Table', 'TableColumn', 'TableStyleInfo', 'Tables', 'TableList']
