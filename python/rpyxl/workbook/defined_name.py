"""Compatibility import path for workbook defined names."""
from .._models import DefinedName, DefinedNames

DefinedNameDict = DefinedNames
DefinedNameList = DefinedNames

__all__ = ['DefinedName', 'DefinedNames', 'DefinedNameDict', 'DefinedNameList']
