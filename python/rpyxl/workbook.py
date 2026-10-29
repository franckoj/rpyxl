from rpyxl._core import (
    Workbook as _CoreWorkbook,
    WriteOnlyWorkbook as _WriteOnlyWorkbook,
    ReadOnlyWorkbook as _ReadOnlyWorkbook,
)

class _WorkbookMeta(type):
    def __instancecheck__(cls, instance):
        return isinstance(instance, (_CoreWorkbook, _WriteOnlyWorkbook, _ReadOnlyWorkbook))

class Workbook(metaclass=_WorkbookMeta):
    """
    Workbook factory providing full openpyxl compatibility for standard,
    streaming write-only, and read-only workbooks.
    """
    def __new__(cls, write_only=False, *args, **kwargs):
        if write_only:
            return _WriteOnlyWorkbook(*args, **kwargs)
        return _CoreWorkbook(*args, **kwargs)

__all__ = ["Workbook"]

