import os
from pathlib import Path
from rpyxl._core import load_workbook as _load_workbook

def load_workbook(
    filename,
    read_only=False,
    keep_vba=False,
    data_only=False,
    keep_links=True,
    rich_text=False,
    **kwargs,
):
    """
    Open the given filename and return the workbook.

    :param filename: the path to open or a file-like object
    :type filename: string or a path-like object
    :param read_only: speed up reading, data_only=True is usually desirable
    :type read_only: bool
    :param keep_vba: preserve VBA content (default: False)
    :param data_only: controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet
    :type data_only: bool
    :param keep_links: whether links to external workbooks should be preserved
    :param rich_text: if True, rich text formatting is preserved
    :return: a :class:`Workbook` or :class:`ReadOnlyWorkbook` object
    """
    if isinstance(filename, (str, Path, os.PathLike)):
        filepath = os.fspath(filename)
    else:
        filepath = str(filename)

    if not os.path.exists(filepath):
        raise FileNotFoundError(f"[Errno 2] No such file or directory: '{filepath}'")

    try:
        return _load_workbook(filepath, read_only=bool(read_only), data_only=bool(data_only))
    except (ValueError, IOError, OSError, Exception) as e:
        msg = str(e)
        if any(
            marker in msg
            for marker in (
                "ZIP bomb",
                "Invalid XLSX",
                "Invalid zip",
                "Cannot find xl/workbook.xml",
                "Cannot open xl/workbook.xml",
                "not found in archive",
                "security violation",
            )
        ):
            from rpyxl.utils.exceptions import InvalidFileException
            raise InvalidFileException(msg) from e
        raise

__all__ = ["load_workbook"]
