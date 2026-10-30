
import os
from shutil import copyfileobj
from tempfile import TemporaryFile

from .._core import load_workbook as _load_workbook
from ..workbook import Workbook


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
    :type filename: string, path-like object, or seekable binary file-like object
    :param read_only: speed up reading, data_only=True is usually desirable
    :type read_only: bool
    :param keep_vba: preserve VBA content (default: False)
    :param data_only: controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet
    :type data_only: bool
    :param keep_links: whether links to external workbooks should be preserved
    :param rich_text: if True, rich text formatting is preserved
    :return: a :class:`Workbook` or :class:`ReadOnlyWorkbook` object
    """
    if kwargs:
        unexpected = next(iter(kwargs))
        raise TypeError(f"load_workbook() got an unexpected keyword argument {unexpected!r}")
    if keep_vba:
        raise NotImplementedError("VBA preservation is not implemented")
    if rich_text:
        raise NotImplementedError("Rich-text run preservation is not implemented")

    # Native loading snapshots binary streams, including for lazy read-only access.
    # The caller retains ownership; the stream is never closed by rpyxl.
    try:
        native = _load_workbook(filename, read_only=bool(read_only), data_only=bool(data_only))
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
            from ..utils.exceptions import InvalidFileException
            raise InvalidFileException(msg) from e
        raise

    # Validate with Rust before opening any metadata archive. Keep our own
    # disk-backed snapshot for deferred metadata access to caller-owned streams;
    # closing or repositioning the caller stream cannot invalidate the wrapper.
    source = (
        os.path.abspath(os.fsdecode(os.fspath(filename)))
        if isinstance(filename, (str, bytes, os.PathLike)) else filename
    )
    owns_source = False
    try:
        if not isinstance(filename, (str, bytes, os.PathLike)):
            position = filename.tell()
            source = TemporaryFile(mode='w+b')
            owns_source = True
            try:
                filename.seek(0)
                copyfileobj(filename, source)
                source.seek(0)
            finally:
                filename.seek(position)
        return Workbook(_native=native, _source=source, _owns_source=owns_source)
    except Exception:
        native.close()
        if owns_source:
            source.close()
        raise

__all__ = ["load_workbook"]
