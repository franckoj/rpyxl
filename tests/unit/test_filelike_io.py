"""File-like contracts shared by the native fast path and metadata overlay."""
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO, StringIO
from zipfile import ZipFile

import pytest
from rpyxl.styles import Font
from rpyxl.utils.exceptions import InvalidFileException

from rpyxl import Workbook, load_workbook


def _book(metadata=False, extra_sheet=False):
    book = Workbook()
    book.active.title = 'Data'
    book.active.append(['Name', 'Amount'])
    book.active.append(['Example', 12.5])
    if metadata:
        book.active['B2'].font = Font(bold=True)
        book.active['B2'].number_format = '0.0000'
        book.properties.title = 'Stream metadata'
    if extra_sheet:
        sheet = book.create_sheet('Stale')
        for index in range(32):
            sheet.append([index, 'previous contents %d' % index])
    return book


@pytest.fixture
def source_bytes():
    book = _book(metadata=True, extra_sheet=True)
    try:
        with BytesIO() as output:
            book.save(output)
            return output.getvalue()
    finally:
        book.close()


@pytest.mark.parametrize('read_only', [False, True], ids=['editable', 'read-only'])
@pytest.mark.parametrize('position', [0, 7, 'end'])
def test_input_position_and_caller_ownership(source_bytes, read_only, position):
    with BytesIO(source_bytes) as source:
        position = len(source_bytes) if position == 'end' else position
        source.seek(position)
        book = load_workbook(source, read_only=read_only)
        try:
            assert source.tell() == position
            assert not source.closed
            assert list(book['Data'].values) == [('Name', 'Amount'), ('Example', 12.5)]
        finally:
            book.close()
        assert not source.closed
        assert source.tell() == position
        # Closing a workbook must not invalidate a caller-owned input.
        source.seek(0)
        assert source.read(2) == b'PK'


@pytest.mark.parametrize('read_only', [False, True], ids=['editable', 'read-only'])
def test_input_can_be_replaced_and_closed_before_deferred_reads(source_bytes, read_only):
    source = BytesIO(source_bytes)
    book = load_workbook(source, read_only=read_only)
    try:
        source.seek(0)
        source.write(b'not the workbook any more')
        source.truncate()
        source.close()
        # Exercise multiple sheets and repeated lazy iterations after closure,
        # rather than only the single styled cell in test_wrapper_metadata.
        for _ in range(2):
            assert list(book['Data'].values) == [('Name', 'Amount'), ('Example', 12.5)]
            assert len(list(book['Stale'].values)) == 32
        if read_only:
            with pytest.warns(RuntimeWarning, match='metadata for all sheets'):
                assert book['Data']['B2'].font.bold
        else:
            assert book['Data']['B2'].font.bold
        assert book['Data']['B2'].number_format == '0.0000'
        assert book.properties.title == 'Stream metadata'
    finally:
        book.close()
        source.close()


@pytest.mark.parametrize('read_only', [False, True], ids=['editable', 'read-only'])
def test_read_only_snapshot_survives_closed_real_file(tmp_path, source_bytes, read_only):
    path = tmp_path / 'source.xlsx'
    path.write_bytes(source_bytes)
    with path.open('rb') as source:
        source.seek(11)
        book = load_workbook(source, read_only=read_only)
        assert source.tell() == 11
    try:
        assert source.closed
        assert list(book['Data'].values)[1] == ('Example', 12.5)
        if read_only:
            with pytest.warns(RuntimeWarning, match='metadata for all sheets'):
                assert book.properties.title == 'Stream metadata'
        else:
            assert book.properties.title == 'Stream metadata'
    finally:
        book.close()


@pytest.mark.parametrize('metadata', [False, True], ids=['native', 'overlay'])
@pytest.mark.parametrize('destination', ['bytesio', 'binary-file'])
def test_output_replaces_truncates_and_can_be_reused(tmp_path, metadata, destination):
    target = BytesIO() if destination == 'bytesio' else (tmp_path / 'output.xlsx').open('w+b')
    large = _book(metadata=metadata, extra_sheet=True)
    small = _book(metadata=metadata)
    try:
        target.write(b'obsolete trailing contents' * 4096)
        target.seek(17)
        large.save(target)
        assert not target.closed
        first_size = target.seek(0, 2)
        assert first_size < len(b'obsolete trailing contents') * 4096
        target.seek(19)
        small.save(target)
        assert not target.closed
        second_size = target.seek(0, 2)
        assert second_size < first_size
        target.seek(0)
        payload = target.read()
        assert payload.startswith(b'PK\x03\x04')
        # A readable ZIP alone would not detect stale bytes after its directory.
        assert payload[-22:-18] == b'PK\x05\x06'
        assert payload[-2:] == b'\x00\x00'
        with ZipFile(BytesIO(payload)) as archive:
            assert archive.testzip() is None
        for read_only in (False, True):
            with BytesIO(payload) as source:
                loaded = load_workbook(source, read_only=read_only)
                try:
                    assert loaded.sheetnames == ['Data']
                    assert list(loaded.active.values) == [('Name', 'Amount'), ('Example', 12.5)]
                finally:
                    loaded.close()
    finally:
        large.close()
        small.close()
        target.close()


@pytest.mark.parametrize('metadata', [False, True], ids=['values-edit', 'style-edit'])
def test_loaded_book_can_save_over_its_input_stream(source_bytes, metadata):
    with BytesIO(source_bytes) as stream:
        book = load_workbook(stream)
        try:
            book['Data']['B2'] = 99
            if metadata:
                book['Data']['B2'].font = Font(italic=True)
            stream.seek(13)
            book.save(stream)
            assert not stream.closed
        finally:
            book.close()
        loaded = load_workbook(stream)
        try:
            assert loaded['Data']['B2'].value == 99
            assert loaded['Data']['B2'].font.bold is (not metadata)
            if metadata:
                assert loaded['Data']['B2'].font.italic
        finally:
            loaded.close()


@pytest.mark.parametrize('read_only', [False, True], ids=['editable', 'read-only'])
def test_relative_path_load_survives_cwd_change(tmp_path, monkeypatch, source_bytes, read_only):
    original = tmp_path / 'original'
    other = tmp_path / 'other'
    original.mkdir()
    other.mkdir()
    (original / 'source.xlsx').write_bytes(source_bytes)
    monkeypatch.chdir(original)
    book = load_workbook('source.xlsx', read_only=read_only)
    try:
        monkeypatch.chdir(other)
        for _ in range(2):
            assert list(book['Data'].values)[1] == ('Example', 12.5)
            assert len(list(book['Stale'].values)) == 32
        if read_only:
            with pytest.warns(RuntimeWarning, match='metadata for all sheets'):
                assert book['Data']['B2'].font.bold
        else:
            assert book['Data']['B2'].font.bold
        assert book.properties.title == 'Stream metadata'
    finally:
        book.close()


def test_failed_binary_input_preserves_position_and_ownership():
    with BytesIO(b'not an xlsx archive') as source:
        source.seek(5)
        with pytest.raises((InvalidFileException, ValueError, OSError)):
            load_workbook(source)
        assert source.tell() == 5
        assert not source.closed


def test_text_input_is_rejected_without_taking_ownership():
    with StringIO('not binary') as source:
        source.seek(3)
        with pytest.raises(TypeError, match='binary'):
            load_workbook(source)
        assert source.tell() == 3
        assert not source.closed


@pytest.mark.parametrize('read_only', [False, True], ids=['editable', 'read-only'])
@pytest.mark.parametrize('transport', ['path', 'bytesio'])
def test_threaded_reads_equal_serial_reads(tmp_path, source_bytes, read_only, transport):
    path = tmp_path / 'shared.xlsx'
    path.write_bytes(source_bytes)

    def read_once(index):
        # Each task owns its workbook and stream; the path alone may be shared.
        source = path if transport == 'path' else BytesIO(source_bytes)
        try:
            book = load_workbook(source, read_only=read_only)
            try:
                return (index, tuple(book.sheetnames),
                        tuple((sheet.title, tuple(sheet.values)) for sheet in book.worksheets))
            finally:
                book.close()
        finally:
            if isinstance(source, BytesIO):
                source.close()

    serial = [read_once(index) for index in range(4)]
    with ThreadPoolExecutor(max_workers=2) as pool:
        threaded = list(pool.map(read_once, range(4)))
    assert threaded == serial
