"""Single-pass metadata serialization and caller-tree preservation contracts."""

from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from rpyxl import NativeWorkbook, _metadata
from rpyxl._models import MetadataBook, MetadataSheet

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
IGNORABLE = _metadata.q("Ignorable", MC)
EXTENSIONS = (
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "http://schemas.microsoft.com/office/drawing/2010/main",
    "http://schemas.microsoft.com/office/excel/2006/main",
)


@pytest.fixture(autouse=True)
def isolated_namespace_registry(monkeypatch):
    # Namespace registrations made by a test must not leak into other suites.
    monkeypatch.setattr(ET, "_namespace_map", ET._namespace_map.copy())


def _snapshot(root):
    return [
        (
            node,
            node.tag,
            node.attrib,
            tuple(node.items()),
            node.text,
            node.tail,
            tuple(node),
        )
        for node in root.iter()
    ]


def _assert_unchanged(snapshot):
    for node, tag, attributes, items, text, tail, children in snapshot:
        assert node.tag is tag
        assert node.attrib is attributes
        assert tuple(node.items()) == items
        assert node.text is text
        assert node.tail is tail
        assert tuple(node) == children


def _assert_extensions(data, expected):
    namespaces = dict(pair for event, pair in ET.iterparse(BytesIO(data), events=("start-ns",)))
    root = ET.fromstring(data)
    prefixes = root.get(IGNORABLE, "").split()
    assert len(prefixes) == len(set(prefixes))
    assert {namespaces[prefix] for prefix in prefixes} == set(expected)
    assert all(node.get(IGNORABLE) is None for node in root.iter() if node is not root)
    return root, namespaces


def _stale_tree():
    root = ET.Element(
        _metadata.q("worksheet"),
        {"before": "one", IGNORABLE: "old14 old15", "after": 'two > "three"'},
    )
    root.text = "root text"
    child = ET.SubElement(root, _metadata.q("extLst"))
    child.attrib = {"first": "x", ET.QName(IGNORABLE): None, "last": "y"}
    child.tail = "tail text"
    ET.SubElement(child, _metadata.q("ext"), {IGNORABLE: ""})
    return root


def test_plain_dump_serializes_original_once_without_copying(monkeypatch):
    root = ET.Element(_metadata.q("worksheet"), {"custom": "kept"})
    ET.SubElement(root, _metadata.q("sheetData"))
    expected = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    snapshot = _snapshot(root)
    original = ET.tostring
    calls = []

    def serialize(node, **kwargs):
        assert node is root
        calls.append(node)
        return original(node, **kwargs)

    def forbidden(*args, **kwargs):
        raise AssertionError("dump must not copy the tree or register namespaces")

    monkeypatch.setattr(_metadata, "deepcopy", forbidden)
    monkeypatch.setattr(ET, "register_namespace", forbidden)
    monkeypatch.setattr(ET, "tostring", serialize)
    registry = ET._namespace_map.copy()
    assert _metadata.dump(root) == expected
    assert calls == [root]
    assert ET._namespace_map == registry
    _assert_unchanged(snapshot)


def test_stale_attributes_restore_dictionary_identity_order_and_aliases(monkeypatch):
    root = _stale_tree()
    # Distinct nodes may share the very same metadata attribute dictionary.
    ET.SubElement(root, _metadata.q("shared")).attrib = root[0].attrib
    snapshot = _snapshot(root)
    original = ET.tostring

    def serialize(node, **kwargs):
        assert all(IGNORABLE not in child.attrib for child in node.iter())
        for _, _, attributes, items, *_ in snapshot:
            assert tuple(attributes.items()) == items
        return original(node, **kwargs)

    monkeypatch.setattr(ET, "tostring", serialize)
    data = _metadata.dump(root)
    parsed, _ = _assert_extensions(data, ())
    assert parsed.attrib == {"before": "one", "after": 'two > "three"'}
    _assert_unchanged(snapshot)


@pytest.mark.parametrize("failure", [TypeError, RuntimeError, MemoryError])
def test_serialization_failure_restores_all_stale_attributes(monkeypatch, failure):
    root = _stale_tree()
    snapshot = _snapshot(root)

    def fail(node, **kwargs):
        assert all(IGNORABLE not in child.attrib for child in node.iter())
        raise failure("serialization failed")

    with monkeypatch.context() as patch:
        patch.setattr(ET, "tostring", fail)
        with pytest.raises(failure, match="serialization failed"):
            _metadata.dump(root)
    _assert_unchanged(snapshot)
    _assert_extensions(_metadata.dump(root), ())
    _assert_unchanged(snapshot)


def test_real_serialization_error_restores_attributes():
    root = _stale_tree()
    root[0].set("invalid", object())
    snapshot = _snapshot(root)
    with pytest.raises(TypeError, match="cannot serialize"):
        _metadata.dump(root)
    _assert_unchanged(snapshot)


@pytest.mark.parametrize("uri", EXTENSIONS)
@pytest.mark.parametrize("prefix", [None, "custom", "mc", "éxt", ""])
def test_extensions_follow_emitted_prefixes_and_default_namespace(uri, prefix):
    if prefix is not None:
        ET.register_namespace(prefix, uri)
    root = _stale_tree()
    extension = ET.SubElement(root[0], _metadata.q("feature", uri), {"value": "kept"})
    extension.text = "extension content"
    snapshot = _snapshot(root)
    registry = ET._namespace_map.copy()
    data = _metadata.dump(root)
    parsed, _ = _assert_extensions(data, (uri,))
    retained = parsed.find(".//" + extension.tag)
    assert retained is not None
    assert retained.attrib == {"value": "kept"}
    assert retained.text == "extension content"
    assert ET._namespace_map == registry
    _assert_unchanged(snapshot)


@pytest.mark.parametrize("usage", ["tag", "attribute", "qname-key", "qname-value"])
def test_namespace_discovery_includes_qname_and_attribute_only_usage(usage):
    uri = EXTENSIONS[0]
    name = _metadata.q("feature", uri)
    root = ET.Element(_metadata.q("worksheet"))
    if usage == "tag":
        ET.SubElement(root, ET.QName(name))
    elif usage == "attribute":
        root.set(name, "enabled")
    elif usage == "qname-key":
        root.set(ET.QName(name), "enabled")
    else:
        root.set("type", ET.QName(name))
    snapshot = _snapshot(root)
    parsed, namespaces = _assert_extensions(_metadata.dump(root), (uri,))
    if usage == "qname-value":
        prefix, local = parsed.get("type").split(":")
        assert namespaces[prefix] == uri
        assert local == "feature"
    _assert_unchanged(snapshot)


def test_existing_mc_prefix_is_reused_and_multiple_extensions_are_ignorable():
    ET.register_namespace("compat", MC)
    root = ET.Element(_metadata.q("worksheet"), {_metadata.q("Other", MC): "kept"})
    for uri in EXTENSIONS:
        ET.SubElement(root, _metadata.q("feature", uri))
    data = _metadata.dump(root)
    parsed, _ = _assert_extensions(data, EXTENSIONS)
    assert parsed.get(_metadata.q("Other", MC)) == "kept"
    assert b" compat:Ignorable=" in data
    assert b"xmlns:mc=" not in data


def test_locally_added_prefixes_do_not_collide_with_registered_names():
    ET.register_namespace("mc", EXTENSIONS[0])
    ET.register_namespace("mc1", "urn:other")
    root = ET.Element(_metadata.q("worksheet"), {_metadata.q("flag", "urn:other"): "yes"})
    ET.SubElement(root, _metadata.q("feature", EXTENSIONS[0]))
    data = _metadata.dump(root)
    _assert_extensions(data, (EXTENSIONS[0],))
    assert b" mc2:Ignorable=" in data


def test_default_extension_alias_avoids_an_existing_ext_prefix():
    ET.register_namespace("", EXTENSIONS[0])
    ET.register_namespace("ext", "urn:other")
    root = ET.Element(_metadata.q("worksheet"), {_metadata.q("flag", "urn:other"): "yes"})
    ET.SubElement(root, _metadata.q("feature", EXTENSIONS[0]))
    parsed, _ = _assert_extensions(_metadata.dump(root), (EXTENSIONS[0],))
    assert parsed.get(IGNORABLE) == "ext1"


def test_registration_rebound_after_serialization_does_not_change_ignorable(
    monkeypatch,
):
    uri = EXTENSIONS[0]
    ET.register_namespace("original", uri)
    root = ET.Element(_metadata.q("worksheet"))
    ET.SubElement(root, _metadata.q("feature", uri))
    original = ET.tostring
    register = ET.register_namespace
    calls = []

    def serialize(node, **kwargs):
        data = original(node, **kwargs)
        calls.append(node)
        register("original", "urn:rebound")
        register("later", uri)
        return data

    monkeypatch.setattr(ET, "tostring", serialize)
    parsed, namespaces = _assert_extensions(_metadata.dump(root), (uri,))
    assert calls == [root]
    assert parsed.get(IGNORABLE) == "original"
    assert namespaces["original"] == uri
    assert ET._namespace_map[uri] == "later"


def test_repeated_dump_tracks_registration_changes_without_mutating_tree():
    root = _stale_tree()
    ET.SubElement(root, _metadata.q("feature", EXTENSIONS[0]))
    snapshot = _snapshot(root)
    first = _metadata.dump(root)
    assert _metadata.dump(root) == first
    ET.register_namespace("rebound", EXTENSIONS[0])
    second = _metadata.dump(root)
    parsed, _ = _assert_extensions(second, (EXTENSIONS[0],))
    assert parsed.get(IGNORABLE) == "rebound"
    assert _metadata.dump(root) == second
    _assert_unchanged(snapshot)


def test_namespace_like_comments_text_and_attributes_are_not_declarations():
    fake = f'xmlns:fake="{EXTENSIONS[0]}"'
    root = ET.Element(_metadata.q("worksheet"), {"custom": fake})
    root.text = fake
    root.append(ET.Comment(fake))
    root.append(ET.ProcessingInstruction("producer", fake))
    parsed, _ = _assert_extensions(_metadata.dump(root), ())
    assert parsed.get("custom") == fake
    assert parsed.text == fake


def test_empty_extension_root_remains_well_formed():
    root = ET.Element(_metadata.q("feature", EXTENSIONS[0]))
    parsed, _ = _assert_extensions(_metadata.dump(root), (EXTENSIONS[0],))
    assert parsed.tag == root.tag
    assert len(parsed) == 0


def test_repeated_overlay_preserves_metadata_snapshots_and_native_cell_payloads():
    book = NativeWorkbook()
    book.active["A1"] = ' text & < > " '
    book.active["B1"] = "=1+2"
    book.active["C1"] = 42
    source = BytesIO()
    book.save(source)
    template = _stale_tree()
    # Loaded XML has string keys/values; QName/None handling is tested directly.
    template[0].attrib = {"first": "x", IGNORABLE: "old14", "last": "y"}
    ET.SubElement(template[0], _metadata.q("feature", EXTENSIONS[0]), {"custom": "kept"})
    sheet = MetadataSheet()
    sheet._xml = template
    metadata = MetadataBook(sheets={"Sheet": sheet})
    snapshot = _snapshot(template)
    with ZipFile(source) as archive:
        native = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    payload = ET.tostring(native.find(_metadata.q("sheetData")))
    results = []
    for _ in range(2):
        output = BytesIO()
        _metadata.apply_metadata(source, output, metadata, native_styles=True)
        with ZipFile(output) as archive:
            data = archive.read("xl/worksheets/sheet1.xml")
        parsed, _ = _assert_extensions(data, (EXTENSIONS[0],))
        assert ET.tostring(parsed.find(_metadata.q("sheetData"))) == payload
        assert parsed.find(".//" + _metadata.q("feature", EXTENSIONS[0])).get("custom") == "kept"
        _assert_unchanged(snapshot)
        results.append(data)
    assert results[0] == results[1]
