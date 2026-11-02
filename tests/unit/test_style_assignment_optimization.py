"""Repeated assignment must save work without trusting mutable source identity."""

import gc
import weakref
from copy import copy, deepcopy
from xml.etree import ElementTree as ET

import pytest
from rpyxl._style_pool import FrozenStyle, StylePool
from rpyxl.styles import (
    STYLE_INTERNAL_FIELDS,
    Alignment,
    Border,
    Color,
    Font,
    PatternFill,
    Side,
)


def _styles():
    side = Side("thin", Color(theme=1))
    return (
        Font(b=True, color=Color(theme=1)),
        PatternFill("solid", fgColor=Color(theme=2), bgColor=Color(theme=3)),
        Border(left=side, right=side, top=side, bottom=side),
        Alignment(horizontal="center", wrapText=True),
    )


@pytest.mark.parametrize("source", _styles())
def test_unchanged_assignment_reuses_tokens_but_still_checks_children(
    source, monkeypatch
):
    pool = StylePool()
    frozen = pool.assign(source)
    token = source._style_cache[0]
    visits = []
    original = Color._style_token

    def record(self, ancestors=None):
        visits.append(self)
        return original(self, ancestors)

    monkeypatch.setattr(Color, "_style_token", record)
    for _ in range(20):
        assert pool.assign(source) is frozen
        assert source._style_cache[0] is token
    assert bool(visits) is (type(source) is not Alignment)


def test_repeated_leaf_assignment_needs_no_ancestor_set(monkeypatch):
    pool = StylePool()
    source = Alignment(wrapText=True)
    frozen = pool.assign(source)

    def unexpected_set():
        raise AssertionError("A tracked leaf cannot form a cycle")

    monkeypatch.setattr("rpyxl.styles.set", unexpected_set, raising=False)
    assert pool.assign(source) is frozen


def test_cached_assignment_does_not_enter_freeze_again(monkeypatch):
    pool = StylePool()
    source = Border(left=Side("thin", Color(theme=1)))
    frozen = pool.assign(source)

    def unexpected_freeze(*args):
        raise AssertionError("Unchanged assignments must use the root snapshot cache")

    monkeypatch.setattr(StylePool, "_freeze", unexpected_freeze)
    for _ in range(20):
        assert pool.assign(source) is frozen
        assert pool.assign(frozen) is frozen
    assert StylePool().assign(source) is frozen

    assert type(frozen.fingerprint) is tuple


@pytest.mark.parametrize(
    "make,mutate,read,expected",
    [
        (
            lambda: Font(color=Color(theme=1)),
            lambda x: setattr(x, "b", True),
            lambda x: x.bold,
            True,
        ),
        (
            lambda: Font(color=Color(theme=1)),
            lambda x: setattr(x.color, "tint", 0.5),
            lambda x: x.color.tint,
            0.5,
        ),
        (
            lambda: PatternFill("solid", fgColor=Color(theme=1)),
            lambda x: setattr(x, "start_color", Color(theme=4)),
            lambda x: x.fgColor.theme,
            4,
        ),
        (
            lambda: Border(left=Side("thin", Color(theme=1))),
            lambda x: setattr(x.left, "border_style", "dashed"),
            lambda x: x.left.style,
            "dashed",
        ),
        (
            lambda: Border(left=Side("thin", Color(theme=1))),
            lambda x: setattr(x.left.color, "theme", 4),
            lambda x: x.left.color.theme,
            4,
        ),
        (
            lambda: Border(left=Side("thin", Color(theme=1))),
            lambda x: setattr(x, "left", Side("double")),
            lambda x: x.left.style,
            "double",
        ),
        (
            Alignment,
            lambda x: setattr(x, "wrapText", True),
            lambda x: x.wrap_text,
            True,
        ),
    ],
)
def test_mutation_invalidates_reused_token(make, mutate, read, expected):
    pool = StylePool()
    source = make()
    old = pool.assign(source)
    old_value = read(old)
    old_token = source._style_cache[0]
    mutate(source)
    new = pool.assign(source)
    assert source._style_cache[0] is not old_token
    assert new is not old
    assert read(new) == expected
    assert read(old) == old_value
    assert pool.assign(source) is new


def test_shared_children_and_replacements_are_checked_in_every_parent_and_pool():
    first_pool, second_pool = StylePool(), StylePool()
    side = Side("thin", Color(theme=1))
    first = Border(left=side, right=side)
    second = Border(bottom=side)
    old_first = first_pool.assign(first)
    old_second = second_pool.assign(second)
    first.left = Side("double", Color(theme=3))
    replaced = first_pool.assign(first)
    side.color.tint = 0.5
    changed_first = first_pool.assign(first)
    changed_second = second_pool.assign(second)
    assert changed_first.left.style == "double"
    assert changed_first.right.color.tint == changed_second.bottom.color.tint == 0.5
    assert old_first.left is old_first.right
    assert old_first.right.color.tint == old_second.bottom.color.tint == 0
    assert replaced.left is changed_first.left
    assert replaced.right.color.tint == 0


@pytest.mark.parametrize("location", ["root", "side", "color"])
def test_opaque_addition_after_caching_and_recovery(location):
    pool = StylePool()
    source = Border(left=Side("thin", Color(theme=1)))
    old = pool.assign(source)
    target = {"root": source, "side": source.left, "color": source.left.color}[location]
    target._xml = ET.fromstring('<extension value="first"/>')
    first = pool.assign(source)
    target._xml.set("value", "second")
    second = pool.assign(source)
    first_target = {"root": first, "side": first.left, "color": first.left.color}[
        location
    ]
    second_target = {"root": second, "side": second.left, "color": second.left.color}[
        location
    ]
    assert type(first) is type(second) is Border
    assert first_target._xml.get("value") == "first"
    assert second_target._xml.get("value") == "second"
    del target._xml
    assert pool.assign(source) is old


def test_vars_edits_before_tracking_and_frozen_dict_protection():
    # Direct dictionary writes after tracking starts remain unsupported. Before
    # activation the complete dictionary, including extras, must be validated.
    source = Font(color=Color(theme=1))
    vars(source)["bold"] = True
    vars(source.color)["tint"] = 0.5
    vars(source)["extra"] = [1]
    pool = StylePool()
    first = pool.assign(source)
    vars(source)["extra"].append(2)
    second = pool.assign(source)
    assert first.bold and first.color.tint == 0.5
    assert first.extra == [1] and second.extra == [1, 2]
    del vars(source)["extra"]
    frozen = pool.assign(source)
    assert type(frozen) is FrozenStyle
    with pytest.raises(TypeError):
        vars(frozen)["bold"] = False
    with pytest.raises(TypeError):
        vars(frozen.color)["tint"] = 1


def test_subclass_and_cycle_still_use_independent_deepcopies():
    class CustomFont(Font):
        pass

    pool = StylePool()
    source = CustomFont(color=Color(theme=1))
    first = pool.assign(source)
    source.color.tint = 0.5
    second = pool.assign(source)
    assert type(first) is type(second) is CustomFont
    assert first.color.tint == 0 and second.color.tint == 0.5
    source = Font()
    pool.assign(source)
    source.color = source
    first, second = pool.assign(source), pool.assign(source)
    assert first is not second and first.color is first and second.color is second


def test_copy_on_write_and_plain_copies_remain_independent():
    pool = StylePool()
    source = Border(left=Side("thin", Color(theme=1)))
    frozen = pool.assign(source)
    cells = [frozen, frozen]
    view = pool.view(lambda: cells[0], lambda x: cells.__setitem__(0, x), lambda: None)
    view.left.color.tint = 0.5
    assert cells[0].left.color.tint == 0.5
    assert cells[1] is frozen and frozen.left.color.tint == 0
    assert source.left.color.tint == 0
    for model in (copy(source), frozen.to_model(), copy(view)):
        assert not (vars(model).keys() & STYLE_INTERNAL_FIELDS)
    copied = deepcopy(source)
    copied.left.color.tint = 0.75
    assert pool.assign(copied).left.color.tint == 0.75
    assert pool.assign(source) is frozen


def test_hash_collisions_preserve_value_interning(monkeypatch):
    monkeypatch.setattr(FrozenStyle, "__hash__", lambda self: 0)
    pool = StylePool()
    first = pool.assign(Border(left=Side("thin", Color(theme=1))))
    second = pool.assign(Border(left=Side("thin", Color(theme=2))))
    assert hash(first.fingerprint) == hash(second.fingerprint)
    assert first is not second
    assert pool.assign(Border(left=Side("thin", Color(theme=1)))) is first
    assert pool.assign(Border(left=Side("thin", Color(theme=2)))) is second
    assert pool.assign(Border(left=Side("thin", Color(theme=1.0)))) is not first


def test_caching_does_not_keep_sources_pools_or_obsolete_snapshots_alive():
    pool = StylePool()
    source = Border(left=Side("thin", Color(theme=1)))
    old = pool.assign(source)
    source_ref, old_ref, pool_ref = (
        weakref.ref(source),
        weakref.ref(old),
        weakref.ref(pool),
    )
    source.left.color.tint = 0.5
    current = pool.assign(source)
    del old
    gc.collect()
    assert old_ref() is None
    del pool
    gc.collect()
    assert pool_ref() is None
    current_ref = weakref.ref(current)
    del source, current
    gc.collect()
    assert source_ref() is None and current_ref() is None
