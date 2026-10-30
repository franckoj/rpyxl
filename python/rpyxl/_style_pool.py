"""Workbook-local immutable style interning and cell-bound copy-on-write views.

Assign through ``pool.assign(value)`` before storing a cell component. Getters
can return ``pool.view(get_current, set_current, check_writable)``; callbacks
read/write that component on its owning cell, and check_writable takes no args.
Opaque assignments deliberately remain ordinary, independently copied models.

Generic model snapshots must exclude ``styles.STYLE_INTERNAL_FIELDS`` and
special-case FrozenStyle using its already cached ``fingerprint``. Tracking
state is private, and no pool or cell is retained by a source model's cache.
"""

from copy import deepcopy
from types import MappingProxyType
from weakref import WeakValueDictionary

from .styles import _STYLE_PRIMITIVES, _STYLE_TYPES, _style_fields


class FrozenStyle:
    """An immutable style value; nested styles are immutable values as well.

    ``fields(instance)`` and model properties/aliases work without thawing.
    ``isinstance(instance, Font)`` (and the other models) is supported through
    __class__; use ``type(instance)`` to distinguish the immutable wrapper.
    """

    __slots__ = ("__weakref__", "_fingerprint", "_hash", "_model_type", "_values")

    def __init__(self, model_type, values):
        try:
            object.__getattribute__(self, "_values")
        except AttributeError:
            pass
        else:
            raise AttributeError("FrozenStyle is immutable")
        if model_type not in _STYLE_TYPES:
            raise TypeError("Unsupported style model")
        if type(values) is not tuple or len(values) != len(_style_fields(model_type)):
            raise TypeError("Style values must be a tuple matching the model fields")
        if any(
            type(value) not in _STYLE_PRIMITIVES and type(value) is not FrozenStyle
            for value in values
        ):
            raise TypeError(
                "Frozen styles may contain only primitives or frozen styles"
            )
        object.__setattr__(self, "_model_type", model_type)
        object.__setattr__(self, "_values", values)
        # Python equates True, 1, and 1.0; interning must not substitute their
        # field types. Nested frozen values already have type-sensitive keys.
        fingerprint = (model_type, tuple((type(value), value) for value in values))
        object.__setattr__(self, "_fingerprint", fingerprint)
        object.__setattr__(self, "_hash", hash(fingerprint))

    @property
    def fingerprint(self):
        """A cached, hashable key preserving model and primitive field types."""
        return self._fingerprint

    @property
    def __class__(self):
        return self._model_type

    @property
    def __dataclass_fields__(self):
        return MappingProxyType(self._model_type.__dataclass_fields__)

    @property
    def __dict__(self):
        return MappingProxyType(
            dict(zip(_style_fields(self._model_type), self._values))
        )

    def __getattr__(self, name):
        model_type = self._model_type
        index = _style_fields(model_type).get(name)
        if index is not None:
            return self._values[index]
        descriptor = getattr(model_type, name, None)
        if isinstance(descriptor, property):
            return descriptor.__get__(self, model_type)
        raise AttributeError(name)

    def __setattr__(self, name, value):
        raise AttributeError(
            "FrozenStyle is immutable; edit a cell view or call to_model()"
        )

    def __delattr__(self, name):
        raise AttributeError("FrozenStyle is immutable")

    def __hash__(self):
        return self._hash

    def __eq__(self, other):
        if type(other) is _StyleView:
            other = other._current()
        if type(other) is FrozenStyle:
            return self._fingerprint == other._fingerprint
        if type(other) is self._model_type:
            return all(
                value == getattr(other, name)
                for name, value in zip(_style_fields(self._model_type), self._values)
            )
        return NotImplemented

    def __repr__(self):
        values = ", ".join(
            f"{name}={value!r}"
            for name, value in zip(_style_fields(self._model_type), self._values)
        )
        return f"FrozenStyle({self._model_type.__name__}, {values})"

    def __copy__(self):
        return self

    def __deepcopy__(self, memo):
        return self

    def to_model(self):
        """Return a fully independent, normal mutable model, without normalization."""
        model = self._model_type.__new__(self._model_type)
        for name, value in zip(_style_fields(self._model_type), self._values):
            setattr(
                model, name, value.to_model() if type(value) is FrozenStyle else value
            )
        return model


class StylePool:
    """Intern styles by value within one workbook, without retaining dead styles.

    Each mutable source retains at most its most recent immutable snapshot.
    Sources never retain this pool, a workbook, or a cell. Weak interning also
    lets obsolete values disappear after sources and cells release them.
    """

    def __init__(self):
        self._styles = WeakValueDictionary()

    def _intern(self, frozen):
        existing = self._styles.get(frozen.fingerprint)
        if existing is not None:
            return existing
        self._styles[frozen.fingerprint] = frozen
        return frozen

    def _freeze(self, value, token):
        cached = value._style_cache
        if cached is not None and cached[0] == token:
            return self._intern(cached[1])
        values = []
        children = dict(token[1])
        for name in _style_fields(type(value)):
            item = getattr(value, name)
            if type(item) in _STYLE_TYPES:
                item = self._freeze(item, children[name])
            values.append(item)
        frozen = self._intern(FrozenStyle(type(value), tuple(values)))
        object.__setattr__(value, "_style_cache", (token, frozen))
        return frozen

    def assign(self, value):
        """Snapshot a style input, or deepcopy opaque input; None stays None.

        Exact built-in style models with primitive/style fields are supported.
        Subclasses, unknown attributes (including _xml), mutable field values,
        and cycles are conservatively uncached. A view is resolved immediately,
        so assignment never stores a live link to another cell.
        """
        if type(value) is _StyleView:
            value = value._current()
        if value is None:
            return None
        if type(value) is FrozenStyle:
            return self._intern(value)
        if type(value) in _STYLE_TYPES:
            token = value._style_token()
            if token is not None:
                return self._freeze(value, token)
        return deepcopy(value)

    def view(self, get_current, set_current, check_writable):
        """Return a live cell view for frozen roots, otherwise the stored value.

        ``get_current()`` returns the stored component; ``set_current(value)``
        stores an already assigned snapshot. ``check_writable()`` must raise
        when writing is disallowed. Opaque roots preserve legacy writable-model
        behavior instead of pretending their extra state is safely immutable.
        """
        current = get_current()
        if type(current) is not FrozenStyle:
            return current
        return _StyleView(self, get_current, set_current, check_writable)


class _StyleView:
    __slots__ = ("_check_writable", "_get_current", "_path", "_pool", "_set_current")
    __hash__ = None

    def __init__(self, pool, get_current, set_current, check_writable, path=()):
        object.__setattr__(self, "_pool", pool)
        object.__setattr__(self, "_get_current", get_current)
        object.__setattr__(self, "_set_current", set_current)
        object.__setattr__(self, "_check_writable", check_writable)
        object.__setattr__(self, "_path", path)

    def _current(self):
        value = self._get_current()
        for name in self._path:
            value = getattr(value, name)
        return value

    @property
    def __class__(self):
        return self._current().__class__

    @property
    def __dataclass_fields__(self):
        return self._current().__dataclass_fields__

    @property
    def __dict__(self):
        return MappingProxyType(dict(vars(self._current())))

    def __getattr__(self, name):
        value = getattr(self._current(), name)
        if type(value) is FrozenStyle:
            return _StyleView(
                self._pool,
                self._get_current,
                self._set_current,
                self._check_writable,
                self._path + (name,),
            )
        return value

    def _edit(self, name, value=None, *, delete=False):
        if name.startswith("_"):
            raise AttributeError("Style view internals are read-only")
        self._check_writable()
        root = self._get_current()
        model = root.to_model() if type(root) is FrozenStyle else deepcopy(root)
        target = model
        for part in self._path:
            target = getattr(target, part)
        if delete:
            delattr(target, name)
        else:
            if type(value) in (FrozenStyle, _StyleView):
                value = (
                    value.to_model() if type(value) is FrozenStyle else deepcopy(value)
                )
            setattr(target, name, value)
        self._set_current(self._pool.assign(model))

    def __setattr__(self, name, value):
        self._edit(name, value)

    def __delattr__(self, name):
        self._edit(name, delete=True)

    def __copy__(self):
        current = self._current()
        return current.to_model() if type(current) is FrozenStyle else deepcopy(current)

    def __deepcopy__(self, memo):
        result = self.__copy__()
        memo[id(self)] = result
        return result

    def __eq__(self, other):
        if type(other) is _StyleView:
            other = other._current()
        return self._current() == other

    def __repr__(self):
        return f"StyleView({self._current()!r})"


def unwrap_style(value):
    """Resolve a cell view without copying its immutable value."""
    return value._current() if type(value) is _StyleView else value


__all__ = ["FrozenStyle", "StylePool", "unwrap_style"]
