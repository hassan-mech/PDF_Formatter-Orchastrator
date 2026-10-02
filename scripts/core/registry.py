"""
scripts/core/registry.py — Generic Plugin/Stage Registry

Provides type-safe, generic Registry[T] for registration and retrieval of pipeline stages:
- Inspector
- Extractor
- OcrEngine
- Builder
- Verifier

NO concrete third-party libraries may be imported here.
"""

from __future__ import annotations

from typing import Callable, Dict, Generic, Iterator, List, Optional, Type, TypeVar

T = TypeVar("T")


class Registry(Generic[T]):
    """Generic registry mapping string identifiers to component classes/factories."""

    def __init__(self, name: str):
        self._name = name
        self._entries: Dict[str, Type[T]] = {}

    @property
    def name(self) -> str:
        return self._name

    def register(self, name: str) -> Callable[[Type[T]], Type[T]]:
        """Decorator to register a component implementation under a given name."""
        def decorator(cls: Type[T]) -> Type[T]:
            if name in self._entries:
                raise ValueError(
                    f"Component '{name}' already registered in registry '{self._name}'"
                )
            self._entries[name] = cls
            return cls
        return decorator

    def get(self, name: str) -> Type[T]:
        """Retrieve a component implementation by name."""
        if name not in self._entries:
            available = ", ".join(self._entries.keys()) or "none"
            raise KeyError(
                f"No component '{name}' found in registry '{self._name}'. Available: [{available}]"
            )
        return self._entries[name]

    def has(self, name: str) -> bool:
        """Check if an entry is registered."""
        return name in self._entries

    def list(self) -> List[str]:
        """List all registered component names."""
        return sorted(list(self._entries.keys()))

    def __contains__(self, name: str) -> bool:
        return name in self._entries

    def __len__(self) -> int:
        return len(self._entries)

    def __iter__(self) -> Iterator[str]:
        return iter(self._entries)

    def __repr__(self) -> str:
        return f"<Registry '{self._name}' entries={list(self._entries.keys())}>"


# Global stage registries
inspector_registry: Registry[Any] = Registry("inspectors")
extractor_registry: Registry[Any] = Registry("extractors")
ocr_registry: Registry[Any] = Registry("ocr_engines")
builder_registry: Registry[Any] = Registry("builders")
verifier_registry: Registry[Any] = Registry("verifiers")
