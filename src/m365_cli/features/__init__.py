"""Pluggable features. Every sub-package exposing a module-level `FEATURE` is discovered automatically.

Add a feature: create `features/<name>/` with `FEATURE = Feature(...)` in its `__init__.py`.
Remove/disable one: delete the package, or list it in `disabled_features` in config.toml.
"""
