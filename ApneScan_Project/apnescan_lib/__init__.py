"""ApneScan internal library — stateless, UI-free engine code.

This package holds pure/stateless helpers that were extracted out of the single
``apnescan.py`` file so they can be maintained, unit-tested and reasoned about in
isolation. Nothing here imports PyQt or touches application state, so every
module is importable and testable on its own.

``apnescan.py`` remains the application entry point and imports these back into
its namespace, so runtime behaviour is unchanged.
"""
