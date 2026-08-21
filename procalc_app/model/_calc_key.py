"""App-embedded AES-256 key for .calc project files.

This key stops a .calc file from being opened in Notepad/Excel/a generic
text or hex editor, or read as JSON by another program — not a real
secret. Every copy of Procalc ships the same key, so no password is ever
asked for or needed to open a project. Anyone who extracts this constant
from a copy of the app (trivial, even from a compiled/frozen build) can
decrypt any .calc file. This is a deliberate, accepted trade-off (see
model/project.py) — do not treat this as protecting sensitive data.
"""

CALC_KEY = bytes((
    0xf2, 0xbc, 0x93, 0x34, 0x15, 0x34, 0x66, 0x0d,
    0x03, 0xdd, 0x8e, 0xdd, 0xf4, 0xf4, 0x7c, 0x90,
    0xfb, 0x13, 0x50, 0xf7, 0xcb, 0xf1, 0x80, 0x01,
    0xdc, 0x24, 0x8b, 0x93, 0x46, 0xb7, 0xe2, 0x67,
))
