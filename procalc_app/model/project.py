"""Encrypted, portable project file (.calc).

Carries everything needed to reopen and re-run a Procalc project: the
circuit grid rows, unit system + overrides, all document-header metadata
(Project/Client/Site/Circuit Name/Prep-Chk-Appr By/Revision/Page, with the
client logo embedded as base64), and a *resolved* HMB stream snapshot —
built via ``engine_api.resolve_stream_for_snapshot``/``stream_snapshot_to_json``
— so the project can be reopened and re-run even if the original HMB
workbook has since moved or been deleted (only ``hmb_path`` itself is kept
as a label; the engine's resolver never has to read it again once every
referenced stream is in the snapshot — see ``hydraulics_XOM._make_stream_resolver``).

Encrypted with AES-256-GCM under an app-embedded key (``_calc_key.py``) —
this stops a .calc file opening in Notepad/Excel/a generic text or JSON
viewer, not a real secret; no password is ever asked for. See
``_calc_key.py`` for the accepted trade-off.

This module does not import the engine directly — only ``engine_api``,
per the app's single-seam rule.
"""
from __future__ import annotations

import dataclasses
import json
import os

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

from ._calc_key import CALC_KEY

_MAGIC = b"PCALC1"
_NONCE_LEN = 12


@dataclasses.dataclass
class Project:
    grid_rows: list = dataclasses.field(default_factory=list)
    hmb_path: str | None = None
    case: str = "Case 1"
    flash_mode: str = "isothermal"
    unit_system: str = "FPS"
    unit_overrides: dict = dataclasses.field(default_factory=dict)
    meta: dict = dataclasses.field(default_factory=dict)
    # JSON-safe form (engine_api.stream_snapshot_to_json) — a mapping of
    # "hmb||case||stream" -> {"sp": {...}, "feed": {...}|None}.
    stream_snapshot: dict = dataclasses.field(default_factory=dict)
    client_logo_b64: str | None = None

    def save(self, path: str) -> None:
        payload = json.dumps(dataclasses.asdict(self)).encode("utf-8")
        aesgcm = AESGCM(CALC_KEY)
        nonce = os.urandom(_NONCE_LEN)
        ciphertext = aesgcm.encrypt(nonce, payload, None)
        with open(path, "wb") as f:
            f.write(_MAGIC + nonce + ciphertext)

    @classmethod
    def load(cls, path: str) -> "Project":
        with open(path, "rb") as f:
            data = f.read()
        if not data.startswith(_MAGIC):
            raise ValueError(f"'{os.path.basename(path)}' is not a Procalc project file.")
        nonce = data[len(_MAGIC):len(_MAGIC) + _NONCE_LEN]
        ciphertext = data[len(_MAGIC) + _NONCE_LEN:]
        aesgcm = AESGCM(CALC_KEY)
        try:
            payload = aesgcm.decrypt(nonce, ciphertext, None)
        except Exception as exc:
            raise ValueError(
                f"'{os.path.basename(path)}' could not be opened — corrupted, "
                "or not a Procalc project file.") from exc
        d = json.loads(payload.decode("utf-8"))
        field_names = {f.name for f in dataclasses.fields(cls)}
        return cls(**{k: v for k, v in d.items() if k in field_names})
