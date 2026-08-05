"""Short-lived storage for an uploaded file between the "choose sheet" step
and the "confirm import" step of the admin upload flow."""
from __future__ import annotations

import uuid
from pathlib import Path

UPLOAD_DIR = Path(__file__).resolve().parent.parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


def save(content: bytes, filename: str) -> str:
    token = uuid.uuid4().hex
    (UPLOAD_DIR / f"{token}.bin").write_bytes(content)
    (UPLOAD_DIR / f"{token}.name").write_text(filename, encoding="utf-8")
    return token


def load(token: str) -> tuple[bytes, str] | None:
    data_path = UPLOAD_DIR / f"{token}.bin"
    name_path = UPLOAD_DIR / f"{token}.name"
    if not data_path.exists():
        return None
    filename = name_path.read_text(encoding="utf-8") if name_path.exists() else "upload.xlsx"
    return data_path.read_bytes(), filename


def delete(token: str) -> None:
    for suffix in (".bin", ".name"):
        p = UPLOAD_DIR / f"{token}{suffix}"
        if p.exists():
            p.unlink()
