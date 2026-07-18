"""PDF → text for statement imports (SPEC §6.17). pypdf, pure Python.

The only file in importers/ that touches the filesystem. Layout mode
keeps column spacing so the CAS line parser sees rows the way the PDF
prints them. AES-encrypted statements decrypt via the `cryptography`
package, which the app already ships for the privacy Lock.
"""

from __future__ import annotations

from pathlib import Path


class NeedsPassword(Exception):
    """The PDF is encrypted and no password was given."""


class WrongPassword(Exception):
    """The PDF is encrypted and this password doesn't open it."""


def extract_text(path: str | Path, password: str | None = None) -> str:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    if reader.is_encrypted:
        if not password:
            raise NeedsPassword(Path(path).name)
        if not reader.decrypt(password):
            raise WrongPassword(Path(path).name)
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text(extraction_mode="layout"))
        except Exception:                      # noqa: BLE001 — a page that
            pages.append(page.extract_text())  # breaks layout mode still
    return "\n".join(pages)                    # yields its plain text


def looks_like_pdf(path: str | Path) -> bool:
    try:
        with open(path, "rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False
