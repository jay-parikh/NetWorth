"""Publish the plain-English release notes onto the GitHub Releases.

Each release's notes live in ``docs/release-notes/<tag>.md``: the first line
is ``# <release title>``, everything after it is the release body. The
release workflow uses the file automatically when a new tag is pushed; this
script back-fills or edits the notes of releases that already exist.

Auth: a GitHub token with write access to the repo, taken from the
``GITHUB_TOKEN`` environment variable or the file
``~/.config/networth/github-token``.

Usage::

    python scripts/publish_release_notes.py [--repo jay-parikh/NetWorth] [--dry-run]
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.request
from pathlib import Path

API = "https://api.github.com"
NOTES_DIR = Path(__file__).resolve().parent.parent / "docs" / "release-notes"


def _token() -> str:
    tok = os.environ.get("GITHUB_TOKEN", "").strip()
    if not tok:
        f = Path.home() / ".config" / "networth" / "github-token"
        if f.exists():
            tok = f.read_text(encoding="utf-8").strip()
    if not tok:
        sys.exit("No GitHub token found: set GITHUB_TOKEN or put one in "
                 "~/.config/networth/github-token")
    return tok


def _ssl_context():
    try:  # corporate/VM cert stores: use the OS trust store when available
        import ssl

        import truststore
        return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except Exception:
        return None


def _call(url: str, token: str, data: dict | None = None,
          method: str = "GET"):
    req = urllib.request.Request(
        url, method=method,
        data=json.dumps(data).encode() if data is not None else None,
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github+json",
                 "User-Agent": "networth-release-notes"})
    with urllib.request.urlopen(req, context=_ssl_context()) as resp:
        return json.load(resp)


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Publish docs/release-notes/*.md to GitHub Releases")
    ap.add_argument("--repo", default="jay-parikh/NetWorth")
    ap.add_argument("--dry-run", action="store_true",
                    help="show what would change without writing")
    args = ap.parse_args()
    token = _token()

    releases = {r["tag_name"]: r for r in _call(
        f"{API}/repos/{args.repo}/releases?per_page=100", token)}

    for path in sorted(NOTES_DIR.glob("v*.md")):
        tag = path.stem
        lines = path.read_text(encoding="utf-8").splitlines()
        title = (lines[0].removeprefix("# ").strip()
                 if lines and lines[0].startswith("# ") else tag)
        body = "\n".join(lines[1:]).strip() + "\n"
        rel = releases.get(tag)
        if rel is None:
            print(f"  - {tag}: no GitHub release yet — the release workflow "
                  f"will use this file when the tag is pushed")
            continue
        if args.dry_run:
            print(f"  - {tag}: would set title {title!r} "
                  f"and a {len(body)}-char body")
            continue
        _call(rel["url"], token, {"name": title, "body": body}, "PATCH")
        print(f"  ✓ {tag}: published — {title!r}")


if __name__ == "__main__":
    main()
