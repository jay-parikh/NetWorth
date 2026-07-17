"""Privacy helpers (SPEC §3.19) — the Mask's password hash and the Lock's
whole-file encryption.

Two honest layers, one password:

* Mask — display-level ••• formats + sheet protection. The password is
  never stored; a PBKDF2 fingerprint persisted inside the workbook lets the
  updater verify it. A curtain against shoulder-surfing, not security.
* Lock — standard OOXML Agile encryption (Excel's native "password to
  open", AES). The at-rest file is ciphertext; decryption success IS the
  password check, so nothing extra is stored. No recovery by design.

Plaintext never touches disk on the Lock path: the workbook is built into
memory, encrypted in memory, self-verified, and only ciphertext is written.
"""

from __future__ import annotations

import hashlib
import hmac
import io
import os

PBKDF2_ITERATIONS = 200_000
_SCHEME = "pbkdf2-sha256"

# Legacy OLE/CFB container magic — what an encrypted OOXML file starts with
# (a plain xlsx is a PK zip).
_CFB_MAGIC = b"\xd0\xcf\x11\xe0"


def hash_password(password: str) -> str:
    """`pbkdf2-sha256$<iter>$<salt-hex>$<hash-hex>` — safe to store."""
    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt,
                                 PBKDF2_ITERATIONS)
    return f"{_SCHEME}${PBKDF2_ITERATIONS}${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        scheme, iters, salt_hex, hash_hex = stored.split("$")
        if scheme != _SCHEME:
            return False
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                                     bytes.fromhex(salt_hex), int(iters))
        return hmac.compare_digest(digest.hex(), hash_hex)
    except (ValueError, AttributeError):
        return False


def sheet_password(stored_hash: str) -> str:
    """Deterministic sheet-protection password for masked builds, derived
    from the stored fingerprint (the real password may be unavailable — the
    user pressed Enter to keep the mask). Excel's sheet protection is a
    weak legacy hash anyway; the REAL check is verify_password above."""
    return stored_hash[-12:] if stored_hash else "networth"


def is_encrypted(source) -> bool:
    """True when the file/bytes are an encrypted OOXML container."""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source[:4]) == _CFB_MAGIC
    with open(source, "rb") as f:
        return f.read(4) == _CFB_MAGIC


class WrongPassword(Exception):
    pass


def encrypt_workbook(plain: bytes, password: str) -> bytes:
    """Encrypt xlsx bytes (Agile encryption) and SELF-VERIFY the result by
    decrypting it back — msoffcrypto's encryption is marked experimental
    upstream, so no ciphertext leaves here unproven."""
    from msoffcrypto.format.ooxml import OOXMLFile

    out = io.BytesIO()
    OOXMLFile(io.BytesIO(plain)).encrypt(password, out)
    cipher = out.getvalue()
    if not is_encrypted(cipher):
        raise RuntimeError("encryption produced a non-encrypted container")
    if decrypt_workbook(cipher, password).getvalue() != plain:
        raise RuntimeError("encryption self-check failed to round-trip")
    return cipher


def decrypt_workbook(source, password: str) -> io.BytesIO:
    """Decrypt an encrypted workbook (path or bytes) into a BytesIO.
    Raises WrongPassword on a bad password."""
    import msoffcrypto
    from msoffcrypto.exceptions import DecryptionError, InvalidKeyError

    if isinstance(source, (bytes, bytearray)):
        stream = io.BytesIO(bytes(source))
    else:
        stream = io.BytesIO(open(source, "rb").read())
    try:
        f = msoffcrypto.OfficeFile(stream)
        f.load_key(password=password)
        out = io.BytesIO()
        f.decrypt(out)
    except (InvalidKeyError, DecryptionError) as e:
        raise WrongPassword(str(e)) from e
    out.seek(0)
    return out
