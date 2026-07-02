from __future__ import annotations

import base64
import hashlib
import hmac
import secrets

PBKDF2_ALGORITHM = "sha256"
PBKDF2_ITERATIONS = 260_000
SALT_BYTES = 16


def hash_password(password: str) -> str:
    if not password:
        raise ValueError("Senha nao pode ser vazia.")

    salt = secrets.token_bytes(SALT_BYTES)
    derived = hashlib.pbkdf2_hmac(
        PBKDF2_ALGORITHM,
        password.encode("utf-8"),
        salt,
        PBKDF2_ITERATIONS,
    )
    salt_b64 = base64.urlsafe_b64encode(salt).decode("ascii").rstrip("=")
    hash_b64 = base64.urlsafe_b64encode(derived).decode("ascii").rstrip("=")
    return f"pbkdf2_{PBKDF2_ALGORITHM}${PBKDF2_ITERATIONS}${salt_b64}${hash_b64}"


def verify_password(password: str, stored_hash: str) -> bool:
    if not password or not stored_hash:
        return False

    try:
        scheme, iterations_text, salt_b64, hash_b64 = stored_hash.split("$", 3)
        if scheme != f"pbkdf2_{PBKDF2_ALGORITHM}":
            return False

        iterations = int(iterations_text)
        salt = _decode_base64(salt_b64)
        expected = _decode_base64(hash_b64)
        derived = hashlib.pbkdf2_hmac(
            PBKDF2_ALGORITHM,
            password.encode("utf-8"),
            salt,
            iterations,
        )
        return hmac.compare_digest(derived, expected)
    except (TypeError, ValueError):
        return False


def _decode_base64(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding)
