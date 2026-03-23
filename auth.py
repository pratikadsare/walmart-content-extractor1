from __future__ import annotations

from typing import Tuple

APPROVED_DOMAIN = "@pattern.com"
DEFAULT_PASSWORD = "Pratik@123"

# Add or remove approved users here. Email matching is case-insensitive.
ALLOWED_USERS = [
    "pratik.adsare@pattern.com",
    "ajinkya.sable@pattern.com"
    # "kunal.adsare@pattern.com",
]


def normalize_email(email: str) -> str:
    return (email or "").strip().lower()


def is_allowed_domain(email: str) -> bool:
    return normalize_email(email).endswith(APPROVED_DOMAIN)


def is_approved_user(email: str) -> bool:
    normalized = normalize_email(email)
    return normalized in {normalize_email(item) for item in ALLOWED_USERS}


def get_display_name(email: str) -> str:
    normalized = normalize_email(email)
    local = normalized.split("@", 1)[0] if "@" in normalized else normalized
    first = local.split(".", 1)[0] if local else ""
    if not first:
        return "User"
    return first[:1].upper() + first[1:]


def authenticate_user(email: str, password: str) -> Tuple[bool, str]:
    normalized = normalize_email(email)
    if not normalized:
        return False, "Please enter your email address."
    if not is_allowed_domain(normalized):
        return False, "Only approved @pattern.com users can access this tool."
    if not is_approved_user(normalized):
        return False, "This email is not approved yet. Please contact Pratik Adsare."
    if (password or "") != DEFAULT_PASSWORD:
        return False, "Incorrect password."
    return True, ""
