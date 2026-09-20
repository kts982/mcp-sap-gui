"""
Transaction policy - which transaction codes an agent may start.

What is dangerous depends on who runs the server and against which system, so
the list is not hard-coded: a preset picks a starting point, and a policy file
adds block patterns and allow exceptions on top.

This is a guardrail against agent mistakes, NOT a security boundary. A blocked
transaction can often be reached another way (a menu path, a screen the user
left open), and the effective boundary is the SAP user's authorizations. Pair
the policy with a minimal-authorization account.

Rules:

- ``block`` and ``allow`` hold transaction codes or glob patterns
  (``SU*``, ``/SCWM/*``, ``*``).
- ``allow`` always wins over ``block``: a transaction is permitted when it
  matches an allow pattern, or matches no block pattern.
- The ``strict`` preset blocks ``*``, so only allow-listed codes can start.

The policy is read once at startup. It cannot be changed from inside a
session: there is deliberately no tool for it.

This module must not import the server module (see ConfirmationMiddleware's
docstring for the double-module trap under ``python -m``).
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from fnmatch import fnmatchcase
from typing import Any, Callable, Dict, List, Optional

POLICY_FILE_ENV = "MCP_SAP_GUI_POLICY_FILE"

# Shared by every preset except "strict".
_BASE_BLOCK = [
    "SU01", "SU10", "SU01D",  # User administration
    "PFCG", "SU53",           # Role administration
    "SM21", "ST22",           # System logs / dumps
    "SE16N",                  # Direct table maintenance
    "SE38", "SA38", "SE80",   # ABAP editor / program execution
    "SE37",                   # Function module test (runs any function)
    "STMS",                   # Transport management
    "SCC4",                   # Client administration
    "RZ10", "RZ11",           # Profile parameters
    "SM36",                   # Background jobs
    "SM49", "SM69",           # OS command execution
    "SM59",                   # RFC destination config
    "STRUST", "SICF",         # Trust manager / ICF services
    "SM01",                   # Lock transactions
    "SM18", "SM19",           # Security audit log: delete / configure
]

PRESETS: Dict[str, Dict[str, List[str]]] = {
    # Consultants and key users: customizing and reports, no basis/dev tools.
    "default": {"block": list(_BASE_BLOCK), "allow": []},
    # ABAP developers testing their own programs: run reports (SA38; SAP
    # recommends it over SE38), browse the dictionary and the repository.
    "abap-dev": {"block": list(_BASE_BLOCK), "allow": ["SA38", "SE11", "SE80"]},
    # Allowlist only: nothing starts unless the policy file allows it.
    "strict": {"block": ["*"], "allow": []},
}
DEFAULT_PRESET = "default"

_GLOB_CHARS = "*?"
_PATTERN_RE = re.compile(r"^[A-Z0-9_/*?]+$")


class PolicyError(ValueError):
    """The policy preset or file is invalid. Startup must fail closed."""


def normalize_pattern(entry: Any, normalize_code: Callable[[str], str]) -> str:
    """Normalize one block/allow entry.

    A plain code goes through the server's transaction normalizer (strips
    ``/n`` and friends, validates the shape); a glob pattern is uppercased and
    checked against the characters a transaction code can contain.
    """
    if not isinstance(entry, str) or not entry.strip():
        raise PolicyError(f"Policy entries must be non-empty strings, got {entry!r}")
    value = entry.strip().upper()
    if not any(ch in value for ch in _GLOB_CHARS):
        try:
            return normalize_code(value)
        except ValueError as exc:
            raise PolicyError(str(exc)) from exc
    if not _PATTERN_RE.fullmatch(value):
        raise PolicyError(
            f"Invalid transaction pattern {entry!r}: use letters, digits, "
            "'_', '/', and the wildcards '*' and '?'."
        )
    return value


def _dedupe(values: List[str]) -> List[str]:
    return list(dict.fromkeys(values))


@dataclass
class TransactionPolicy:
    """Effective transaction policy: block patterns plus allow exceptions."""

    block: List[str] = field(default_factory=list)
    allow: List[str] = field(default_factory=list)
    preset: str = DEFAULT_PRESET
    source: str = "preset"
    file_sha256: Optional[str] = None

    def is_blocked_pattern(self, tcode: str) -> bool:
        return any(fnmatchcase(tcode, pattern) for pattern in self.block)

    def is_allowed_exception(self, tcode: str) -> bool:
        return any(fnmatchcase(tcode, pattern) for pattern in self.allow)

    def permits(self, tcode: str) -> bool:
        """Return True when *tcode* (already normalized) may be started."""
        return self.is_allowed_exception(tcode) or not self.is_blocked_pattern(tcode)

    def describe(self) -> Dict[str, Any]:
        """The effective policy, for the startup log and the audit log."""
        info: Dict[str, Any] = {
            "preset": self.preset,
            "source": self.source,
            "block": list(self.block),
            "allow": list(self.allow),
        }
        if self.file_sha256:
            info["file_sha256"] = self.file_sha256
        return info

    def how_to_allow(self, tcode: str) -> str:
        """Text for a 'blocked' error: what the USER can do about it."""
        return (
            f" (transaction policy preset '{self.preset}'). Only the user can "
            f"change this: allow {tcode} in a policy file (--policy-file) or "
            "pick another --policy-preset, then restart the server. It cannot "
            "be changed from inside the session - tell the user instead of "
            "looking for another way in."
        )


def default_policy_path() -> str:
    """Per-user policy file, outside any project workspace."""
    base = os.environ.get("APPDATA") or os.path.join(
        os.path.expanduser("~"), ".config"
    )
    return os.path.join(base, "mcp-sap-gui", "policy.json")


def resolve_policy_file(cli_path: Optional[str]) -> Optional[str]:
    """Pick the policy file: --policy-file, then the env var, then the
    per-user default when it exists. An explicitly named file must exist."""
    explicit = cli_path or os.environ.get(POLICY_FILE_ENV)
    if explicit:
        path = os.path.abspath(os.path.expanduser(explicit))
        if not os.path.isfile(path):
            raise PolicyError(f"Policy file not found: {path}")
        return path
    default = default_policy_path()
    return default if os.path.isfile(default) else None


def load_policy(
    normalize_code: Callable[[str], str],
    *,
    preset: Optional[str] = None,
    policy_file: Optional[str] = None,
) -> TransactionPolicy:
    """Build the effective policy from a preset and an optional JSON file.

    File format (every key optional)::

        {
          "preset": "abap-dev",
          "block": ["ZHR*"],
          "allow": ["SM59"]
        }

    ``block`` and ``allow`` are ADDED to the preset's lists. A preset named on
    the command line wins over the one in the file. Anything invalid raises
    PolicyError: a policy that cannot be read must stop the server rather
    than fall back to something more permissive.
    """
    data: Dict[str, Any] = {}
    file_sha256 = None
    if policy_file:
        try:
            with open(policy_file, "rb") as handle:
                raw = handle.read()
            data = json.loads(raw.decode("utf-8-sig"))
        except (OSError, ValueError) as exc:
            raise PolicyError(f"Cannot read policy file {policy_file}: {exc}") from exc
        if not isinstance(data, dict):
            raise PolicyError(f"Policy file {policy_file} must contain a JSON object")
        unknown = sorted(set(data) - {"preset", "block", "allow"})
        if unknown:
            raise PolicyError(
                f"Unknown keys in policy file {policy_file}: {', '.join(unknown)}"
            )
        file_sha256 = hashlib.sha256(raw).hexdigest()

    preset_name = preset or data.get("preset") or DEFAULT_PRESET
    if preset_name not in PRESETS:
        raise PolicyError(
            f"Unknown policy preset {preset_name!r}. Choose from: "
            + ", ".join(sorted(PRESETS))
        )

    lists: Dict[str, List[str]] = {}
    for key in ("block", "allow"):
        extra = data.get(key, [])
        if not isinstance(extra, list):
            raise PolicyError(f"'{key}' in policy file {policy_file} must be a list")
        lists[key] = _dedupe([
            normalize_pattern(entry, normalize_code)
            for entry in [*PRESETS[preset_name][key], *extra]
        ])

    return TransactionPolicy(
        block=lists["block"],
        allow=lists["allow"],
        preset=preset_name,
        source=policy_file or "preset",
        file_sha256=file_sha256,
    )
