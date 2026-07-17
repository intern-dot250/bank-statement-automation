"""Resolve writable base directories for both normal (local/server)
environments and read-only serverless deployments (e.g. Vercel), where
only /tmp is writable and the deployed code directory is read-only.

Vercel automatically sets the VERCEL environment variable at runtime,
giving a clean, explicit signal to use instead of probing the filesystem.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path


def is_serverless() -> bool:
    """True if running on Vercel (or any environment that sets VERCEL)."""
    return bool(os.environ.get("VERCEL"))


def base_data_dir(preferred: Path) -> Path:
    """Return a writable base directory for input/output/log/state files.

    Args:
        preferred: The normal, non-serverless base directory (typically
            the project's SCRIPT_DIR).

    Returns:
        `preferred` unchanged in a normal environment. On a serverless
        deployment, a directory under the system temp directory instead
        — note this is ephemeral and may be wiped between invocations;
        it only guarantees the write itself won't crash on import/use.
    """
    if is_serverless():
        fallback = Path(tempfile.gettempdir()) / "bank_statement_automation"
        return fallback
    return preferred
