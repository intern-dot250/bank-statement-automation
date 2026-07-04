"""Unlock a password-protected PDF and save a decrypted copy.

Usage:
    python unlock_pdf.py
    python unlock_pdf.py --password "MySecret123"
    python unlock_pdf.py --input "other.pdf" --output "out/unlocked.pdf"
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.errors import (
    EmptyFileError,
    FileNotDecryptedError,
    PdfReadError,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DEFAULT_INPUT = Path("PDF_password_protected.pdf")
DEFAULT_OUTPUT = Path("output/unlocked_statement.pdf")
LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"

# -----------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
log = logging.getLogger("unlock_pdf")


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------
def decrypt_pdf(input_path: Path, output_path: Path, password: str) -> None:
    """Decrypt ``input_path`` with ``password`` and write it to ``output_path``.

    Raises:
        FileNotFoundError:  input file is missing.
        ValueError:         password is wrong / file uses a different cipher.
        PdfReadError:       file is not a valid PDF.
    """
    log.info("Reading encrypted PDF: %s", input_path)

    # 1) Open the file. ``strict=False`` is friendlier to slightly malformed PDFs.
    try:
        reader = PdfReader(str(input_path), strict=False)
    except EmptyFileError as exc:
        raise PdfReadError("File is empty.") from exc

    # 2) Validate that the file is actually encrypted.
    if not reader.is_encrypted:
        log.warning("Input PDF is NOT encrypted — copying it through unchanged.")

    # 3) Attempt to decrypt. ``decrypt`` accepts str or bytes and returns an
    #    int (0 = failure, 1 = user-password matched, 2 = owner-password only).
    log.info("Attempting decryption…")
    result = reader.decrypt(password if password else "")
    if result == 0:
        raise ValueError("Incorrect password — could not decrypt the PDF.")

    log.info("Decryption successful (code=%d).", result)

    # 4) Reassemble pages into a new, unencrypted writer.
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # 5) Ensure the destination directory exists.
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # If the output file already exists, delete it to avoid PermissionError on Windows.
    if output_path.is_file():
        try:
            output_path.unlink()
            log.info("Deleted existing unlocked PDF: %s", output_path)
        except OSError as exc:
            log.warning("Could not delete existing unlocked PDF (%s): %s", output_path, exc)

    log.info("Writing unlocked PDF to: %s", output_path)
    with output_path.open("wb") as fh:
        writer.write(fh)

    log.info("Done. %d page(s) written.", len(reader.pages))


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Decrypt a password-protected PDF.",
    )
    parser.add_argument(
        "-i", "--input",
        type=Path, default=DEFAULT_INPUT,
        help=f"Path to the encrypted PDF (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "-o", "--output",
        type=Path, default=DEFAULT_OUTPUT,
        help=f"Path for the unlocked PDF (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "-p", "--password",
        default="",
        help="Password for the PDF (will prompt if omitted).",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    # Securely prompt if no password was supplied on the CLI.
    if not args.password:
        import getpass
        try:
            args.password = getpass.getpass("PDF password: ")
        except (EOFError, KeyboardInterrupt):
            log.error("No password provided.")
            return 2

    try:
        decrypt_pdf(args.input, args.output, args.password)
    except FileNotFoundError:
        log.exception("Input file not found: %s", args.input)
        return 1
    except ValueError:
        # Wrong password path — keep message concise (no traceback noise).
        log.error("Wrong password. Aborting.")
        return 1
    except (PdfReadError, FileNotDecryptedError) as exc:
        log.error("PDF is corrupted or unreadable: %s", exc)
        return 1
    except OSError as exc:
        log.error("Filesystem error: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())