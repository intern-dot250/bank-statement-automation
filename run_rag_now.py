"""One-shot script to run RAG AI classification on existing ? rows in the sheet.

Usage:
    set GROQ_API_KEY=gsk_...
    py -3 run_rag_now.py
"""

import os
import sys
from pathlib import Path

# Set your Groq API key here if not already in environment
# os.environ["GROQ_API_KEY"] = "gsk_..."

if not os.environ.get("GROQ_API_KEY"):
    print("ERROR: GROQ_API_KEY not set.")
    print("Run:  set GROQ_API_KEY=gsk_your_key_here")
    sys.exit(1)

from rag_classifier import run_rag_classifier

print("=" * 60)
print("RAG AI Classifier — processing existing ? rows")
print("=" * 60)
print()

resolved, unknown = run_rag_classifier()

print()
print("=" * 60)
if resolved > 0:
    print(f"Done. {resolved} rows classified (orange text in sheet).")
else:
    print("Done. No ? rows found — sheet is fully classified.")
print("=" * 60)
