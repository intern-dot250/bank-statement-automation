"""Bank Statement Processor — Desktop GUI Application.

A Tkinter-based desktop application for processing YES BANK PDF statements.
Employees can select a PDF, enter the password, and upload to Google Sheets
without needing to use the command line.

Usage:
    py app.py
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import (
    END,
    FALSE,
    WORD,
    Button,
    Entry,
    Frame,
    Label,
    LabelFrame,
    PhotoImage,
    Scrollbar,
    Text,
    Tk,
    filedialog,
    messagebox,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
INPUT_PDF_PATH = SCRIPT_DIR / "input" / "current.pdf"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def load_config() -> dict:
    """Load config.json."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_selected_file_types() -> list[str]:
    """Return supported file filter for file dialog."""
    return [
        ("PDF Files", "*.pdf"),
        ("All Files", "*.*"),
    ]


# ---------------------------------------------------------------------------
# App class
# ---------------------------------------------------------------------------
class BankStatementApp:
    """Main application window for the Bank Statement Processor."""

    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Bank Statement Processor — YES BANK")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f4f8")

        self._build_ui()

        # State
        self.selected_file: Path | None = None
        self.is_processing = FALSE

    # ------------------------------------------------------------------
    # UI Construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """Build all UI widgets and layout."""
        # ── Header ──────────────────────────────────────────────────
        header_frame = Frame(self.root, bg="#1a73e8", pady=12)
        header_frame.pack(fill="x")

        title_label = Label(
            header_frame,
            text="Bank Statement Processor",
            font=("Segoe UI", 16, "bold"),
            fg="white",
            bg="#1a73e8",
        )
        title_label.pack()

        subtitle_label = Label(
            header_frame,
            text="Extract transactions from YES BANK PDF statements → Google Sheets",
            font=("Segoe UI", 9),
            fg="#c8d9fa",
            bg="#1a73e8",
        )
        subtitle_label.pack()

        # ── Main content ───────────────────────────────────────────
        content_frame = Frame(self.root, bg="#f0f4f8", padx=24, pady=16)
        content_frame.pack(fill="both", expand=True)

        # ── Input section ───────────────────────────────────────────
        input_group = LabelFrame(
            content_frame,
            text="1. Select PDF File",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f4f8",
            fg="#1a73e8",
            padx=12,
            pady=10,
        )
        input_group.pack(fill="x", pady=(0, 12))

        file_row = Frame(input_group, bg="#f0f4f8")
        file_row.pack(fill="x")

        self.file_entry = Entry(
            file_row,
            font=("Segoe UI", 10),
            bg="white",
            fg="#555",
            relief="solid",
            bd=1,
            insertbackground="#1a73e8",
        )
        self.file_entry.insert(0, "No file selected")
        self.file_entry.config(state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.select_btn = Button(
            file_row,
            text="Browse...",
            font=("Segoe UI", 10),
            bg="#1a73e8",
            fg="white",
            activebackground="#1557b0",
            activeforeground="white",
            relief="flat",
            padx=12,
            command=self._on_select_file,
        )
        self.select_btn.pack(side="left")

        # ── Password section ────────────────────────────────────────
        password_group = LabelFrame(
            content_frame,
            text="2. Enter PDF Password",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f4f8",
            fg="#1a73e8",
            padx=12,
            pady=10,
        )
        password_group.pack(fill="x", pady=(0, 12))

        pw_row = Frame(password_group, bg="#f0f4f8")
        pw_row.pack(fill="x")

        self.password_entry = Entry(
            pw_row,
            font=("Segoe UI", 10),
            bg="white",
            fg="#555",
            relief="solid",
            bd=1,
            insertbackground="#1a73e8",
            show="•",
        )
        self.password_entry.pack(side="left", fill="x", expand=True)

        self.show_pw_btn = Button(
            pw_row,
            text="Show",
            font=("Segoe UI", 9),
            bg="#e8eaed",
            fg="#444",
            activebackground="#d2d5d9",
            relief="flat",
            padx=10,
            command=self._on_toggle_password,
        )
        self.show_pw_btn.pack(side="left", padx=(8, 0))

        # ── Process button ───────────────────────────────────────────
        self.process_btn = Button(
            content_frame,
            text="▶  Process Statement",
            font=("Segoe UI", 11, "bold"),
            bg="#34a853",
            fg="white",
            activebackground="#2d8e47",
            activeforeground="white",
            relief="flat",
            padx=24,
            pady=8,
            state="disabled",
            command=self._on_process,
        )
        self.process_btn.pack(pady=(0, 12))

        # ── Status section ──────────────────────────────────────────
        status_group = LabelFrame(
            content_frame,
            text="Status",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f4f8",
            fg="#1a73e8",
            padx=12,
            pady=10,
        )
        status_group.pack(fill="both", expand=True, pady=(0, 8))

        # Scrollable text area for logs
        text_frame = Frame(status_group, bg="white")
        text_frame.pack(fill="both", expand=True)

        self.log_text = Text(
            text_frame,
            font=("Consolas", 9),
            bg="white",
            fg="#2c2c2c",
            relief="solid",
            bd=1,
            state="disabled",
            wrap=WORD,
            height=14,
            insertbackground="#1a73e8",
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(text_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        self.status_label = Label(
            content_frame,
            text="Ready — select a PDF to begin.",
            font=("Segoe UI", 9),
            fg="#5f6368",
            bg="#f0f4f8",
            anchor="w",
        )
        self.status_label.pack(fill="x", pady=(4, 0))

        # ── Footer ───────────────────────────────────────────────────
        footer_frame = Frame(self.root, bg="#e8eaed", pady=6)
        footer_frame.pack(fill="x", side="bottom")

        footer_label = Label(
            footer_frame,
            text="YES BANK Statement Automation  |  Secure PDF Processing",
            font=("Segoe UI", 8),
            fg="#5f6368",
            bg="#e8eaed",
        )
        footer_label.pack()

    # ------------------------------------------------------------------
    # Callbacks
    # ------------------------------------------------------------------
    def _on_select_file(self) -> None:
        """Handle file selection."""
        filepath = filedialog.askopenfilename(
            title="Select Bank Statement PDF",
            filetypes=get_selected_file_types(),
            initialdir=SCRIPT_DIR / "input",
        )

        if not filepath:
            return

        self.selected_file = Path(filepath)

        # Show relative path in entry
        try:
            relative = self.selected_file.relative_to(SCRIPT_DIR)
            display_path = str(relative)
        except ValueError:
            display_path = str(self.selected_file)

        self.file_entry.config(state="normal")
        self.file_entry.delete(0, END)
        self.file_entry.insert(0, display_path)
        self.file_entry.config(state="readonly")

        self._update_process_btn()

    def _on_toggle_password(self) -> None:
        """Toggle password visibility."""
        current = self.password_entry.cget("show")
        if current == "":
            self.password_entry.config(show="•")
            self.show_pw_btn.config(text="Show")
        else:
            self.password_entry.config(show="")
            self.show_pw_btn.config(text="Hide")

    def _on_process(self) -> None:
        """Trigger the full pipeline in a background thread."""
        if self.is_processing:
            return

        password = self.password_entry.get().strip()
        if not password:
            messagebox.showwarning(
                "Password Required",
                "Please enter the PDF password before processing.",
            )
            return

        self._set_processing(TRUE)
        self._log_info("Starting pipeline...")
        self._set_status("Processing... please wait.", "blue")

        # Run pipeline in background thread so UI stays responsive
        thread = threading.Thread(
            target=self._run_pipeline_thread,
            args=(password,),
            daemon=True,
        )
        thread.start()

    def _run_pipeline_thread(self, password: str) -> None:
        """Background thread that runs the pipeline."""
        try:
            # Copy selected file to input/current.pdf
            self.root.after(0, lambda: self._log_info("Copying file to input/..."))
            INPUT_PDF_PATH.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self.selected_file, INPUT_PDF_PATH)

            # Run pipeline
            result = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT_DIR / "run_pipeline.py"),
                    "--password", password,
                ],
                capture_output=True,
                text=True,
            )

            # Parse output
            output = (result.stdout or "") + (result.stderr or "")
            for line in output.strip().split("\n"):
                if line.strip():
                    if "| ERROR" in line or "| CRITICAL" in line:
                        self.root.after(0, lambda l=line: self._log_error(l))
                    else:
                        self.root.after(0, lambda l=line: self._log_info(l))

            if result.returncode == 0:
                self.root.after(0, lambda: self._on_success())
            else:
                self.root.after(0, lambda: self._on_failure("Pipeline returned an error."))

        except Exception as exc:
            self.root.after(0, lambda: self._on_failure(str(exc)))

    def _on_success(self) -> None:
        """Handle successful pipeline completion."""
        self._set_processing(FALSE)
        self._set_status(
            "✓ Success — Statement uploaded to Google Sheets!",
            "green",
        )
        messagebox.showinfo(
            "Success",
            "Statement processed and uploaded to Google Sheets successfully.",
        )

    def _on_failure(self, message: str) -> None:
        """Handle pipeline failure."""
        self._set_processing(FALSE)
        self._set_status(f"✗ Failed — {message}", "red")
        messagebox.showerror(
            "Processing Failed",
            f"The statement could not be processed.\n\n{message}\n\n"
            "Check the log above for details. The file has been moved to the 'failed' folder.",
        )

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------
    def _update_process_btn(self) -> None:
        """Enable/disable process button based on file selection."""
        has_file = self.selected_file is not None and self.selected_file.exists()
        self.process_btn.config(state="normal" if has_file else "disabled")

    def _set_processing(self, is_processing: bool) -> None:
        """Update UI to reflect processing state."""
        self.is_processing = is_processing
        self.select_btn.config(state="disabled" if is_processing else "normal")
        self.password_entry.config(state="disabled" if is_processing else "normal")
        self.show_pw_btn.config(state="disabled" if is_processing else "normal")
        self.process_btn.config(
            state="disabled",
            text="⏳ Processing..." if is_processing else "▶  Process Statement",
        )

    def _log_info(self, message: str) -> None:
        """Append an info line to the log text area."""
        self._append_log(message, "black")

    def _log_error(self, message: str) -> None:
        """Append an error line to the log text area."""
        self._append_log(message, "red")

    def _append_log(self, message: str, color: str) -> None:
        """Append a colored line to the log text area."""
        self.log_text.config(state="normal")
        self.log_text.tag_config("error", foreground="red")
        self.log_text.tag_config("success", foreground="green")

        tag = "error" if color == "red" else ("success" if color == "green" else None)

        self.log_text.insert(END, message + "\n", tag)
        self.log_text.see(END)
        self.log_text.config(state="disabled")

    def _set_status(self, message: str, color: str) -> None:
        """Update the status label."""
        self.status_label.config(text=message, fg=color)

    # ------------------------------------------------------------------
    # Entry point
    # ------------------------------------------------------------------
    def run(self) -> None:
        """Start the Tkinter event loop."""
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    try:
        config = load_config()
    except FileNotFoundError:
        root = Tk()
        root.withdraw()
        messagebox.showerror(
            "Configuration Error",
            f"config.json not found.\n\n"
            f"Please create a config.json file in:\n{SCRIPT_DIR}",
        )
        root.destroy()
        return 1
    except json.JSONDecodeError:
        root = Tk()
        root.withdraw()
        messagebox.showerror(
            "Configuration Error",
            f"config.json is not valid JSON.\n\n"
            f"Please check the file at:\n{CONFIG_PATH}",
        )
        root.destroy()
        return 1

    root = Tk()
    app = BankStatementApp(root)
    app.run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
