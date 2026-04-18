# -*- coding: utf-8 -*-
"""
XTRCTR - PDF Page Extractor & Compressor
A modern GUI tool for extracting, splitting, and compressing PDF pages.
"""

import os
import re
import sys
import shutil
import zipfile
import subprocess
import threading
import tempfile
from pathlib import Path
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter


# ============================================================
# THEME & CONSTANTS
# ============================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT_PRIMARY = "#6C63FF"
ACCENT_SECONDARY = "#A78BFA"
ACCENT_SUCCESS = "#34D399"
ACCENT_WARNING = "#FBBF24"
ACCENT_DANGER = "#F87171"
BG_DARK = "#0F0F1A"
BG_CARD = "#1A1A2E"
BG_INPUT = "#16213E"
TEXT_PRIMARY = "#E2E8F0"
TEXT_SECONDARY = "#94A3B8"
TEXT_MUTED = "#64748B"

MAX_SIZE_MB = 2.0  # Threshold for aggressive compression


# ============================================================
# PDF COMPRESSION (pypdf-based, no Ghostscript needed)
# ============================================================
def compress_pdf_pypdf(input_path: str, output_path: str) -> bool:
    """Compress a PDF using pypdf's built-in compression."""
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        for page in reader.pages:
            page.compress_content_streams()
            writer.add_page(page)
        with open(output_path, "wb") as f:
            writer.write(f)
        return True
    except Exception:
        # If compression fails, just copy the original
        shutil.copy2(input_path, output_path)
        return False


def compress_pdf_ghostscript(input_path: str, output_path: str, level: str = "/ebook") -> bool:
    """Compress a PDF using Ghostscript if available on the system."""
    gs_paths = [
        shutil.which("gswin64c"),
        shutil.which("gswin32c"),
        shutil.which("gs"),
    ]
    gs_exe = next((p for p in gs_paths if p), None)

    if not gs_exe:
        return False

    try:
        command = [
            gs_exe,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={level}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_path}",
            input_path,
        ]
        subprocess.run(command, check=True, capture_output=True, timeout=120)
        return True
    except Exception:
        return False


def smart_compress(input_path: str, output_path: str, max_mb: float = MAX_SIZE_MB) -> dict:
    """
    Try the best available compression. Returns info dict.
    Strategy: Try Ghostscript /ebook → /screen fallback → pypdf fallback.
    """
    original_size = os.path.getsize(input_path)
    info = {
        "original_size": original_size,
        "final_size": original_size,
        "method": "none",
        "success": True,
    }

    # Try Ghostscript first (better compression)
    if compress_pdf_ghostscript(input_path, output_path, "/ebook"):
        info["method"] = "ghostscript-ebook"
        info["final_size"] = os.path.getsize(output_path)

        # If still too large, try aggressive
        if info["final_size"] > max_mb * 1024 * 1024:
            if compress_pdf_ghostscript(input_path, output_path, "/screen"):
                info["method"] = "ghostscript-screen"
                info["final_size"] = os.path.getsize(output_path)
    else:
        # Fall back to pypdf compression
        if compress_pdf_pypdf(input_path, output_path):
            info["method"] = "pypdf"
            info["final_size"] = os.path.getsize(output_path)
        else:
            # Last resort: copy as-is
            shutil.copy2(input_path, output_path)
            info["method"] = "copy"

    return info


# ============================================================
# PAGE RANGE PARSER WITH VALIDATION
# ============================================================
def parse_page_input(text: str, total_pages: int) -> tuple[list[tuple[int, int]], list[str]]:
    """
    Parse page input like '1,3,5-10,all' into a list of (start, end) tuples.
    Auto-clamps out-of-range values instead of erroring.
    Returns (ranges, warnings) where warnings lists any clamping that occurred.
    Raises ValueError only on truly invalid input (empty, non-numeric, etc.).
    """
    text = text.strip()
    if not text:
        raise ValueError("Page input cannot be empty.")

    # Support 'all' keyword
    if text.lower() == "all":
        return [(1, total_pages)], []

    ranges = []
    warnings = []
    parts = [p.strip() for p in text.split(",")]

    for part in parts:
        if not part:
            continue

        # Support 'all' mixed with other parts
        if part.lower() == "all":
            ranges.append((1, total_pages))
            continue

        if "-" in part:
            pieces = part.split("-", maxsplit=1)
            if len(pieces) != 2 or not pieces[0].strip() or not pieces[1].strip():
                raise ValueError(f"Invalid range: '{part}'. Use format like '5-10'.")
            try:
                start = int(pieces[0].strip())
                end = int(pieces[1].strip())
            except ValueError:
                raise ValueError(f"Non-numeric value in range: '{part}'.")

            if start > end:
                raise ValueError(f"Start page ({start}) is greater than end page ({end}) in '{part}'.")
            if start < 1:
                start = 1
                warnings.append(f"Clamped start page to 1 in '{part}'.")
            if end > total_pages:
                warnings.append(f"Clamped '{part}' → {start}-{total_pages} (PDF has {total_pages} pages).")
                end = total_pages
            ranges.append((start, end))
        else:
            try:
                page = int(part)
            except ValueError:
                raise ValueError(f"Non-numeric page: '{part}'.")
            if page < 1:
                warnings.append(f"Skipped page {page} (must be ≥ 1).")
                continue
            if page > total_pages:
                warnings.append(f"Skipped page {page} (PDF only has {total_pages} pages).")
                continue
            ranges.append((page, page))

    if not ranges:
        raise ValueError(f"No valid pages in range. PDF has {total_pages} pages.")

    return ranges, warnings


# ============================================================
# SMART TITLE EXTRACTION
# ============================================================
def sanitize_filename(name: str, max_len: int = 60) -> str:
    """
    Clean a string for safe use as a Windows filename.
    Removes illegal characters, normalizes whitespace, and truncates.
    """
    # Remove characters illegal in Windows filenames
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    # Replace tabs, newlines, and multiple spaces with a single space
    name = re.sub(r'\s+', ' ', name).strip()
    # Remove leading/trailing dots (Windows doesn't like them)
    name = name.strip('. ')
    # Truncate to max length, breaking at a word boundary if possible
    if len(name) > max_len:
        truncated = name[:max_len].rsplit(' ', 1)[0]
        name = truncated if truncated else name[:max_len]
    return name.strip()


def extract_title_from_pages(reader: PdfReader, start_page: int, end_page: int) -> str | None:
    """
    Extract a title from the pages in the given range.
    Tries multiple strategies:
      1. PDF bookmarks/outlines that point to pages in this range
      2. Text heuristic: first meaningful line from the first page
    Returns a sanitized title string, or None if nothing useful found.
    """
    # Strategy 1: Check PDF bookmarks/outlines for titles mapping to this range
    try:
        outlines = reader.outline
        if outlines:
            title = _search_outlines_for_range(reader, outlines, start_page, end_page)
            if title:
                clean = sanitize_filename(title)
                if clean and len(clean) >= 3:
                    return clean
    except Exception:
        pass

    # Strategy 2: Text heuristic — scan first few pages for a heading
    for page_idx in range(start_page - 1, min(end_page, start_page + 1)):  # Check first 2 pages max
        try:
            page = reader.pages[page_idx]
            text = page.extract_text() or ""
        except Exception:
            continue

        title = _find_title_in_text(text)
        if title:
            return title

    return None


def _search_outlines_for_range(
    reader: PdfReader, outlines, start_page: int, end_page: int
) -> str | None:
    """
    Recursively search PDF outlines/bookmarks for an entry whose
    destination page falls within [start_page, end_page].
    """
    for item in outlines:
        if isinstance(item, list):
            result = _search_outlines_for_range(reader, item, start_page, end_page)
            if result:
                return result
        else:
            try:
                dest_page = reader.get_destination_page_number(item)
                # dest_page is 0-indexed; our range is 1-indexed
                if start_page - 1 <= dest_page <= end_page - 1:
                    title = str(item.title).strip()
                    if title and len(title) >= 2:
                        return title
            except Exception:
                continue
    return None


def _find_title_in_text(text: str) -> str | None:
    """
    Heuristic: find the first meaningful line in extracted text.
    Skips blank lines, page numbers, dates, and very short fragments.
    """
    lines = text.split('\n')

    for line in lines:
        line = line.strip()

        # Skip empty lines
        if not line:
            continue

        # Skip lines that are just numbers (page numbers)
        if re.match(r'^\d+$', line):
            continue

        # Skip very short lines (< 4 chars) — likely noise
        if len(line) < 4:
            continue

        # Skip lines that look like dates only
        if re.match(r'^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}$', line):
            continue

        # Skip lines that are mostly numbers/punctuation (e.g. ISBN, codes)
        alpha_ratio = sum(c.isalpha() for c in line) / max(len(line), 1)
        if alpha_ratio < 0.3:
            continue

        # This looks like a real title/heading
        clean = sanitize_filename(line)
        if clean and len(clean) >= 3:
            return clean

    return None


# ============================================================
# MAIN APPLICATION
# ============================================================
class XTRCTRApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("XTRCTR — PDF Page Extractor & Compressor")
        self.geometry("900x720")
        self.minsize(750, 600)
        self.configure(fg_color=BG_DARK)

        # State
        self.loaded_pdfs: list[str] = []
        self.output_dir: str = ""
        self.is_processing = False

        self._build_ui()

    # --------------------------------------------------------
    # UI CONSTRUCTION
    # --------------------------------------------------------
    def _build_ui(self):
        # Main container with padding
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=24, pady=16)

        # ---- Header ----
        header = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))

        title_label = ctk.CTkLabel(
            header,
            text="✦ XTRCTR",
            font=ctk.CTkFont(family="Segoe UI", size=32, weight="bold"),
            text_color=ACCENT_PRIMARY,
        )
        title_label.pack(side="left")

        subtitle = ctk.CTkLabel(
            header,
            text="PDF Page Extractor & Compressor",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=TEXT_SECONDARY,
        )
        subtitle.pack(side="left", padx=(14, 0), pady=(10, 0))

        # ---- Step 1: Select PDFs ----
        self._make_section_header(self.main_frame, "① Select PDF Files")

        file_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        file_card.pack(fill="x", pady=(0, 14))

        file_inner = ctk.CTkFrame(file_card, fg_color="transparent")
        file_inner.pack(fill="x", padx=16, pady=14)

        self.btn_select = ctk.CTkButton(
            file_inner,
            text="📁  Browse PDF Files",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_SECONDARY,
            height=42,
            corner_radius=10,
            command=self._select_files,
        )
        self.btn_select.pack(side="left")

        self.file_label = ctk.CTkLabel(
            file_inner,
            text="No files selected",
            font=ctk.CTkFont(size=13),
            text_color=TEXT_MUTED,
            wraplength=500,
            anchor="w",
        )
        self.file_label.pack(side="left", padx=(16, 0), fill="x", expand=True)

        # ---- Step 2: Page Selection ----
        self._make_section_header(self.main_frame, "② Enter Pages to Extract")

        page_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        page_card.pack(fill="x", pady=(0, 14))

        page_inner = ctk.CTkFrame(page_card, fg_color="transparent")
        page_inner.pack(fill="x", padx=16, pady=14)

        self.page_entry = ctk.CTkEntry(
            page_inner,
            placeholder_text="e.g.  1, 3, 5-10, 15-20",
            font=ctk.CTkFont(size=14),
            fg_color=BG_INPUT,
            border_color=ACCENT_PRIMARY,
            height=42,
            corner_radius=10,
        )
        self.page_entry.pack(fill="x")

        hint = ctk.CTkLabel(
            page_inner,
            text="Comma-separated pages or ranges. Each selection becomes a separate PDF.",
            font=ctk.CTkFont(size=11),
            text_color=TEXT_MUTED,
        )
        hint.pack(anchor="w", pady=(6, 0))

        # ---- Step 3: Output ----
        self._make_section_header(self.main_frame, "③ Output Settings")

        out_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        out_card.pack(fill="x", pady=(0, 14))

        out_inner = ctk.CTkFrame(out_card, fg_color="transparent")
        out_inner.pack(fill="x", padx=16, pady=14)

        self.btn_output = ctk.CTkButton(
            out_inner,
            text="📂  Choose Output Folder",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#334155",
            hover_color="#475569",
            height=42,
            corner_radius=10,
            command=self._select_output,
        )
        self.btn_output.pack(side="left")

        self.output_label = ctk.CTkLabel(
            out_inner,
            text="Default: same folder as input PDF",
            font=ctk.CTkFont(size=13),
            text_color=TEXT_MUTED,
            wraplength=400,
            anchor="w",
        )
        self.output_label.pack(side="left", padx=(16, 0), fill="x", expand=True)

        # ---- Compress toggle + ZIP toggle ----
        opts_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        opts_frame.pack(fill="x", pady=(0, 10))

        self.compress_var = ctk.BooleanVar(value=True)
        self.compress_check = ctk.CTkCheckBox(
            opts_frame,
            text="Compress extracted PDFs",
            variable=self.compress_var,
            font=ctk.CTkFont(size=13),
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_SECONDARY,
            corner_radius=6,
        )
        self.compress_check.pack(side="left", padx=(0, 24))

        self.zip_var = ctk.BooleanVar(value=True)
        self.zip_check = ctk.CTkCheckBox(
            opts_frame,
            text="Bundle output into ZIP",
            variable=self.zip_var,
            font=ctk.CTkFont(size=13),
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_SECONDARY,
            corner_radius=6,
        )
        self.zip_check.pack(side="left")

        # ---- Extract Button ----
        self.btn_extract = ctk.CTkButton(
            self.main_frame,
            text="⚡  EXTRACT & COMPRESS",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_SECONDARY,
            height=50,
            corner_radius=12,
            command=self._start_extraction,
        )
        self.btn_extract.pack(fill="x", pady=(4, 14))

        # ---- Progress ----
        self.progress_bar = ctk.CTkProgressBar(
            self.main_frame,
            progress_color=ACCENT_SUCCESS,
            fg_color=BG_CARD,
            height=6,
            corner_radius=3,
        )
        self.progress_bar.pack(fill="x", pady=(0, 6))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="Ready",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
        )
        self.status_label.pack(anchor="w")

        # ---- Log Area ----
        self.log_box = ctk.CTkTextbox(
            self.main_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=BG_CARD,
            text_color=TEXT_PRIMARY,
            corner_radius=12,
            height=140,
            state="disabled",
        )
        self.log_box.pack(fill="both", expand=True, pady=(8, 0))

    # --------------------------------------------------------
    # HELPERS
    # --------------------------------------------------------
    def _make_section_header(self, parent, text: str):
        lbl = ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=TEXT_PRIMARY,
        )
        lbl.pack(anchor="w", pady=(8, 4))

    def _log(self, msg: str, tag: str = ""):
        self.log_box.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, text: str, color: str = TEXT_MUTED):
        self.status_label.configure(text=text, text_color=color)

    def _set_progress(self, value: float):
        self.progress_bar.set(value)

    def _toggle_ui(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self.btn_select.configure(state=state)
        self.btn_output.configure(state=state)
        self.btn_extract.configure(state=state)
        self.page_entry.configure(state=state)

    def _prompt_filename(self, suggested_name: str, page_label: str) -> str:
        """
        Show a dialog asking the user to confirm or edit the filename.
        Runs on the main thread; the worker thread waits via threading.Event.
        Returns the final filename (without .pdf extension).
        """
        result = {"name": suggested_name}
        event = threading.Event()

        def _show_dialog():
            # Custom dialog with pre-filled text
            win = ctk.CTkToplevel(self)
            win.title("📝 Choose Filename")
            win.geometry("500x210")
            win.configure(fg_color=BG_DARK)
            win.resizable(False, False)
            win.transient(self)
            win.grab_set()
            win.focus_force()

            # Center on parent
            win.update_idletasks()
            x = self.winfo_x() + (self.winfo_width() - 500) // 2
            y = self.winfo_y() + (self.winfo_height() - 210) // 2
            win.geometry(f"+{x}+{y}")

            lbl = ctk.CTkLabel(
                win,
                text=f"Filename for {page_label}:",
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color=TEXT_PRIMARY,
            )
            lbl.pack(padx=20, pady=(18, 4), anchor="w")

            hint = ctk.CTkLabel(
                win,
                text="Edit the name below or press Enter to accept",
                font=ctk.CTkFont(size=11),
                text_color=TEXT_MUTED,
            )
            hint.pack(padx=20, pady=(0, 8), anchor="w")

            entry = ctk.CTkEntry(
                win,
                font=ctk.CTkFont(size=14),
                fg_color=BG_INPUT,
                border_color=ACCENT_PRIMARY,
                height=42,
                corner_radius=10,
                width=460,
            )
            entry.pack(padx=20)
            entry.insert(0, suggested_name)
            entry.select_range(0, "end")
            entry.focus_force()

            def _on_confirm(_event=None):
                val = entry.get().strip()
                if val:
                    result["name"] = sanitize_filename(val)
                win.destroy()
                event.set()

            def _on_close():
                # User closed dialog — keep suggestion
                win.destroy()
                event.set()

            entry.bind("<Return>", _on_confirm)
            win.protocol("WM_DELETE_WINDOW", _on_close)

            btn = ctk.CTkButton(
                win,
                text="✓  Confirm",
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=ACCENT_PRIMARY,
                hover_color=ACCENT_SECONDARY,
                height=40,
                corner_radius=10,
                command=_on_confirm,
            )
            btn.pack(padx=20, pady=(14, 16))

        self.after(0, _show_dialog)
        event.wait()  # Block worker thread until dialog closes
        return result["name"]

    # --------------------------------------------------------
    # FILE SELECTION
    # --------------------------------------------------------
    def _select_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if filepaths:
            self.loaded_pdfs = list(filepaths)
            names = [os.path.basename(f) for f in self.loaded_pdfs]
            display = ", ".join(names)
            if len(display) > 80:
                display = display[:77] + "..."
            self.file_label.configure(
                text=f"📎 {len(self.loaded_pdfs)} file(s): {display}",
                text_color=ACCENT_SUCCESS,
            )
            self._log(f"Selected {len(self.loaded_pdfs)} PDF(s): {', '.join(names)}")

            # Show page counts for each PDF
            for fp in self.loaded_pdfs:
                try:
                    r = PdfReader(fp)
                    self._log(f"  📄 {os.path.basename(fp)} → {len(r.pages)} pages")
                except Exception:
                    self._log(f"  ⚠️ {os.path.basename(fp)} → could not read page count")

    def _select_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir = folder
            self.output_label.configure(
                text=f"📂 {folder}",
                text_color=ACCENT_SUCCESS,
            )
            self._log(f"Output folder: {folder}")

    # --------------------------------------------------------
    # EXTRACTION LOGIC
    # --------------------------------------------------------
    def _start_extraction(self):
        if self.is_processing:
            return

        # Validate inputs
        if not self.loaded_pdfs:
            messagebox.showwarning("No Files", "Please select at least one PDF file.")
            return

        page_text = self.page_entry.get().strip()
        if not page_text:
            messagebox.showwarning("No Pages", "Please enter the pages you want to extract.")
            return

        # Verify files still exist
        missing = [f for f in self.loaded_pdfs if not os.path.isfile(f)]
        if missing:
            messagebox.showerror("Missing Files", f"These files no longer exist:\n{chr(10).join(missing)}")
            return

        self.is_processing = True
        self._toggle_ui(False)
        self._set_status("Processing...", ACCENT_WARNING)
        self._set_progress(0)

        # Run in thread to keep GUI responsive
        thread = threading.Thread(target=self._extraction_worker, args=(page_text,), daemon=True)
        thread.start()

    def _extraction_worker(self, page_text: str):
        try:
            total_pdfs = len(self.loaded_pdfs)
            all_results = []

            for pdf_idx, pdf_path in enumerate(self.loaded_pdfs):
                pdf_name = os.path.basename(pdf_path)
                base_name = os.path.splitext(pdf_name)[0]
                self._log(f"━━━ Processing: {pdf_name} ━━━")
                self.after(0, self._set_status, f"Reading {pdf_name}...", ACCENT_WARNING)

                try:
                    reader = PdfReader(pdf_path)
                except Exception as e:
                    self._log(f"❌ Failed to read {pdf_name}: {e}")
                    continue

                total_pages = len(reader.pages)
                self._log(f"  Total pages: {total_pages}")

                # Parse pages
                try:
                    ranges, warnings = parse_page_input(page_text, total_pages)
                    for w in warnings:
                        self._log(f"  ⚠️ {w}")
                except ValueError as e:
                    self._log(f"❌ Page input error for {pdf_name}: {e}")
                    continue

                # Determine output directory
                if self.output_dir:
                    out_dir = self.output_dir
                else:
                    out_dir = os.path.dirname(pdf_path)

                os.makedirs(out_dir, exist_ok=True)

                output_files = []
                temp_files = []

                for r_idx, (start, end) in enumerate(ranges):
                    # Smart title detection for suggestion
                    detected_title = extract_title_from_pages(reader, start, end)

                    if detected_title:
                        suggested = detected_title
                        self._log(f"  🔍 Auto-detected: \"{detected_title}\"")
                    elif start == end:
                        suggested = f"{base_name}_page_{start}"
                    else:
                        suggested = f"{base_name}_pages_{start}-{end}"

                    # Show naming dialog — user can accept or edit
                    if start == end:
                        page_label = f"page {start} of {pdf_name}"
                    else:
                        page_label = f"pages {start}–{end} of {pdf_name}"

                    chosen_name = self._prompt_filename(suggested, page_label)
                    out_name = f"{chosen_name}.pdf"

                    # Prevent filename collisions
                    final_path = os.path.join(out_dir, out_name)
                    counter = 1
                    while os.path.exists(final_path):
                        stem = os.path.splitext(out_name)[0]
                        final_path = os.path.join(out_dir, f"{stem}_{counter}.pdf")
                        counter += 1

                    # Extract pages
                    writer = PdfWriter()
                    for i in range(start - 1, end):
                        writer.add_page(reader.pages[i])

                    if self.compress_var.get():
                        # Write to temp, then compress to final
                        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                            tmp_path = tmp.name
                            writer.write(tmp)
                        temp_files.append(tmp_path)

                        self.after(0, self._set_status, f"Compressing {out_name}...", ACCENT_WARNING)
                        comp_info = smart_compress(tmp_path, final_path)

                        orig_kb = comp_info["original_size"] / 1024
                        final_kb = comp_info["final_size"] / 1024
                        method = comp_info["method"]
                        if orig_kb > 0:
                            ratio = (1 - final_kb / orig_kb) * 100
                        else:
                            ratio = 0
                        self._log(
                            f"  ✅ {os.path.basename(final_path)} — "
                            f"{final_kb:.0f} KB ({ratio:.0f}% reduced, {method})"
                        )
                    else:
                        with open(final_path, "wb") as f:
                            writer.write(f)
                        size_kb = os.path.getsize(final_path) / 1024
                        self._log(f"  ✅ {os.path.basename(final_path)} — {size_kb:.0f} KB")

                    output_files.append(final_path)

                    # Update progress
                    progress = ((pdf_idx + (r_idx + 1) / len(ranges)) / total_pdfs)
                    self.after(0, self._set_progress, min(progress, 0.95))

                # Clean up temp files
                for tmp in temp_files:
                    try:
                        os.unlink(tmp)
                    except OSError:
                        pass

                # ZIP if requested
                if self.zip_var.get() and output_files:
                    zip_name = f"{base_name}_extracted.zip"
                    zip_path = os.path.join(out_dir, zip_name)
                    # Prevent collision
                    counter = 1
                    while os.path.exists(zip_path):
                        zip_path = os.path.join(out_dir, f"{base_name}_extracted_{counter}.zip")
                        counter += 1

                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in output_files:
                            zf.write(f, arcname=os.path.basename(f))

                    zip_kb = os.path.getsize(zip_path) / 1024
                    self._log(f"  📦 ZIP created: {os.path.basename(zip_path)} ({zip_kb:.0f} KB)")
                    all_results.append(zip_path)

                    # Optionally remove individual files when zipped
                    # (keeping them for now so user has both)

                all_results.extend(output_files)

            self.after(0, self._set_progress, 1.0)

            if all_results:
                # Open the output folder
                result_dir = os.path.dirname(all_results[0])
                self._log(f"\n🎉 All done! Files saved to: {result_dir}")
                self.after(0, self._set_status, "✅ Complete!", ACCENT_SUCCESS)
                self.after(100, lambda: self._open_folder(result_dir))
            else:
                self._log("⚠️ No files were produced. Check errors above.")
                self.after(0, self._set_status, "⚠️ No output", ACCENT_WARNING)

        except Exception as e:
            self._log(f"❌ Unexpected error: {e}")
            self.after(0, self._set_status, f"Error: {e}", ACCENT_DANGER)

        finally:
            self.is_processing = False
            self.after(0, self._toggle_ui, True)

    def _open_folder(self, path: str):
        """Open the output folder in Explorer."""
        try:
            os.startfile(path)
        except Exception:
            pass


# ============================================================
# ENTRY POINT
# ============================================================
def main():
    app = XTRCTRApp()
    app.mainloop()


if __name__ == "__main__":
    main()
