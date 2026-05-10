# -*- coding: utf-8 -*-
"""
XTRCTR - PDF Page Extractor & Compressor
A modern GUI tool for extracting, splitting, and compressing PDF pages.
(With Local Tesseract OCR support for scanned PDFs)
"""

import os
import re
import sys
import shutil
import zipfile
import subprocess
import threading
import tempfile
import base64
import io
import json
from pathlib import Path
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF for PDF rendering
from PIL import Image

# Import local OCR utility
try:
    import ocr_utils
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


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
CONFIG_FILE = "config_ocr.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"use_local_ocr": True}

def save_config(config):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)
    except Exception:
        pass


# ============================================================
# PDF COMPRESSION (pypdf-based, no Ghostscript needed)
# ============================================================
def compress_pdf_pypdf(input_path: str, output_path: str) -> bool:
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
        shutil.copy2(input_path, output_path)
        return False


def compress_pdf_fitz_rasterize(input_path: str, output_path: str, dpi: int = 150, quality: int = 70) -> bool:
    try:
        doc = fitz.open(input_path)
        new_doc = fitz.open()
        
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG', quality=quality, optimize=True)
            img_data = img_byte_arr.getvalue()
            
            new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img_data)
            
        new_doc.save(output_path, garbage=4, deflate=True)
        new_doc.close()
        doc.close()
        return True
    except Exception:
        return False


def compress_pdf_ghostscript(input_path: str, output_path: str, level: str = "/ebook", dpi: int = None) -> bool:
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
            "-dDetectDuplicateImages=true",
            "-dSubsetFonts=true",
            "-dColorImageDownsampleType=/Bicubic",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dMonoImageDownsampleType=/Bicubic",
        ]
        
        if dpi:
            command.extend([
                f"-dColorImageResolution={dpi}",
                f"-dGrayImageResolution={dpi}",
                f"-dMonoImageResolution={dpi}",
                "-dDownsampleColorImages=true",
                "-dDownsampleGrayImages=true",
                "-dDownsampleMonoImages=true",
            ])

        command.extend([
            f"-sOutputFile={output_path}",
            input_path,
        ])
        
        subprocess.run(command, check=True, capture_output=True, timeout=120)
        return True
    except Exception:
        return False


def smart_compress(input_path: str, output_path: str, max_mb: float = MAX_SIZE_MB) -> dict:
    max_bytes = max_mb * 1024 * 1024
    original_size = os.path.getsize(input_path)
    
    info = {
        "original_size": original_size,
        "final_size": original_size,
        "method": "none",
        "success": True,
    }

    if original_size <= max_bytes:
        shutil.copy2(input_path, output_path)
        info["method"] = "none (already small)"
        return info

    if compress_pdf_ghostscript(input_path, output_path, "/ebook"):
        info["method"] = "ghostscript-ebook"
        info["final_size"] = os.path.getsize(output_path)

        if info["final_size"] > max_bytes:
            if compress_pdf_ghostscript(input_path, output_path, "/screen"):
                info["method"] = "ghostscript-screen"
                info["final_size"] = os.path.getsize(output_path)

                if info["final_size"] > max_bytes:
                    for dpi in [60, 50, 40, 30, 20]:
                        if compress_pdf_ghostscript(input_path, output_path, "/screen", dpi=dpi):
                            info["method"] = f"ghostscript-aggressive-{dpi}dpi"
                            info["final_size"] = os.path.getsize(output_path)
                            if info["final_size"] <= max_bytes:
                                return info
    else:
        if compress_pdf_pypdf(input_path, output_path):
            info["method"] = "pypdf"
            info["final_size"] = os.path.getsize(output_path)
    
    if info["final_size"] > max_bytes:
        for dpi in [150, 120, 100, 80, 60, 40]:
            if compress_pdf_fitz_rasterize(input_path, output_path, dpi=dpi, quality=60):
                info["method"] = f"fitz-rasterize-{dpi}dpi"
                info["final_size"] = os.path.getsize(output_path)
                if info["final_size"] <= max_bytes:
                    break
                    
    if info["final_size"] > max_bytes:
        if compress_pdf_fitz_rasterize(input_path, output_path, dpi=30, quality=40):
            info["method"] = "fitz-rasterize-30dpi-low-qual"
            info["final_size"] = os.path.getsize(output_path)

    return info


# ============================================================
# PAGE RANGE PARSER WITH VALIDATION
# ============================================================
def parse_page_input(text: str, total_pages: int) -> tuple[list[tuple[int, int]], list[str]]:
    text = text.strip()
    if not text:
        raise ValueError("Page input cannot be empty.")

    if text.lower() == "all":
        return [(1, total_pages)], []

    ranges = []
    warnings = []
    parts = [p.strip() for p in text.split(",")]

    for part in parts:
        if not part:
            continue

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
# SMART TITLE EXTRACTION & OCR
# ============================================================
def sanitize_filename(name: str, max_len: int = 60) -> str:
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    name = name.strip('. ')
    if len(name) > max_len:
        truncated = name[:max_len].rsplit(' ', 1)[0]
        name = truncated if truncated else name[:max_len]
    return name.strip()

def _find_title_in_text(text: str) -> str | None:
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if re.match(r'^\d+$', line):
            continue
        if len(line) < 4:
            continue
        if re.match(r'^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}$', line):
            continue
        alpha_ratio = sum(c.isalpha() for c in line) / max(len(line), 1)
        if alpha_ratio < 0.3:
            continue

        clean = sanitize_filename(line)
        if clean and len(clean) >= 3:
            return clean

    return None

def extract_title_with_tesseract(pdf_path: str, page_num: int) -> str | None:
    """Uses local Tesseract OCR to find the title by reading the first line."""
    if not OCR_AVAILABLE:
        return None
    
    text = ocr_utils.extract_text_with_tesseract(pdf_path, page_num)
    if text:
        return _find_title_in_text(text)
    return None

def extract_title_from_pages(reader: PdfReader, start_page: int, end_page: int) -> str | None:
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

    for page_idx in range(start_page - 1, min(end_page, start_page + 1)):
        try:
            page = reader.pages[page_idx]
            text = page.extract_text() or ""
        except Exception:
            continue

        title = _find_title_in_text(text)
        if title:
            return title

    return None


def _search_outlines_for_range(reader: PdfReader, outlines, start_page: int, end_page: int) -> str | None:
    for item in outlines:
        if isinstance(item, list):
            result = _search_outlines_for_range(reader, item, start_page, end_page)
            if result:
                return result
        else:
            try:
                dest_page = reader.get_destination_page_number(item)
                if start_page - 1 <= dest_page <= end_page - 1:
                    title = str(item.title).strip()
                    if title and len(title) >= 2:
                        return title
            except Exception:
                continue
    return None

# ============================================================
# MAIN APPLICATION
# ============================================================
class XTRCTRApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("XTRCTR — PDF Extractor (with Local Tesseract OCR)")
        self.geometry("900x820")
        self.minsize(750, 700)
        self.configure(fg_color=BG_DARK)

        self.loaded_pdfs: list[str] = []
        self.file_items: list[dict] = []
        self.output_dir: str = ""
        self.is_processing = False
        
        self.current_preview_path = None
        self.preview_cache = {} 
        
        self.config = load_config()
        self.use_local_ocr = self.config.get("use_local_ocr", True)

        self._build_ui()

    def _build_ui(self):
        self.root_container = ctk.CTkFrame(self, fg_color="transparent")
        self.root_container.pack(fill="both", expand=True)

        self.left_panel = ctk.CTkFrame(self.root_container, fg_color="transparent")
        self.left_panel.pack(side="left", fill="both", expand=True)

        self.main_frame = ctk.CTkScrollableFrame(
            self.left_panel, 
            fg_color="transparent",
            scrollbar_button_color=ACCENT_PRIMARY,
            scrollbar_button_hover_color=ACCENT_SECONDARY
        )
        self.main_frame.pack(fill="both", expand=True, padx=12, pady=12)

        self.preview_panel = ctk.CTkFrame(
            self.root_container, 
            fg_color=BG_CARD, 
            width=0, 
            corner_radius=0,
            border_width=1,
            border_color=BG_DARK
        )

        self.preview_header = ctk.CTkFrame(self.preview_panel, fg_color="transparent")
        self.preview_header.pack(fill="x", padx=10, pady=10)

        self.preview_title = ctk.CTkLabel(
            self.preview_header,
            text="预览 | Preview",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=ACCENT_SECONDARY
        )
        self.preview_title.pack(side="left")

        self.btn_close_preview = ctk.CTkButton(
            self.preview_header,
            text="✕",
            width=28,
            height=28,
            fg_color="transparent",
            hover_color=BG_DARK,
            text_color=TEXT_MUTED,
            command=self._close_preview
        )
        self.btn_close_preview.pack(side="right")

        self.preview_scroll = ctk.CTkScrollableFrame(
            self.preview_panel,
            fg_color="transparent",
            scrollbar_button_color=ACCENT_PRIMARY,
            scrollbar_button_hover_color=ACCENT_SECONDARY
        )
        self.preview_scroll.pack(fill="both", expand=True, padx=5, pady=5)

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
            text="PDF Page Extractor + Local OCR",
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
            text="➕  Add PDF Files",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_SECONDARY,
            height=42,
            width=160,
            corner_radius=10,
            command=self._select_files,
        )
        self.btn_select.pack(side="left")

        self.btn_clear = ctk.CTkButton(
            file_inner,
            text="🗑️ Clear All",
            font=ctk.CTkFont(size=12),
            fg_color="#334155",
            hover_color=ACCENT_DANGER,
            height=32,
            width=100,
            corner_radius=8,
            command=self._clear_files,
        )
        self.btn_clear.pack_forget()

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
        self.step2_header = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.step2_header.pack(fill="x", pady=(8, 4))
        
        self._make_section_header(self.step2_header, "② Enter Pages to Extract", pack=False)
        self.btn_apply_all = ctk.CTkButton(
            self.step2_header,
            text="📋 Apply first to all",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="transparent",
            text_color=ACCENT_SECONDARY,
            hover_color=BG_CARD,
            height=24,
            width=120,
            command=self._apply_first_range_to_all,
        )
        
        self.page_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        self.page_card.pack(fill="x", pady=(0, 14))

        self.files_scroll_frame = ctk.CTkScrollableFrame(
            self.page_card,
            fg_color="transparent",
            height=200,
            scrollbar_button_color=ACCENT_PRIMARY,
            scrollbar_button_hover_color=ACCENT_SECONDARY,
        )
        self.files_scroll_frame.pack(fill="x", padx=4, pady=4)

        self.empty_files_label = ctk.CTkLabel(
            self.files_scroll_frame,
            text="Please select PDF files first...",
            font=ctk.CTkFont(size=13, slant="italic"),
            text_color=TEXT_MUTED,
        )
        self.empty_files_label.pack(pady=40)

        # ---- Step 3: OCR Settings ----
        self._make_section_header(self.main_frame, "③ Local OCR Settings")
        ocr_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        ocr_card.pack(fill="x", pady=(0, 14))
        
        ocr_inner = ctk.CTkFrame(ocr_card, fg_color="transparent")
        ocr_inner.pack(fill="x", padx=16, pady=14)

        self.ocr_var = ctk.BooleanVar(value=self.use_local_ocr)
        chk = ctk.CTkCheckBox(
            ocr_inner, 
            text="Enable Tesseract OCR to read text from scanned PDFs locally",
            variable=self.ocr_var,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_SECONDARY,
            command=self._save_ocr_setting
        )
        chk.pack(anchor="w")
        
        if not OCR_AVAILABLE or not os.path.exists(r'C:\Program Files\Tesseract-OCR\tesseract.exe'):
            self.ocr_var.set(False)
            chk.configure(state="disabled", text="Enable Tesseract OCR (NOT INSTALLED - Please install Tesseract)")

        # ---- Step 4: Output ----
        self._make_section_header(self.main_frame, "④ Output Settings")

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


    def _save_ocr_setting(self):
        self.config["use_local_ocr"] = self.ocr_var.get()
        save_config(self.config)

    def _make_section_header(self, parent, text: str, pack: bool = True):
        lbl = ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=TEXT_PRIMARY,
        )
        if pack:
            lbl.pack(anchor="w", pady=(8, 4))
        else:
            lbl.pack(side="left", pady=(8, 4))

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
        self.btn_clear.configure(state=state)
        self.btn_output.configure(state=state)
        self.btn_extract.configure(state=state)
        for item in self.file_items:
            item["entry"].configure(state=state)
        self.btn_apply_all.configure(state=state)

    def _prompt_filename(self, suggested_name: str, page_label: str) -> str:
        result = {"name": suggested_name}
        event = threading.Event()

        def _show_dialog():
            win = ctk.CTkToplevel(self)
            win.title("📝 Choose Filename")
            win.geometry("500x210")
            win.configure(fg_color=BG_DARK)
            win.resizable(False, False)
            win.transient(self)
            win.grab_set()
            win.focus_force()

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
        event.wait()
        return result["name"]

    def _get_pdf_page_image(self, doc: fitz.Document, page_num: int) -> ctk.CTkImage | None:
        try:
            page = doc[page_num]
            zoom = 1.2
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            display_w = 280
            aspect = pix.height / pix.width
            display_h = int(display_w * aspect)
            
            return ctk.CTkImage(light_image=img, dark_image=img, size=(display_w, display_h))
        except Exception:
            return None

    def _toggle_preview(self, pdf_path: str, btn_toggle: ctk.CTkButton):
        if self.current_preview_path == pdf_path:
            self._close_preview()
            return

        if not self.preview_panel.winfo_ismapped():
            self.preview_panel.pack(side="right", fill="both", expand=False)
            self.preview_panel.configure(width=340)

        self.current_preview_path = pdf_path
        self.preview_title.configure(text=f"👁️ {os.path.basename(pdf_path)}")
        
        for widget in self.preview_scroll.winfo_children():
            widget.destroy()

        threading.Thread(target=self._load_full_preview_worker, args=(pdf_path,), daemon=True).start()

    def _close_preview(self):
        self.preview_panel.pack_forget()
        self.current_preview_path = None

    def _load_full_preview_worker(self, pdf_path: str):
        try:
            doc = fitz.open(pdf_path)
            total = len(doc)
            
            for i in range(total):
                if self.current_preview_path != pdf_path:
                    break 
                
                img = self._get_pdf_page_image(doc, i)
                if img:
                    self.after(0, self._add_page_to_preview, i + 1, img)
            
            doc.close()
        except Exception as e:
            self.after(0, self._log, f"Error loading preview: {e}")

    def _add_page_to_preview(self, page_num: int, img: ctk.CTkImage):
        page_container = ctk.CTkFrame(self.preview_scroll, fg_color="transparent")
        page_container.pack(pady=10, fill="x")

        num_lbl = ctk.CTkLabel(
            page_container, 
            text=f"Page {page_num}", 
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=ACCENT_PRIMARY
        )
        num_lbl.pack()

        img_lbl = ctk.CTkLabel(page_container, image=img, text="")
        img_lbl.pack()

    def _select_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Add PDF Files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if filepaths:
            added_count = 0
            for fp in filepaths:
                if fp not in self.loaded_pdfs:
                    self.loaded_pdfs.append(fp)
                    added_count += 1
            
            if added_count == 0:
                return

            names = [os.path.basename(f) for f in self.loaded_pdfs]
            display = ", ".join(names)
            if len(display) > 80:
                display = display[:77] + "..."
            
            self.file_label.configure(
                text=f"📎 {len(self.loaded_pdfs)} file(s): {display}",
                text_color=ACCENT_SUCCESS,
            )
            self.btn_clear.pack(side="left", padx=(12, 0))
            self._log(f"Added {added_count} PDF(s). Total: {len(self.loaded_pdfs)}")

            for fp in filepaths:
                try:
                    r = PdfReader(fp)
                    self._log(f"  📄 {os.path.basename(fp)} → {len(r.pages)} pages")
                except Exception:
                    pass
            
            self._update_file_list_ui()

    def _clear_files(self):
        if not self.loaded_pdfs:
            return
        
        self.loaded_pdfs = []
        self.file_label.configure(text="No files selected", text_color=TEXT_MUTED)
        self.btn_clear.pack_forget()
        self._log("Cleared all files.")
        self._update_file_list_ui()

    def _remove_file(self, pdf_path: str):
        if pdf_path in self.loaded_pdfs:
            self.loaded_pdfs.remove(pdf_path)
            self._log(f"Removed: {os.path.basename(pdf_path)}")
            
            if not self.loaded_pdfs:
                self._clear_files()
            else:
                names = [os.path.basename(f) for f in self.loaded_pdfs]
                display = ", ".join(names)
                if len(display) > 80:
                    display = display[:77] + "..."
                self.file_label.configure(text=f"📎 {len(self.loaded_pdfs)} file(s): {display}")
                self._update_file_list_ui()

    def _select_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir = folder
            self.output_label.configure(
                text=f"📂 {folder}",
                text_color=ACCENT_SUCCESS,
            )
            self._log(f"Output folder: {folder}")

    def _update_file_list_ui(self):
        for widget in self.files_scroll_frame.winfo_children():
            widget.destroy()
        self.file_items = []

        if not self.loaded_pdfs:
            self.empty_files_label = ctk.CTkLabel(
                self.files_scroll_frame,
                text="Please select PDF files first...",
                font=ctk.CTkFont(size=13, slant="italic"),
                text_color=TEXT_MUTED,
            )
            self.empty_files_label.pack(pady=40)
            self.btn_apply_all.pack_forget()
            return

        self.btn_apply_all.pack(side="right", padx=(0, 4), pady=(8, 4))

        for idx, pdf_path in enumerate(self.loaded_pdfs):
            name = os.path.basename(pdf_path)
            try:
                reader = PdfReader(pdf_path)
                page_count = len(reader.pages)
            except Exception:
                page_count = "?"

            container = ctk.CTkFrame(self.files_scroll_frame, fg_color="transparent")
            container.pack(fill="x", pady=4)

            item_frame = ctk.CTkFrame(container, fg_color="transparent")
            item_frame.pack(fill="x")

            btn_remove = ctk.CTkButton(
                item_frame,
                text="❌",
                font=ctk.CTkFont(size=10),
                fg_color="transparent",
                text_color=ACCENT_DANGER,
                hover_color=BG_DARK,
                width=24,
                height=24,
                corner_radius=4,
                command=lambda p=pdf_path: self._remove_file(p),
            )
            btn_remove.pack(side="left", padx=(4, 0))

            btn_preview = ctk.CTkButton(
                item_frame,
                text="👁️",
                font=ctk.CTkFont(size=14),
                fg_color="transparent",
                text_color=ACCENT_SECONDARY,
                hover_color=BG_DARK,
                width=28,
                height=28,
                corner_radius=6,
            )
            btn_preview.configure(command=lambda p=pdf_path, b=btn_preview: self._toggle_preview(p, b))
            btn_preview.pack(side="left", padx=(4, 0))

            name_label = ctk.CTkLabel(
                item_frame,
                text=f"📄 {name}",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=TEXT_PRIMARY,
                anchor="w",
                width=160,
            )
            name_label.pack(side="left", padx=(8, 4))

            btn_folder = ctk.CTkButton(
                item_frame,
                text="📂",
                font=ctk.CTkFont(size=12),
                fg_color="transparent",
                text_color=ACCENT_PRIMARY,
                hover_color=BG_DARK,
                width=28,
                height=28,
                corner_radius=6,
                command=lambda p=pdf_path: self._select_output_for_file(p),
            )
            btn_folder.pack(side="left")

            folder_label = ctk.CTkLabel(
                item_frame,
                text="Default",
                font=ctk.CTkFont(size=10),
                text_color=TEXT_MUTED,
                width=80,
                anchor="w",
            )
            folder_label.pack(side="left", padx=(2, 8))

            count_label = ctk.CTkLabel(
                item_frame,
                text=f"({page_count} pgs)",
                font=ctk.CTkFont(size=11),
                text_color=TEXT_MUTED,
                width=55,
            )
            count_label.pack(side="left")

            entry = ctk.CTkEntry(
                item_frame,
                placeholder_text="e.g. 1, 3-5, all",
                font=ctk.CTkFont(size=13),
                fg_color=BG_INPUT,
                border_color=ACCENT_PRIMARY,
                height=32,
                corner_radius=8,
            )
            entry.pack(side="right", padx=(8, 8), fill="x", expand=True)
            entry.insert(0, "all")

            self.file_items.append({
                "path": pdf_path,
                "entry": entry,
                "pages": page_count,
                "output_dir": "", 
                "folder_label": folder_label,
            })

    def _apply_first_range_to_all(self):
        if not self.file_items:
            return
        
        first_val = self.file_items[0]["entry"].get().strip()
        for item in self.file_items[1:]:
            item["entry"].delete(0, "end")
            item["entry"].insert(0, first_val)
        self._log(f"Applied range '{first_val}' to all files.")

    def _select_output_for_file(self, pdf_path: str):
        folder = filedialog.askdirectory(title=f"Output Folder for {os.path.basename(pdf_path)}")
        if folder:
            for item in self.file_items:
                if item["path"] == pdf_path:
                    item["output_dir"] = folder
                    display = os.path.basename(folder)
                    if not display: 
                        display = folder
                    item["folder_label"].configure(text=display, text_color=ACCENT_SUCCESS)
                    self._log(f"Target folder for {os.path.basename(pdf_path)}: {folder}")
                    break

    def _start_extraction(self):
        if self.is_processing:
            return

        if not self.file_items:
            messagebox.showwarning("No Files", "Please select at least one PDF file.")
            return

        extraction_plan = []
        for item in self.file_items:
            path = item["path"]
            page_text = item["entry"].get().strip()
            out_dir_override = item["output_dir"]
            
            if not page_text:
                messagebox.showwarning(
                    "Missing Input", 
                    f"Please enter pages for {os.path.basename(path)} or remove it."
                )
                return
            
            extraction_plan.append({
                "path": path,
                "pages": page_text,
                "out_dir": out_dir_override
            })

        missing = [p["path"] for p in extraction_plan if not os.path.isfile(p["path"])]
        if missing:
            messagebox.showerror("Missing Files", f"These files no longer exist:\n{chr(10).join(missing)}")
            return

        self.is_processing = True
        self._toggle_ui(False)
        self._set_status("Processing...", ACCENT_WARNING)
        self._set_progress(0)

        thread = threading.Thread(target=self._extraction_worker, args=(extraction_plan,), daemon=True)
        thread.start()

    def _extraction_worker(self, extraction_plan: list[dict]):
        try:
            total_pdfs = len(extraction_plan)
            all_results = []
            
            use_ocr = self.ocr_var.get() and OCR_AVAILABLE

            for pdf_idx, item in enumerate(extraction_plan):
                pdf_path = item["path"]
                page_text = item["pages"]
                pdf_out_override = item["out_dir"]
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

                try:
                    ranges, warnings = parse_page_input(page_text, total_pages)
                    for w in warnings:
                        self._log(f"  ⚠️ {w}")
                except ValueError as e:
                    self._log(f"❌ Page input error for {pdf_name}: {e}")
                    continue

                if pdf_out_override:
                    out_dir = pdf_out_override
                elif self.output_dir:
                    out_dir = self.output_dir
                else:
                    out_dir = os.path.dirname(pdf_path)

                os.makedirs(out_dir, exist_ok=True)

                output_files = []
                temp_files = []

                for r_idx, (start, end) in enumerate(ranges):
                    detected_title = None
                    if use_ocr:
                        self.after(0, self._set_status, f"Tesseract OCR reading {pdf_name}...", ACCENT_SECONDARY)
                        detected_title = extract_title_with_tesseract(pdf_path, start)
                        if detected_title:
                            self._log(f"  🧠 OCR Extracted: \"{detected_title}\"")
                    
                    if not detected_title:
                        detected_title = extract_title_from_pages(reader, start, end)
                        if detected_title:
                            self._log(f"  🔍 Standard Auto-detected: \"{detected_title}\"")

                    if detected_title:
                        suggested = detected_title
                    elif start == end:
                        suggested = f"{base_name}_page_{start}"
                    else:
                        suggested = f"{base_name}_pages_{start}-{end}"

                    if start == end:
                        page_label = f"page {start} of {pdf_name}"
                    else:
                        page_label = f"pages {start}–{end} of {pdf_name}"

                    self.after(0, self._set_status, "Awaiting User Input...", ACCENT_WARNING)
                    chosen_name = self._prompt_filename(suggested, page_label)
                    out_name = f"{chosen_name}.pdf"

                    final_path = os.path.join(out_dir, out_name)
                    counter = 1
                    while os.path.exists(final_path):
                        stem = os.path.splitext(out_name)[0]
                        final_path = os.path.join(out_dir, f"{stem}_{counter}.pdf")
                        counter += 1

                    self.after(0, self._set_status, f"Extracting {out_name}...", ACCENT_WARNING)
                    writer = PdfWriter()
                    for i in range(start - 1, end):
                        writer.add_page(reader.pages[i])

                    if self.compress_var.get():
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

                    progress = ((pdf_idx + (r_idx + 1) / len(ranges)) / total_pdfs)
                    self.after(0, self._set_progress, min(progress, 0.95))

                for tmp in temp_files:
                    try:
                        os.unlink(tmp)
                    except OSError:
                        pass

                if self.zip_var.get() and output_files:
                    zip_name = f"{base_name}_extracted.zip"
                    zip_path = os.path.join(out_dir, zip_name)
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

                all_results.extend(output_files)

            self.after(0, self._set_progress, 1.0)

            if all_results:
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
        try:
            os.startfile(path)
        except Exception:
            pass


def main():
    app = XTRCTRApp()
    app.mainloop()


if __name__ == "__main__":
    main()
