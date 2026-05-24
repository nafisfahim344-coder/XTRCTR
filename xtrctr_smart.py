# -*- coding: utf-8 -*-
"""
XTRCTR Smart — Auto-Detect PDF Section Extractor
Automatically detects chapters/sections and extracts them with no manual page input.
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
import fitz
from PIL import Image
import json
from ai_utils import XTRCTRAI

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

ACCENT_PRIMARY   = "#6C63FF"
ACCENT_SECONDARY = "#A78BFA"
ACCENT_SUCCESS   = "#34D399"
ACCENT_WARNING   = "#FBBF24"
ACCENT_DANGER    = "#F87171"
BG_DARK  = "#0F0F1A"
BG_CARD  = "#1A1A2E"
BG_INPUT = "#16213E"
TEXT_PRIMARY   = "#E2E8F0"
TEXT_SECONDARY = "#94A3B8"
TEXT_MUTED     = "#64748B"

MAX_SIZE_MB = 2.0


# ============================================================
# COMPRESSION (identical to original)
# ============================================================
def compress_pdf_pypdf(input_path, output_path):
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


def compress_pdf_fitz_rasterize(input_path, output_path, dpi=150, quality=70):
    import io
    try:
        doc = fitz.open(input_path)
        new_doc = fitz.open()
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=quality, optimize=True)
            new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=buf.getvalue())
        new_doc.save(output_path, garbage=4, deflate=True)
        new_doc.close(); doc.close()
        return True
    except Exception:
        return False


def compress_pdf_ghostscript(input_path, output_path, level="/ebook", dpi=None):
    gs_paths = [shutil.which("gswin64c"), shutil.which("gswin32c"), shutil.which("gs")]
    gs_exe = next((p for p in gs_paths if p), None)
    if not gs_exe:
        return False
    try:
        cmd = [gs_exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
               f"-dPDFSETTINGS={level}", "-dNOPAUSE", "-dQUIET", "-dBATCH",
               "-dDetectDuplicateImages=true", "-dSubsetFonts=true"]
        if dpi:
            cmd += [f"-dColorImageResolution={dpi}", f"-dGrayImageResolution={dpi}",
                    "-dDownsampleColorImages=true", "-dDownsampleGrayImages=true"]
        cmd += [f"-sOutputFile={output_path}", input_path]
        subprocess.run(cmd, check=True, capture_output=True, timeout=120)
        return True
    except Exception:
        return False


def smart_compress(input_path, output_path, max_mb=MAX_SIZE_MB):
    max_bytes = max_mb * 1024 * 1024
    orig_size = os.path.getsize(input_path)
    info = {"original_size": orig_size, "final_size": orig_size, "method": "none", "success": True}
    if orig_size <= max_bytes:
        shutil.copy2(input_path, output_path)
        info["method"] = "none (already small)"
        return info
    if compress_pdf_ghostscript(input_path, output_path, "/ebook"):
        info.update(method="ghostscript-ebook", final_size=os.path.getsize(output_path))
        if info["final_size"] > max_bytes:
            if compress_pdf_ghostscript(input_path, output_path, "/screen"):
                info.update(method="ghostscript-screen", final_size=os.path.getsize(output_path))
                if info["final_size"] > max_bytes:
                    for dpi in [60, 50, 40, 30, 20]:
                        if compress_pdf_ghostscript(input_path, output_path, "/screen", dpi=dpi):
                            info.update(method=f"ghostscript-{dpi}dpi", final_size=os.path.getsize(output_path))
                            if info["final_size"] <= max_bytes:
                                return info
    else:
        compress_pdf_pypdf(input_path, output_path)
        info.update(method="pypdf", final_size=os.path.getsize(output_path))
    if info["final_size"] > max_bytes:
        for dpi in [150, 120, 100, 80, 60, 40, 30]:
            if compress_pdf_fitz_rasterize(input_path, output_path, dpi=dpi, quality=60):
                info.update(method=f"fitz-{dpi}dpi", final_size=os.path.getsize(output_path))
                if info["final_size"] <= max_bytes:
                    break
    return info


# ============================================================
# FILENAME SANITIZER
# ============================================================
def sanitize(name, max_len=60):
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    name = re.sub(r'\s+', ' ', name).strip().strip('. ')
    if len(name) > max_len:
        t = name[:max_len].rsplit(' ', 1)[0]
        name = t if t else name[:max_len]
    return name.strip()


# ============================================================
# AI DOCUMENT CLASSIFICATION & SPLITTING ENGINE
# ============================================================

def auto_detect_sections(pdf_path, api_key=None, target_program="Auto-Detect (Let AI Decide)", progress_callback=None):
    """
    Run Document Splitting using Tesseract + Gemini.
    Returns: (sections: list[dict], method: str)
    """
    try:
        reader = PdfReader(pdf_path)
        total = len(reader.pages)
    except Exception:
        return [], "error"

    base = sanitize(os.path.splitext(os.path.basename(pdf_path))[0])

    if api_key and OCR_AVAILABLE:
        if progress_callback:
            progress_callback("Running local Tesseract OCR on all pages...", ACCENT_WARNING)
            
        # Step 1: Extract OCR Text from ALL pages
        full_text = ""
        for pg in range(1, total + 1):
            text = ocr_utils.extract_text_with_tesseract(pdf_path, pg)
            if text:
                full_text += f"\n--- PAGE {pg} ---\n{text}\n"

        if full_text.strip():
            if progress_callback:
                progress_callback("Sending text to Gemini for document splitting...", ACCENT_SECONDARY)
                
            # Step 2: Send to Gemini for Classification
            ai = XTRCTRAI(api_key)
            result_list, msg = ai.classify_document(full_text, target_program)
            
            if result_list and isinstance(result_list, list):
                secs = []
                for result in result_list:
                    if "official_pdf_name" in result:
                        confidence = result.get("confidence", 0)
                        official_name = result["official_pdf_name"]
                        start_page = int(result.get("start_page", 1))
                        end_page = int(result.get("end_page", total))
                        
                        # Clamp pages
                        start_page = max(1, min(start_page, total))
                        end_page = max(start_page, min(end_page, total))
                        
                        # Format name based on confidence and flags
                        if confidence >= 90 and not result.get("flag_for_manual_review"):
                            title = official_name
                        else:
                            reason = result.get("flag_reason", "Low Confidence")
                            title = f"⚠️ REVIEW: {official_name} [{reason}]"

                        secs.append({
                            "title": sanitize(title), 
                            "start": start_page, 
                            "end": end_page, 
                            "signal": f"gemini-classifier (Conf: {confidence}%)"
                        })
                
                if secs:
                    # Sort by start page just to be clean
                    secs.sort(key=lambda x: x["start"])
                    return secs, "AI Classifier & Splitter"

    # Fallback to single section if AI fails or is not configured
    return [{"title": base, "start": 1, "end": total, "signal": "fallback"}], "fallback"


# ============================================================
# AI CONFIGURATION
# ============================================================
CONFIG_FILE = "config.json"

def save_config(config):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)
    except Exception:
        pass

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "api_key": "",
        "target_program": "Auto-Detect (Let AI Decide)"
    }


# ============================================================
# MAIN APPLICATION — XTRCTR SMART
# ============================================================
class XTRCTRSmartApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("XTRCTR Smart — AI Document Splitter")
        self.geometry("1250x820")
        self.minsize(900, 700)
        self.configure(fg_color=BG_DARK)

        self.loaded_pdfs: list[str] = []
        self.file_items: list[dict] = []
        self.output_dir: str = ""
        self.is_processing = False
        
        # Preview Panel State
        self.current_preview_id = None
        self.preview_cache = {}

        # Load Config
        self.config = load_config()
        self.ai_key = self.config.get("api_key", "")
        self.target_program = self.config.get("target_program", "Auto-Detect (Let AI Decide)")

        self._build_ui()

    # --------------------------------------------------------
    # UI BUILD
    # --------------------------------------------------------
    def _build_ui(self):
        # Root container for Split View
        self.root_container = ctk.CTkFrame(self, fg_color="transparent")
        self.root_container.pack(fill="both", expand=True)

        # ---- Left Side: Controls ----
        self.left_panel = ctk.CTkFrame(self.root_container, fg_color="transparent")
        self.left_panel.pack(side="left", fill="both", expand=True)

        self.main_frame = ctk.CTkScrollableFrame(
            self.left_panel, fg_color="transparent",
            scrollbar_button_color=ACCENT_PRIMARY,
            scrollbar_button_hover_color=ACCENT_SECONDARY,
        )
        self.main_frame.pack(fill="both", expand=True, padx=14, pady=14)

        # Header
        hdr = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(hdr, text="✦ XTRCTR Smart",
                     font=ctk.CTkFont(family="Segoe UI", size=30, weight="bold"),
                     text_color=ACCENT_PRIMARY).pack(side="left")
        ctk.CTkLabel(hdr, text="Bangladeshi University Application AI Splitter",
                     font=ctk.CTkFont(size=13), text_color=TEXT_SECONDARY).pack(side="left", padx=(12,0), pady=(10,0))

        # Step 1 — File selection
        self._section_label("① Add Merged Application PDFs")
        file_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        file_card.pack(fill="x", pady=(0,14))
        
        prog_inner = ctk.CTkFrame(file_card, fg_color="transparent")
        prog_inner.pack(fill="x", padx=16, pady=(14, 0))
        
        ctk.CTkLabel(prog_inner, text="Target Program:", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=(0, 10))
        
        self.program_var = ctk.StringVar(value=self.target_program)
        self.program_menu = ctk.CTkOptionMenu(
            prog_inner, 
            variable=self.program_var,
            values=[
                "Auto-Detect (Let AI Decide)", 
                "EAP/KLP PROGRAMME (Set 1)", 
                "BACHELOR'S / MASTER'S (Set 2)"
            ],
            font=ctk.CTkFont(size=13),
            fg_color=BG_INPUT,
            button_color=ACCENT_PRIMARY,
            button_hover_color=ACCENT_SECONDARY,
            width=280,
            command=self._save_program_setting
        )
        self.program_menu.pack(side="left")
        
        fi = ctk.CTkFrame(file_card, fg_color="transparent")
        fi.pack(fill="x", padx=16, pady=14)

        ctk.CTkButton(fi, text="➕  Add PDFs",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      fg_color=ACCENT_PRIMARY, hover_color=ACCENT_SECONDARY,
                      height=42, width=140, corner_radius=10,
                      command=self._add_files).pack(side="left")

        self.btn_clear = ctk.CTkButton(fi, text="🗑️ Clear All",
                                       font=ctk.CTkFont(size=12),
                                       fg_color="#334155", hover_color=ACCENT_DANGER,
                                       height=32, width=90, corner_radius=8,
                                       command=self._clear_all)
        self.btn_clear.pack_forget()

        self.file_count_lbl = ctk.CTkLabel(fi, text="No files added",
                                           font=ctk.CTkFont(size=13),
                                           text_color=TEXT_MUTED, anchor="w")
        self.file_count_lbl.pack(side="left", padx=(14,0), fill="x", expand=True)

        # Step 2 — File list + section review cards
        self._section_label("② Review AI Splits & Classifications")
        self.list_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        self.list_card.pack(fill="x", pady=(0,14))

        self.list_scroll = ctk.CTkScrollableFrame(
            self.list_card, fg_color="transparent", height=340,
            scrollbar_button_color=ACCENT_PRIMARY,
            scrollbar_button_hover_color=ACCENT_SECONDARY,
        )
        self.list_scroll.pack(fill="x", padx=6, pady=6)

        self.empty_lbl = ctk.CTkLabel(self.list_scroll,
                                      text="Add merged PDFs above — the AI will split, classify, and name them.",
                                      font=ctk.CTkFont(size=13, slant="italic"),
                                      text_color=TEXT_MUTED)
        self.empty_lbl.pack(pady=40)

        # Step 3 — Output + options
        self._section_label("③ Output Settings")
        out_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        out_card.pack(fill="x", pady=(0,14))
        oi = ctk.CTkFrame(out_card, fg_color="transparent")
        oi.pack(fill="x", padx=16, pady=14)

        ctk.CTkButton(oi, text="📂  Output Folder",
                      font=ctk.CTkFont(size=13, weight="bold"),
                      fg_color="#334155", hover_color="#475569",
                      height=38, corner_radius=10,
                      command=self._pick_output).pack(side="left")

        self.out_lbl = ctk.CTkLabel(oi, text="Default: same folder as each PDF",
                                    font=ctk.CTkFont(size=12), text_color=TEXT_MUTED, anchor="w")
        self.out_lbl.pack(side="left", padx=(12,0), fill="x", expand=True)

        # Step 4 — AI Settings
        self._section_label("④ Document AI Configuration")
        ai_card = ctk.CTkFrame(self.main_frame, fg_color=BG_CARD, corner_radius=12)
        ai_card.pack(fill="x", pady=(0,14))
        ai_inner = ctk.CTkFrame(ai_card, fg_color="transparent")
        ai_inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(ai_inner, text="Gemini API Key:", font=ctk.CTkFont(size=12)).pack(side="left")
        self.key_entry = ctk.CTkEntry(ai_inner, placeholder_text="Paste your Google API key here...",
                                      fg_color=BG_INPUT, border_color=ACCENT_PRIMARY,
                                      height=32, width=300, show="*")
        self.key_entry.pack(side="left", padx=10)
        self.key_entry.insert(0, self.ai_key)

        ctk.CTkButton(ai_inner, text="💾 Save",
                      font=ctk.CTkFont(size=11, weight="bold"),
                      fg_color=ACCENT_SUCCESS, hover_color="#059669",
                      height=32, width=70, corner_radius=8,
                      command=self._save_api_key).pack(side="left")

        # Checkboxes
        opts = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        opts.pack(fill="x", pady=(0,10))
        self.compress_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts, text="Compress output PDFs",
                        variable=self.compress_var, font=ctk.CTkFont(size=13),
                        fg_color=ACCENT_PRIMARY, hover_color=ACCENT_SECONDARY,
                        corner_radius=6).pack(side="left", padx=(0,24))
        self.zip_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts, text="Bundle into ZIP",
                        variable=self.zip_var, font=ctk.CTkFont(size=13),
                        fg_color=ACCENT_PRIMARY, hover_color=ACCENT_SECONDARY,
                        corner_radius=6).pack(side="left")

        # Extract button
        self.btn_extract = ctk.CTkButton(
            self.main_frame, text="⚡  SPLIT, RENAME & COMPRESS",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=ACCENT_PRIMARY, hover_color=ACCENT_SECONDARY,
            height=50, corner_radius=12,
            command=self._start_extraction)
        self.btn_extract.pack(fill="x", pady=(4,12))

        # Progress + status
        self.progress = ctk.CTkProgressBar(self.main_frame,
                                           progress_color=ACCENT_SUCCESS,
                                           fg_color=BG_CARD, height=6, corner_radius=3)
        self.progress.pack(fill="x", pady=(0,6))
        self.progress.set(0)

        self.status_lbl = ctk.CTkLabel(self.main_frame, text="Ready",
                                       font=ctk.CTkFont(size=12), text_color=TEXT_MUTED)
        self.status_lbl.pack(anchor="w")

        # Log
        self.log = ctk.CTkTextbox(self.main_frame,
                                  font=ctk.CTkFont(family="Consolas", size=12),
                                  fg_color=BG_CARD, text_color=TEXT_PRIMARY,
                                  corner_radius=12, height=130, state="disabled")
        self.log.pack(fill="both", expand=True, pady=(8,0))

        # ---- Right Side: Preview Panel (Hidden by default) ----
        self.preview_panel = ctk.CTkFrame(
            self.root_container, 
            fg_color=BG_CARD, 
            width=0, # Start hidden
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

    def _section_label(self, text):
        ctk.CTkLabel(self.main_frame, text=text,
                     font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
                     text_color=TEXT_PRIMARY).pack(anchor="w", pady=(8,4))

    # --------------------------------------------------------
    # LOGGING / STATUS
    # --------------------------------------------------------
    def _log(self, msg):
        self.log.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{ts}] {msg}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _status(self, text, color=TEXT_MUTED):
        self.status_lbl.configure(text=text, text_color=color)

    def _item_status(self, item, text, color=TEXT_MUTED):
        if item.get("status_lbl") and item["status_lbl"].winfo_exists():
            item["status_lbl"].configure(text=text, text_color=color)

    # --------------------------------------------------------
    # INTEGRATED SIDE PREVIEW SYSTEM
    # --------------------------------------------------------
    def _toggle_preview_from_row(self, pdf_path: str, sec: dict):
        try:
            start_page = int(sec["start_var"].get().strip())
            end_page = int(sec["end_var"].get().strip())
        except ValueError:
            start_page = sec["start"]
            end_page = sec["end"]
        self._toggle_preview(pdf_path, start_page, end_page)

    def _toggle_preview(self, pdf_path: str, start_page: int = None, end_page: int = None):
        preview_id = (pdf_path, start_page, end_page)
        if self.current_preview_id == preview_id:
            self._close_preview()
            return

        if not self.preview_panel.winfo_ismapped():
            self.preview_panel.pack(side="right", fill="both", expand=False)
            self.preview_panel.configure(width=340)

        self.current_preview_id = preview_id
        
        base_name = os.path.basename(pdf_path)
        if start_page is not None and end_page is not None:
            self.preview_title.configure(text=f"👁️ {base_name} (p.{start_page}-{end_page})")
        else:
            self.preview_title.configure(text=f"👁️ {base_name}")
        
        for widget in self.preview_scroll.winfo_children():
            widget.destroy()

        threading.Thread(target=self._load_preview_worker, args=(pdf_path, start_page, end_page, preview_id), daemon=True).start()

    def _close_preview(self):
        self.preview_panel.pack_forget()
        self.current_preview_id = None

    def _load_preview_worker(self, pdf_path: str, start_page: int, end_page: int, preview_id: tuple):
        try:
            doc = fitz.open(pdf_path)
            total = len(doc)
            
            start_idx = (start_page - 1) if start_page is not None else 0
            end_idx = end_page if end_page is not None else total
            
            start_idx = max(0, min(start_idx, total - 1))
            end_idx = max(start_idx + 1, min(end_idx, total))
            
            for i in range(start_idx, end_idx):
                if self.current_preview_id != preview_id:
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

    def _select_output_for_file(self, pdf_path: str):
        folder = filedialog.askdirectory(title=f"Output Folder for {os.path.basename(pdf_path)}")
        if folder:
            for item in self.file_items:
                if item["path"] == pdf_path:
                    item["output_dir"] = folder
                    display = os.path.basename(folder)
                    if not display: 
                        display = folder
                    if len(display) > 12:
                        display = display[:9] + "..."
                    item["folder_label"].configure(text=display, text_color=ACCENT_SUCCESS)
                    self._log(f"Target folder for {os.path.basename(pdf_path)}: {folder}")
                    break

    # --------------------------------------------------------
    # FILE MANAGEMENT
    # --------------------------------------------------------
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Add PDF Files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if not paths:
            return
        
        added = 0
        for p in paths:
            if p not in self.loaded_pdfs:
                # Destroy empty label if it exists
                if len(self.loaded_pdfs) == 0 and hasattr(self, "empty_lbl") and self.empty_lbl.winfo_exists():
                    self.empty_lbl.destroy()
                
                self.loaded_pdfs.append(p)
                self._add_pdf_block(p)
                added += 1
                
        if added > 0:
            self._refresh_file_count()
            self.btn_clear.pack(side="left", padx=(10,0))

    def _clear_all(self):
        self.loaded_pdfs.clear()
        
        # Safely remove all wrappers of file items
        for item in self.file_items:
            if item.get("wrapper") and item["wrapper"].winfo_exists():
                item["wrapper"].destroy()
                
        self.file_items.clear()
        self._show_empty_label()
        self._refresh_file_count()
        self.btn_clear.pack_forget()
        self._close_preview()

    def _remove_pdf(self, path):
        self.loaded_pdfs = [p for p in self.loaded_pdfs if p != path]
        
        # Destroy wrapper of removed file item
        for item in list(self.file_items):
            if item["path"] == path:
                if item.get("wrapper") and item["wrapper"].winfo_exists():
                    item["wrapper"].destroy()
                self.file_items.remove(item)
                
        self._refresh_file_count()
        if not self.loaded_pdfs:
            self._show_empty_label()
            self.btn_clear.pack_forget()
            self._close_preview()

    def _show_empty_label(self):
        if hasattr(self, "empty_lbl") and self.empty_lbl.winfo_exists():
            self.empty_lbl.destroy()
            
        self.empty_lbl = ctk.CTkLabel(
            self.list_scroll,
            text="Add merged PDFs above — the AI will split, classify, and name them.",
            font=ctk.CTkFont(size=13, slant="italic"),
            text_color=TEXT_MUTED
        )
        self.empty_lbl.pack(pady=40)

    def _refresh_file_count(self):
        n = len(self.loaded_pdfs)
        if n == 0:
            self.file_count_lbl.configure(text="No files added", text_color=TEXT_MUTED)
        else:
            names = ", ".join(os.path.basename(p) for p in self.loaded_pdfs)
            if len(names) > 80:
                names = names[:77] + "..."
            self.file_count_lbl.configure(text=f"📎 {n} file(s): {names}", text_color=ACCENT_SUCCESS)

    def _pick_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir = folder
            self.out_lbl.configure(text=f"📂 {folder}", text_color=ACCENT_SUCCESS)
            
    def _save_program_setting(self, choice):
        self.target_program = choice
        self.config["target_program"] = choice
        save_config(self.config)
        self._log(f"📋 Target program set to: {choice}")
        
        # Optionally prompt to reclassify if files are already loaded
        if self.file_items:
            if messagebox.askyesno("Re-Classify?", "You changed the target program. Do you want to re-classify the currently loaded PDFs?"):
                for item in self.file_items:
                    self._run_detect(item)

    def _save_api_key(self):
        self.ai_key = self.key_entry.get().strip()
        self.config["api_key"] = self.ai_key
        save_config(self.config)
        self._log("✨ API Key saved successfully.")
        messagebox.showinfo("Saved", "Gemini API Key saved. AI classification is ready!")

    # --------------------------------------------------------
    # LIST UI + SECTION CARDS
    # --------------------------------------------------------
    def _rebuild_list_ui(self):
        """Destroy and recreate the full file list."""
        for w in self.list_scroll.winfo_children():
            w.destroy()
        self.file_items = []

        if not self.loaded_pdfs:
            self.empty_lbl = ctk.CTkLabel(
                self.list_scroll,
                text="Add merged PDFs above — the AI will split, classify, and name them.",
                font=ctk.CTkFont(size=13, slant="italic"), text_color=TEXT_MUTED)
            self.empty_lbl.pack(pady=40)
            return

        for pdf_path in self.loaded_pdfs:
            self._add_pdf_block(pdf_path)

    def _add_pdf_block(self, pdf_path):
        """Add one PDF's row + section card to the list."""
        name = os.path.basename(pdf_path)
        try:
            pg_count = len(PdfReader(pdf_path).pages)
        except Exception:
            pg_count = "?"

        # Wrapper
        wrapper = ctk.CTkFrame(self.list_scroll, fg_color="transparent")
        wrapper.pack(fill="x", pady=(4,2))

        # --- Top row ---
        row = ctk.CTkFrame(wrapper, fg_color=BG_INPUT, corner_radius=8)
        row.pack(fill="x")

        # Remove button
        ctk.CTkButton(row, text="✕",
                      font=ctk.CTkFont(size=10),
                      fg_color="transparent", text_color=ACCENT_DANGER,
                      hover_color=BG_DARK, width=26, height=26, corner_radius=4,
                      command=lambda p=pdf_path: self._remove_pdf(p)).pack(side="left", padx=(6,0), pady=4)

        # Name + page count
        ctk.CTkLabel(row, text=f"📄 {name}",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_PRIMARY, anchor="w").pack(side="left", padx=(8,4), pady=4)
                     
        # Quick Preview Button
        preview_btn = ctk.CTkButton(
            row, text="👁️",
            font=ctk.CTkFont(size=12),
            fg_color="transparent", text_color=ACCENT_SECONDARY,
            hover_color=BG_DARK, width=30, height=26,
            corner_radius=6,
            command=lambda p=pdf_path: self._toggle_preview(p))
        preview_btn.pack(side="left", padx=(4,0))
                     
        ctk.CTkLabel(row, text=f"({pg_count} pages)",
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left", padx=(4,0))

        # Per-file folder select
        folder_btn = ctk.CTkButton(
            row, text="📂",
            font=ctk.CTkFont(size=12),
            fg_color="transparent", text_color=ACCENT_PRIMARY,
            hover_color=BG_DARK, width=28, height=26,
            corner_radius=6,
            command=lambda p=pdf_path: self._select_output_for_file(p)
        )
        folder_btn.pack(side="left", padx=(4,0))

        folder_label = ctk.CTkLabel(
            row, text="Default",
            font=ctk.CTkFont(size=10),
            text_color=TEXT_MUTED, width=80, anchor="w"
        )
        folder_label.pack(side="left", padx=(2, 8))

        # Detect button (manual trigger)
        item = {
            "path": pdf_path,
            "wrapper": wrapper,
            "sections": [],
            "section_frame": None,
            "output_dir": "",
            "folder_label": folder_label,
        }
        self.file_items.append(item)

        detect_btn = ctk.CTkButton(
            row, text="🔄 Re-Classify",
            font=ctk.CTkFont(size=11),
            fg_color="transparent", text_color=ACCENT_SECONDARY,
            hover_color=BG_DARK, width=90, height=26,
            command=lambda i=item: self._run_detect(i))
        detect_btn.pack(side="right", padx=(0,8))

        # Status indicator label
        item["status_lbl"] = ctk.CTkLabel(
            row, text="🤖 Initializing...",
            font=ctk.CTkFont(size=11), text_color=ACCENT_WARNING)
        item["status_lbl"].pack(side="right", padx=(0,6))

        # Auto-run detection
        self._run_detect(item)

    def _run_detect(self, item: dict):
        """Run detection in background thread for one file."""
        self._item_status(item, "🤖 OCR Engine Starting...", ACCENT_WARNING)
        
        # Remove old card
        if item.get("section_frame") and item["section_frame"].winfo_exists():
            item["section_frame"].destroy()
            item["section_frame"] = None
        item["sections"] = []

        def _worker():
            def cb(text, color):
                self.after(0, self._item_status, item, text, color)
                
            secs, method = auto_detect_sections(item["path"], self.ai_key, self.target_program, cb)
            self.after(0, self._on_detect_done, item, secs, method)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_detect_done(self, item: dict, secs: list, method: str):
        """Called on main thread when detection completes."""
        item["sections"] = secs
        n = len(secs)
        
        status_color = ACCENT_SUCCESS
        status_text = f"✅ {n} Docs Split"
        
        # Check if there's a flag in the title
        if any("⚠️ REVIEW" in s["title"] for s in secs):
            status_color = ACCENT_WARNING
            status_text = f"⚠️ {n} Docs (Needs Review)"
            
        self._item_status(item, status_text, status_color)
            
        self._log(f"  📄 {os.path.basename(item['path'])}: Split into {n} documents  [{method}]")
        self._build_section_card(item, method)

    def _build_section_card(self, item: dict, method: str):
        """Build/replace the review card below a file row."""
        if item.get("section_frame") and item["section_frame"].winfo_exists():
            item["section_frame"].destroy()

        card = ctk.CTkFrame(item["wrapper"], fg_color=BG_CARD,
                            corner_radius=10, border_width=1,
                            border_color=ACCENT_PRIMARY)
        card.pack(fill="x", padx=10, pady=(2,6))
        item["section_frame"] = card

        # Header
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(8,4))
        ctk.CTkLabel(hdr,
                     text=f"Analyzed via: {method}",
                     font=ctk.CTkFont(size=10), text_color=TEXT_MUTED).pack(side="left")

        # Column headers
        col = ctk.CTkFrame(card, fg_color="transparent")
        col.pack(fill="x", padx=10, pady=(0,2))
        ctk.CTkLabel(col, text="#", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=TEXT_MUTED, width=26, anchor="w").pack(side="left")
        ctk.CTkLabel(col, text="Output Filename  (click to edit)",
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=TEXT_MUTED, anchor="w").pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(col, text="Pages", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=TEXT_MUTED, width=80, anchor="center").pack(side="left")
        ctk.CTkLabel(col, text="", width=30).pack(side="right")

        # Rows frame
        rows_f = ctk.CTkFrame(card, fg_color="transparent")
        rows_f.pack(fill="x", padx=8, pady=(0,4))
        item["rows_frame"] = rows_f

        for sec in item["sections"]:
            self._add_section_row(rows_f, item, sec)

        # Footer
        foot = ctk.CTkFrame(card, fg_color="transparent")
        foot.pack(fill="x", padx=10, pady=(2,8))
        ctk.CTkButton(foot, text="＋ Add Document Split",
                      font=ctk.CTkFont(size=11),
                      fg_color="transparent", text_color=ACCENT_SUCCESS,
                      hover_color=BG_DARK, border_width=1, border_color=ACCENT_SUCCESS,
                      height=26, width=130, corner_radius=6,
                      command=lambda i=item: self._add_manual_section(i)).pack(side="left")

    def _add_section_row(self, parent, item: dict, sec: dict):
        """One editable section row."""
        
        # If it's flagged for review, highlight the background a bit
        is_review = "⚠️ REVIEW" in sec["title"]
        row_color = "#451a03" if is_review else BG_INPUT # Deep amber/orange tint for review
        
        row = ctk.CTkFrame(parent, fg_color=row_color, corner_radius=6)
        row.pack(fill="x", pady=2)

        idx = item["sections"].index(sec) + 1
        ctk.CTkLabel(row, text=str(idx), font=ctk.CTkFont(size=11),
                     text_color=TEXT_MUTED, width=26, anchor="center").pack(side="left", padx=(6,4))

        # Editable filename
        name_var = ctk.StringVar(value=sec["title"])
        
        entry = ctk.CTkEntry(
            row, textvariable=name_var,
            font=ctk.CTkFont(size=12, weight="bold" if is_review else "normal"),
            fg_color=BG_DARK, border_width=1, border_color=ACCENT_PRIMARY if not is_review else ACCENT_WARNING,
            height=28, text_color=ACCENT_WARNING if is_review else TEXT_PRIMARY,
            corner_radius=6
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0,4))
        sec["name_var"] = name_var  # extraction reads this

        # Editable Pages (Start - End)
        pg_frame = ctk.CTkFrame(row, fg_color="transparent")
        pg_frame.pack(side="left", padx=(0,4))
        
        start_var = ctk.StringVar(value=str(sec['start']))
        end_var = ctk.StringVar(value=str(sec['end']))
        
        ctk.CTkLabel(pg_frame, text="p.", font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left")
        
        start_entry = ctk.CTkEntry(pg_frame, textvariable=start_var, font=ctk.CTkFont(size=11),
                                   width=26, height=24, fg_color=BG_DARK, border_width=0, justify="center")
        start_entry.pack(side="left", padx=2)
        
        ctk.CTkLabel(pg_frame, text="–", font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(side="left", padx=2)
        
        end_entry = ctk.CTkEntry(pg_frame, textvariable=end_var, font=ctk.CTkFont(size=11),
                                 width=26, height=24, fg_color=BG_DARK, border_width=0, justify="center")
        end_entry.pack(side="left", padx=2)
        
        sec["start_var"] = start_var
        sec["end_var"] = end_var

        # Quick Preview Button for Split Row
        preview_btn = ctk.CTkButton(
            row, text="👁️",
            font=ctk.CTkFont(size=12),
            fg_color="transparent", text_color=ACCENT_SECONDARY,
            hover_color=BG_DARK, width=28, height=24,
            corner_radius=4,
            command=lambda p=item["path"], s=sec: self._toggle_preview_from_row(p, s)
        )
        preview_btn.pack(side="left", padx=(2, 4))

        # Remove
        def _rm(s=sec, r=row):
            if s in item["sections"]:
                item["sections"].remove(s)
            r.destroy()

        ctk.CTkButton(row, text="✕", font=ctk.CTkFont(size=10),
                      fg_color="transparent", text_color=ACCENT_DANGER,
                      hover_color=BG_DARK, width=28, height=28, corner_radius=4,
                      command=_rm).pack(side="right", padx=(0,4))

    def _add_manual_section(self, item: dict):
        total = item["sections"][-1]["end"] if item["sections"] else 1
        new_sec = {"title": "New Document", "start": total, "end": total, "signal": "manual"}
        item["sections"].append(new_sec)
        if item.get("rows_frame") and item["rows_frame"].winfo_exists():
            self._add_section_row(item["rows_frame"], item, new_sec)

    # --------------------------------------------------------
    # EXTRACTION
    # --------------------------------------------------------
    def _start_extraction(self):
        if self.is_processing:
            return
        if not self.file_items:
            messagebox.showwarning("No Files", "Please add at least one PDF.")
            return

        # Check all files have sections
        for item in self.file_items:
            if not item["sections"]:
                messagebox.showwarning(
                    "Not Ready",
                    f"{os.path.basename(item['path'])} has not been classified yet.\n"
                    "Wait for detection to finish or click Re-Classify.")
                return

        self.is_processing = True
        self.btn_extract.configure(state="disabled")
        self._status("Processing…", ACCENT_WARNING)
        self.progress.set(0)

        plan = [{"path": i["path"], "sections": list(i["sections"]),
                 "out_dir": i.get("output_dir", "") or self.output_dir} for i in self.file_items]

        threading.Thread(target=self._worker, args=(plan,), daemon=True).start()

    def _worker(self, plan: list):
        try:
            total_pdfs = len(plan)
            all_results = []

            for pdf_idx, item in enumerate(plan):
                pdf_path = item["path"]
                pdf_name = os.path.basename(pdf_path)
                base = os.path.splitext(pdf_name)[0]
                self._log(f"━━━ {pdf_name} ━━━")
                self.after(0, self._status, f"Reading {pdf_name}…", ACCENT_WARNING)

                try:
                    reader = PdfReader(pdf_path)
                    total_pages = len(reader.pages)
                except Exception as e:
                    self._log(f"  ❌ Cannot read: {e}")
                    continue

                out_dir = item["out_dir"] or os.path.dirname(pdf_path)
                os.makedirs(out_dir, exist_ok=True)

                sections = item["sections"]
                output_files = []
                temp_files = []

                for s_idx, sec in enumerate(sections):
                    # Get the final name from the StringVar (user may have edited it)
                    name_var = sec.get("name_var")
                    if name_var:
                        try:
                            chosen = name_var.get().strip()
                            # Automatically strip out the review warning if the user didn't change it manually
                            if "⚠️ REVIEW:" in chosen:
                                chosen = chosen.split("⚠️ REVIEW:")[1].split("[")[0].strip()
                        except Exception:
                            chosen = ""
                    else:
                        chosen = ""
                        
                    chosen = sanitize(chosen) if chosen else sanitize(sec["title"]) or f"{base}_part_{s_idx+1}"

                    out_name = f"{chosen}.pdf"
                    final_path = os.path.join(out_dir, out_name)
                    # Collision avoidance
                    c = 1
                    while os.path.exists(final_path):
                        final_path = os.path.join(out_dir, f"{chosen}_{c}.pdf")
                        c += 1

                    try:
                        start = max(1, int(sec["start_var"].get()))
                        end = min(int(sec["end_var"].get()), total_pages)
                        # Ensure logic
                        if start > end:
                            start, end = end, start
                    except ValueError:
                        start = max(1, sec["start"])
                        end = min(sec["end"], total_pages)

                    writer = PdfWriter()
                    for i in range(start - 1, end):
                        writer.add_page(reader.pages[i])

                    if self.compress_var.get():
                        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                            tmp_path = tmp.name
                            writer.write(tmp)
                        temp_files.append(tmp_path)
                        self.after(0, self._status, f"Compressing {out_name}…", ACCENT_WARNING)
                        ci = smart_compress(tmp_path, final_path)
                        kb = ci["final_size"] / 1024
                        self._log(f"  ✅ {out_name}  ({kb:.0f} KB, {ci['method']})")
                    else:
                        with open(final_path, "wb") as f:
                            writer.write(f)
                        kb = os.path.getsize(final_path) / 1024
                        self._log(f"  ✅ {out_name}  ({kb:.0f} KB)")

                    output_files.append(final_path)
                    prog = (pdf_idx + (s_idx+1)/len(sections)) / total_pdfs
                    self.after(0, self.progress.set, min(prog, 0.95))

                # Cleanup temps
                for t in temp_files:
                    try: os.unlink(t)
                    except OSError: pass

                # ZIP
                if self.zip_var.get() and output_files:
                    zip_path = os.path.join(out_dir, f"{base}_extracted_documents.zip")
                    c = 1
                    while os.path.exists(zip_path):
                        zip_path = os.path.join(out_dir, f"{base}_extracted_documents_{c}.zip")
                        c += 1
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in output_files:
                            zf.write(f, arcname=os.path.basename(f))
                    kb = os.path.getsize(zip_path) / 1024
                    self._log(f"  📦 ZIP: {os.path.basename(zip_path)}  ({kb:.0f} KB)")
                    all_results.append(zip_path)

                all_results.extend(output_files)

            self.after(0, self.progress.set, 1.0)

            if all_results:
                result_dir = os.path.dirname(all_results[0])
                self._log(f"\n🎉 Done! Saved to: {result_dir}")
                self.after(0, self._status, "✅ Complete!", ACCENT_SUCCESS)
                self.after(200, lambda: self._open_folder(result_dir))
            else:
                self._log("⚠️ No output produced.")
                self.after(0, self._status, "⚠️ No output", ACCENT_WARNING)

        except Exception as e:
            self._log(f"❌ Unexpected error: {e}")
            self.after(0, self._status, f"Error: {e}", ACCENT_DANGER)
        finally:
            self.is_processing = False
            self.after(0, self.btn_extract.configure, {"state": "normal"})

    def _open_folder(self, path):
        try:
            os.startfile(path)
        except Exception:
            pass


# ============================================================
# ENTRY POINT
# ============================================================
def main():
    app = XTRCTRSmartApp()
    app.mainloop()


if __name__ == "__main__":
    main()
