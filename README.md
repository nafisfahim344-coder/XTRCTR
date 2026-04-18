# XTRCTR — PDF Page Extractor & Compressor

XTRCTR is a powerful, modern, and standalone Windows utility for extracting specific pages or ranges from PDF files and compressing them on the fly.

## ✨ Features

- **Smart Title Detection**: Automatically detects headings from extracted pages to suggest intelligent filenames.
- **Manual Naming**: Interactive dialog lets you confirm or edit every filename during extraction.
- **High-Quality Compression**: Uses two-tier compression (Ghostscript fallback to pypdf) to keep file sizes small without sacrificing quality.
- **Page Auto-Clamping**: Never worry about out-of-range errors; XTRCTR automatically handles overflow ranges.
- **Bulk Processing**: Load multiple PDFs and process them in one go.
- **Portable**: Single `.exe` with no installation or Python required.

## 🚀 Getting Started

1. Download `XTRCTR_Portable.exe` from the [Releases](https://github.com/nafisfahim344-coder/XTRCTR/releases) page.
2. Run it—no installation needed!
3. Select your PDFs, enter the pages (e.g., `1, 3, 5-10`, or `all`), and hit **EXTRACT & COMPRESS**.

## 🛠️ Built With

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modern GUI framework.
- [pypdf](https://github.com/py-pdf/pypdf) - PDF manipulation.
- [PyInstaller](https://www.pyinstaller.org/) - Standalone EXE packaging.

## 📄 License

MIT License. See `LICENSE` for details.
