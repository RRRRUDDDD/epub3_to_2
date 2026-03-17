import json
import sys
import zipfile
import re
import threading
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Final
from xml.sax.saxutils import escape as xml_escape, quoteattr as xml_quoteattr

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from lxml import etree

# Conditional import: tkinterdnd2 for drag-and-drop support
try:
    from tkinterdnd2 import TkinterDnD
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

# Conditional import: ctypes.windll is Windows-only
if sys.platform == "win32":
    from ctypes import windll, byref, sizeof, c_int


class EpubConverter:
    """Handles the core logic of converting EPUB 3 metadata and navigation to EPUB 2."""

    NAMESPACES: Final[Dict[str, str]] = {
        'n': 'urn:oasis:names:tc:opendocument:xmlns:container',
        'opf': 'http://www.idpf.org/2007/opf',
        'xhtml': 'http://www.w3.org/1999/xhtml',
        'epub': 'http://www.idpf.org/2007/ops',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'ncx': 'http://www.daisy.org/z3986/2005/ncx/'
    }

    def __init__(self) -> None:
        self._play_order: int = 1

    def _get_metadata(self, opf_xml: etree._Element, tag: str) -> str:
        """Extracts Dublin Core metadata text from OPF XML."""
        results = opf_xml.xpath(f'//dc:{tag}/text()', namespaces=self.NAMESPACES)
        return str(results[0]).strip() if results else ""

    def _parse_nav_level(self, element: etree._Element, nav_rel_dir: Path) -> List[Dict]:
        """Recursively parses EPUB 3 nav elements into a dictionary structure."""
        items = []
        for li in element.xpath('./xhtml:li', namespaces=self.NAMESPACES):
            a_tags = li.xpath('./xhtml:a', namespaces=self.NAMESPACES)
            if not a_tags:
                continue
            anchor = a_tags[0]
            text = "".join(anchor.itertext()).strip()
            href = anchor.get('href')
            src = (nav_rel_dir / href).as_posix() if nav_rel_dir.name else href
            item = {'text': text, 'src': src, 'children': []}
            sub_ols = li.xpath('./xhtml:ol', namespaces=self.NAMESPACES)
            if sub_ols:
                item['children'] = self._parse_nav_level(sub_ols[0], nav_rel_dir)
            items.append(item)
        return items

    def _build_ncx_points(self, nav_items: List[Dict]) -> str:
        """Converts internal nav structure to EPUB 2 NCX XML string."""
        xml_parts: List[str] = []
        for item in nav_items:
            p_id = f"nav_{self._play_order}"
            # #1: XML-escape text and attribute values to prevent malformed XML
            safe_text = xml_escape(item["text"])
            safe_src = xml_quoteattr(item["src"])
            xml = f'<navPoint id="{p_id}" playOrder="{self._play_order}">\n'
            xml += f'  <navLabel><text>{safe_text}</text></navLabel>\n'
            xml += f'  <content src={safe_src}/>\n'
            self._play_order += 1
            if item['children']:
                xml += self._build_ncx_points(item['children'])
            xml += '</navPoint>\n'
            xml_parts.append(xml)
        return "".join(xml_parts)

    def convert(self, input_path: Path, output_path: Path) -> None:
        """Processes a single EPUB file from version 3 to 2."""
        self._play_order = 1
        with zipfile.ZipFile(input_path, 'r') as zin:
            container_xml = etree.fromstring(zin.read('META-INF/container.xml'))
            # #5: XPath result null check with friendly error
            xpath_query = '//n:rootfile/@full-path'
            rootfiles = container_xml.xpath(xpath_query, namespaces=self.NAMESPACES)
            if not rootfiles:
                raise ValueError(
                    f"Invalid EPUB: no rootfile in container.xml ({input_path.name})"
                )
            opf_rel_path = rootfiles[0]
            opf_path = Path(opf_rel_path)
            # #7: Cache posix path to avoid repeated conversion
            opf_posix = opf_path.as_posix()
            opf_xml = etree.fromstring(zin.read(opf_posix))

            title = self._get_metadata(opf_xml, 'title') or "Untitled"
            author = self._get_metadata(opf_xml, 'creator')

            # 解析导航
            nav_points: List[Dict] = []
            nav_xpath = '//opf:item[contains(@properties, "nav")]'
            nav_meta = opf_xml.xpath(nav_xpath, namespaces=self.NAMESPACES)
            if nav_meta:
                nav_href = nav_meta[0].get('href')
                nav_full_path = (opf_path.parent / nav_href).as_posix()
                nav_xml = etree.fromstring(zin.read(nav_full_path))
                nav_node = (nav_xml.xpath('//xhtml:nav[@epub:type="toc"]',
                                          namespaces=self.NAMESPACES) or
                            nav_xml.xpath('//xhtml:nav', namespaces=self.NAMESPACES))
                if nav_node and nav_node[0].xpath('./xhtml:ol',
                                                  namespaces=self.NAMESPACES):
                    ol_node = nav_node[0].xpath('./xhtml:ol',
                                                namespaces=self.NAMESPACES)[0]
                    nav_points = self._parse_nav_level(ol_node, Path(nav_href).parent)

            # 构建 NCX (uses xml_escape internally now)
            ncx_content = self._build_ncx_points(nav_points)
            # #1: Escape title and author in NCX document header
            safe_title = xml_escape(title)
            safe_author = xml_escape(author) if author else ""
            doc_author = (f"<docAuthor><text>{safe_author}</text></docAuthor>"
                          if author else "")
            ncx_str = f'''<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
<head><meta name="dtb:uid" content="auto-gen"/><meta name="dtb:depth" content="3"/></head>
<docTitle><text>{safe_title}</text></docTitle>{doc_author}
<navMap>\n{ncx_content}</navMap></ncx>'''

            # 写入新文件
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                # #3: Write mimetype with clean ZipInfo (no extra fields, STORED)
                if 'mimetype' in zin.namelist():
                    mime_info = zipfile.ZipInfo('mimetype')
                    mime_info.compress_type = zipfile.ZIP_STORED
                    mime_info.extra = b''
                    zout.writestr(mime_info, zin.read('mimetype'))
                for item in zin.infolist():
                    if item.filename == 'mimetype':
                        continue
                    data = zin.read(item.filename)
                    if item.filename == opf_posix:
                        # #4: UTF-8 decode with error tolerance
                        content = data.decode('utf-8', errors='replace')
                        content = content.replace('version="3.0"', 'version="2.0"')
                        content = re.sub(
                            r'\s*properties="[^"]*nav[^"]*"', '', content
                        )
                        if 'toc.ncx' not in content:
                            item_ncx = ('<item id="ncx" href="toc.ncx" '
                                        'media-type="application/x-dtbncx+xml"/>')
                            content = content.replace(
                                '</manifest>', f'{item_ncx}\n</manifest>'
                            )
                            content = re.sub(
                                r'<spine([^>]*)>', r'<spine\1 toc="ncx">', content
                            )
                        # #8: Preserve original ZipInfo metadata for modified OPF
                        opf_info = zipfile.ZipInfo(
                            item.filename, date_time=item.date_time
                        )
                        opf_info.compress_type = item.compress_type
                        zout.writestr(opf_info, content.encode('utf-8'))
                    else:
                        zout.writestr(item, data)
                zout.writestr(
                    (opf_path.parent / 'toc.ncx').as_posix(),
                    ncx_str.encode('utf-8')
                )


class FluentGUI:
    """Fluent 2 styled GUI for the EPUB converter."""

    # -- Fluent 2 Design Tokens --
    BLUE: Final[str] = "#0078D4"
    BLUE_HOVER: Final[str] = "#106EBE"
    CARD_BORDER: Final[str] = "#E0E0E0"
    TEXT_PRIMARY = ("#1A1A1A", "#E0E0E0")  # (light, dark) tuple
    TEXT_SECONDARY = ("#616161", "#9E9E9E")
    # Transparent-color key for DWM Mica punch-through.
    _MICA_KEY: Final[str] = "#010101"
    _CONFIG_FILE: Final[Path] = Path.home() / ".epub3_to_2.json"

    def __init__(self, app: ctk.CTk) -> None:
        self.app = app
        self.app.title("EPUB 3 to 2 Converter")
        self.app.geometry("760x720")
        self.app.minsize(600, 520)

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.converter = EpubConverter()
        self.status_var = tk.StringVar(value="Ready")

        self._file_count: int = 0
        self._processed: int = 0
        self._failed: int = 0  # #13: Track failure count
        self._cancel_flag = threading.Event()
        self._lock = threading.Lock()  # #2: Thread synchronization

        # #15: Subfolder output option
        self._use_subfolder = tk.BooleanVar(value=False)

        self.log_area: ctk.CTkTextbox = None  # type: ignore[assignment]
        self.run_btn: ctk.CTkButton = None  # type: ignore[assignment]
        self.cancel_btn: ctk.CTkButton = None  # type: ignore[assignment]
        self.progress_bar: ctk.CTkProgressBar = None  # type: ignore[assignment]
        self.progress_label: ctk.CTkLabel = None  # type: ignore[assignment]
        self.file_count_label: ctk.CTkLabel = None  # type: ignore[assignment]
        self.output_entry: ctk.CTkEntry = None  # type: ignore[assignment]
        self.output_browse_btn: ctk.CTkButton = None  # type: ignore[assignment]

        # #20: Track appearance mode for Mica updates
        self._last_appearance_mode: str = ctk.get_appearance_mode().lower()

        self._load_config()
        self._build_ui()
        self._apply_mica()
        self._bind_shortcuts()  # #19: Keyboard shortcuts
        self._start_theme_watcher()  # #20: Theme change detection
        # Sync subfolder UI state from loaded config
        if self._use_subfolder.get():
            self.output_entry.configure(state="disabled")
            self.output_browse_btn.configure(state="disabled")
        # Show file count if source dir was restored from config
        if self.input_dir.get():
            self._update_file_count()

    # ------------------------------------------------------------------ Config
    def _load_config(self) -> None:
        """Load last-used directories from config file."""
        try:
            data = json.loads(self._CONFIG_FILE.read_text(encoding="utf-8"))
            if data.get("input_dir"):
                self.input_dir.set(data["input_dir"])
            if data.get("output_dir"):
                self.output_dir.set(data["output_dir"])
            if "use_subfolder" in data:
                self._use_subfolder.set(data["use_subfolder"])
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            pass

    def _save_config(self) -> None:
        """Persist current directories to config file."""
        data = {
            "input_dir": self.input_dir.get(),
            "output_dir": self.output_dir.get(),
            "use_subfolder": self._use_subfolder.get(),
        }
        try:
            self._CONFIG_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except OSError:
            pass

    # ------------------------------------------------------------------ Mica
    def _apply_mica(self) -> None:
        """Apply Windows 11 Mica backdrop via DWM API (requires 22H2+, build 22621)."""
        if sys.platform != "win32":
            return
        try:
            build = int(sys.getwindowsversion().build)  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            return
        if build < 22621:
            return

        try:
            self.app.update_idletasks()
            hwnd = windll.user32.GetParent(self.app.winfo_id())

            # DWMWA_USE_IMMERSIVE_DARK_MODE (20) — sync with system theme
            is_dark = ctk.get_appearance_mode().lower() == "dark"
            windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, byref(c_int(int(is_dark))), sizeof(c_int)
            )

            # DWMWA_SYSTEMBACKDROP_TYPE (38) — 2 = Mica, 3 = Acrylic, 4 = Mica Alt
            windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 38, byref(c_int(2)), sizeof(c_int)
            )

            # Slight alpha nudge to trigger DWM composition
            self.app.attributes("-alpha", 0.992)

            # Punch-through: mark a key color as transparent so Mica shows behind
            self.app.wm_attributes("-transparentcolor", self._MICA_KEY)
            self.app.configure(fg_color=self._MICA_KEY)
        except (OSError, AttributeError, tk.TclError):
            pass  # Graceful fallback: DWM unavailable or blocked by policy

    # #20: Theme change detection
    def _start_theme_watcher(self) -> None:
        """Periodically check for appearance mode changes and update Mica."""
        current = ctk.get_appearance_mode().lower()
        if current != self._last_appearance_mode:
            self._last_appearance_mode = current
            self._apply_mica()
        self.app.after(2000, self._start_theme_watcher)

    # ------------------------------------------------------------------ Fonts
    @staticmethod
    def _resolve_font(preferred: str, fallback: str) -> str:
        """Return preferred font if available, otherwise fallback."""
        try:
            available = tk.font.families()  # type: ignore[attr-defined]
            return preferred if preferred in available else fallback
        except (AttributeError, tk.TclError):
            return fallback

    def _font(self, size: int = 13, weight: str = "normal"):
        """Returns a Fluent-aware font tuple with fallback."""
        if not hasattr(self, "_cached_font_family"):
            self._cached_font_family = self._resolve_font(
                "Segoe UI Variable", "Segoe UI"
            )
        return (self._cached_font_family, size, weight)

    def _mono_font(self, size: int = 12):
        """Returns a monospace font tuple with fallback."""
        if not hasattr(self, "_cached_mono_family"):
            self._cached_mono_family = self._resolve_font(
                "Cascadia Code", "Consolas"
            )
        return (self._cached_mono_family, size)

    # ------------------------------------------------------------------ Shortcuts
    # #19: Keyboard shortcuts
    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts for common actions."""
        self.app.bind(
            '<Return>',
            lambda _: (self._start()
                        if str(self.run_btn.cget("state")) != "disabled" else None)
        )
        self.app.bind(
            '<Escape>',
            lambda _: (self._cancel()
                        if str(self.cancel_btn.cget("state")) != "disabled" else None)
        )

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        """Constructs the full Fluent 2 interface."""
        self.app.grid_columnconfigure(0, weight=1)
        self.app.grid_rowconfigure(2, weight=1)

        # -- Title area --
        title_frame = ctk.CTkFrame(self.app, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="ew", padx=32, pady=(28, 4))

        ctk.CTkLabel(
            title_frame, text="EPUB 3 \u279c 2 Converter",
            font=self._font(22, "bold"), text_color=self.TEXT_PRIMARY,
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_frame,
            text="Batch downgrade EPUB 3 navigation & metadata for legacy readers",
            font=self._font(12), text_color=self.TEXT_SECONDARY,
        ).pack(anchor="w", pady=(2, 0))

        # -- Path settings card --
        path_card = ctk.CTkFrame(self.app, corner_radius=8, border_width=1,
                                 border_color=self.CARD_BORDER)
        path_card.grid(row=1, column=0, sticky="ew", padx=32, pady=(16, 0))
        path_card.grid_columnconfigure(1, weight=1)

        self._add_path_row(path_card, 0, "Source directory", self.input_dir)
        self._add_path_row(path_card, 1, "Output directory", self.output_dir,
                           is_output=True)

        # #15: Subfolder output checkbox
        ctk.CTkCheckBox(
            path_card, text="Output to source subdirectory (epub2_output/)",
            font=self._font(11), variable=self._use_subfolder,
            command=self._on_subfolder_toggle,
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=16, pady=(8, 4))

        # File count hint (shown after source directory is selected)
        self.file_count_label = ctk.CTkLabel(
            path_card, text="", font=self._font(11),
            text_color=self.TEXT_SECONDARY,
        )
        self.file_count_label.grid(
            row=5, column=0, columnspan=3, sticky="w", padx=16, pady=(4, 12)
        )

        # -- Log card --
        log_card = ctk.CTkFrame(self.app, corner_radius=8, border_width=1,
                                border_color=self.CARD_BORDER)
        log_card.grid(row=2, column=0, sticky="nsew", padx=32, pady=(12, 0))
        log_card.grid_rowconfigure(1, weight=1)
        log_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            log_card, text="Conversion Log",
            font=self._font(12, "bold"), text_color=self.TEXT_SECONDARY,
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))

        self.log_area = ctk.CTkTextbox(
            log_card, font=self._mono_font(), corner_radius=6,
            state="disabled", wrap="word", height=200,
        )
        self.log_area.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        # -- Progress + action area --
        action_frame = ctk.CTkFrame(self.app, fg_color="transparent")
        action_frame.grid(row=3, column=0, sticky="ew", padx=32, pady=(12, 0))
        action_frame.grid_columnconfigure(0, weight=1)

        progress_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        progress_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        progress_row.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(
            progress_row, height=6, corner_radius=3,
            progress_color=self.BLUE,
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_row, text="", font=self._font(11),
            text_color=self.TEXT_SECONDARY,
        )
        self.progress_label.grid(row=0, column=1, padx=(10, 0))

        # Buttons row
        btn_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_row.grid(row=1, column=0, pady=(0, 4))

        self.run_btn = ctk.CTkButton(
            btn_row, text="Start Batch Conversion",
            font=self._font(14, "bold"), height=44, corner_radius=4,
            fg_color=self.BLUE, hover_color=self.BLUE_HOVER,
            command=self._start,
        )
        self.run_btn.pack(side="left", padx=(0, 8))

        self.cancel_btn = ctk.CTkButton(
            btn_row, text="Cancel", width=90,
            font=self._font(13), height=44, corner_radius=4,
            fg_color=("gray86", "gray28"), hover_color=("gray78", "gray36"),
            text_color=self.TEXT_PRIMARY,
            command=self._cancel, state="disabled",
        )
        self.cancel_btn.pack(side="left")

        # -- Status bar --
        ctk.CTkLabel(
            self.app, textvariable=self.status_var,
            font=self._font(11), text_color=self.TEXT_SECONDARY,
            anchor="w",
        ).grid(row=4, column=0, sticky="ew", padx=36, pady=(8, 14))

    def _add_path_row(self, parent, row: int, label: str, var: tk.StringVar,
                      *, is_output: bool = False) -> None:
        """Creates a labeled path entry + browse button inside a card."""
        top_pad = (16, 4) if row == 0 else (8, 4)
        bot_pad = (0, 0)

        ctk.CTkLabel(
            parent, text=label, font=self._font(12), text_color=self.TEXT_SECONDARY,
        ).grid(row=row * 2, column=0, columnspan=3, sticky="w", padx=16, pady=top_pad)

        entry = ctk.CTkEntry(
            parent, textvariable=var, font=self._font(12),
            height=36, corner_radius=4,
        )
        entry.grid(row=row * 2 + 1, column=0, columnspan=2, sticky="ew",
                   padx=(16, 8), pady=bot_pad)

        btn = ctk.CTkButton(
            parent, text="Browse", width=80, height=36, corner_radius=4,
            fg_color=("gray86", "gray28"), hover_color=("gray78", "gray36"),
            text_color=self.TEXT_PRIMARY, font=self._font(12),
            command=lambda: self._browse(var),
        )
        btn.grid(row=row * 2 + 1, column=2, sticky="e", padx=(0, 16), pady=bot_pad)

        # Store output row widgets for subfolder toggle
        if is_output:
            self.output_entry = entry
            self.output_browse_btn = btn

        # #12: Drag-and-drop support
        if _HAS_DND:
            self._setup_dnd(entry, var)

    # ------------------------------------------------------------------ Drag & Drop
    # #12: Drag-and-drop support via tkinterdnd2
    def _setup_dnd(self, widget, var: tk.StringVar) -> None:
        """Register drag-and-drop for directory paths on a widget."""
        try:
            # Access the underlying tkinter entry for CTkEntry
            target = getattr(widget, '_entry', widget)
            # Register via Tcl commands (tkdnd loaded by DnDCTk root)
            target.tk.call('tkdnd::drop_target', 'register', target._w, 'DND_Files')

            def on_drop(data):
                path = data.strip('{}')
                dropped = Path(path)
                # Accept directories directly; for files, use parent directory
                target_dir = dropped if dropped.is_dir() else dropped.parent
                var.set(str(target_dir))
                self._save_config()
                if var is self.input_dir:
                    self._update_file_count()
                    if self._use_subfolder.get():
                        self.output_dir.set(str(target_dir / "epub2_output"))
                return ''

            tcl_cmd = target.register(on_drop)
            target.tk.call('tkdnd::bind', target._w, '<<Drop>>', f'{tcl_cmd} %D')
        except (AttributeError, tk.TclError):
            pass  # Graceful degradation if DnD unavailable

    # ------------------------------------------------------------------ Subfolder
    # #15: Output subfolder toggle
    def _on_subfolder_toggle(self) -> None:
        """Toggle between manual output directory and automatic subfolder mode."""
        if self._use_subfolder.get():
            src = self.input_dir.get()
            if src:
                self.output_dir.set(str(Path(src) / "epub2_output"))
            self.output_entry.configure(state="disabled")
            self.output_browse_btn.configure(state="disabled")
        else:
            self.output_entry.configure(state="normal")
            self.output_browse_btn.configure(state="normal")
        self._save_config()

    # ------------------------------------------------------------------ Dialogs
    def _browse(self, var: tk.StringVar) -> None:
        """Opens directory dialog and updates the given StringVar."""
        path = filedialog.askdirectory()
        if path:
            var.set(path)
            self._save_config()
            # Update file count when source dir changes
            if var is self.input_dir:
                self._update_file_count()
                # #15: Auto-update subfolder output path
                if self._use_subfolder.get():
                    self.output_dir.set(str(Path(path) / "epub2_output"))

    def _update_file_count(self) -> None:
        """Scans source directory and shows .epub file count."""
        src = self.input_dir.get()
        if src and Path(src).is_dir():
            # #9: Use generator to avoid building full list in memory
            count = sum(1 for _ in Path(src).glob("*.epub"))
            self.file_count_label.configure(
                text=f"{count} .epub file(s) found" if count else "No .epub files found"
            )
        else:
            self.file_count_label.configure(text="")

    # ------------------------------------------------------------------ Log
    def _log(self, msg: str) -> None:
        """Appends a timestamped message to the log textbox (thread-safe via after)."""
        # #16: Add timestamp to log entries
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {msg}\n"
        def _append():
            self.log_area.configure(state="normal")
            self.log_area.insert("end", formatted)
            self.log_area.see("end")
            self.log_area.configure(state="disabled")
        self.app.after(0, _append)

    # ------------------------------------------------------------------ Run
    def _start(self) -> None:
        """Validates inputs and launches conversion in a worker thread."""
        if not self.input_dir.get() or not self.output_dir.get():
            messagebox.showwarning("Notice", "Please select both directories first.")
            return

        self._cancel_flag.clear()
        self.run_btn.configure(state="disabled", text="Converting...")
        self.cancel_btn.configure(state="normal")
        # #14: Set progress bar to indeterminate mode during file scanning
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        self.progress_label.configure(text="Scanning...")
        self.status_var.set("Converting...")
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self._save_config()
        threading.Thread(target=self._run, daemon=True).start()

    def _cancel(self) -> None:
        """Signals the worker thread to stop after the current file."""
        self._cancel_flag.set()
        self.cancel_btn.configure(state="disabled")
        self.status_var.set("Cancelling...")

    def _run(self) -> None:
        """Background worker: iterates over EPUB files and converts them."""
        in_p = Path(self.input_dir.get())
        out_p = Path(self.output_dir.get())
        out_p.mkdir(parents=True, exist_ok=True)
        files = list(in_p.glob("*.epub"))

        # #2: Thread-safe state update
        with self._lock:
            self._file_count = len(files)
            self._processed = 0
            self._failed = 0

        # #14: Switch from indeterminate to determinate mode
        self.app.after(0, self._switch_to_determinate)

        if not files:
            self._log("[!] No .epub files found in source directory")
        else:
            for epub_file in files:
                if self._cancel_flag.is_set():
                    self._log("[!] Cancelled by user.")
                    break
                try:
                    self.converter.convert(epub_file, out_p / epub_file.name)
                    self._log(f"  OK   {epub_file.name}")
                except Exception as err:  # pylint: disable=broad-except
                    self._log(f"  FAIL {epub_file.name}  ->  {err}")
                    with self._lock:
                        self._failed += 1
                with self._lock:
                    self._processed += 1
                self._update_progress()
            else:
                # #13: Summary with success/fail counts
                with self._lock:
                    ok = self._processed - self._failed
                    fail = self._failed
                self._log(f"\nAll tasks completed. ({ok} succeeded, {fail} failed)")

        self.app.after(0, self._on_finish)

    # #14: Switch progress bar mode
    def _switch_to_determinate(self) -> None:
        """Switch progress bar from indeterminate to determinate mode."""
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(0)

    def _update_progress(self) -> None:
        """Updates progress bar and label from worker thread."""
        with self._lock:
            if self._file_count == 0:
                return
            ratio = self._processed / self._file_count
            text = f"{self._processed}/{self._file_count}"
        # #18: Update window title with progress
        self.app.after(0, lambda r=ratio, t=text: self._set_progress(r, t))

    def _set_progress(self, ratio: float, text: str) -> None:
        """Applies progress values on the main thread."""
        self.progress_bar.set(ratio)
        self.progress_label.configure(text=text)
        # #18: Show progress in window title
        self.app.title(f"Converting... ({text}) - EPUB 3 to 2 Converter")

    def _on_finish(self) -> None:
        """Restores UI state after conversion completes."""
        cancelled = self._cancel_flag.is_set()
        self.run_btn.configure(state="normal", text="Start Batch Conversion")
        self.cancel_btn.configure(state="disabled")
        # #18: Restore window title
        self.app.title("EPUB 3 to 2 Converter")

        with self._lock:
            ok = self._processed - self._failed
            fail = self._failed
            total = self._file_count

        if cancelled:
            # #17: Cancelled summary with completion details
            self.status_var.set(f"Cancelled: {ok}/{total} completed, {fail} failed")
        else:
            self.status_var.set(
                f"Done: {ok} succeeded, {fail} failed" if fail else "Done"
            )
            self.progress_bar.set(1.0)
            # #13: Include failure info in completion dialog
            if fail:
                messagebox.showwarning(
                    "Complete",
                    f"Conversion finished.\n{ok} succeeded, {fail} failed.\n"
                    "Check the log for details."
                )
            else:
                messagebox.showinfo("Complete", "All conversion tasks finished.")


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass

    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")

    import tkinter.font  # noqa: E402 — needed for font detection

    # #12: Use TkinterDnD-enabled root if tkinterdnd2 is available
    if _HAS_DND:
        try:
            class _DnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
                """CTk root with TkinterDnD drag-and-drop support."""
                def __init__(self, *args, **kwargs):
                    super().__init__(*args, **kwargs)
                    self.TkdndVersion = TkinterDnD._require(self)
            root = _DnDCTk()
        except Exception:  # pylint: disable=broad-except
            root = ctk.CTk()
    else:
        root = ctk.CTk()

    FluentGUI(root)
    root.mainloop()
