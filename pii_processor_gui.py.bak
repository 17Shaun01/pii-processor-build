#!/usr/bin/env python3
"""
==============================================================================
  Legal Document PII Anonymizer & Restorer  —  GUI Edition
  ─────────────────────────────────────────────────────────
  A desktop application for legal professionals to safely anonymize
  client documents before sending to cloud AI services, and to restore
  the original PII afterwards.

  Built with Python / Tkinter (no external GUI dependencies).
  Core engine: Microsoft Presidio + spaCy en_core_web_lg.
==============================================================================
"""

import json
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Optional, Tuple

# ── PII engine imports ────────────────────────────────────────────────────────
try:
    from presidio_analyzer import AnalyzerEngine
    from presidio_analyzer.nlp_engine import NlpEngineProvider
except ImportError:
    messagebox.showerror(
        "Missing Dependency",
        "presidio-analyzer is not installed.\n\nRun:\n  pip install presidio-analyzer presidio-anonymizer spacy\n  python -m spacy download en_core_web_lg",
    )
    sys.exit(1)

try:
    from docx import Document as DocxDocument
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

try:
    from pdfminer.high_level import extract_text as pdf_extract_text
    PDF_OK = True
except ImportError:
    PDF_OK = False


# ─────────────────────────────────────────────────────────────────────────────
#  Entity configuration
# ─────────────────────────────────────────────────────────────────────────────

ENTITY_LABELS: Dict[str, str] = {
    "PERSON":            "PERSON",
    "EMAIL_ADDRESS":     "EMAIL",
    "PHONE_NUMBER":      "PHONE",
    "LOCATION":          "LOCATION",
    "DATE_TIME":         "DATE",
    "US_SSN":            "SSN",
    "US_PASSPORT":       "PASSPORT",
    "US_DRIVER_LICENSE": "DRIVER_LICENSE",
    "US_ITIN":           "ITIN",
    "US_BANK_NUMBER":    "BANK_ACCOUNT",
    "CREDIT_CARD":       "CREDIT_CARD",
    "IBAN_CODE":         "IBAN",
    "IP_ADDRESS":        "IP_ADDRESS",
    "URL":               "URL",
    "CRYPTO":            "CRYPTO_ADDRESS",
    "MEDICAL_LICENSE":   "MEDICAL_LICENSE",
    "UK_NHS":            "NHS_NUMBER",
    "NRP":               "NATIONAL_ID",
}
ALL_ENTITIES = list(ENTITY_LABELS.keys())
DEFAULT_CONFIDENCE = 0.60

# ─────────────────────────────────────────────────────────────────────────────
#  Colour palette
# ─────────────────────────────────────────────────────────────────────────────

DARK_BG      = "#1e1e2e"
PANEL_BG     = "#2a2a3e"
ACCENT       = "#7c6af7"        # purple
ACCENT_HOVER = "#9b8df9"
SUCCESS      = "#50fa7b"
WARNING      = "#f1fa8c"
DANGER       = "#ff5555"
TEXT_MAIN    = "#cdd6f4"
TEXT_DIM     = "#6c7086"
BORDER       = "#45475a"
ENTRY_BG     = "#313244"
TAG_PERSON   = "#ff79c6"
TAG_EMAIL    = "#8be9fd"
TAG_PHONE    = "#50fa7b"
TAG_LOCATION = "#ffb86c"
TAG_DATE     = "#f1fa8c"
TAG_ID       = "#ff5555"
TAG_OTHER    = "#bd93f9"

ENTITY_COLOURS = {
    "PERSON":   TAG_PERSON,
    "EMAIL":    TAG_EMAIL,
    "PHONE":    TAG_PHONE,
    "LOCATION": TAG_LOCATION,
    "DATE":     TAG_DATE,
    "SSN":      TAG_ID,
    "PASSPORT": TAG_ID,
    "DRIVER_LICENSE": TAG_ID,
    "IBAN":     TAG_ID,
    "BANK_ACCOUNT": TAG_ID,
    "CREDIT_CARD":  TAG_ID,
}


# ─────────────────────────────────────────────────────────────────────────────
#  File I/O helpers
# ─────────────────────────────────────────────────────────────────────────────

def read_file(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read()
    elif ext == ".docx":
        if not DOCX_OK:
            raise RuntimeError("python-docx not installed. Run: pip install python-docx")
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs)
    elif ext == ".pdf":
        if not PDF_OK:
            raise RuntimeError("pdfminer.six not installed. Run: pip install pdfminer.six")
        return pdf_extract_text(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def write_file(path: str, text: str) -> None:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        if not DOCX_OK:
            raise RuntimeError("python-docx not installed.")
        doc = DocxDocument()
        for line in text.split("\n"):
            doc.add_paragraph(line)
        doc.save(path)
    else:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)


# ─────────────────────────────────────────────────────────────────────────────
#  PII engine
# ─────────────────────────────────────────────────────────────────────────────

class PIIEngine:
    _instance = None

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        provider = NlpEngineProvider(nlp_configuration={
            "nlp_engine_name": "spacy",
            "models": [{"lang_code": "en", "model_name": "en_core_web_sm"}],
        })
        self.analyzer = AnalyzerEngine(nlp_engine=provider.create_engine())

    @staticmethod
    def _resolve_overlaps(results):
        sorted_r = sorted(results, key=lambda r: (r.score, r.end - r.start), reverse=True)
        kept = []
        for cand in sorted_r:
            if not any(cand.start < a.end and cand.end > a.start for a in kept):
                kept.append(cand)
        kept.sort(key=lambda r: r.start)
        return kept

    def anonymize(self, text: str, confidence: float = DEFAULT_CONFIDENCE,
                  entities: Optional[List[str]] = None) -> Tuple[str, Dict, List]:
        if entities is None:
            entities = ALL_ENTITIES
        raw = self.analyzer.analyze(text=text, language="en",
                                    entities=entities, score_threshold=confidence)
        resolved = self._resolve_overlaps(raw)

        mapping: Dict[str, str] = {}
        entity_counts: Dict[str, int] = {}
        value_to_ph: Dict[str, str] = {}
        detections: List[dict] = []

        for result in sorted(resolved, key=lambda r: r.start, reverse=True):
            original = text[result.start:result.end]
            label = ENTITY_LABELS.get(result.entity_type, result.entity_type)
            if original in value_to_ph:
                ph = value_to_ph[original]
            else:
                entity_counts[label] = entity_counts.get(label, 0) + 1
                ph = f"{{{{{label}_{entity_counts[label]}}}}}"
                mapping[ph] = original
                value_to_ph[original] = ph
            detections.append({"placeholder": ph, "original": original,
                                "label": label, "score": result.score,
                                "start": result.start, "end": result.end})
            text = text[:result.start] + ph + text[result.end:]

        return text, mapping, detections

    @staticmethod
    def restore(text: str, mapping: Dict[str, str]) -> str:
        for ph, orig in mapping.items():
            text = text.replace(ph, orig)
        return text


# ─────────────────────────────────────────────────────────────────────────────
#  Reusable styled widgets
# ─────────────────────────────────────────────────────────────────────────────

def styled_button(parent, text, command, bg=ACCENT, fg="white",
                  width=None, pady=8):
    kw = dict(text=text, command=command, bg=bg, fg=fg,
              font=("Segoe UI", 10, "bold"), relief="flat",
              activebackground=ACCENT_HOVER, activeforeground="white",
              cursor="hand2", pady=pady, bd=0)
    if width:
        kw["width"] = width
    btn = tk.Button(parent, **kw)
    btn.bind("<Enter>", lambda e: btn.config(bg=ACCENT_HOVER))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    return btn


def section_label(parent, text):
    return tk.Label(parent, text=text, bg=PANEL_BG, fg=ACCENT,
                    font=("Segoe UI", 9, "bold"))


def dim_label(parent, text):
    return tk.Label(parent, text=text, bg=PANEL_BG, fg=TEXT_DIM,
                    font=("Segoe UI", 9))


def entry_row(parent, label_text, var, browse_cmd=None):
    row = tk.Frame(parent, bg=PANEL_BG)
    tk.Label(row, text=label_text, bg=PANEL_BG, fg=TEXT_MAIN,
             font=("Segoe UI", 9), width=18, anchor="w").pack(side="left")
    ent = tk.Entry(row, textvariable=var, bg=ENTRY_BG, fg=TEXT_MAIN,
                   insertbackground=TEXT_MAIN, relief="flat",
                   font=("Segoe UI", 9), bd=4)
    ent.pack(side="left", fill="x", expand=True)
    if browse_cmd:
        tk.Button(row, text="Browse", command=browse_cmd,
                  bg=BORDER, fg=TEXT_MAIN, font=("Segoe UI", 8),
                  relief="flat", cursor="hand2", padx=6).pack(side="left", padx=(4, 0))
    return row


# ─────────────────────────────────────────────────────────────────────────────
#  Main Application Window
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Legal Document PII Anonymizer & Restorer")
        self.geometry("1100x780")
        self.minsize(900, 640)
        self.configure(bg=DARK_BG)
        self._engine_ready = False
        self._engine_loading = False
        self._build_ui()
        self._load_engine_async()

    # ── engine loading ────────────────────────────────────────────────────────

    def _load_engine_async(self):
        self._engine_loading = True
        self._set_status("Loading NLP engine (first launch may take ~30 s)…", WARNING)
        threading.Thread(target=self._load_engine, daemon=True).start()

    def _load_engine(self):
        try:
            PIIEngine.get()
            self._engine_ready = True
            self.after(0, lambda: self._set_status("Engine ready.", SUCCESS))
        except Exception as exc:
            self.after(0, lambda: self._set_status(f"Engine error: {exc}", DANGER))

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── header ────────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=PANEL_BG, pady=14)
        header.pack(fill="x")
        tk.Label(header, text="⚖  Legal Document PII Anonymizer & Restorer",
                 bg=PANEL_BG, fg=TEXT_MAIN,
                 font=("Segoe UI", 15, "bold")).pack(side="left", padx=20)
        self._status_lbl = tk.Label(header, text="", bg=PANEL_BG,
                                    fg=WARNING, font=("Segoe UI", 9))
        self._status_lbl.pack(side="right", padx=20)

        # ── notebook tabs ─────────────────────────────────────────────────────
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook",         background=DARK_BG, borderwidth=0)
        style.configure("TNotebook.Tab",     background=PANEL_BG, foreground=TEXT_DIM,
                         font=("Segoe UI", 10, "bold"), padding=[18, 8])
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(8, 0))

        self._tab_anon   = tk.Frame(nb, bg=DARK_BG)
        self._tab_restore = tk.Frame(nb, bg=DARK_BG)
        self._tab_about  = tk.Frame(nb, bg=DARK_BG)

        nb.add(self._tab_anon,    text="  🔒  Anonymize  ")
        nb.add(self._tab_restore, text="  🔓  Restore  ")
        nb.add(self._tab_about,   text="  ℹ  About  ")

        self._build_anonymize_tab()
        self._build_restore_tab()
        self._build_about_tab()

        # ── bottom status bar ─────────────────────────────────────────────────
        bar = tk.Frame(self, bg=PANEL_BG, height=28)
        bar.pack(fill="x", side="bottom")
        self._progress = ttk.Progressbar(bar, mode="indeterminate", length=200)
        self._progress.pack(side="right", padx=12, pady=4)
        self._bar_lbl = tk.Label(bar, text="", bg=PANEL_BG, fg=TEXT_DIM,
                                 font=("Segoe UI", 8))
        self._bar_lbl.pack(side="left", padx=12)

    # ── Anonymize tab ─────────────────────────────────────────────────────────

    def _build_anonymize_tab(self):
        tab = self._tab_anon
        left = tk.Frame(tab, bg=DARK_BG, width=320)
        left.pack(side="left", fill="y", padx=(12, 6), pady=12)
        left.pack_propagate(False)
        right = tk.Frame(tab, bg=DARK_BG)
        right.pack(side="left", fill="both", expand=True, padx=(0, 12), pady=12)

        # ── left panel: settings ──────────────────────────────────────────────
        panel = tk.Frame(left, bg=PANEL_BG, bd=0, relief="flat")
        panel.pack(fill="both", expand=True, pady=(0, 8))

        section_label(panel, "INPUT / OUTPUT").pack(anchor="w", padx=14, pady=(14, 4))

        self._anon_input  = tk.StringVar()
        self._anon_output = tk.StringVar()
        self._anon_map    = tk.StringVar()

        for lbl, var, cmd in [
            ("Input document",   self._anon_input,  self._browse_input),
            ("Anonymized output", self._anon_output, self._browse_anon_out),
            ("Mapping file",     self._anon_map,    self._browse_anon_map),
        ]:
            row = entry_row(panel, lbl, var, cmd)
            row.pack(fill="x", padx=14, pady=3)

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=10)
        section_label(panel, "DETECTION SETTINGS").pack(anchor="w", padx=14, pady=(0, 6))

        conf_row = tk.Frame(panel, bg=PANEL_BG)
        conf_row.pack(fill="x", padx=14, pady=(0, 4))
        tk.Label(conf_row, text="Min. confidence", bg=PANEL_BG, fg=TEXT_MAIN,
                 font=("Segoe UI", 9), width=18, anchor="w").pack(side="left")
        self._confidence = tk.DoubleVar(value=DEFAULT_CONFIDENCE)
        conf_scale = tk.Scale(conf_row, from_=0.3, to=1.0, resolution=0.05,
                              orient="horizontal", variable=self._confidence,
                              bg=PANEL_BG, fg=TEXT_MAIN, troughcolor=ENTRY_BG,
                              highlightthickness=0, font=("Segoe UI", 8),
                              activebackground=ACCENT, length=120)
        conf_scale.pack(side="left")
        self._conf_lbl = tk.Label(conf_row, text="0.60", bg=PANEL_BG, fg=ACCENT,
                                  font=("Segoe UI", 9, "bold"), width=4)
        self._conf_lbl.pack(side="left")
        self._confidence.trace_add("write",
            lambda *_: self._conf_lbl.config(text=f"{self._confidence.get():.2f}"))

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=10)
        section_label(panel, "ENTITY TYPES TO DETECT").pack(anchor="w", padx=14, pady=(0, 6))

        self._entity_vars: Dict[str, tk.BooleanVar] = {}
        scroll_frame = tk.Frame(panel, bg=PANEL_BG)
        scroll_frame.pack(fill="x", padx=14)
        for i, (etype, label) in enumerate(ENTITY_LABELS.items()):
            var = tk.BooleanVar(value=True)
            self._entity_vars[etype] = var
            col = i % 2
            row_idx = i // 2
            cb = tk.Checkbutton(scroll_frame, text=label, variable=var,
                                bg=PANEL_BG, fg=TEXT_MAIN, selectcolor=ENTRY_BG,
                                activebackground=PANEL_BG, activeforeground=TEXT_MAIN,
                                font=("Segoe UI", 8), anchor="w")
            cb.grid(row=row_idx, column=col, sticky="w", pady=1)

        sel_row = tk.Frame(panel, bg=PANEL_BG)
        sel_row.pack(fill="x", padx=14, pady=(6, 0))
        tk.Button(sel_row, text="Select All", command=self._select_all_entities,
                  bg=BORDER, fg=TEXT_MAIN, font=("Segoe UI", 8), relief="flat",
                  cursor="hand2", padx=6).pack(side="left", padx=(0, 4))
        tk.Button(sel_row, text="Clear All", command=self._clear_all_entities,
                  bg=BORDER, fg=TEXT_MAIN, font=("Segoe UI", 8), relief="flat",
                  cursor="hand2", padx=6).pack(side="left")

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=10)
        styled_button(panel, "🔒  Run Anonymization", self._run_anonymize,
                      width=28).pack(padx=14, pady=(0, 14))

        # ── right panel: preview ──────────────────────────────────────────────
        self._build_preview_area(right, "anon")

    # ── Restore tab ───────────────────────────────────────────────────────────

    def _build_restore_tab(self):
        tab = self._tab_restore
        left = tk.Frame(tab, bg=DARK_BG, width=320)
        left.pack(side="left", fill="y", padx=(12, 6), pady=12)
        left.pack_propagate(False)
        right = tk.Frame(tab, bg=DARK_BG)
        right.pack(side="left", fill="both", expand=True, padx=(0, 12), pady=12)

        panel = tk.Frame(left, bg=PANEL_BG, bd=0)
        panel.pack(fill="both", expand=True, pady=(0, 8))

        section_label(panel, "INPUT / OUTPUT").pack(anchor="w", padx=14, pady=(14, 4))

        self._rest_input  = tk.StringVar()
        self._rest_map    = tk.StringVar()
        self._rest_output = tk.StringVar()

        for lbl, var, cmd in [
            ("Anonymized document", self._rest_input,  self._browse_rest_input),
            ("Mapping file",        self._rest_map,    self._browse_rest_map),
            ("Restored output",     self._rest_output, self._browse_rest_out),
        ]:
            row = entry_row(panel, lbl, var, cmd)
            row.pack(fill="x", padx=14, pady=3)

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=14)

        info_box = tk.Frame(panel, bg=ENTRY_BG, bd=0)
        info_box.pack(fill="x", padx=14, pady=(0, 10))
        tk.Label(info_box,
                 text="Workflow:\n\n"
                      "1. Anonymize your document\n"
                      "2. Send the anonymized text to your\n"
                      "   cloud AI (ChatGPT, Claude, etc.)\n"
                      "3. Save the AI's response to a file\n"
                      "4. Load it here as 'Anonymized document'\n"
                      "5. Click Restore — done!",
                 bg=ENTRY_BG, fg=TEXT_DIM, font=("Segoe UI", 9),
                 justify="left", padx=12, pady=10).pack()

        styled_button(panel, "🔓  Restore Original PII", self._run_restore,
                      width=28).pack(padx=14, pady=(0, 14))

        self._build_preview_area(right, "rest")

    # ── About tab ─────────────────────────────────────────────────────────────

    def _build_about_tab(self):
        tab = self._tab_about
        frame = tk.Frame(tab, bg=DARK_BG)
        frame.pack(expand=True)

        tk.Label(frame, text="⚖", bg=DARK_BG, fg=ACCENT,
                 font=("Segoe UI", 48)).pack(pady=(40, 0))
        tk.Label(frame, text="Legal Document PII Anonymizer & Restorer",
                 bg=DARK_BG, fg=TEXT_MAIN,
                 font=("Segoe UI", 16, "bold")).pack()
        tk.Label(frame, text="Protect attorney-client privilege before using cloud AI",
                 bg=DARK_BG, fg=TEXT_DIM, font=("Segoe UI", 10)).pack(pady=(4, 24))

        info = [
            ("Engine",    "Microsoft Presidio + spaCy en_core_web_lg"),
            ("Entities",  f"{len(ENTITY_LABELS)} PII types detected"),
            ("Formats",   ".txt  |  .docx  |  .pdf  (input)    .txt  |  .docx  (output)"),
            ("Privacy",   "All processing is 100% local — no data leaves your machine"),
        ]
        for key, val in info:
            row = tk.Frame(frame, bg=PANEL_BG, pady=8, padx=20)
            row.pack(fill="x", padx=60, pady=3)
            tk.Label(row, text=f"{key}:", bg=PANEL_BG, fg=ACCENT,
                     font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
            tk.Label(row, text=val, bg=PANEL_BG, fg=TEXT_MAIN,
                     font=("Segoe UI", 9)).pack(side="left")

    # ── shared preview area ───────────────────────────────────────────────────

    def _build_preview_area(self, parent, prefix):
        top = tk.Frame(parent, bg=DARK_BG)
        top.pack(fill="both", expand=True)

        # Two text boxes side by side
        left_frame = tk.Frame(top, bg=DARK_BG)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
        right_frame = tk.Frame(top, bg=DARK_BG)
        right_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))

        lbl_in  = "Original Document" if prefix == "anon" else "Anonymized Document"
        lbl_out = "Anonymized Output" if prefix == "anon" else "Restored Document"

        for frame, lbl, attr in [
            (left_frame,  lbl_in,  f"_{prefix}_txt_in"),
            (right_frame, lbl_out, f"_{prefix}_txt_out"),
        ]:
            header = tk.Frame(frame, bg=PANEL_BG)
            header.pack(fill="x")
            tk.Label(header, text=lbl, bg=PANEL_BG, fg=ACCENT,
                     font=("Segoe UI", 9, "bold"), pady=6, padx=10).pack(side="left")
            if attr.endswith("_out"):
                copy_btn = tk.Button(header, text="Copy", bg=BORDER, fg=TEXT_MAIN,
                                     font=("Segoe UI", 8), relief="flat", cursor="hand2",
                                     padx=6, command=lambda a=attr: self._copy_text(a))
                copy_btn.pack(side="right", padx=6, pady=4)
                save_btn = tk.Button(header, text="Save As…", bg=BORDER, fg=TEXT_MAIN,
                                     font=("Segoe UI", 8), relief="flat", cursor="hand2",
                                     padx=6, command=lambda a=attr: self._save_text(a))
                save_btn.pack(side="right", padx=(0, 4), pady=4)

            txt = scrolledtext.ScrolledText(
                frame, bg=ENTRY_BG, fg=TEXT_MAIN, insertbackground=TEXT_MAIN,
                font=("Consolas", 9), relief="flat", wrap="word",
                selectbackground=ACCENT, selectforeground="white",
            )
            txt.pack(fill="both", expand=True)
            setattr(self, attr, txt)

        # Configure highlight tags on output box
        out_widget = getattr(self, f"_{prefix}_txt_out")
        for label, colour in ENTITY_COLOURS.items():
            out_widget.tag_configure(label, foreground=colour, font=("Consolas", 9, "bold"))

        # Detection table below
        tbl_frame = tk.Frame(parent, bg=PANEL_BG, height=160)
        tbl_frame.pack(fill="x", pady=(8, 0))
        tbl_frame.pack_propagate(False)

        tk.Label(tbl_frame, text="Detection Log", bg=PANEL_BG, fg=ACCENT,
                 font=("Segoe UI", 9, "bold"), pady=6, padx=10).pack(anchor="w")

        cols = ("Placeholder", "Original Value", "Entity Type", "Confidence")
        tv = ttk.Treeview(tbl_frame, columns=cols, show="headings", height=4)
        style = ttk.Style()
        style.configure("Treeview",
                         background=ENTRY_BG, foreground=TEXT_MAIN,
                         fieldbackground=ENTRY_BG, rowheight=22,
                         font=("Consolas", 8))
        style.configure("Treeview.Heading",
                         background=PANEL_BG, foreground=ACCENT,
                         font=("Segoe UI", 8, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", ACCENT)])

        for col in cols:
            tv.heading(col, text=col)
            tv.column(col, width=180 if col == "Original Value" else 130, anchor="w")

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tv.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        setattr(self, f"_{prefix}_table", tv)

    # ── browse helpers ────────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select input document",
            filetypes=[("Documents", "*.txt *.docx *.pdf"), ("All files", "*.*")])
        if path:
            self._anon_input.set(path)
            base = os.path.splitext(path)[0]
            self._anon_output.set(base + "_anonymized.txt")
            self._anon_map.set(base + "_mapping.json")
            self._load_preview(path, self._anon_txt_in)

    def _browse_anon_out(self):
        path = filedialog.asksaveasfilename(
            title="Save anonymized document",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("Word document", "*.docx")])
        if path:
            self._anon_output.set(path)

    def _browse_anon_map(self):
        path = filedialog.asksaveasfilename(
            title="Save mapping file",
            defaultextension=".json",
            filetypes=[("JSON file", "*.json")])
        if path:
            self._anon_map.set(path)

    def _browse_rest_input(self):
        path = filedialog.askopenfilename(
            title="Select anonymized document",
            filetypes=[("Documents", "*.txt *.docx"), ("All files", "*.*")])
        if path:
            self._rest_input.set(path)
            base = os.path.splitext(path)[0]
            self._rest_output.set(base + "_restored.txt")
            self._load_preview(path, self._rest_txt_in)

    def _browse_rest_map(self):
        path = filedialog.askopenfilename(
            title="Select mapping file",
            filetypes=[("JSON file", "*.json"), ("All files", "*.*")])
        if path:
            self._rest_map.set(path)

    def _browse_rest_out(self):
        path = filedialog.asksaveasfilename(
            title="Save restored document",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("Word document", "*.docx")])
        if path:
            self._rest_output.set(path)

    # ── preview helpers ───────────────────────────────────────────────────────

    def _load_preview(self, path: str, widget: scrolledtext.ScrolledText):
        try:
            text = read_file(path)
            widget.config(state="normal")
            widget.delete("1.0", "end")
            widget.insert("1.0", text)
            widget.config(state="disabled")
        except Exception as exc:
            messagebox.showerror("Read Error", str(exc))

    def _set_output_text(self, widget: scrolledtext.ScrolledText,
                         text: str, detections: Optional[List[dict]] = None):
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        if detections:
            self._highlight_placeholders(widget, text, detections)
        widget.config(state="disabled")

    def _highlight_placeholders(self, widget, text: str, detections: List[dict]):
        """Colour-highlight each placeholder in the output text widget."""
        import re
        for ph_match in re.finditer(r"\{\{([A-Z_]+)_\d+\}\}", text):
            ph = ph_match.group(0)
            label = ph_match.group(1)
            tag = label if label in ENTITY_COLOURS else "OTHER"
            start_idx = f"1.0 + {ph_match.start()} chars"
            end_idx   = f"1.0 + {ph_match.end()} chars"
            widget.tag_add(tag, start_idx, end_idx)

    def _copy_text(self, attr: str):
        widget = getattr(self, attr)
        text = widget.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(text)
        self._set_status("Copied to clipboard.", SUCCESS)

    def _save_text(self, attr: str):
        widget = getattr(self, attr)
        text = widget.get("1.0", "end-1c")
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("Word document", "*.docx")])
        if path:
            try:
                write_file(path, text)
                self._set_status(f"Saved → {path}", SUCCESS)
            except Exception as exc:
                messagebox.showerror("Save Error", str(exc))

    # ── entity checkbox helpers ───────────────────────────────────────────────

    def _select_all_entities(self):
        for v in self._entity_vars.values():
            v.set(True)

    def _clear_all_entities(self):
        for v in self._entity_vars.values():
            v.set(False)

    # ── status helpers ────────────────────────────────────────────────────────

    def _set_status(self, msg: str, colour: str = TEXT_MAIN):
        self._status_lbl.config(text=msg, fg=colour)
        self._bar_lbl.config(text=msg)

    def _start_spinner(self, msg: str):
        self._bar_lbl.config(text=msg)
        self._progress.start(12)

    def _stop_spinner(self):
        self._progress.stop()

    # ── populate detection table ──────────────────────────────────────────────

    def _populate_table(self, tv: ttk.Treeview, detections: List[dict]):
        for row in tv.get_children():
            tv.delete(row)
        seen = set()
        for d in sorted(detections, key=lambda x: x["start"]):
            key = (d["placeholder"], d["original"])
            if key not in seen:
                seen.add(key)
                tv.insert("", "end", values=(
                    d["placeholder"],
                    d["original"],
                    d["label"],
                    f"{d['score']:.2f}",
                ))

    # ── run anonymize ─────────────────────────────────────────────────────────

    def _run_anonymize(self):
        if not self._engine_ready:
            messagebox.showwarning("Engine Loading",
                                   "The NLP engine is still loading. Please wait a moment.")
            return
        inp  = self._anon_input.get().strip()
        out  = self._anon_output.get().strip()
        mapf = self._anon_map.get().strip()
        if not inp:
            messagebox.showwarning("Missing Input", "Please select an input document.")
            return
        if not out:
            messagebox.showwarning("Missing Output", "Please specify an output file path.")
            return
        if not mapf:
            messagebox.showwarning("Missing Mapping", "Please specify a mapping file path.")
            return

        selected_entities = [e for e, v in self._entity_vars.items() if v.get()]
        if not selected_entities:
            messagebox.showwarning("No Entities", "Please select at least one entity type.")
            return

        confidence = self._confidence.get()
        self._start_spinner("Anonymizing document…")
        self._set_status("Anonymizing…", WARNING)

        def task():
            try:
                text = read_file(inp)
                engine = PIIEngine.get()
                anon_text, mapping, detections = engine.anonymize(
                    text, confidence=confidence, entities=selected_entities)
                write_file(out, anon_text)
                with open(mapf, "w", encoding="utf-8") as fh:
                    json.dump(mapping, fh, indent=4, ensure_ascii=False)
                self.after(0, lambda: self._on_anonymize_done(
                    anon_text, mapping, detections, out, mapf))
            except Exception as exc:
                self.after(0, lambda: self._on_error(str(exc)))

        threading.Thread(target=task, daemon=True).start()

    def _on_anonymize_done(self, anon_text, mapping, detections, out, mapf):
        self._stop_spinner()
        self._set_output_text(self._anon_txt_out, anon_text, detections)
        self._populate_table(self._anon_table, detections)
        n_unique = len(mapping)
        n_total  = len(detections)
        self._set_status(
            f"Done — {n_unique} unique PII items replaced ({n_total} total occurrences).",
            SUCCESS)
        messagebox.showinfo(
            "Anonymization Complete",
            f"Anonymization complete!\n\n"
            f"  Unique PII items replaced : {n_unique}\n"
            f"  Total occurrences         : {n_total}\n\n"
            f"Anonymized document → {out}\n"
            f"Mapping file        → {mapf}\n\n"
            f"You can now safely send the anonymized text to your cloud AI.",
        )

    # ── run restore ───────────────────────────────────────────────────────────

    def _run_restore(self):
        inp  = self._rest_input.get().strip()
        mapf = self._rest_map.get().strip()
        out  = self._rest_output.get().strip()
        if not inp:
            messagebox.showwarning("Missing Input", "Please select an anonymized document.")
            return
        if not mapf:
            messagebox.showwarning("Missing Mapping", "Please select a mapping file.")
            return
        if not out:
            messagebox.showwarning("Missing Output", "Please specify an output file path.")
            return

        self._start_spinner("Restoring PII…")
        self._set_status("Restoring…", WARNING)

        def task():
            try:
                text = read_file(inp)
                with open(mapf, "r", encoding="utf-8") as fh:
                    mapping = json.load(fh)
                restored = PIIEngine.restore(text, mapping)
                write_file(out, restored)
                self.after(0, lambda: self._on_restore_done(restored, mapping, out))
            except Exception as exc:
                self.after(0, lambda: self._on_error(str(exc)))

        threading.Thread(target=task, daemon=True).start()

    def _on_restore_done(self, restored, mapping, out):
        self._stop_spinner()
        self._set_output_text(self._rest_txt_out, restored)
        # Populate table with mapping entries
        fake_detections = [
            {"placeholder": ph, "original": orig, "label": ph.strip("{}").rsplit("_", 1)[0],
             "score": 1.0, "start": 0, "end": 0}
            for ph, orig in mapping.items()
        ]
        self._populate_table(self._rest_table, fake_detections)
        self._set_status(
            f"Done — {len(mapping)} placeholders restored.", SUCCESS)
        messagebox.showinfo(
            "Restoration Complete",
            f"Restoration complete!\n\n"
            f"  Placeholders restored : {len(mapping)}\n\n"
            f"Restored document → {out}",
        )

    # ── error handler ─────────────────────────────────────────────────────────

    def _on_error(self, msg: str):
        self._stop_spinner()
        self._set_status(f"Error: {msg}", DANGER)
        messagebox.showerror("Error", msg)


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
