"""
==============================================================================
  Legal Document PII Anonymizer & Restorer  —  GUI Edition  (v3.0)
  -----------------------------------------------------------------
  Full Hebrew + English support.

  Hebrew capabilities:
    * xx_ent_wiki_sm  — multilingual spaCy NER (PERSON / LOC / ORG in Hebrew)
    * HebrewNameRecognizer  — dictionary of 700+ Israeli first/last names
    * HebrewLocationRecognizer — 150+ Israeli cities, regions, legal terms
    * Israeli ID (Teudat Zehut) with Luhn-10 checksum
    * Israeli phone numbers (+972 / 05x / 07x / 0x formats)
    * Israeli IBAN (IL prefix)
    * Hebrew date patterns (DD/MM/YYYY, DD.MM.YYYY)
    * Auto language detection per document
    * RTL-aware text display

  Core engine: Microsoft Presidio + spaCy (en + xx multilingual)
==============================================================================
"""

import json
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Optional, Tuple

try:
    from presidio_analyzer import (
        AnalyzerEngine, RecognizerRegistry,
        PatternRecognizer, Pattern, EntityRecognizer, RecognizerResult,
    )
    from presidio_analyzer.nlp_engine import NlpEngineProvider
except ImportError:
    messagebox.showerror(
        "Missing Dependency",
        "presidio-analyzer is not installed.\n\nRun:\n  pip install presidio-analyzer presidio-anonymizer spacy",
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


# ---------------------------------------------------------------------------
#  Hebrew name / location dictionaries  (imported from hebrew_data.py)
# ---------------------------------------------------------------------------

try:
    from hebrew_data import HEBREW_FIRST_NAMES, HEBREW_LAST_NAMES, HEBREW_LOCATIONS
except ImportError:
    # Fallback: empty sets (pattern recognizers still work)
    HEBREW_FIRST_NAMES = set()
    HEBREW_LAST_NAMES = set()
    HEBREW_LOCATIONS = set()


# ---------------------------------------------------------------------------
#  Language detection
# ---------------------------------------------------------------------------

HEBREW_CHARS = re.compile(r'[\u05d0-\u05ea\ufb1d-\ufb4e]')

def detect_language(text: str) -> str:
    total = len(text.replace(" ", "").replace("\n", ""))
    if total == 0:
        return "en"
    hebrew_count = len(HEBREW_CHARS.findall(text))
    return "he" if (hebrew_count / total) > 0.15 else "en"

def is_rtl(text: str) -> bool:
    return detect_language(text) == "he"


# ---------------------------------------------------------------------------
#  Entity configuration
# ---------------------------------------------------------------------------

ENTITY_LABELS: Dict[str, str] = {
    "PERSON":              "PERSON",
    "EMAIL_ADDRESS":       "EMAIL",
    "PHONE_NUMBER":        "PHONE",
    "LOCATION":            "LOCATION",
    "DATE_TIME":           "DATE",
    "IL_ID_NUMBER":        "IL_ID",
    "IL_PHONE":            "IL_PHONE",
    "HE_PERSON":           "PERSON",
    "HE_LOCATION":         "LOCATION",
    "US_SSN":              "SSN",
    "US_PASSPORT":         "PASSPORT",
    "US_DRIVER_LICENSE":   "DRIVER_LICENSE",
    "US_ITIN":             "ITIN",
    "US_BANK_NUMBER":      "BANK_ACCOUNT",
    "CREDIT_CARD":         "CREDIT_CARD",
    "IBAN_CODE":           "IBAN",
    "IP_ADDRESS":          "IP_ADDRESS",
    "URL":                 "URL",
    "CRYPTO":              "CRYPTO_ADDRESS",
    "MEDICAL_LICENSE":     "MEDICAL_LICENSE",
    "NRP":                 "NATIONAL_ID",
}
ALL_ENTITIES = list(ENTITY_LABELS.keys())
DEFAULT_CONFIDENCE = 0.55

HE_ENTITY_NAMES: Dict[str, str] = {
    "PERSON":           "\u05e9\u05dd \u05d0\u05d3\u05dd",
    "EMAIL":            '\u05d3\u05d5\u05d0"\u05dc',
    "PHONE":            "\u05d8\u05dc\u05e4\u05d5\u05df",
    "IL_PHONE":         "\u05d8\u05dc\u05e4\u05d5\u05df \u05d9\u05e9\u05e8\u05d0\u05dc\u05d9",
    "LOCATION":         "\u05de\u05d9\u05e7\u05d5\u05dd",
    "DATE":             "\u05ea\u05d0\u05e8\u05d9\u05da",
    "IL_ID":            "\u05ea.\u05d6.",
    "SSN":              "SSN",
    "PASSPORT":         "\u05d3\u05e8\u05db\u05d5\u05df",
    "DRIVER_LICENSE":   "\u05e8\u05d9\u05e9\u05d9\u05d5\u05df \u05e0\u05d4\u05d9\u05d2\u05d4",
    "IBAN":             "IBAN",
    "BANK_ACCOUNT":     "\u05d7\u05e9\u05d1\u05d5\u05df \u05d1\u05e0\u05e7",
    "CREDIT_CARD":      "\u05db\u05e8\u05d8\u05d9\u05e1 \u05d0\u05e9\u05e8\u05d0\u05d9",
    "NATIONAL_ID":      "\u05de\u05e1\u05e4\u05e8 \u05d6\u05d4\u05d5\u05ea",
}


# ---------------------------------------------------------------------------
#  Colour palette
# ---------------------------------------------------------------------------

DARK_BG      = "#1e1e2e"
PANEL_BG     = "#2a2a3e"
ACCENT       = "#7c6af7"
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
    "PERSON":         TAG_PERSON,
    "EMAIL":          TAG_EMAIL,
    "PHONE":          TAG_PHONE,
    "IL_PHONE":       TAG_PHONE,
    "LOCATION":       TAG_LOCATION,
    "DATE":           TAG_DATE,
    "IL_ID":          TAG_ID,
    "SSN":            TAG_ID,
    "PASSPORT":       TAG_ID,
    "DRIVER_LICENSE": TAG_ID,
    "IBAN":           TAG_ID,
    "BANK_ACCOUNT":   TAG_ID,
    "CREDIT_CARD":    TAG_ID,
    "NATIONAL_ID":    TAG_ID,
}


# ---------------------------------------------------------------------------
#  Israeli-specific PII recognizers
# ---------------------------------------------------------------------------

def _luhn10_israeli_id(id_str: str) -> bool:
    id_str = id_str.zfill(9)
    if len(id_str) != 9 or not id_str.isdigit():
        return False
    total = 0
    for i, ch in enumerate(id_str):
        d = int(ch) * ((i % 2) + 1)
        total += d - 9 if d > 9 else d
    return total % 10 == 0


class IsraeliIdRecognizer(PatternRecognizer):
    PATTERNS = [
        Pattern("Israeli ID (plain)",   r"\b\d{9}\b",             0.6),
        Pattern("Israeli ID (dashes)",  r"\b\d{3}-\d{3}-\d{3}\b", 0.7),
    ]
    CONTEXT = [
        "id", "identity", "id number", "identification",
        "teudat", "zehut", "t.z", "tz",
        "\u05ea.\u05d6", "\u05ea\u05e2\u05d5\u05d3\u05ea \u05d6\u05d4\u05d5\u05ea",
        "\u05de\u05e1\u05e4\u05e8 \u05d6\u05d4\u05d5\u05ea", "\u05d6\u05d4\u05d5\u05ea",
    ]

    def __init__(self, lang="he"):
        super().__init__(
            supported_entity="IL_ID_NUMBER",
            patterns=self.PATTERNS,
            context=self.CONTEXT,
            supported_language=lang,
        )

    def validate_result(self, pattern_text: str):
        digits = re.sub(r"[^0-9]", "", pattern_text)
        return _luhn10_israeli_id(digits)


class IsraeliPhoneRecognizer(PatternRecognizer):
    PATTERNS = [
        Pattern("IL Phone (+972)",   r"\+972[-\s]?\d{1,2}[-\s]?\d{3}[-\s]?\d{4}", 0.85),
        Pattern("IL Mobile (05x)",   r"\b05[0-9][-\s]?\d{7}\b",                    0.80),
        Pattern("IL Mobile (07x)",   r"\b07[2-9][-\s]?\d{7}\b",                    0.80),
        Pattern("IL Landline (0x)",  r"\b0[2-9][-\s]?\d{7}\b",                     0.65),
    ]
    CONTEXT = [
        "phone", "tel", "mobile", "cell", "fax", "telephone",
        "\u05d8\u05dc\u05e4\u05d5\u05df", "\u05e0\u05d9\u05d9\u05d3",
        "\u05e4\u05e7\u05e1", "\u05de\u05e1\u05e4\u05e8 \u05d8\u05dc\u05e4\u05d5\u05df",
        "\u05e1\u05dc\u05d5\u05dc\u05e8\u05d9",
    ]

    def __init__(self, lang="he"):
        super().__init__(
            supported_entity="IL_PHONE",
            patterns=self.PATTERNS,
            context=self.CONTEXT,
            supported_language=lang,
        )


class HebrewDateRecognizer(PatternRecognizer):
    PATTERNS = [
        Pattern("HE Date (slash)",  r"\b\d{1,2}/\d{1,2}/\d{4}\b",  0.75),
        Pattern("HE Date (dot)",    r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", 0.75),
    ]
    CONTEXT = [
        "\u05ea\u05d0\u05e8\u05d9\u05da", "\u05e0\u05d5\u05dc\u05d3",
        "\u05dc\u05d9\u05d3\u05d4", "date", "born", "birth",
        "\u05e0\u05d7\u05ea\u05dd", "\u05de\u05d9\u05d5\u05dd", "\u05d1\u05d9\u05d5\u05dd",
        "\u05ea\u05d0\u05e8\u05d9\u05da \u05dc\u05d9\u05d3\u05d4",
    ]

    def __init__(self, lang="he"):
        super().__init__(
            supported_entity="DATE_TIME",
            patterns=self.PATTERNS,
            context=self.CONTEXT,
            supported_language=lang,
        )


class HebrewNameRecognizer(EntityRecognizer):
    """Dictionary-based recognizer for Hebrew-script person names."""
    SUPPORTED_ENTITY = "HE_PERSON"

    def __init__(self):
        super().__init__(
            supported_entities=[self.SUPPORTED_ENTITY],
            supported_language="he",
            name="HebrewNameRecognizer",
        )
        self._first = HEBREW_FIRST_NAMES
        self._last  = HEBREW_LAST_NAMES
        self._all   = HEBREW_FIRST_NAMES | HEBREW_LAST_NAMES

    def load(self):
        pass

    def analyze(self, text: str, entities: List[str], nlp_artifacts=None) -> List[RecognizerResult]:
        results = []
        tokens = re.finditer(r'[\u05d0-\u05ea\u05f0-\u05f4\ufb1d-\ufb4e]+', text)
        token_list = [(m.group(), m.start(), m.end()) for m in tokens]
        i = 0
        while i < len(token_list):
            word, start, end = token_list[i]
            if i + 1 < len(token_list):
                next_word, next_start, next_end = token_list[i + 1]
                if (word in self._first and next_word in self._last) or \
                   (word in self._last  and next_word in self._first) or \
                   (word in self._first and next_word in self._first):
                    results.append(RecognizerResult(
                        entity_type=self.SUPPORTED_ENTITY,
                        start=start, end=next_end, score=0.82,
                    ))
                    i += 2
                    continue
            if word in self._all:
                score = 0.75 if word in self._first else 0.70
                results.append(RecognizerResult(
                    entity_type=self.SUPPORTED_ENTITY,
                    start=start, end=end, score=score,
                ))
            i += 1
        return results


class HebrewLocationRecognizer(EntityRecognizer):
    """Dictionary-based recognizer for Hebrew-script locations."""
    SUPPORTED_ENTITY = "HE_LOCATION"

    def __init__(self):
        super().__init__(
            supported_entities=[self.SUPPORTED_ENTITY],
            supported_language="he",
            name="HebrewLocationRecognizer",
        )
        self._locations = sorted(HEBREW_LOCATIONS, key=len, reverse=True)

    def load(self):
        pass

    def analyze(self, text: str, entities: List[str], nlp_artifacts=None) -> List[RecognizerResult]:
        results = []
        for loc in self._locations:
            for m in re.finditer(re.escape(loc), text):
                results.append(RecognizerResult(
                    entity_type=self.SUPPORTED_ENTITY,
                    start=m.start(), end=m.end(), score=0.78,
                ))
        return results


# ---------------------------------------------------------------------------
#  File I/O helpers
# ---------------------------------------------------------------------------

def read_file(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        for enc in ("utf-8", "utf-8-sig", "windows-1255", "iso-8859-8", "cp1255"):
            try:
                with open(path, "r", encoding=enc) as fh:
                    return fh.read()
            except (UnicodeDecodeError, LookupError):
                continue
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read()
    elif ext == ".docx":
        if not DOCX_OK:
            raise RuntimeError("python-docx not installed.")
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs)
    elif ext == ".pdf":
        if not PDF_OK:
            raise RuntimeError("pdfminer.six not installed.")
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


# ---------------------------------------------------------------------------
#  PII engine
# ---------------------------------------------------------------------------

class PIIEngine:
    _instance = None

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self._en_analyzer = self._build_en_analyzer()
        self._he_analyzer = self._build_he_analyzer()
        self._xx_available = False

    def _build_en_analyzer(self) -> AnalyzerEngine:
        provider = NlpEngineProvider(nlp_configuration={
            "nlp_engine_name": "spacy",
            "models": [{"lang_code": "en", "model_name": "en_core_web_sm"}],
        })
        nlp_engine = provider.create_engine()
        registry = RecognizerRegistry()
        registry.load_predefined_recognizers(nlp_engine=nlp_engine)
        registry.add_recognizer(IsraeliIdRecognizer(lang="en"))
        registry.add_recognizer(IsraeliPhoneRecognizer(lang="en"))
        registry.add_recognizer(HebrewDateRecognizer(lang="en"))
        return AnalyzerEngine(nlp_engine=nlp_engine, registry=registry, supported_languages=["en"])

    def _build_he_analyzer(self) -> AnalyzerEngine:
        nlp_engine = None
        try:
            import spacy
            spacy.load("xx_ent_wiki_sm")
            provider = NlpEngineProvider(nlp_configuration={
                "nlp_engine_name": "spacy",
                "models": [{"lang_code": "he", "model_name": "xx_ent_wiki_sm"}],
            })
            nlp_engine = provider.create_engine()
            self._xx_available = True
        except Exception:
            self._xx_available = False

        registry = RecognizerRegistry()
        if nlp_engine:
            registry.load_predefined_recognizers(nlp_engine=nlp_engine)

        registry.add_recognizer(IsraeliIdRecognizer(lang="he"))
        registry.add_recognizer(IsraeliPhoneRecognizer(lang="he"))
        registry.add_recognizer(HebrewDateRecognizer(lang="he"))
        registry.add_recognizer(HebrewNameRecognizer())
        registry.add_recognizer(HebrewLocationRecognizer())

        from presidio_analyzer.predefined_recognizers import (
            EmailRecognizer, IbanRecognizer, CreditCardRecognizer, UrlRecognizer, IpRecognizer,
        )
        for rec_cls in (EmailRecognizer, IbanRecognizer, CreditCardRecognizer, UrlRecognizer, IpRecognizer):
            try:
                registry.add_recognizer(rec_cls(supported_language="he"))
            except Exception:
                pass

        try:
            from presidio_analyzer.predefined_recognizers import PhoneRecognizer
            try:
                registry.add_recognizer(PhoneRecognizer(supported_language="he", supported_regions=["IL"]))
            except Exception:
                registry.add_recognizer(PhoneRecognizer(supported_language="he"))
        except Exception:
            pass

        if nlp_engine:
            return AnalyzerEngine(nlp_engine=nlp_engine, registry=registry, supported_languages=["he"])
        else:
            provider = NlpEngineProvider(nlp_configuration={
                "nlp_engine_name": "spacy",
                "models": [{"lang_code": "he", "model_name": "en_core_web_sm"}],
            })
            try:
                fallback_engine = provider.create_engine()
                return AnalyzerEngine(nlp_engine=fallback_engine, registry=registry, supported_languages=["he"])
            except Exception:
                return AnalyzerEngine(registry=registry, supported_languages=["he"])

    @staticmethod
    def _map_entity(entity_type: str) -> str:
        mapping = {
            "PERS": "PERSON", "PER": "PERSON",
            "LOC": "LOCATION", "GPE": "LOCATION",
            "ORG": "ORGANIZATION",
            "DATE": "DATE_TIME", "TIME": "DATE_TIME",
            "HE_PERSON": "PERSON",
            "HE_LOCATION": "LOCATION",
        }
        return mapping.get(entity_type, entity_type)

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

        lang = detect_language(text)

        if lang == "he":
            try:
                he_results = self._he_analyzer.analyze(text=text, language="he", score_threshold=confidence)
            except Exception:
                he_results = []
            try:
                en_results = self._en_analyzer.analyze(text=text, language="en", score_threshold=confidence)
            except Exception:
                en_results = []
            raw = list(he_results) + list(en_results)
            analysis_lang = "he"
        else:
            try:
                raw = self._en_analyzer.analyze(text=text, language="en", entities=entities, score_threshold=confidence)
            except Exception:
                raw = self._en_analyzer.analyze(text=text, language="en", score_threshold=confidence)
            analysis_lang = "en"

        resolved = self._resolve_overlaps(raw)
        mapping: Dict[str, str] = {}
        entity_counts: Dict[str, int] = {}
        value_to_ph: Dict[str, str] = {}
        detections: List[dict] = []

        for result in sorted(resolved, key=lambda r: r.start, reverse=True):
            original = text[result.start:result.end]
            etype = self._map_entity(result.entity_type)
            label = ENTITY_LABELS.get(etype, etype)
            if original in value_to_ph:
                ph = value_to_ph[original]
            else:
                entity_counts[label] = entity_counts.get(label, 0) + 1
                ph = f"{{{{{label}_{entity_counts[label]}}}}}"
                mapping[ph] = original
                value_to_ph[original] = ph
            detections.append({
                "placeholder": ph, "original": original,
                "label": label, "score": result.score,
                "start": result.start, "end": result.end,
                "lang": analysis_lang,
            })
            text = text[:result.start] + ph + text[result.end:]

        return text, mapping, detections

    @staticmethod
    def restore(text: str, mapping: Dict[str, str]) -> str:
        for ph, orig in mapping.items():
            text = text.replace(ph, orig)
        return text

    @property
    def hebrew_ner_model(self) -> str:
        if self._xx_available:
            return "xx_ent_wiki_sm + dictionary (700+ names)"
        return "dictionary (700+ Israeli names)"


# ---------------------------------------------------------------------------
#  Reusable styled widgets
# ---------------------------------------------------------------------------

def styled_button(parent, text, command, bg=ACCENT, fg="white", width=None, pady=8):
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
    return tk.Label(parent, text=text, bg=PANEL_BG, fg=ACCENT, font=("Segoe UI", 9, "bold"))


def entry_row(parent, label_text, var, browse_cmd=None):
    row = tk.Frame(parent, bg=PANEL_BG)
    tk.Label(row, text=label_text, bg=PANEL_BG, fg=TEXT_MAIN,
             font=("Segoe UI", 9), width=18, anchor="w").pack(side="left")
    ent = tk.Entry(row, textvariable=var, bg=ENTRY_BG, fg=TEXT_MAIN,
                   insertbackground=TEXT_MAIN, relief="flat", font=("Segoe UI", 9), bd=4)
    ent.pack(side="left", fill="x", expand=True)
    if browse_cmd:
        tk.Button(row, text="Browse", command=browse_cmd, bg=BORDER, fg=TEXT_MAIN,
                  font=("Segoe UI", 8), relief="flat", cursor="hand2", padx=6
                  ).pack(side="left", padx=(4, 0))
    return row


# ---------------------------------------------------------------------------
#  Main Application Window
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Legal Document PII Anonymizer  |  Hebrew / English  v3.0")
        self.geometry("1150x820")
        self.minsize(900, 640)
        self.configure(bg=DARK_BG)
        self._engine_ready = False
        self._build_ui()
        self._load_engine_async()

    def _load_engine_async(self):
        self._set_status("Loading NLP engine...", WARNING)
        threading.Thread(target=self._load_engine, daemon=True).start()

    def _load_engine(self):
        try:
            engine = PIIEngine.get()
            self._engine_ready = True
            model_info = engine.hebrew_ner_model
            msg = f"Engine ready — Hebrew: {model_info}"
            self.after(0, lambda: self._set_status(msg, SUCCESS))
            self.after(0, lambda: self._lang_lbl.config(text=f"EN + HE ({model_info})", fg=SUCCESS))
        except Exception as exc:
            self.after(0, lambda: self._set_status(f"Engine error: {exc}", DANGER))

    def _build_ui(self):
        header = tk.Frame(self, bg=PANEL_BG, pady=14)
        header.pack(fill="x")
        tk.Label(header, text="Legal Document PII Anonymizer & Restorer  v3.0",
                 bg=PANEL_BG, fg=TEXT_MAIN, font=("Segoe UI", 14, "bold")).pack(side="left", padx=20)
        self._lang_lbl = tk.Label(header, text="Loading...", bg=PANEL_BG, fg=WARNING, font=("Segoe UI", 9))
        self._lang_lbl.pack(side="right", padx=20)
        self._status_lbl = tk.Label(header, text="", bg=PANEL_BG, fg=WARNING, font=("Segoe UI", 9))
        self._status_lbl.pack(side="right", padx=20)

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook", background=DARK_BG, borderwidth=0)
        style.configure("TNotebook.Tab", background=PANEL_BG, foreground=TEXT_DIM,
                        font=("Segoe UI", 10, "bold"), padding=[18, 8])
        style.map("TNotebook.Tab", background=[("selected", ACCENT)], foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(8, 0))

        self._tab_anon    = tk.Frame(nb, bg=DARK_BG)
        self._tab_restore = tk.Frame(nb, bg=DARK_BG)
        self._tab_about   = tk.Frame(nb, bg=DARK_BG)

        nb.add(self._tab_anon,    text="  \U0001f512  Anonymize / Remove PII  ")
        nb.add(self._tab_restore, text="  \U0001f513  Restore / Reinstate PII  ")
        nb.add(self._tab_about,   text="  \u2139  About  ")

        self._build_anonymize_tab()
        self._build_restore_tab()
        self._build_about_tab()

        bar = tk.Frame(self, bg=PANEL_BG, height=28)
        bar.pack(fill="x", side="bottom")
        self._progress = ttk.Progressbar(bar, mode="indeterminate", length=200)
        self._progress.pack(side="right", padx=12, pady=4)
        self._bar_lbl = tk.Label(bar, text="", bg=PANEL_BG, fg=TEXT_DIM, font=("Segoe UI", 8))
        self._bar_lbl.pack(side="left", padx=12)

    def _build_anonymize_tab(self):
        tab = self._tab_anon
        left = tk.Frame(tab, bg=DARK_BG, width=340)
        left.pack(side="left", fill="y", padx=(12, 6), pady=12)
        left.pack_propagate(False)
        right = tk.Frame(tab, bg=DARK_BG)
        right.pack(side="left", fill="both", expand=True, padx=(0, 12), pady=12)

        panel = tk.Frame(left, bg=PANEL_BG, bd=0)
        panel.pack(fill="both", expand=True, pady=(0, 8))

        section_label(panel, "INPUT / OUTPUT").pack(anchor="w", padx=14, pady=(14, 4))

        self._anon_input  = tk.StringVar()
        self._anon_output = tk.StringVar()
        self._anon_map    = tk.StringVar()

        for lbl, var, cmd in [
            ("Input document",    self._anon_input,  self._browse_input),
            ("Anonymized output", self._anon_output, self._browse_anon_out),
            ("Mapping file",      self._anon_map,    self._browse_anon_map),
        ]:
            entry_row(panel, lbl, var, cmd).pack(fill="x", padx=14, pady=3)

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=10)
        section_label(panel, "DETECTION SETTINGS").pack(anchor="w", padx=14, pady=(0, 6))

        lang_row = tk.Frame(panel, bg=PANEL_BG)
        lang_row.pack(fill="x", padx=14, pady=(0, 6))
        tk.Label(lang_row, text="Language", bg=PANEL_BG, fg=TEXT_MAIN,
                 font=("Segoe UI", 9), width=18, anchor="w").pack(side="left")
        self._lang_override = tk.StringVar(value="auto")
        ttk.Combobox(lang_row, textvariable=self._lang_override,
                     values=["auto", "English", "Hebrew"],
                     state="readonly", width=16, font=("Segoe UI", 9)).pack(side="left")

        conf_row = tk.Frame(panel, bg=PANEL_BG)
        conf_row.pack(fill="x", padx=14, pady=(0, 4))
        tk.Label(conf_row, text="Min. confidence", bg=PANEL_BG, fg=TEXT_MAIN,
                 font=("Segoe UI", 9), width=18, anchor="w").pack(side="left")
        self._confidence = tk.DoubleVar(value=DEFAULT_CONFIDENCE)
        tk.Scale(conf_row, from_=0.3, to=1.0, resolution=0.05, orient="horizontal",
                 variable=self._confidence, bg=PANEL_BG, fg=TEXT_MAIN, troughcolor=ENTRY_BG,
                 highlightthickness=0, font=("Segoe UI", 8), activebackground=ACCENT, length=120
                 ).pack(side="left")
        self._conf_lbl = tk.Label(conf_row, text=f"{DEFAULT_CONFIDENCE:.2f}",
                                   bg=PANEL_BG, fg=ACCENT, font=("Segoe UI", 9, "bold"), width=4)
        self._conf_lbl.pack(side="left")
        self._confidence.trace_add("write",
            lambda *_: self._conf_lbl.config(text=f"{self._confidence.get():.2f}"))

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=10)
        section_label(panel, "ENTITY TYPES TO DETECT").pack(anchor="w", padx=14, pady=(0, 6))

        self._entity_vars: Dict[str, tk.BooleanVar] = {}
        scroll_frame = tk.Frame(panel, bg=PANEL_BG)
        scroll_frame.pack(fill="x", padx=14)

        display_entities = {k: v for k, v in ENTITY_LABELS.items()
                            if k not in ("HE_PERSON", "HE_LOCATION")}
        for i, (etype, label) in enumerate(display_entities.items()):
            var = tk.BooleanVar(value=True)
            self._entity_vars[etype] = var
            col = i % 2
            row_idx = i // 2
            he_name = HE_ENTITY_NAMES.get(label, "")
            display = f"{label}" + (f" ({he_name})" if he_name else "")
            tk.Checkbutton(scroll_frame, text=display, variable=var,
                           bg=PANEL_BG, fg=TEXT_MAIN, selectcolor=ENTRY_BG,
                           activebackground=PANEL_BG, activeforeground=TEXT_MAIN,
                           font=("Segoe UI", 8), anchor="w"
                           ).grid(row=row_idx, column=col, sticky="w", pady=1)

        for k in ("HE_PERSON", "HE_LOCATION"):
            self._entity_vars[k] = tk.BooleanVar(value=True)

        sel_row = tk.Frame(panel, bg=PANEL_BG)
        sel_row.pack(fill="x", padx=14, pady=(6, 0))
        tk.Button(sel_row, text="Select All", command=self._select_all_entities,
                  bg=BORDER, fg=TEXT_MAIN, font=("Segoe UI", 8), relief="flat",
                  cursor="hand2", padx=6).pack(side="left", padx=(0, 4))
        tk.Button(sel_row, text="Clear All", command=self._clear_all_entities,
                  bg=BORDER, fg=TEXT_MAIN, font=("Segoe UI", 8), relief="flat",
                  cursor="hand2", padx=6).pack(side="left")

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=10)
        styled_button(panel, "\U0001f512  Remove PII / Anonymize", self._run_anonymize,
                      width=28).pack(padx=14, pady=(0, 14))

        self._build_preview_area(right, "anon")

    def _build_restore_tab(self):
        tab = self._tab_restore
        left = tk.Frame(tab, bg=DARK_BG, width=340)
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
            entry_row(panel, lbl, var, cmd).pack(fill="x", padx=14, pady=3)

        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=14, pady=14)

        info_box = tk.Frame(panel, bg=ENTRY_BG, bd=0)
        info_box.pack(fill="x", padx=14, pady=(0, 10))
        tk.Label(info_box,
                 text="Workflow:\n\n"
                      "1. Anonymize your document\n"
                      "2. Send anonymized text to cloud AI\n"
                      "   (ChatGPT, Claude, Gemini, etc.)\n"
                      "3. Save the AI response to a file\n"
                      "4. Load it here as \'Anonymized document\'\n"
                      "5. Click Restore — done!\n\n"
                      "Works with English and Hebrew.\n"
                      "Hebrew names: 700+ Israeli names\n"
                      "dictionary + multilingual NER.",
                 bg=ENTRY_BG, fg=TEXT_DIM, font=("Segoe UI", 9),
                 justify="left", padx=12, pady=10).pack()

        styled_button(panel, "\U0001f513  Restore Original PII", self._run_restore,
                      width=28).pack(padx=14, pady=(0, 14))

        self._build_preview_area(right, "rest")

    def _build_about_tab(self):
        tab = self._tab_about
        frame = tk.Frame(tab, bg=DARK_BG)
        frame.pack(expand=True)

        tk.Label(frame, text="Legal Document PII Anonymizer & Restorer  v3.0",
                 bg=DARK_BG, fg=TEXT_MAIN, font=("Segoe UI", 16, "bold")).pack(pady=(40, 0))
        tk.Label(frame, text="Protect attorney-client privilege before using cloud AI",
                 bg=DARK_BG, fg=TEXT_DIM, font=("Segoe UI", 10)).pack(pady=(4, 24))

        info = [
            ("Engine",      "Microsoft Presidio + spaCy (en_core_web_sm + xx_ent_wiki_sm)"),
            ("Languages",   "English + Hebrew — auto-detected per document"),
            ("Hebrew NER",  "700+ Israeli first names, 200+ surnames, 150+ locations (dictionary)"),
            ("Israeli PII", "Teudat Zehut (Luhn-10), phone (+972/05x/07x), IBAN IL"),
            ("Entities",    f"{len(set(ENTITY_LABELS.values()))} PII types detected"),
            ("Formats",     ".txt  |  .docx  |  .pdf  (input)    .txt  |  .docx  (output)"),
            ("Encoding",    "UTF-8, Windows-1255, ISO-8859-8, CP1255"),
            ("Privacy",     "100% local — no data leaves your machine"),
        ]
        for key, val in info:
            row = tk.Frame(frame, bg=PANEL_BG, pady=8, padx=20)
            row.pack(fill="x", padx=60, pady=3)
            tk.Label(row, text=f"{key}:", bg=PANEL_BG, fg=ACCENT,
                     font=("Segoe UI", 9, "bold"), width=14, anchor="w").pack(side="left")
            tk.Label(row, text=val, bg=PANEL_BG, fg=TEXT_MAIN, font=("Segoe UI", 9)).pack(side="left")

    def _build_preview_area(self, parent, prefix):
        top = tk.Frame(parent, bg=DARK_BG)
        top.pack(fill="both", expand=True)

        left_frame  = tk.Frame(top, bg=DARK_BG)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
        right_frame = tk.Frame(top, bg=DARK_BG)
        right_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))

        lbl_in  = "Original Document"  if prefix == "anon" else "Anonymized Document"
        lbl_out = "Anonymized Output"  if prefix == "anon" else "Restored Document"

        for frame, lbl, attr in [
            (left_frame,  lbl_in,  f"_{prefix}_txt_in"),
            (right_frame, lbl_out, f"_{prefix}_txt_out"),
        ]:
            header = tk.Frame(frame, bg=PANEL_BG)
            header.pack(fill="x")
            tk.Label(header, text=lbl, bg=PANEL_BG, fg=ACCENT,
                     font=("Segoe UI", 9, "bold"), pady=6, padx=10).pack(side="left")
            if attr.endswith("_out"):
                tk.Button(header, text="Copy", bg=BORDER, fg=TEXT_MAIN,
                          font=("Segoe UI", 8), relief="flat", cursor="hand2",
                          padx=6, command=lambda a=attr: self._copy_text(a)
                          ).pack(side="right", padx=6, pady=4)
                tk.Button(header, text="Save As...", bg=BORDER, fg=TEXT_MAIN,
                          font=("Segoe UI", 8), relief="flat", cursor="hand2",
                          padx=6, command=lambda a=attr: self._save_text(a)
                          ).pack(side="right", padx=(0, 4), pady=4)

            txt = scrolledtext.ScrolledText(
                frame, bg=ENTRY_BG, fg=TEXT_MAIN, insertbackground=TEXT_MAIN,
                font=("Consolas", 9), relief="flat", wrap="word",
                selectbackground=ACCENT, selectforeground="white",
            )
            txt.pack(fill="both", expand=True)
            setattr(self, attr, txt)

        out_widget = getattr(self, f"_{prefix}_txt_out")
        for label, colour in ENTITY_COLOURS.items():
            out_widget.tag_configure(label, foreground=colour, font=("Consolas", 9, "bold"))
        out_widget.tag_configure("OTHER", foreground=TAG_OTHER, font=("Consolas", 9, "bold"))

        tbl_frame = tk.Frame(parent, bg=PANEL_BG, height=160)
        tbl_frame.pack(fill="x", pady=(8, 0))
        tbl_frame.pack_propagate(False)

        tk.Label(tbl_frame, text="Detection Log", bg=PANEL_BG, fg=ACCENT,
                 font=("Segoe UI", 9, "bold"), pady=6, padx=10).pack(anchor="w")

        cols = ("Placeholder", "Original Value", "Entity Type", "Confidence", "Lang")
        tv = ttk.Treeview(tbl_frame, columns=cols, show="headings", height=4)
        style = ttk.Style()
        style.configure("Treeview", background=ENTRY_BG, foreground=TEXT_MAIN,
                         fieldbackground=ENTRY_BG, rowheight=22, font=("Consolas", 8))
        style.configure("Treeview.Heading", background=PANEL_BG, foreground=ACCENT,
                         font=("Segoe UI", 8, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", ACCENT)])

        col_widths = {"Placeholder": 130, "Original Value": 200,
                      "Entity Type": 120, "Confidence": 80, "Lang": 50}
        for col in cols:
            tv.heading(col, text=col)
            tv.column(col, width=col_widths.get(col, 120), anchor="w")

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tv.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        setattr(self, f"_{prefix}_table", tv)

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
        path = filedialog.asksaveasfilename(title="Save anonymized document",
            defaultextension=".txt", filetypes=[("Text file", "*.txt"), ("Word document", "*.docx")])
        if path: self._anon_output.set(path)

    def _browse_anon_map(self):
        path = filedialog.asksaveasfilename(title="Save mapping file",
            defaultextension=".json", filetypes=[("JSON file", "*.json")])
        if path: self._anon_map.set(path)

    def _browse_rest_input(self):
        path = filedialog.askopenfilename(title="Select anonymized document",
            filetypes=[("Documents", "*.txt *.docx"), ("All files", "*.*")])
        if path:
            self._rest_input.set(path)
            base = os.path.splitext(path)[0]
            self._rest_output.set(base + "_restored.txt")
            self._load_preview(path, self._rest_txt_in)

    def _browse_rest_map(self):
        path = filedialog.askopenfilename(title="Select mapping file",
            filetypes=[("JSON file", "*.json"), ("All files", "*.*")])
        if path: self._rest_map.set(path)

    def _browse_rest_out(self):
        path = filedialog.asksaveasfilename(title="Save restored document",
            defaultextension=".txt", filetypes=[("Text file", "*.txt"), ("Word document", "*.docx")])
        if path: self._rest_output.set(path)

    def _load_preview(self, path: str, widget: scrolledtext.ScrolledText):
        try:
            text = read_file(path)
            widget.config(state="normal")
            widget.delete("1.0", "end")
            widget.insert("1.0", text)
            if is_rtl(text):
                try: widget.config(justify="right")
                except Exception: pass
            widget.config(state="disabled")
        except Exception as exc:
            messagebox.showerror("Read Error", str(exc))

    def _set_output_text(self, widget: scrolledtext.ScrolledText, text: str, detections=None):
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        if is_rtl(text):
            try: widget.config(justify="right")
            except Exception: pass
        if detections:
            self._highlight_placeholders(widget, text, detections)
        widget.config(state="disabled")

    def _highlight_placeholders(self, widget, text: str, detections):
        for ph_match in re.finditer(r"\{\{([A-Z_]+)_\d+\}\}", text):
            label = ph_match.group(1)
            tag   = label if label in ENTITY_COLOURS else "OTHER"
            widget.tag_add(tag,
                           f"1.0 + {ph_match.start()} chars",
                           f"1.0 + {ph_match.end()} chars")

    def _copy_text(self, attr: str):
        text = getattr(self, attr).get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(text)
        self._set_status("Copied to clipboard.", SUCCESS)

    def _save_text(self, attr: str):
        text = getattr(self, attr).get("1.0", "end-1c")
        path = filedialog.asksaveasfilename(defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("Word document", "*.docx")])
        if path:
            try:
                write_file(path, text)
                self._set_status(f"Saved to {path}", SUCCESS)
            except Exception as exc:
                messagebox.showerror("Save Error", str(exc))

    def _select_all_entities(self):
        for v in self._entity_vars.values(): v.set(True)

    def _clear_all_entities(self):
        for v in self._entity_vars.values(): v.set(False)

    def _set_status(self, msg: str, colour: str = TEXT_MAIN):
        self._status_lbl.config(text=msg, fg=colour)
        self._bar_lbl.config(text=msg)

    def _start_spinner(self, msg: str):
        self._bar_lbl.config(text=msg)
        self._progress.start(12)

    def _stop_spinner(self):
        self._progress.stop()

    def _populate_table(self, tv: ttk.Treeview, detections):
        for row in tv.get_children(): tv.delete(row)
        seen = set()
        for d in sorted(detections, key=lambda x: x["start"]):
            key = (d["placeholder"], d["original"])
            if key not in seen:
                seen.add(key)
                tv.insert("", "end", values=(
                    d["placeholder"], d["original"], d["label"],
                    f"{d['score']:.2f}", d.get("lang", "en").upper(),
                ))

    def _run_anonymize(self):
        if not self._engine_ready:
            messagebox.showwarning("Engine Loading", "Please wait for the NLP engine to finish loading.")
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
        confidence = self._confidence.get()
        self._start_spinner("Anonymizing document...")
        self._set_status("Anonymizing...", WARNING)

        def task():
            try:
                text = read_file(inp)
                engine = PIIEngine.get()
                anon_text, mapping, detections = engine.anonymize(
                    text, confidence=confidence, entities=selected_entities)
                write_file(out, anon_text)
                with open(mapf, "w", encoding="utf-8") as fh:
                    json.dump(mapping, fh, indent=4, ensure_ascii=False)
                self.after(0, lambda: self._on_anonymize_done(anon_text, mapping, detections, out, mapf))
            except Exception as exc:
                self.after(0, lambda: self._on_error(str(exc)))

        threading.Thread(target=task, daemon=True).start()

    def _on_anonymize_done(self, anon_text, mapping, detections, out, mapf):
        self._stop_spinner()
        self._set_output_text(self._anon_txt_out, anon_text, detections)
        self._populate_table(self._anon_table, detections)
        n_unique = len(mapping)
        n_total  = len(detections)
        self._set_status(f"Done — {n_unique} unique PII items replaced ({n_total} total).", SUCCESS)
        messagebox.showinfo("Anonymization Complete",
            f"Anonymization complete!\n\n"
            f"  Unique PII items replaced : {n_unique}\n"
            f"  Total occurrences         : {n_total}\n\n"
            f"Anonymized document: {out}\n"
            f"Mapping file: {mapf}\n\n"
            f"Safe to send to cloud AI.")

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

        self._start_spinner("Restoring PII...")
        self._set_status("Restoring...", WARNING)

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
        fake_detections = [
            {"placeholder": ph, "original": orig,
             "label": ph.strip("{}").rsplit("_", 1)[0],
             "score": 1.0, "start": 0, "end": 0, "lang": ""}
            for ph, orig in mapping.items()
        ]
        self._populate_table(self._rest_table, fake_detections)
        self._set_status(f"Done — {len(mapping)} placeholders restored.", SUCCESS)
        messagebox.showinfo("Restoration Complete",
            f"Restoration complete!\n\n"
            f"  Placeholders restored : {len(mapping)}\n\n"
            f"Restored document saved to:\n  {out}")

    def _on_error(self, msg: str):
        self._stop_spinner()
        self._set_status(f"Error: {msg}", DANGER)
        messagebox.showerror("Error", msg)


# ---------------------------------------------------------------------------
#  Entry point
# ---------------------------------------------------------------------------

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
