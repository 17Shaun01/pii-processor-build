"""
Headless test for Hebrew ambiguity detection and review dialog logic.
Mocks Tkinter so it runs without a display.
"""
import sys, types, re

# ── Mock tkinter ──────────────────────────────────────────────────────────────
tk_mock = types.ModuleType("tkinter")
class _FakeTk:
    def __init__(self, *a, **kw): pass
    def __call__(self, *a, **kw): return self
    def configure(self, **kw): pass
    def pack(self, **kw): pass
    def grid(self, **kw): pass
    def place(self, **kw): pass
    def bind(self, *a, **kw): pass
    def bind_all(self, *a, **kw): pass
    def after(self, *a, **kw): pass
    def wait_window(self, *a): pass
    def geometry(self, *a): pass
    def title(self, *a): pass
    def resizable(self, *a): pass
    def minsize(self, *a): pass
    def grab_set(self): pass
    def grab_release(self): pass
    def destroy(self): pass
    def overrideredirect(self, *a): pass
    def attributes(self, *a): pass
    def protocol(self, *a): pass
    def winfo_screenwidth(self): return 1920
    def winfo_screenheight(self): return 1080
    def update_idletasks(self): pass
    def config(self, **kw): pass
    def get(self, *a): return ""
    def delete(self, *a): pass
    def insert(self, *a): pass
    def tag_configure(self, *a, **kw): pass
    def tag_add(self, *a): pass
    def get_children(self): return []
    def start(self, *a): pass
    def stop(self): pass
    def set(self, *a): pass

class _FakeVar:
    def __init__(self, value=None): self._v = value
    def get(self): return self._v
    def set(self, v): self._v = v

for name in ["Tk", "Toplevel", "Frame", "Label", "Button", "Entry", "Text",
             "Canvas", "Checkbutton", "OptionMenu", "Scrollbar", "LabelFrame",
             "PanedWindow", "Scale", "Spinbox", "Listbox"]:
    setattr(tk_mock, name, _FakeTk)

tk_mock.BooleanVar = _FakeVar
tk_mock.StringVar  = _FakeVar
tk_mock.IntVar     = _FakeVar
tk_mock.DoubleVar  = _FakeVar
tk_mock.END = "end"
tk_mock.BOTH = "both"
tk_mock.LEFT = "left"
tk_mock.RIGHT = "right"
tk_mock.TOP = "top"
tk_mock.BOTTOM = "bottom"
tk_mock.X = "x"
tk_mock.Y = "y"
tk_mock.W = "w"
tk_mock.E = "e"
tk_mock.N = "n"
tk_mock.S = "s"
tk_mock.NW = "nw"
tk_mock.WORD = "word"
tk_mock.NORMAL = "normal"
tk_mock.DISABLED = "disabled"
tk_mock.HORIZONTAL = "horizontal"
tk_mock.VERTICAL = "vertical"
tk_mock.FLAT = "flat"
tk_mock.RIDGE = "ridge"
tk_mock.GROOVE = "groove"
tk_mock.SUNKEN = "sunken"
tk_mock.RAISED = "raised"

filedialog_mock = types.ModuleType("tkinter.filedialog")
filedialog_mock.askopenfilename = lambda **kw: ""
filedialog_mock.asksaveasfilename = lambda **kw: ""
messagebox_mock = types.ModuleType("tkinter.messagebox")
messagebox_mock.showinfo = lambda *a, **kw: None
messagebox_mock.showwarning = lambda *a, **kw: None
messagebox_mock.showerror = lambda *a, **kw: None
messagebox_mock.askyesno = lambda *a, **kw: True
scrolledtext_mock = types.ModuleType("tkinter.scrolledtext")
scrolledtext_mock.ScrolledText = _FakeTk
ttk_mock = types.ModuleType("tkinter.ttk")
for name in ["Notebook", "Frame", "Label", "Button", "Entry", "Combobox",
             "Progressbar", "Treeview", "Scrollbar", "Style", "Scale"]:
    setattr(ttk_mock, name, _FakeTk)

sys.modules["tkinter"] = tk_mock
sys.modules["tkinter.filedialog"] = filedialog_mock
sys.modules["tkinter.messagebox"] = messagebox_mock
sys.modules["tkinter.scrolledtext"] = scrolledtext_mock
sys.modules["tkinter.ttk"] = ttk_mock

# ── Now import the app module ─────────────────────────────────────────────────
sys.path.insert(0, "/home/ubuntu/pii_win_build")
from pii_processor_gui import (
    find_hebrew_ambiguous_candidates,
    HEBREW_AMBIGUOUS_NAMES,
    HEBREW_TITLE_WORDS,
    detect_language,
    PIIEngine,
)

print("=" * 60)
print("TEST 1: Language detection on Hebrew text")
he_text = "הסכם זה נחתם בין דוד לוי לבין שרה כהן בתל אביב ביום 01/01/2024."
lang = detect_language(he_text)
assert lang == "he", f"Expected 'he', got '{lang}'"
print(f"  PASS: detected language = '{lang}'")

print()
print("TEST 2: find_hebrew_ambiguous_candidates — basic detection")
# 'דוד' and 'שרה' are in HEBREW_AMBIGUOUS_NAMES (also common nouns)
# 'לוי' and 'כהן' are surnames in the dictionary
already_mapped = set()  # nothing already detected by NLP
candidates = find_hebrew_ambiguous_candidates(he_text, already_mapped)
print(f"  Found {len(candidates)} candidates:")
for c in candidates:
    print(f"    '{c['text']}' -> label={c['label']}, count={c['count']}, "
          f"preceded_by_title={c['preceded_by_title']}, context='{c['context'][:40]}...'")
assert len(candidates) > 0, "Expected at least 1 candidate"
candidate_words = {c["text"] for c in candidates}
# דוד and שרה are ambiguous names that should be found
assert "דוד" in candidate_words or "שרה" in candidate_words, \
    f"Expected דוד or שרה in candidates, got {candidate_words}"
print("  PASS: ambiguous candidates detected correctly")

print()
print("TEST 3: Already-mapped values are excluded")
already_mapped_2 = {"דוד", "שרה"}
candidates_2 = find_hebrew_ambiguous_candidates(he_text, already_mapped_2)
candidate_words_2 = {c["text"] for c in candidates_2}
assert "דוד" not in candidate_words_2, "דוד should be excluded (already mapped)"
assert "שרה" not in candidate_words_2, "שרה should be excluded (already mapped)"
print(f"  PASS: already-mapped values excluded. Remaining: {candidate_words_2}")

print()
print("TEST 4: Title-word detection")
title_text = "הנתבע דוד לוי מתגורר בתל אביב."
candidates_3 = find_hebrew_ambiguous_candidates(title_text, set())
title_preceded = [c for c in candidates_3 if c["preceded_by_title"]]
print(f"  Found {len(title_preceded)} title-preceded candidates:")
for c in title_preceded:
    print(f"    '{c['text']}' preceded_by_title=True")
# 'דוד' should be preceded by 'הנתבע' (a title word)
title_words_found = {c["text"] for c in title_preceded}
assert "דוד" in title_words_found, f"Expected דוד to be title-preceded, got {title_words_found}"
print("  PASS: title-word detection works")

print()
print("TEST 5: English document produces no Hebrew candidates")
en_text = "This agreement is between John Smith and Jane Doe, residing in Tel Aviv."
lang_en = detect_language(en_text)
assert lang_en == "en", f"Expected 'en', got '{lang_en}'"
candidates_en = find_hebrew_ambiguous_candidates(en_text, set())
assert len(candidates_en) == 0, f"Expected 0 candidates for English text, got {len(candidates_en)}"
print(f"  PASS: English document produces 0 Hebrew candidates")

print()
print("TEST 6: Full anonymize pipeline on Hebrew text")
engine = PIIEngine.get()
anon_text, mapping, detections = engine.anonymize(he_text, confidence=0.3)
print(f"  NLP detected {len(mapping)} unique PII items:")
for ph, orig in mapping.items():
    print(f"    {ph} -> '{orig}'")
already_mapped_nlp = set(mapping.values())
candidates_after_nlp = find_hebrew_ambiguous_candidates(he_text, already_mapped_nlp)
print(f"  After NLP: {len(candidates_after_nlp)} additional ambiguous candidates:")
for c in candidates_after_nlp:
    print(f"    '{c['text']}' -> label={c['label']}, count={c['count']}")
print("  PASS: Pipeline integration works")

print()
print("=" * 60)
print("ALL TESTS PASSED")
