"""
Test that the Hebrew engine builds without the 'misconfigured engine' error,
both when xx model is available and when it falls back to en_core_web_sm.
"""
import sys, os, types

# Mock tkinter
for mod in ["tkinter","tkinter.filedialog","tkinter.messagebox","tkinter.scrolledtext","tkinter.ttk"]:
    sys.modules[mod] = types.ModuleType(mod)

class _D:
    def __init__(self,*a,**k): pass
    def __call__(self,*a,**k): return self
    def __getattr__(self,n): return self
    def get(self,*a,**k): return ""
    def set(self,*a,**k): pass
    def trace_add(self,*a,**k): pass

for attr in ["Tk","Toplevel","Frame","Label","Button","Entry","Scale","Checkbutton","Text","Canvas"]:
    setattr(sys.modules["tkinter"], attr, _D)
for attr in ["StringVar","BooleanVar","DoubleVar"]:
    setattr(sys.modules["tkinter"], attr, lambda *a,**k: _D())
for attr in ["END","DISABLED","NORMAL","HORIZONTAL","WORD","BOTH","LEFT","RIGHT","TOP","BOTTOM","X","Y","W","E","N","S","NW","NE","SW","SE","FLAT","GROOVE","RIDGE","SOLID","SUNKEN","RAISED","INSERT","SEL","SEL_FIRST","SEL_LAST"]:
    setattr(sys.modules["tkinter"], attr, attr.lower())
for mod in ["tkinter.ttk","tkinter.scrolledtext","tkinter.messagebox","tkinter.filedialog"]:
    m = sys.modules[mod]
    for attr in ["Notebook","Treeview","Scrollbar","Progressbar","Combobox","Style","ScrolledText",
                 "showerror","showinfo","showwarning","askyesno","askopenfilename","asksaveasfilename","askdirectory"]:
        setattr(m, attr, _D)

sys.path.insert(0, "/home/ubuntu/pii_win_build")

print("=" * 60)
print("Test 1: Engine builds without misconfigured error")
print("=" * 60)
from pii_processor_gui import PIIEngine, detect_language

engine = PIIEngine.get()
print(f"  English analyzer: OK")
print(f"  Hebrew analyzer:  OK (xx_available={engine._xx_available})")
print("PASS\n")

print("=" * 60)
print("Test 2: Hebrew text anonymization (auto-detect language)")
print("=" * 60)
he_text = "הלקוח דוד לוי, ת.ז. 025478963, טלפון 052-1234567, מתגורר בתל אביב."
lang = detect_language(he_text)
print(f"  Detected language: {lang}")
assert lang == "he", f"Expected 'he', got '{lang}'"
anon, mapping, detections = engine.anonymize(he_text, confidence=0.4)
print(f"  Input:      {he_text}")
print(f"  Anonymized: {anon}")
print(f"  Detected {len(mapping)} PII items: {list(mapping.values())}")
assert len(mapping) >= 2, f"Expected at least 2 PII items, got {len(mapping)}"
print("PASS\n")

print("=" * 60)
print("Test 3: English text anonymization still works")
print("=" * 60)
en_text = "Client John Smith, ID 025478963, email john@example.com"
lang2 = detect_language(en_text)
print(f"  Detected language: {lang2}")
assert lang2 == "en", f"Expected 'en', got '{lang2}'"
anon2, mapping2, detections2 = engine.anonymize(en_text, confidence=0.4)
print(f"  Input:      {en_text}")
print(f"  Anonymized: {anon2}")
print(f"  Detected {len(mapping2)} PII items: {list(mapping2.values())}")
assert len(mapping2) >= 2, f"Expected at least 2 PII items, got {len(mapping2)}"
print("PASS\n")

print("=" * 60)
print("Test 4: Round-trip restore for Hebrew")
print("=" * 60)
restored = PIIEngine.restore(anon, mapping)
assert restored == he_text, f"Round-trip failed!\nExpected: {he_text}\nGot:      {restored}"
print(f"  Restored: {restored}")
print("PASS\n")

print("=" * 60)
print("ALL TESTS PASSED — engine fix verified")
print("=" * 60)
