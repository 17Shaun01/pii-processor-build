"""
Headless test for the three new features:
1. Custom PII manual entry (apply_custom_pii)
2. Project PII file save/load
3. Batch processing logic (engine + custom PII)
"""
import sys, os, json, tempfile, shutil

# Mock tkinter before importing GUI
import types
for mod in ["tkinter", "tkinter.filedialog", "tkinter.messagebox",
            "tkinter.scrolledtext", "tkinter.ttk"]:
    sys.modules[mod] = types.ModuleType(mod)

class _Dummy:
    def __init__(self, *a, **k): pass
    def __call__(self, *a, **k): return self
    def __getattr__(self, n): return self
    def pack(self, *a, **k): pass
    def grid(self, *a, **k): pass
    def config(self, *a, **k): pass
    def bind(self, *a, **k): pass
    def insert(self, *a, **k): pass
    def delete(self, *a, **k): pass
    def get(self, *a, **k): return ""
    def set(self, *a, **k): pass
    def trace_add(self, *a, **k): pass
    def after(self, *a, **k): pass
    def withdraw(self, *a, **k): pass
    def deiconify(self, *a, **k): pass
    def mainloop(self, *a, **k): pass
    def update(self, *a, **k): pass
    def lift(self, *a, **k): pass
    def destroy(self, *a, **k): pass
    def winfo_screenwidth(self): return 1920
    def winfo_screenheight(self): return 1080
    def geometry(self, *a, **k): pass
    def overrideredirect(self, *a, **k): pass
    def attributes(self, *a, **k): pass
    def configure(self, *a, **k): pass
    def start(self, *a, **k): pass
    def stop(self, *a, **k): pass
    def pack_propagate(self, *a, **k): pass
    def tag_configure(self, *a, **k): pass
    def tag_add(self, *a, **k): pass
    def see(self, *a, **k): pass
    def get_children(self): return []
    def selection(self): return []
    def item(self, *a, **k): return {"values": []}
    def heading(self, *a, **k): pass
    def column(self, *a, **k): pass
    def yview(self, *a, **k): pass
    def __iter__(self): return iter([])

tk_mock = sys.modules["tkinter"]
tk_mock.Tk = _Dummy
tk_mock.Toplevel = _Dummy
tk_mock.Frame = _Dummy
tk_mock.Label = _Dummy
tk_mock.Button = _Dummy
tk_mock.Entry = _Dummy
tk_mock.Scale = _Dummy
tk_mock.Checkbutton = _Dummy
tk_mock.StringVar = lambda *a, **k: _Dummy()
tk_mock.BooleanVar = lambda *a, **k: _Dummy()
tk_mock.DoubleVar = lambda *a, **k: _Dummy()
tk_mock.HORIZONTAL = "horizontal"
tk_mock.END = "end"
tk_mock.DISABLED = "disabled"
tk_mock.NORMAL = "normal"

ttk_mock = sys.modules["tkinter.ttk"]
ttk_mock.Notebook = _Dummy
ttk_mock.Treeview = _Dummy
ttk_mock.Scrollbar = _Dummy
ttk_mock.Progressbar = _Dummy
ttk_mock.Combobox = _Dummy
ttk_mock.Style = _Dummy

st_mock = sys.modules["tkinter.scrolledtext"]
st_mock.ScrolledText = _Dummy

mb_mock = sys.modules["tkinter.messagebox"]
mb_mock.showerror = lambda *a, **k: None
mb_mock.showinfo  = lambda *a, **k: None
mb_mock.showwarning = lambda *a, **k: None
mb_mock.askyesno = lambda *a, **k: True

fd_mock = sys.modules["tkinter.filedialog"]
fd_mock.askopenfilename = lambda *a, **k: ""
fd_mock.asksaveasfilename = lambda *a, **k: ""
fd_mock.askdirectory = lambda *a, **k: ""

sys.path.insert(0, os.path.dirname(__file__))

from pii_processor_gui import (
    PIIEngine, apply_custom_pii, load_project_pii, save_project_pii,
    get_project_pii_path, read_file, write_file, CUSTOM_PII_FILENAME
)

print("=" * 60)
print("Test 1: apply_custom_pii — manual entry")
print("=" * 60)

text = "The client Moshe Cohen signed the agreement. Moshe Cohen will pay 5000 NIS."
custom_entries = [
    {"text": "Moshe Cohen", "label": "PERSON"},
    {"text": "5000 NIS",    "label": "CUSTOM"},
]
mapping = {}
entity_counts = {}
value_to_ph = {}
detections = []

result = apply_custom_pii(text, custom_entries, mapping, entity_counts, value_to_ph, detections)
print(f"Input:  {text}")
print(f"Output: {result}")
print(f"Mapping: {json.dumps(mapping, ensure_ascii=False, indent=2)}")
assert "Moshe Cohen" not in result, "Name should be replaced"
assert "5000 NIS" not in result, "Amount should be replaced"
assert "{{PERSON_1}}" in result, "PERSON placeholder should appear"
assert "{{CUSTOM_1}}" in result, "CUSTOM placeholder should appear"
# Both occurrences of Moshe Cohen should use the same placeholder
assert result.count("{{PERSON_1}}") == 2, "Both occurrences should use same placeholder"
print("PASS\n")

print("=" * 60)
print("Test 2: Project PII file save and load")
print("=" * 60)

tmpdir = tempfile.mkdtemp()
try:
    entries = [
        {"text": "דוד לוי",       "label": "PERSON"},
        {"text": "רחוב הרצל 12",  "label": "LOCATION"},
        {"text": "IL12-0123",      "label": "IBAN"},
    ]
    save_project_pii(tmpdir, entries)
    pii_path = get_project_pii_path(tmpdir)
    assert os.path.exists(pii_path), f"PII file should exist at {pii_path}"
    print(f"Saved to: {pii_path}")

    loaded = load_project_pii(tmpdir)
    assert len(loaded) == 3, f"Should load 3 entries, got {len(loaded)}"
    assert loaded[0]["text"] == "דוד לוי", "Hebrew name should round-trip correctly"
    assert loaded[1]["label"] == "LOCATION", "Label should be preserved"
    print(f"Loaded {len(loaded)} entries correctly (including Hebrew characters)")
    print("PASS\n")
finally:
    shutil.rmtree(tmpdir)

print("=" * 60)
print("Test 3: Batch processing — multiple files with custom PII")
print("=" * 60)

tmpdir = tempfile.mkdtemp()
try:
    # Create 3 test files
    docs = {
        "contract_a.txt": "Client: Sarah Johnson, ID: 025478963, Phone: 052-1234567",
        "contract_b.txt": "Defendant: John Smith signed on 01/01/2024. Email: john@example.com",
        "contract_c.txt": "The firm of ABC Partners agrees to pay $10,000 to Jane Doe.",
    }
    for fname, content in docs.items():
        with open(os.path.join(tmpdir, fname), "w", encoding="utf-8") as f:
            f.write(content)

    # Save a project custom PII file
    custom = [{"text": "ABC Partners", "label": "ORGANIZATION"}]
    save_project_pii(tmpdir, custom)

    # Simulate batch processing
    import re
    from pii_processor_gui import apply_custom_pii, load_project_pii
    engine = PIIEngine.get()
    folder_custom = load_project_pii(tmpdir)
    out_dir = os.path.join(tmpdir, "out")
    os.makedirs(out_dir)

    results = {}
    for fname in docs:
        fpath = os.path.join(tmpdir, fname)
        text = read_file(fpath)
        anon_text, mapping, detections = engine.anonymize(text, confidence=0.4)
        if folder_custom:
            ec = {}
            vp = {ph: orig for ph, orig in mapping.items()}
            for ph in mapping:
                m = re.match(r'\{\{([A-Z_]+)_(\d+)\}\}', ph)
                if m:
                    lbl, num = m.group(1), int(m.group(2))
                    ec[lbl] = max(ec.get(lbl, 0), num)
            anon_text = apply_custom_pii(anon_text, folder_custom, mapping, ec, vp, detections)
        out_path = os.path.join(out_dir, os.path.splitext(fname)[0] + "_anonymized.txt")
        write_file(out_path, anon_text)
        map_path = os.path.join(out_dir, os.path.splitext(fname)[0] + "_mapping.json")
        with open(map_path, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False)
        results[fname] = (anon_text, mapping)
        print(f"  {fname}: {len(mapping)} PII items replaced")

    # Verify contract_c has ABC Partners replaced
    anon_c, map_c = results["contract_c.txt"]
    assert "ABC Partners" not in anon_c, "Custom PII 'ABC Partners' should be replaced"
    print(f"  Custom PII 'ABC Partners' correctly replaced in contract_c.txt")

    # Verify restoration works
    for fname, (anon_text, mapping) in results.items():
        original = docs[fname]
        restored = PIIEngine.restore(anon_text, mapping)
        assert restored == original, f"Round-trip failed for {fname}"
    print("  All 3 files restored perfectly (round-trip verified)")
    print("PASS\n")
finally:
    shutil.rmtree(tmpdir)

print("=" * 60)
print("ALL TESTS PASSED")
print("=" * 60)
