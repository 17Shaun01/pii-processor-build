"""
Headless test for the translation map feature (v0.5.4).
Tests: generate_translation_map, _get_translation_map_path, restore with translated map.
"""
import sys, os, json, re
sys.path.insert(0, "/home/ubuntu/pii_win_build")

# Minimal tkinter mock so the GUI module imports without a display
import types
tk_mock = types.ModuleType("tkinter")
tk_mock.Frame = tk_mock.Label = tk_mock.Button = tk_mock.Entry = tk_mock.Radiobutton = object
tk_mock.StringVar = lambda value="": type("SV", (), {"get": lambda s: value, "set": lambda s,v: None})()
tk_mock.BooleanVar = lambda value=True: type("BV", (), {"get": lambda s: value, "set": lambda s,v: None})()
tk_mock.IntVar = lambda value=0: type("IV", (), {"get": lambda s: value, "set": lambda s,v: None})()
tk_mock.scrolledtext = types.ModuleType("scrolledtext")
tk_mock.scrolledtext.ScrolledText = object
tk_mock.filedialog = types.ModuleType("filedialog")
tk_mock.messagebox = types.ModuleType("messagebox")
tk_mock.ttk = types.ModuleType("ttk")
tk_mock.ttk.Notebook = tk_mock.ttk.Treeview = tk_mock.ttk.Scrollbar = tk_mock.ttk.Progressbar = tk_mock.ttk.Combobox = tk_mock.ttk.Style = object
tk_mock.Tk = tk_mock.Toplevel = object
tk_mock.BOTH = tk_mock.LEFT = tk_mock.RIGHT = tk_mock.TOP = tk_mock.BOTTOM = tk_mock.X = tk_mock.Y = tk_mock.W = tk_mock.E = tk_mock.N = tk_mock.S = ""
tk_mock.END = "end"
sys.modules["tkinter"] = tk_mock
sys.modules["tkinter.filedialog"] = tk_mock.filedialog
sys.modules["tkinter.messagebox"] = tk_mock.messagebox
sys.modules["tkinter.ttk"] = tk_mock.ttk
sys.modules["tkinter.scrolledtext"] = tk_mock.scrolledtext

# Import only the helper functions we need
import importlib.util
spec = importlib.util.spec_from_file_location("gui", "/home/ubuntu/pii_win_build/pii_processor_gui.py")
# We only need the module-level functions, not the App class
# Execute the module up to the class definitions
import ast, textwrap

# Read and extract just the helper functions
with open("/home/ubuntu/pii_win_build/pii_processor_gui.py", "r", encoding="utf-8") as f:
    source = f.read()

# Execute the module-level code (functions and constants only, skip class App)
exec_globals = {"__name__": "__test__", "__file__": "/home/ubuntu/pii_win_build/pii_processor_gui.py"}
# We'll import the specific functions directly
exec(compile(source, "pii_processor_gui.py", "exec"), exec_globals)

generate_translation_map = exec_globals["generate_translation_map"]
_get_translation_map_path = exec_globals["_get_translation_map_path"]
PIIEngine = exec_globals["PIIEngine"]

PASS = "\033[92mPASS\033[0m"
FAIL = "\033[91mFAIL\033[0m"
errors = 0

# ---- Test 1: _get_translation_map_path ----
p1 = _get_translation_map_path("/docs/contract_mapping.json")
expected1 = "/docs/contract_mapping_translated.json"
if p1 == expected1:
    print(f"{PASS}  Test 1: _get_translation_map_path  ->  {p1}")
else:
    print(f"{FAIL}  Test 1: got {p1!r}, expected {expected1!r}")
    errors += 1

# ---- Test 2: generate_translation_map — Hebrew names ----
mapping = {
    "{{PERSON_1}}": "דוד לוי",
    "{{PERSON_2}}": "שרה כהן",
    "{{LOCATION_1}}": "תל אביב",
    "{{IL_ID_1}}": "025478963",
    "{{PHONE_1}}": "052-1234567",
    "{{EMAIL_1}}": "david@example.com",
}
trans_map = generate_translation_map(mapping)

# Person 1: David Levy
t1 = trans_map.get("{{PERSON_1}}", {}).get("translated", "")
if "David" in t1 or "david" in t1.lower():
    print(f"{PASS}  Test 2a: PERSON_1 דוד לוי -> '{t1}'")
else:
    print(f"{FAIL}  Test 2a: PERSON_1 דוד לוי -> '{t1}' (expected David ...)")
    errors += 1

# Person 2: Sarah Cohen
t2 = trans_map.get("{{PERSON_2}}", {}).get("translated", "")
if t2:
    print(f"{PASS}  Test 2b: PERSON_2 שרה כהן -> '{t2}'")
else:
    print(f"{FAIL}  Test 2b: PERSON_2 שרה כהן -> empty (expected transliteration)")
    errors += 1

# Location: Tel Aviv
t3 = trans_map.get("{{LOCATION_1}}", {}).get("translated", "")
if "Tel Aviv" in t3 or "tel aviv" in t3.lower():
    print(f"{PASS}  Test 2c: LOCATION_1 תל אביב -> '{t3}'")
else:
    print(f"{FAIL}  Test 2c: LOCATION_1 תל אביב -> '{t3}' (expected Tel Aviv)")
    errors += 1

# ID: should be copied as-is
t4 = trans_map.get("{{IL_ID_1}}", {}).get("translated", "")
if t4 == "025478963":
    print(f"{PASS}  Test 2d: IL_ID_1 copied as-is -> '{t4}'")
else:
    print(f"{FAIL}  Test 2d: IL_ID_1 -> '{t4}' (expected '025478963')")
    errors += 1

# Phone: should be copied as-is
t5 = trans_map.get("{{PHONE_1}}", {}).get("translated", "")
if t5 == "052-1234567":
    print(f"{PASS}  Test 2e: PHONE_1 copied as-is -> '{t5}'")
else:
    print(f"{FAIL}  Test 2e: PHONE_1 -> '{t5}' (expected '052-1234567')")
    errors += 1

# ---- Test 3: Restore with translated map ----
anon_text = "The defendant {{PERSON_1}} residing at {{LOCATION_1}} with ID {{IL_ID_1}} called {{PHONE_1}}."
# Build a flat restore mapping from the translation map
restore_mapping = {}
for ph, info in trans_map.items():
    translated = info.get("translated", "").strip()
    original = info.get("original", "")
    restore_mapping[ph] = translated if translated and translated != "?" else original

restored = PIIEngine.restore(anon_text, restore_mapping)
if "{{" not in restored:
    print(f"{PASS}  Test 3: All placeholders restored in translated document")
    print(f"         Result: {restored}")
else:
    print(f"{FAIL}  Test 3: Unreplaced placeholders remain: {restored}")
    errors += 1

# ---- Test 4: Entity structure in translation map ----
for ph, info in trans_map.items():
    if not all(k in info for k in ("original", "translated", "entity")):
        print(f"{FAIL}  Test 4: Missing keys in {ph}: {info}")
        errors += 1
        break
else:
    print(f"{PASS}  Test 4: All translation map entries have correct structure")

print()
if errors == 0:
    print(f"\033[92mAll tests passed!\033[0m")
else:
    print(f"\033[91m{errors} test(s) failed.\033[0m")
    sys.exit(1)
