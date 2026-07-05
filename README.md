# Legal Document PII Anonymizer & Restorer — v0.5.4

**Hebrew & English | Windows GUI Application**

A fully local, privacy-first tool for Israeli legal professionals. It extracts all personal identification information (PII) from legal documents, replaces each item with a neutral placeholder variable, and allows you to safely send the sanitized document to any cloud AI service (ChatGPT, Claude, Gemini, DeepL, etc.). Once the AI returns its output — whether processed, summarized, or **translated** — the tool restores every original value back into the document in the correct language.

**No data ever leaves your computer during the anonymization phase.**

---

## Quick Start (Windows)

1. Unzip the downloaded archive.
2. Double-click **`PII_Processor.exe`**.
3. A loading splash screen will appear while the NLP engine initialises.
4. The main application window opens automatically when ready.

> **First launch note:** The very first time you run the `.exe`, Windows Defender will scan it. This may take 15–45 seconds. Every subsequent launch will be significantly faster (5–12 seconds) because Windows caches the trust decision.

---

## How It Works

### Step 1 — (Optional) Add Custom PII Entries

Before anonymizing, open the **Custom PII** tab. Type in any name, phrase, or identifier you want to force-replace — for example, a client name the NLP engine might miss — assign it a label, and click **Add Entry**. Click **Save to Project Folder** to save the list alongside your documents so it loads automatically next time.

### Step 2 — Anonymize

1. Open the **Anonymize** tab.
2. Click **Browse** and select your document (`.txt`, `.docx`, or `.pdf`).
3. Choose your output file path.
4. Adjust the confidence threshold and entity types if needed.
5. Click **Run Anonymizer**.

**For Hebrew documents:** After the NLP engine runs, a **Hebrew Review Dialog** may appear listing ambiguous words — Hebrew words that are both common names and everyday nouns (e.g., `דוד` = David / uncle). Review the checklist, tick the words that are genuinely PII, and click **Confirm**.

The tool creates three files:

| File | Description |
|---|---|
| `{filename}_anonymized.txt` | Safe to paste into any cloud AI or translator |
| `{filename}_mapping.json` | Maps every placeholder back to the **original** PII value |
| `{filename}_mapping_translated.json` | Maps every placeholder to a **pre-filled transliteration** — review this in the Translation Map tab before using |

### Step 3 — Process with Cloud AI or Translator

Copy the anonymized document and paste it into ChatGPT, Claude, DeepL, or any other service. Ask it to draft, summarize, translate, or analyze the document. Save the output to a new file.

### Step 4 — (For Translation) Review the Translation Map

If you sent the document for translation, open the **Translation Map** tab:

1. Click **Load Map** and select the `_mapping_translated.json` file.
2. The table shows every placeholder, the original Hebrew value, and the pre-filled English transliteration (e.g., `דוד לוי` → `David Levy`).
3. Edit any cell directly in the table — correct transliterations, adjust name spellings, or update any value that needs changing.
4. Click **Save Map** to write the reviewed file.

> This manual review step is important because PII transliterations can vary (e.g., `שרה` may be `Sarah` or `Sara` depending on the client's preference).

### Step 5 — Restore

1. Open the **Restore** tab.
2. Select the **Restore Mode**:
   - **Restore original PII** — use this when the AI processed the document in the same language (e.g., Hebrew in, Hebrew out).
   - **Restore translated PII** — use this when the document was translated (e.g., Hebrew in, English out).
3. Browse for the **anonymized/translated document** (the AI output file).
4. Browse for the **mapping file** (original or translation map, depending on mode).
5. Choose your output path.
6. Click **Restore**.

Every `{{PLACEHOLDER}}` in the AI's output is replaced with the correct PII value in the target language. The final document is ready to use.

---

## Complete Translation Workflow Example

| Step | Action | File |
|---|---|---|
| 1 | Anonymize Hebrew contract | `contract_anonymized.txt` + `contract_mapping.json` + `contract_mapping_translated.json` |
| 2 | Review translation map | Edit `contract_mapping_translated.json` in Translation Map tab |
| 3 | Upload `contract_anonymized.txt` to DeepL | Get English translation with `{{PERSON_1}}` etc. intact |
| 4 | Save DeepL output | `contract_translated_anon.txt` |
| 5 | Restore with translated map | Load `contract_translated_anon.txt` + `contract_mapping_translated.json` → **final English contract with correct names** |

---

## Batch Processing an Entire Folder

For matters with multiple documents, use the **Batch Folder** tab instead of processing files one at a time:

1. Open the **Batch Folder** tab.
2. Click **Browse Source Folder** and select the folder containing your documents.
3. Click **Browse Output Folder** and select where anonymized files should be saved.
4. Click **Run Batch Anonymize**.

The tool processes every `.txt`, `.docx`, and `.pdf` file in the source folder. Each file receives its own `_anonymized.txt`, `_mapping.json`, and `_mapping_translated.json` output file. A live log shows the result for each file. If a `_custom_pii.json` project file exists in the source folder, it is applied automatically to every file in the batch.

> **Note:** The Hebrew Review Dialog is skipped in batch mode. Ambiguous Hebrew candidates are logged as warnings in the Debug Log tab instead.

---

## Custom PII — Manual Entry & Project Lists

The **Custom PII** tab allows you to manually specify text that should always be anonymized, regardless of what the NLP engine detects. This is particularly useful for:

- Client names that appear in an unusual format or spelling
- Matter-specific terms (company names, property addresses, case references)
- Any text the automatic detection missed

### Adding Entries

| Field | Description |
|---|---|
| **Text** | The exact text to find and replace (case-sensitive) |
| **Label** | The entity type to assign: PERSON, LOCATION, ORGANIZATION, IL_ID, PHONE, EMAIL, DATE, or CUSTOM |

Click **Add Entry** to add it to the list. All occurrences of that text in the document will be replaced with a consistent placeholder (e.g., every instance of `דוד לוי` becomes `{{PERSON_1}}`).

### Project PII Files

Click **Save to Project Folder** to save the current list as `_custom_pii.json` in the same folder as your source documents. The next time you open any document from that folder, the list loads automatically. This allows you to maintain a per-client or per-matter PII list that persists across sessions and applies to every document in that case folder.

---

## Supported PII Entity Types

| Category | Detected Entities |
|---|---|
| **People** | Full names (Hebrew and English), first names, last names |
| **Israeli IDs** | Teudat Zehut (9-digit, Luhn-10 validated) |
| **Phone Numbers** | Israeli mobile (+972, 05x, 07x), Israeli landlines (0x-xxxxxxx) |
| **Locations** | Cities, regions, addresses (Hebrew and English) |
| **Dates** | DD/MM/YYYY, DD.MM.YYYY, and English date formats |
| **Email Addresses** | All standard email formats |
| **Financial** | IBAN (including Israeli IL prefix), bank account numbers, credit cards |
| **Documents** | Passports, driver's licenses, national IDs |
| **Digital** | IP addresses, URLs, cryptocurrency addresses |
| **Medical** | Medical license numbers |
| **Custom** | Any user-defined text via the Custom PII tab |

---

## Hebrew Language Support

The application provides full support for documents written entirely in Hebrew script, including:

- **Multilingual NER model** (`xx_ent_wiki_sm`) — a spaCy model trained on Hebrew Wikipedia, capable of detecting PERSON, LOCATION, and ORGANISATION entities in Hebrew script.
- **Hebrew name dictionary** — 248 common Israeli first names and 210 last names in Hebrew characters, sourced from Israeli Central Bureau of Statistics data.
- **Location dictionary** — 103 entries covering all major Israeli cities, regions, and common legal/institutional terms in Hebrew.
- **Hebrew Review Dialog** — after NLP detection, ambiguous Hebrew words (words that are both names and common nouns) are presented in a checklist for manual confirmation.
- **Translation map with transliterations** — 700+ pre-filled Hebrew-to-English transliterations for names and locations, editable before use.
- **Automatic language detection** — the application detects whether a document is primarily Hebrew or English and applies the appropriate NLP pipeline automatically.
- **RTL text display** — all preview panels switch to right-to-left layout when a Hebrew document is loaded.
- **Hebrew encoding support** — reads UTF-8, Windows-1255, ISO-8859-8, and CP1255 encoded files.

---

## GUI Overview

The application has eight tabs:

### Anonymize Tab
- **File browser** for input document (`.txt`, `.docx`, `.pdf`)
- **Output file** path selector
- **Confidence threshold slider** (0.30–1.00) — lower values catch more PII but may produce false positives; 0.55 is recommended
- **Language override** — Auto (recommended), English only, or Hebrew only
- **Entity type checkboxes** — enable or disable any of the 18 supported PII categories
- **Date granularity setting** — choose how precisely dates are detected: Full (DD/MM/YYYY only), Month+Year, Year-Only, or None; law citations (e.g. "Inheritance Law-1965") are always excluded regardless of this setting
- **Entity Review Dialog toggle** — when enabled, a step-by-step review dialog appears before anonymizing, showing every detected entity with approve/reject/add-to-global/add-to-local options; borderline detections (low confidence) are highlighted in yellow
- **Colour-coded output preview** — each entity type is highlighted in a distinct colour
- **Detection log table** — lists every placeholder, its original value, entity type, confidence score, and detected language

### Restore Tab
- **Restore Mode** radio buttons — original PII or translated PII
- **Anonymized document** browser (the AI/translator output file)
- **Mapping file** browser (`_mapping.json` for original, or `_mapping_translated.json` for translated)
- **Translation map** browser — appears only in "Restore translated PII" mode
- **Output file** path selector
- **Preview panel** showing the restored document with placeholder restoration log

### Translation Map Tab
- **Load Map** button — opens a `_mapping_translated.json` file
- **Editable table** — Placeholder | Entity Type | Original Value | Translated Value (click any cell to edit)
- **Auto-fill suggestions** button — re-runs the transliteration engine on any empty cells
- **Save Map** button — writes the reviewed file back to disk
- Supports direct editing for any PII value that needs a custom transliteration

### Custom PII Tab
- **Text entry field** for the exact phrase to anonymize
- **Label selector** for the entity type to assign
- **Entry list** showing all current custom entries
- **Add / Remove** buttons to manage the list
- **Save to Project Folder** — saves `_custom_pii.json` to the source document folder for automatic reuse

### Batch Folder Tab
- **Source folder** browser — selects the folder of documents to process
- **Output folder** browser — selects where anonymized files are saved
- **Run Batch Anonymize** button
- **Live log** showing each file's status (success, skipped, or error)
- Automatically applies the project `_custom_pii.json` if present in the source folder

### Global Exclusions Tab
- **Persistent cross-matter exclusion list** — items added here are never treated as PII in any document
- **Add / Remove** buttons to manage the list
- **Save** button — writes the list to `~/.pii_processor/global_exclusions.json`
- Items excluded during a run are reported in a **Skip Report popup** at the end of the run (they are not silently dropped)
- Items can also be added to this list directly from the **Entity Review Dialog** during anonymization

### Debug Log Tab
- **Live scrolling log** with colour-coded severity levels (DEBUG / INFO / WARNING / ERROR)
- **Debug Mode toggle** — switches between INFO-only and full DEBUG verbosity
- **Clear**, **Copy Log**, and **Open Log File** buttons
- Log file (`pii_processor.log`) rotates at 2 MB, keeps 3 backups
- Privacy-safe: original PII values are never written to the log

### About Tab
- Version information, NLP engine summary, and **Check for Updates** button

---

## Auto-Updater

The application checks for new versions automatically at startup. If a newer version is available:

1. A green banner appears in the header: **⬆ Update available vX.X.X**
2. Click **Download & Restart** — the new `.exe` is downloaded in the background with a progress bar
3. The application restarts automatically with the new version

You can also check manually at any time via the **About** tab → **↺ Check for Updates**.

> Keep `PII_Processor.exe` in a **writable folder** (e.g., `Documents\PII_Processor\`) so the updater can replace the file in place.

---

## NLP Engine Details

| Component | Details |
|---|---|
| **Framework** | Microsoft Presidio (open-source PII detection) |
| **English NER** | spaCy `en_core_web_sm` — detects names, locations, organisations, dates in English |
| **Hebrew NER** | spaCy `xx_ent_wiki_sm` — multilingual model with Hebrew Wikipedia training |
| **Hebrew dictionary** | Custom `PatternRecognizer` with 700+ Israeli names and 103 locations |
| **Israeli patterns** | Regex-based recognizers for Teudat Zehut, Israeli phone formats, Israeli IBAN |
| **Transliteration** | Built-in dictionary of 700+ Hebrew→English name/location transliterations |
| **RAM usage** | Approximately 400–600 MB while running |
| **Internet required** | No — all processing is fully local (internet only used for update checks) |

---

## File Format Notes

| Format | Read | Write |
|---|---|---|
| `.txt` | Yes (UTF-8, Windows-1255, ISO-8859-8, CP1255) | Yes (UTF-8) |
| `.docx` | Yes | Yes |
| `.pdf` | Yes (text extraction) | No — output as `.txt` |

---

## Output File Naming Convention

For each processed document, the following files are created:

| File | Description |
|---|---|
| `{filename}_anonymized.txt` | The anonymized document, safe for cloud AI or translator |
| `{filename}_mapping.json` | Private mapping of placeholders → original PII values |
| `{filename}_mapping_translated.json` | Mapping of placeholders → pre-filled transliterations (editable) |
| `_custom_pii.json` | (Project folder) The saved custom PII list for this matter |

---

## Privacy & Security

- All PII detection and replacement runs **entirely on your local machine**.
- No document content, no names, no IDs, and no mapping data are transmitted over the internet at any point during anonymization or restoration.
- The `_mapping.json` file contains the original PII values. **Treat it with the same confidentiality as the original document.** Do not share it or upload it to any cloud service.
- The `_mapping_translated.json` file also contains PII values (in transliterated form). Store it securely.
- The `_custom_pii.json` project file also contains real names and terms. Store it securely alongside your matter files.
- The anonymized document is safe to share with cloud AI services, as it contains no real personal data.

---

## Troubleshooting

**The application takes a long time to load on first launch.**
This is expected on the very first run. Windows Defender scans the executable. Subsequent launches will be much faster (5–12 seconds). The splash screen will inform you of progress.

**A name or ID was not detected automatically.**
Try lowering the confidence threshold slider. For names not detected by the NLP engine, use the **Custom PII** tab to add them manually. For Hebrew names not in common use, switch the language override to "Hebrew" to force the Hebrew NER pipeline.

**A Hebrew name was not in the review dialog.**
The review dialog only shows words that appear in the built-in dictionary. For names not in the dictionary, add them manually in the **Custom PII** tab before running the anonymizer.

**The translation map has wrong or missing transliterations.**
Open the **Translation Map** tab, load the `_mapping_translated.json` file, and edit any cell directly. Click **Save Map** when done. The transliterations are pre-filled suggestions only — manual review is always recommended for legal documents.

**"Update check failed" error.**
This is usually a network or firewall issue. The updater uses HTTPS to connect to GitHub. If your network blocks GitHub, you can download new versions manually from the GitHub repository. Check the Debug Log tab for the specific error message.

**The restored document has leftover `{{PLACEHOLDER}}` variables.**
This means the AI or translator modified the placeholder text (e.g., changed `{{PERSON_1}}` to `{{person_1}}`). Check the output file for any altered placeholders and correct them manually before running the restore step.

**PDF text is garbled or missing.**
Some PDFs use image-based scanning rather than embedded text. The tool can only extract text from PDFs that contain actual text layers. For scanned PDFs, convert to `.docx` using Word or Adobe Acrobat first.

**Batch processing skips some files.**
Only `.txt`, `.docx`, and `.pdf` files are processed. Other file types (`.xlsx`, `.msg`, etc.) are skipped and noted in the batch log.

---

## Running from Source

If you prefer to run the Python source directly (e.g., on macOS or Linux):

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Download the spaCy models
python -m spacy download en_core_web_sm
python -m spacy download xx_ent_wiki_sm

# 3. Run the application
python pii_processor_gui.py
```

---

## Version History

| Version | Changes |
|---|---|
| **v0.5.5** | Entity Review Dialog — pre-anonymization step-by-step review of every detected entity (approve/reject/add-to-global/add-to-local); smart date filtering with configurable granularity (Full / Month+Year / Year-Only / None); law citation exclusion (e.g. "Inheritance Law-1965" never treated as a date); possessive trimming ("David's" → only "David" is PII); Global Exclusions tab for cross-matter persistent exclusion list with end-of-run skip report |
| **v0.5.4** | Translation Map tab with editable transliterations; dual-mode Restore (original / translated PII); auto-generated `_mapping_translated.json` on every anonymization; 700+ Hebrew→English transliterations built-in; SSL fix for auto-updater on Windows |
| v0.5.3 | Auto-updater SSL fix; semantic version numbering |
| v5.3 | Hebrew Ambiguity Review Dialog — checklist of ambiguous Hebrew words for manual PII confirmation |
| v5.2 | Debug Log tab; rotating log file; Debug Mode toggle; `requirements.txt` added |
| v5.1 | In-app auto-updater with GitHub Releases integration |
| v5.0 | Custom PII manual entry tab; per-project `_custom_pii.json`; batch folder processing |
| v4.0 | Single-file `.exe` build for fast Windows startup; splash screen |
| v3.0 | Full Hebrew support: multilingual NER, name/location dictionaries, RTL display, Israeli PII patterns |
| v2.0 | GUI application with colour-coded output, detection log, confidence slider, entity type checkboxes |
| v1.0 | Command-line tool, English only |

---

## License

This project is released under the **MIT License**.

```
MIT License

Copyright (c) 2024-2026 17Shaun01

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### Third-Party Licenses

This application bundles the following open-source components:

| Component | License | Notes |
|---|---|---|
| [Microsoft Presidio](https://github.com/microsoft/presidio) | MIT | PII detection engine |
| [spaCy](https://github.com/explosion/spaCy) | MIT | NLP framework |
| [en_core_web_sm](https://github.com/explosion/spacy-models) | MIT | English NER model |
| [xx_ent_wiki_sm](https://github.com/explosion/spacy-models) | MIT | Multilingual NER model (Hebrew) |
| [certifi](https://github.com/certifi/python-certifi) | MPL-2.0 | CA certificate bundle for SSL |
| [python-docx](https://github.com/python-openxml/python-docx) | MIT | Word document support |
| [pdfminer.six](https://github.com/pdfminer/pdfminer.six) | MIT | PDF text extraction |
| Python standard library (tkinter, re, json, etc.) | PSF License | GUI and utilities |
