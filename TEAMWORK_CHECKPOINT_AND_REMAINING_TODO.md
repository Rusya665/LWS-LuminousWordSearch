# LWS 2.0 — Single-Agent Execution Blueprint (Milestones M3 & M4)

**Role of This Document:** Complete, line-by-line execution specification for a **single regular agent** (NO `/teamwork` swarm) to finish the remaining 3 files (`lws_engine/database.py`, `LWS.py`, `web_ui/app.js`) and their unit tests in one pass.

* **Target Repository:** `C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch`
* **Python Test Runner:** `C:\Users\runiza\GoogleProjects\UsefulStuff\.venv\Scripts\python.exe -m pytest`
* **Baseline Status:** Milestones **M1** and **M2** are **100% DONE** on disk (**103 / 103 tests passing**).

---

## Rules for the Executing Agent

1. **DO NOT spawn subagents or use `/teamwork`.** Execute all edits and tests directly in your main thread.
2. **Preserve all 103 existing tests.** Run `C:\Users\runiza\GoogleProjects\UsefulStuff\.venv\Scripts\python.exe -m pytest -q` before and after your changes.
3. Follow the exact specifications in **Step 1**, **Step 2**, **Step 3**, **Step 4**, and **Step 5** below.

---

## What Was Already Completed (M1 & M2 — Do Not Re-Do)

* **`lws_engine/bundler.py` (DONE):** Fixed false-positive `PATENT` classification (`focus2020.pdf` $\neq$ patent), multi-word SI suffix stripping (`_supplementary_information`, `_supp_info`), and `.tif`/`.tiff` image role support.
* **`lws_engine/extractors.py` (DONE):** Fixed DOCX heading sentence indexing and `et al.` sentence-boundary splitting.
* **`lws_engine/thesaurus.py` (DONE):** Replaced silent `except: pass` in WordNet lookup with proper logging.
* **`lws_engine/figure_extractor.py` (DONE):** Added `max_figures` support with safe int coercion, `Path`/`str` handling, `{safe_stem}_p{page}_{idx}.png` filename namespacing, `Chart`/`Scheme`/`Figure 1A-D` caption regexes, CMYK $\rightarrow$ RGB conversion, and harmonized dictionary keys (`figure_index`, `page_number`, `rel_image_path`).
* **`lws_engine/synchronizer.py` (DONE):** Added deleted-file cleanup across SQLite + `references.json` + `.ai_cache/`, `None`-safe sorting in `REFERENCES.md`, and error logging.
* **Existing Passing Test Files (103 tests):**
  * `tests/test_lws_engine.py`
  * `tests/test_m1_core.py`, `tests/test_challenger_m1.py`, `tests/test_challenger_m1_2.py`
  * `tests/test_m2_figures_sync.py`, `tests/test_challenger_m2_figures.py`, `tests/test_challenger_m2_sync.py`

---

## Step 1: Remediate `lws_engine/database.py` (Milestone M3)

File: `C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch\lws_engine\database.py`

### 1.1 Thread Safety, `:memory:` Support, and `RLock` (`lines 8-53`)
* Import `threading`, `logging`, and `re`. Initialize `logger = logging.getLogger(__name__)`.
* In `__init__(self, db_path: str)`:
  * Guard directory creation:
    ```python
    self.db_path = db_path
    if self.db_path != ":memory:":
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    self._lock = threading.RLock()
    self._conn: Optional[sqlite3.Connection] = None
    self._init_db()
    ```
* In `get_connection(self) -> sqlite3.Connection`:
  * Acquire `with self._lock:`.
  * Connect with `check_same_thread=False`:
    ```python
    target_path = ":memory:" if self.db_path == ":memory:" else win_safe_path(self.db_path)
    conn = sqlite3.connect(target_path, timeout=30.0, check_same_thread=False)
    ```
* In `close(self)`:
  * Acquire `with self._lock:`, close `self._conn` (logging any exception via `logger.debug`), and set `self._conn = None`.
  * Note: Calling `get_connection()` after `close()` must transparently reopen `self._conn` (required by `test_challenger_m2_sync.py`).
* Wrap all public database methods (`_init_db`, `upsert_document`, `clear_document_content`, `delete_document_by_rel_path`, `index_sentences`, `index_figures`, `search_fts`, `search_figures`, `get_bundle`, `get_document_by_id`, `get_figures_for_doc`, `list_all_documents`, `save_annotation`, `get_annotations`) in `with self._lock:`.

### 1.2 Add `delete_document_by_rel_path(self, rel_path: str) -> bool` (`after clear_document_content`)
`lws_engine/synchronizer.py:105` checks `if hasattr(self.db, "delete_document_by_rel_path"):`. Implement this method on `LwsDatabase`:
```python
def delete_document_by_rel_path(self, rel_path: str) -> bool:
    """
    Atomically removes a document and its sentences, figures,
    FTS entries, and user annotations by its relative path.
    Returns True if a document was deleted, False if not found.
    """
    if not rel_path:
        return False
    with self._lock:
        conn = self.get_connection()
        with conn:
            cur = conn.cursor()
            norm_p = rel_path.replace("\\", "/")
            cur.execute(
                "SELECT id FROM documents WHERE rel_path = ? OR rel_path = ?",
                (rel_path, norm_p)
            )
            row = cur.fetchone()
            if not row:
                return False
            doc_id = row["id"]

            conn.execute("DELETE FROM sentences_fts WHERE doc_id = ?", (str(doc_id),))
            cur.execute("SELECT id FROM figures WHERE doc_id = ?", (doc_id,))
            fig_ids = [str(r["id"]) for r in cur.fetchall()]
            for fid in fig_ids:
                conn.execute("DELETE FROM figures_fts WHERE figure_id = ?", (fid,))
            conn.execute("DELETE FROM figures_fts WHERE doc_id = ?", (str(doc_id),))
            conn.execute("DELETE FROM figures WHERE doc_id = ?", (doc_id,))
            conn.execute("DELETE FROM user_annotations WHERE doc_id = ?", (doc_id,))
            conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
            return True
```
Also update `clear_document_content(self, doc_id: int)` to include `conn.execute("DELETE FROM figures_fts WHERE doc_id = ?", (str(doc_id),))` so no orphan `figures_fts` rows ever remain.

### 1.3 Harden `index_figures(self, doc_id: int, figures: List[Dict[str, Any]])` (`lines 189-218`)
Accept both canonical keys and alias keys from extractors/tests:
* `figure_index`: `fig.get("figure_index") if fig.get("figure_index") is not None else fig.get("fig_index", 0)` (coerced safely to `int`, default `0`).
* `page_number`: `fig.get("page_number") if fig.get("page_number") is not None else fig.get("page_num", 1)` (coerced safely to `int`, default `1`).
* `rel_image_path`: `str(fig.get("rel_image_path") or fig.get("file_path") or fig.get("image_path") or "").replace("\\", "/")`.
* `caption`: `str(fig.get("caption") or "")`.
* `width`: safe `int(fig.get("width") or 0)`.
* `height`: safe `int(fig.get("height") or 0)`.
* `ocr_text`: `str(fig.get("ocr_text") or "")`.

### 1.4 Fix Boolean & Multi-Word FTS5 Query Parsing in `search_fts` and `search_figures` (`lines 219-328`)
**Current Bug:** `clean_q = query.strip().replace('"', '""'); terms = [f'"{clean_q}"']` wraps the entire query string in quotes, which turns `PDMS AND TPO-L` into a literal phrase search and returns 0 hits when words are non-adjacent.

**Required Implementation (`_build_fts_query(self, query: str, synonyms: Optional[List[str]] = None) -> str`):**
1. Strip `query`. If empty, return `""`.
2. Check if the user wrote an explicit boolean query (contains standalone `AND`, `OR`, `NOT` tokens case-insensitively, or balanced `"..."` quotes):
   * Parse the query into tokens: quoted phrases `"..."`, boolean keywords (`AND`, `OR`, `NOT`), parentheses `(`, `)`, and bare words/hyphenated terms (like `TPO-L`).
   * Wrap bare words/hyphenated terms in double quotes (`"PDMS"`, `"TPO-L"`) so FTS5 never interprets a hyphen as a column subtraction operator (`TPO - L` $\rightarrow$ `no such column: L` error in SQLite FTS5!).
   * Preserve `AND`, `OR`, `NOT` operators between terms (converting lowercase `and`/`or`/`not` to uppercase `AND`/`OR`/`NOT`, and dropping leading/trailing/dangling operators).
3. If the query is a plain single or multi-word query without boolean operators:
   * Match either the exact phrase `"clean_q"` OR all individual words `("word1" AND "word2")`.
   * If `synonyms` are provided, append valid sanitized synonym phrases via `OR`: e.g. `(("word1" AND "word2") OR "syn1" OR "syn2")`.
4. Use `_build_fts_query` in both `search_fts()` and `search_figures()`.
5. In the `except sqlite3.OperationalError as e:` fallback blocks of `search_fts` and `search_figures`, log `logger.debug(f"FTS query fallback for '{query}': {e}")` instead of bare `pass`, and fall back to safe quoted tokens.

---

## Step 2: Remediate `LWS.py` (Milestone M4 — CLI & Pywebview Bridge)

File: `C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch\LWS.py`

1. **`choose_folder_dialog(self)` (`lines 58-101`):**
   * Track `dialog_attempted = False`.
   * If `self.window` is set:
     ```python
     if self.window:
         try:
             dialog_attempted = True
             selected_folders = self.window.create_file_dialog(
                 webview.FileDialog.FOLDER,
                 allow_multiple=True
             )
         except Exception as e:
             logger.warning(f"pywebview folder dialog failed, falling back to tkinter: {e}")
             dialog_attempted = False
             selected_folders = None
     ```
   * Only run the `tkinter` fallback if `not dialog_attempted` (so when a user clicks **Cancel** in the native pywebview dialog, it cleanly returns `added_folders: []` without popping up a second Tkinter window!).
   * Replace `except Exception: pass` on line 87 with `except Exception as e: logger.warning(f"Tkinter folder dialog failed: {e}")`.
2. **`add_watch_folder(self, folder_path: str)` (`lines 103-110`):**
   * Guard empty strings.
   * Replace `if not os.path.exists(win_safe_path(norm_p)):` with `if not os.path.isdir(win_safe_path(norm_p)):` so a regular file path returns `{"status": "error", "message": f"Directory does not exist: {norm_p}"}`.
3. **`sync_folders(self)` (`lines 118-139`):**
   * Wrap the per-folder sync inside the `for f in folders:` loop in a `try / except Exception as e:` block:
     ```python
     for f in folders:
         try:
             res = self.synchronizer.sync_watch_folder(f)
             total_proc += res.get("processed", 0)
             total_skip += res.get("skipped", 0)
             total_err += res.get("errors", 0)
             details.append(res)
         except Exception as e:
             logger.error(f"Error syncing watch folder '{f}': {e}")
             total_err += 1
             details.append({"watch_folder": f, "processed": 0, "skipped": 0, "errors": 1, "error_message": str(e)})
     ```
4. **`search_query(self, query: str, use_synonyms: bool = True, role_filter: str = "ALL", regex_mode: bool = False)` (`lines 141-160`):**
   * Support an optional `regex_mode: bool = False` parameter so the UI regex toggle works end-to-end!
   * When `regex_mode=True` and `query.strip()` is non-empty:
     * Compile `pattern = re.compile(query, re.IGNORECASE)` (if `re.error` is raised, return `{"query": query, "synonyms": {}, "hits": [], "error": f"Invalid regex: {e}"}`).
     * Query sentences from `self.db` (or add a helper `self.db.search_regex(pattern, role_filter=..., limit=80)` on `LwsDatabase` that scans `sentences_fts JOIN documents` and highlights matches with `<mark class="hit-direct">...</mark>`).
5. **`get_figures(self, query: str = "")` (`lines 192-200`):**
   * Change line 198 to:
     ```python
     rel_p = f.get("rel_image_path") or ""
     item["rel_image_path"] = str(rel_p).replace("\\", "/")
     ```
6. **`open_native_file(self, file_path: str) -> Dict[str, Any]` (`lines 207-216`):**
   * Validate non-empty path and check `os.path.exists(win_safe_path(safe_p))`. If missing, return `{"status": "error", "message": f"File does not exist: {safe_p}"}`.
   * Block dangerous executable extensions `{".exe", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".msi", ".scr", ".com", ".hta"}` (if `Path(safe_p).suffix.lower() in BLOCKED_EXTENSIONS`, return `{"status": "error", "message": f"Blocked executable file type: {Path(safe_p).suffix}"}`).
   * Launch viewer inside `try / except Exception as e:` and return `{"status": "ok", "path": safe_p}` (or `{"status": "error", "message": str(e)}`).
7. **CLI `main()` Hardening (`lines 218-271`):**
   * Track `cli_errors = 0`.
   * When processing `args.dir`: if `not os.path.isdir(win_safe_path(clean_p))`, print a clear warning to `sys.stderr` (`print(f"[LWS] Warning: Directory does not exist or is invalid: {clean_p}", file=sys.stderr)`) and increment `cli_errors += 1`.
   * In `if args.headless_sync:`: wrap in `try / finally: db.close()`. Sum `res.get("errors", 0)` across all synced folders into `cli_errors`. Exit with `sys.exit(1 if cli_errors > 0 else 0)`.
   * For GUI mode, wrap `webview.start(debug=False)` in `try / finally: db.close()`.

---

## Step 3: Remediate `web_ui/app.js` (Milestone M4 — Frontend Guards & Wiring)

File: `C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch\web_ui\app.js`

1. **Wire Up `toggleRegexBtn` (`after line 120`):**
   ```javascript
   if (toggleRegexBtn) {
     toggleRegexBtn.addEventListener('click', () => {
       state.isRegex = !state.isRegex;
       toggleRegexBtn.classList.toggle('active', state.isRegex);
       performSearch();
     });
   }
   ```
   And in `performSearch()` (`line 371`), pass `state.isRegex`:
   ```javascript
   const resp = await api.search_query(query, state.useSynonyms, state.activeRole, state.isRegex);
   ```
2. **Wire Up `Ctrl+F` Shortcut (`inside setupEventListeners()`):**
   ```javascript
   document.addEventListener('keydown', (e) => {
     if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
       e.preventDefault();
       if (searchInput) {
         searchInput.focus();
         searchInput.select();
       }
     }
   });
   ```
3. **Initialize & Guard `state.currentBundleDocs` (`lines 3-16` and `150-162`):**
   * Add `currentBundleDocs: [],` to the initial `state` object at line 10.
   * In `openMainNativeBtn` and `openSiNativeBtn` click listeners (`lines 150-162`), guard with `if (!Array.isArray(state.currentBundleDocs) || state.currentBundleDocs.length === 0) return;`.

---

## Step 4: Add Unit & Integration Tests for M3 & M4

Create `C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch\tests\test_m3_m4_database_cli_bridge.py` testing:
1. **M3 Database Tests:**
   * Multi-threaded concurrent read/write/search across 8 threads on a shared `LwsDatabase` instance (`check_same_thread=False` + `RLock`, verifying 0 `ProgrammingError` / `OperationalError`).
   * `:memory:` database initialization and querying.
   * `delete_document_by_rel_path` cascading across `documents`, `sentences_fts`, `figures`, `figures_fts`, and `user_annotations` (including backslash/forward-slash normalization).
   * `index_figures` handling both canonical (`figure_index`, `page_number`, `rel_image_path`) and alias (`fig_index`, `page_num`, `file_path`) keys.
   * Boolean FTS5 search: `PDMS AND TPO-L` (matching non-adjacent words in a sentence), `PDMS OR hydrogel`, `PDMS NOT toxic`, hyphenated terms (`TPO-L`), and malformed FTS5 syntax fallback (`"unclosed quote` or `AND OR`).
2. **M4 CLI & Bridge Tests:**
   * `LwsApi.add_watch_folder` rejecting non-existent paths and regular file paths (`status == "error"`), accepting valid directories (`status == "ok"`).
   * `LwsApi.choose_folder_dialog` returning `added_folders: []` without calling Tkinter when `window.create_file_dialog` returns `None` (user cancelled).
   * `LwsApi.sync_folders` isolating an exception in one bad watch folder while still syncing a second valid watch folder.
   * `LwsApi.open_native_file` blocking `.exe` / `.bat` / `.ps1` files and non-existent files with `status == "error"`.
   * `LwsApi.search_query` with `regex_mode=True` matching regex patterns (e.g. `\b\d+\s*nm\b`) and gracefully handling invalid regexes (`[unclosed`).
   * CLI `--headless-sync` returning exit code `0` on valid directory sync and exit code `1` when passed `--dir /nonexistent/invalid/path`.

---

## Step 5: Cleanup & Verification

1. Run the full test suite:
   ```powershell
   cd C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch
   C:\Users\runiza\GoogleProjects\UsefulStuff\.venv\Scripts\python.exe -m pytest -v
   ```
2. Verify that `C:\Users\runiza\GoogleProjects\OpenCAL-alternative\.agents` and `C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch\.agents` are removed so neither repository is polluted by leftover teamwork scratch files:
   ```powershell
   Remove-Item -Recurse -Force "C:\Users\runiza\GoogleProjects\OpenCAL-alternative\.agents" -ErrorAction SilentlyContinue
   Remove-Item -Recurse -Force "C:\Users\runiza\GoogleProjects\LWS-LuminousWordSearch\.agents" -ErrorAction SilentlyContinue
   ```
