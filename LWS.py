"""
LWS (LuminousWordSearch) 2.0
============================
Modern Literature, Patent & SI Intelligence Suite.
Combines pypdfium2 high-speed extraction, SQLite FTS5 search, scientific & WordNet thesaurus,
automatic figure and caption harvesting, companion document bundling (Main + SI),
and a sleek dark glassmorphic interface built with pywebview.
"""

import os
import sys
import argparse
import subprocess
import logging
import re
import base64
from pathlib import Path
from typing import Dict, Any, List, Optional

import webview

from lws_engine.config import ConfigManager, win_safe_path
from lws_engine.database import LwsDatabase
from lws_engine.synchronizer import LwsSynchronizer
from lws_engine.thesaurus import ThesaurusManager

logger = logging.getLogger(__name__)

BLOCKED_EXTENSIONS = {".exe", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".msi", ".scr", ".com", ".hta"}


class LwsApi:
    """JavaScript API bridge exposed to the Web UI via window.pywebview.api."""

    def __init__(self, config_mgr: ConfigManager, db: LwsDatabase, synchronizer: LwsSynchronizer, thesaurus: ThesaurusManager):
        # Leading underscores are critical: pywebview's API generator recursively inspects
        # all public attributes. Prefixing internal backend objects with '_' prevents pywebview
        # from traversing into Window / COM WebView2 controllers or queues.
        self._config_mgr = config_mgr
        self._db = db
        self._synchronizer = synchronizer
        self._thesaurus = thesaurus
        self._window: Optional[webview.Window] = None

    def set_window(self, window: webview.Window):
        self._window = window

    def get_workspace_state(self) -> Dict[str, Any]:
        """Returns watched folders and the cached list of documents."""
        folders = self._config_mgr.get_watch_folders()
        docs = self._db.list_all_documents(limit=1000)

        folder_stats = []
        for f in folders:
            count = sum(1 for d in docs if d.get("watch_folder") == f)
            folder_stats.append({
                "path": f,
                "exists": os.path.exists(win_safe_path(f)),
                "doc_count": count
            })

        return {
            "watch_folders": folders,
            "folder_stats": folder_stats,
            "documents": docs
        }

    def choose_folder_dialog(self) -> Dict[str, Any]:
        """
        Opens a native Windows Folder Picker dialog (NO browser prompt).
        Supports choosing directory and immediately adds it to watched folders.
        """
        selected_folders = None
        dialog_attempted = False

        if self._window and hasattr(self._window, "create_file_dialog"):
            try:
                dialog_attempted = True
                selected_folders = self._window.create_file_dialog(
                    webview.FileDialog.FOLDER,
                    allow_multiple=True
                )
            except Exception as e:
                logger.warning(f"pywebview folder dialog failed, falling back to tkinter: {e}")
                dialog_attempted = False
                selected_folders = None

        if not dialog_attempted and not selected_folders:
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                chosen = filedialog.askdirectory(title="Select Folder to Watch in LWS")
                root.destroy()
                if chosen:
                    selected_folders = [chosen]
            except Exception as e:
                logger.warning(f"Tkinter folder dialog failed: {e}")

        added = []
        if selected_folders:
            for folder in selected_folders:
                folder_str = str(folder).strip()
                if folder_str and os.path.isdir(win_safe_path(folder_str)):
                    self._config_mgr.add_watch_folder(folder_str)
                    added.append(folder_str)

        return {
            "status": "ok",
            "added_folders": added,
            "watch_folders": self._config_mgr.get_watch_folders()
        }

    def add_watch_folder(self, folder_path: str) -> Dict[str, Any]:
        """Adds a folder path string and registers it."""
        if not folder_path or not folder_path.strip():
            return {"status": "error", "message": "Folder path cannot be empty."}
        norm_p = os.path.abspath(folder_path.strip().strip('"').strip("'"))
        if not os.path.isdir(win_safe_path(norm_p)):
            return {"status": "error", "message": f"Directory does not exist: {norm_p}"}

        self._config_mgr.add_watch_folder(norm_p)
        return {"status": "ok", "watch_folders": self._config_mgr.get_watch_folders()}

    def remove_watch_folder(self, folder_path: str) -> Dict[str, Any]:
        """Removes a watched folder."""
        norm_p = os.path.abspath(folder_path.strip())
        self._config_mgr.remove_watch_folder(norm_p)
        return {"status": "ok", "watch_folders": self._config_mgr.get_watch_folders()}

    def sync_folders(self) -> Dict[str, Any]:
        """Synchronizes all watched directories."""
        folders = self._config_mgr.get_watch_folders()
        total_proc = 0
        total_skip = 0
        total_err = 0
        details = []

        for f in folders:
            try:
                res = self._synchronizer.sync_watch_folder(f)
                total_proc += res.get("processed", 0)
                total_skip += res.get("skipped", 0)
                total_err += res.get("errors", 0)
                details.append(res)
            except Exception as e:
                logger.error(f"Error syncing watch folder '{f}': {e}")
                total_err += 1
                details.append({"watch_folder": f, "processed": 0, "skipped": 0, "errors": 1, "error_message": str(e)})

        return {
            "status": "ok",
            "processed": total_proc,
            "skipped": total_skip,
            "errors": total_err,
            "details": details
        }

    def search_query(
        self,
        query: str,
        use_synonyms: bool = True,
        role_filter: str = "ALL",
        regex_mode: bool = False
    ) -> Dict[str, Any]:
        """
        Executes a search query across all indexed sentences, supporting
        FTS5 with synonym expansion or regular expression mode.
        """
        if not query or not query.strip():
            return {"query": query, "synonyms": {}, "hits": []}

        q = query.strip()
        if regex_mode:
            try:
                pattern = re.compile(q, re.IGNORECASE)
            except re.error as e:
                return {"query": query, "synonyms": {}, "hits": [], "error": f"Invalid regex: {e}"}

            hits = self._db.search_regex(
                pattern=pattern,
                role_filter=role_filter if role_filter != "ALL" else None,
                limit=80
            )
            return {
                "query": query,
                "synonyms": {},
                "hits": hits
            }

        synonym_map = {}
        all_syns = []
        if use_synonyms and q:
            synonym_map = self._thesaurus.expand_query(q, use_wordnet=True, use_scientific=True)
            for s_list in synonym_map.values():
                all_syns.extend(s_list)

        hits = self._db.search_fts(
            query=q,
            synonyms=all_syns if use_synonyms else None,
            role_filter=role_filter if role_filter != "ALL" else None,
            limit=80
        )
        return {
            "query": query,
            "synonyms": synonym_map,
            "hits": hits
        }

    def get_document_bundle(self, bundle_id: str) -> Dict[str, Any]:
        """Retrieves all companion files in a bundle and their markdown text content."""
        bundle_docs = self._db.get_bundle(bundle_id)
        contents = {}
        for doc in bundle_docs:
            md_rel = doc.get("md_path")
            watch_f = doc.get("watch_folder")
            if md_rel and watch_f:
                full_md = os.path.join(watch_f, md_rel)
                if os.path.exists(win_safe_path(full_md)):
                    try:
                        with open(win_safe_path(full_md), "r", encoding="utf-8", errors="replace") as f:
                            contents[doc["id"]] = f.read()
                    except Exception:
                        contents[doc["id"]] = ""

        main_doc_id = bundle_docs[0]["id"] if bundle_docs else None
        notes = ""
        if main_doc_id:
            annos = self._db.get_annotations(main_doc_id)
            if annos:
                notes = annos[-1]["note"]

        return {
            "bundle_id": bundle_id,
            "documents": bundle_docs,
            "contents": contents,
            "notes": notes
        }

    def get_figures(self, query: str = "") -> List[Dict[str, Any]]:
        """Returns figures matching query or recent figures with inline base64 data URIs."""
        figs = self._db.search_figures(query, limit=50)
        formatted = []
        for f in figs:
            item = dict(f)
            rel_p = f.get("rel_image_path") or ""
            item["rel_image_path"] = str(rel_p).replace("\\", "/")
            watch_f = f.get("watch_folder") or ""

            # Construct absolute paths
            full_img_path = ""
            if watch_f and rel_p:
                cand = os.path.join(watch_f, rel_p)
                if os.path.exists(win_safe_path(cand)):
                    full_img_path = os.path.abspath(cand)
            if not full_img_path and os.path.isabs(str(rel_p)) and os.path.exists(win_safe_path(str(rel_p))):
                full_img_path = os.path.abspath(str(rel_p))

            item["full_image_path"] = full_img_path
            doc_rel = f.get("rel_path") or ""
            item["full_doc_path"] = os.path.abspath(os.path.join(watch_f, doc_rel)) if (watch_f and doc_rel) else ""

            # Provide base64 data URI for zero-CORS WebView2 rendering
            item["data_url"] = ""
            if full_img_path and os.path.exists(win_safe_path(full_img_path)):
                try:
                    fsize = os.path.getsize(win_safe_path(full_img_path))
                    if fsize <= 3 * 1024 * 1024:
                        ext = Path(full_img_path).suffix.lower()
                        mime = "image/jpeg" if ext in {".jpg", ".jpeg"} else "image/png"
                        with open(win_safe_path(full_img_path), "rb") as img_file:
                            b64 = base64.b64encode(img_file.read()).decode("ascii")
                            item["data_url"] = f"data:{mime};base64,{b64}"
                except Exception as e:
                    logger.debug(f"Failed to encode figure '{full_img_path}': {e}")

            formatted.append(item)
        return formatted

    def get_figure_data(self, figure_id: int) -> Dict[str, Any]:
        """Loads and returns base64 data URI for an individual figure on demand."""
        try:
            conn = self._db.get_connection()
            cur = conn.cursor()
            cur.execute("""
                SELECT f.*, d.watch_folder, d.rel_path
                FROM figures f
                JOIN documents d ON f.doc_id = d.id
                WHERE f.id = ?
            """, (figure_id,))
            row = cur.fetchone()
            if not row:
                return {"status": "error", "message": "Figure not found"}

            watch_f = row["watch_folder"] or ""
            rel_p = row["rel_image_path"] or ""
            full_img_path = os.path.join(watch_f, rel_p) if (watch_f and rel_p) else ""
            if not os.path.exists(win_safe_path(full_img_path)):
                return {"status": "error", "message": "Figure file not found on disk"}

            ext = Path(full_img_path).suffix.lower()
            mime = "image/jpeg" if ext in {".jpg", ".jpeg"} else "image/png"
            with open(win_safe_path(full_img_path), "rb") as img_file:
                b64 = base64.b64encode(img_file.read()).decode("ascii")
                data_url = f"data:{mime};base64,{b64}"

            return {
                "status": "ok",
                "figure_id": figure_id,
                "data_url": data_url,
                "full_image_path": os.path.abspath(full_img_path),
                "full_doc_path": os.path.abspath(os.path.join(watch_f, row["rel_path"] or ""))
            }
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def save_document_note(self, doc_id: int, note: str) -> Dict[str, Any]:
        """Saves a user annotation for a document."""
        row_id = self._db.save_annotation(doc_id=doc_id, page_number=1, sentence_index=0, note=note)
        return {"status": "ok", "annotation_id": row_id}

    def open_native_file(self, file_path: str) -> Dict[str, Any]:
        """Opens a PDF, Word doc, or folder in the default system viewer safely."""
        if not file_path or not file_path.strip():
            return {"status": "error", "message": "File path cannot be empty."}
        safe_p = os.path.abspath(file_path.strip())
        if not os.path.exists(win_safe_path(safe_p)):
            return {"status": "error", "message": f"File does not exist: {safe_p}"}

        ext = Path(safe_p).suffix.lower()
        if ext in BLOCKED_EXTENSIONS:
            return {"status": "error", "message": f"Blocked executable file type: {ext}"}

        try:
            if os.name == "nt":
                os.startfile(safe_p)
            elif sys.platform == "darwin":
                subprocess.run(["open", safe_p], check=False)
            else:
                subprocess.run(["xdg-open", safe_p], check=False)
            return {"status": "ok", "path": safe_p}
        except Exception as e:
            logger.error(f"Error opening native file '{safe_p}': {e}")
            return {"status": "error", "message": str(e)}


def main():
    parser = argparse.ArgumentParser(description="LWS 2.0 - Luminous Word Search")
    parser.add_argument(
        "--dir",
        type=str,
        nargs="+",
        help="Add or synchronize one or more directories (space-separated or comma-separated)"
    )
    parser.add_argument("--headless-sync", action="store_true", help="Run synchronizer headlessly and exit")
    parser.add_argument("--no-figures", action="store_true", help="Disable figure extraction for faster indexing")
    args = parser.parse_args()

    config_mgr = ConfigManager()
    cli_errors = 0

    # Handle multiple directories from CLI if provided
    if args.dir:
        for entry in args.dir:
            for sub_p in entry.split(","):
                clean_p = sub_p.strip().strip('"').strip("'")
                if clean_p:
                    if os.path.isdir(win_safe_path(clean_p)):
                        config_mgr.add_watch_folder(os.path.abspath(clean_p))
                    else:
                        print(f"[LWS] Warning: Directory does not exist or is invalid: {clean_p}", file=sys.stderr)
                        cli_errors += 1

    db_path = config_mgr.get_database_path()
    db = LwsDatabase(db_path)
    thesaurus = ThesaurusManager(enable_wordnet=True)
    synchronizer = LwsSynchronizer(db=db, extract_figures=not args.no_figures)

    if args.headless_sync:
        print("[LWS] Running headless synchronization...")
        try:
            folders = config_mgr.get_watch_folders()
            for f in folders:
                print(f"[LWS] Syncing: {f}")
                try:
                    res = synchronizer.sync_watch_folder(f)
                    print(f"[LWS] Done: {res['processed']} processed, {res['skipped']} up to date, {res['errors']} errors.")
                    cli_errors += res.get("errors", 0)
                except Exception as e:
                    print(f"[LWS] Error syncing '{f}': {e}", file=sys.stderr)
                    cli_errors += 1
        finally:
            db.close()
        sys.exit(1 if cli_errors > 0 else 0)

    # Launch GUI
    api = LwsApi(config_mgr, db, synchronizer, thesaurus)
    ui_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web_ui")
    index_html = os.path.join(ui_dir, "index.html")

    try:
        window = webview.create_window(
            title="LWS 2.0 — Luminous Word Search",
            url=index_html,
            js_api=api,
            width=1340,
            height=860,
            min_size=(1024, 700),
            background_color="#080B11"
        )
        api.set_window(window)
        webview.start(debug=False)
    finally:
        db.close()


if __name__ == "__main__":
    main()
