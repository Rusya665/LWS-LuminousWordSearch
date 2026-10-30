"""
lws_engine.synchronizer
=======================
Handles recursive multi-folder discovery, incremental change detection,
content & figure extraction, .ai_cache generation, and database indexing.
"""

from typing import Dict, Any, List, Optional, Callable, Set, Tuple
import os
import sys
import json
import logging
import datetime
from pathlib import Path

from .config import win_safe_path, to_file_uri
from .extractors import (
    extract_pdf,
    extract_docx,
    extract_pptx,
    split_into_sentences,
    extract_doi,
    extract_abstract,
    parse_filename_metadata,
)
from .figure_extractor import extract_figures_from_pdf, extract_figures_from_docx, extract_figures_from_pptx
from .bundler import bundle_document_record
from .database import LwsDatabase

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".pptx", ".txt", ".md"}
IGNORED_FILES = {"REFERENCES.md", "README.md", "references.json"}


def _strip_long_prefix(path_str: str) -> str:
    """Strips Windows \\?\\ or //?/ prefix for clean relative path calculations."""
    if path_str.startswith("\\\\?\\") or path_str.startswith("//?/"):
        return path_str[4:]
    return path_str


class LwsSynchronizer:
    """Synchronizes one or more watch directories into SQLite, .ai_cache, and references.json."""

    def __init__(self, db: LwsDatabase, extract_figures: bool = True):
        self.db = db
        self.extract_figures = extract_figures

    def scan_directory(self, watch_folder: str) -> List[Dict[str, Any]]:
        """
        Recursively scans watch_folder for all supported documents,
        preserving nested subfolder hierarchy.
        """
        watch_path_str = _strip_long_prefix(os.path.abspath(watch_folder))
        discovered = []

        if not os.path.exists(win_safe_path(watch_path_str)):
            return discovered

        for root, dirs, files in os.walk(win_safe_path(watch_path_str)):
            # Skip hidden and cache folders
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in {"node_modules", "__pycache__", "venv"}]

            clean_root = _strip_long_prefix(root)

            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if (
                    ext in SUPPORTED_EXTENSIONS
                    and not f.startswith("~$")
                    and not f.startswith(".")
                    and f not in IGNORED_FILES
                ):
                    full_p = os.path.join(clean_root, f)
                    try:
                        st = os.stat(win_safe_path(full_p))
                        # Compute path relative to watch_path
                        rel_path = os.path.relpath(full_p, watch_path_str).replace("\\", "/")
                        category = os.path.dirname(rel_path).replace("\\", "/")
                        if not category:
                            category = "Root"

                        discovered.append({
                            "watch_folder": watch_path_str,
                            "full_path": full_p,
                            "rel_path": rel_path,
                            "category": category,
                            "filename": f,
                            "extension": ext,
                            "file_size": st.st_size,
                            "mtime": st.st_mtime,
                        })
                    except Exception as e:
                        logger.warning(f"Could not stat file '{full_p}': {e}")
                        continue

        return discovered

    def _delete_document_record(self, rel_path: str, doc_id: Optional[int] = None) -> bool:
        """
        Removes a document from SQLite (sentences_fts, figures_fts, figures, user_annotations, documents).
        Delegates to self.db.delete_document_by_rel_path if available; otherwise executes direct DB deletion.
        """
        if hasattr(self.db, "delete_document_by_rel_path"):
            try:
                return self.db.delete_document_by_rel_path(rel_path)
            except Exception as e:
                logger.warning(f"delete_document_by_rel_path failed for '{rel_path}': {e}")

        conn = self.db.get_connection()
        with conn:
            cur = conn.cursor()
            if doc_id is None:
                cur.execute("SELECT id FROM documents WHERE rel_path = ?", (rel_path,))
                row = cur.fetchone()
                if row:
                    doc_id = row["id"]
            if doc_id is not None:
                self.db.clear_document_content(doc_id)
                conn.execute("DELETE FROM user_annotations WHERE doc_id = ?", (doc_id,))
                conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
            else:
                conn.execute("DELETE FROM documents WHERE rel_path = ?", (rel_path,))
        return True

    def _reconcile_deletions(
        self,
        watch_path: str,
        ref_data: Dict[str, Any],
        discovered_rel_paths: Set[str]
    ) -> int:
        """
        Detects files previously recorded in DB or references.json that no longer exist on disk.
        Deletes orphan database records, prunes references.json, and removes orphan .ai_cache markdown & figures.
        """
        conn = self.db.get_connection()
        cur = conn.cursor()
        cur.execute(
            "SELECT id, rel_path, md_path FROM documents WHERE watch_folder = ? COLLATE NOCASE OR watch_folder = ? COLLATE NOCASE OR watch_folder = ? COLLATE NOCASE",
            (watch_path, watch_path.replace("\\", "/"), watch_path.replace("/", "\\"))
        )
        db_rows = cur.fetchall()
        db_docs = {r["rel_path"]: dict(r) for r in db_rows}

        all_known_paths = set(ref_data.keys()) | set(db_docs.keys())
        orphan_paths = [
            p for p in all_known_paths
            if p not in discovered_rel_paths and not os.path.exists(win_safe_path(os.path.join(watch_path, p)))
        ]

        deleted_count = 0
        for p in orphan_paths:
            # 1. Clean markdown cache file
            candidate_md_paths = set()
            if p in db_docs and db_docs[p].get("md_path"):
                candidate_md_paths.add(db_docs[p]["md_path"])
            if p in ref_data and ref_data[p].get("md_cache_rel"):
                candidate_md_paths.add(ref_data[p]["md_cache_rel"])
            candidate_md_paths.add(os.path.join(".ai_cache", os.path.splitext(p)[0] + ".md").replace("\\", "/"))

            for md_rel in candidate_md_paths:
                md_full = os.path.join(watch_path, md_rel)
                if os.path.exists(win_safe_path(md_full)):
                    try:
                        os.remove(win_safe_path(md_full))
                        logger.info(f"Removed orphan cache markdown: {md_full}")
                    except Exception as e:
                        logger.warning(f"Failed to remove orphan cache file '{md_full}': {e}")

            # 2. Clean cached figure files and database entries
            doc_id = db_docs[p]["id"] if p in db_docs else None
            if doc_id is not None:
                try:
                    cur.execute("SELECT rel_image_path FROM figures WHERE doc_id = ?", (doc_id,))
                    fig_rows = cur.fetchall()
                    for fig_row in fig_rows:
                        fig_rel = fig_row["rel_image_path"]
                        if fig_rel:
                            # Verify no other document references this exact image file
                            cur.execute(
                                "SELECT count(*) as cnt FROM figures WHERE rel_image_path = ? AND doc_id != ?",
                                (fig_rel, doc_id)
                            )
                            cnt_row = cur.fetchone()
                            if cnt_row and cnt_row["cnt"] == 0:
                                candidate_fig_paths = [
                                    os.path.join(watch_path, fig_rel),
                                    os.path.join(watch_path, ".ai_cache", fig_rel),
                                    os.path.join(watch_path, ".ai_cache", "figures", os.path.basename(fig_rel)),
                                ]
                                for fig_full in candidate_fig_paths:
                                    if os.path.exists(win_safe_path(fig_full)):
                                        try:
                                            os.remove(win_safe_path(fig_full))
                                            logger.info(f"Removed orphan figure cache: {fig_full}")
                                            break
                                        except Exception as e:
                                            logger.warning(f"Failed to remove orphan figure '{fig_full}': {e}")
                except Exception as e:
                    logger.warning(f"Error querying figures for orphan document '{p}': {e}")

            self._delete_document_record(p, doc_id=doc_id)

            # 3. Clean references.json dict
            if p in ref_data:
                del ref_data[p]

            deleted_count += 1
            logger.info(f"Reconciled deletion of orphan document: {p}")

        # Clean empty directories inside .ai_cache (protecting base cache and figures directories)
        ai_cache_dir = os.path.abspath(os.path.join(watch_path, ".ai_cache"))
        figures_cache_dir = os.path.abspath(os.path.join(ai_cache_dir, "figures"))
        protected_dirs = {
            ai_cache_dir,
            figures_cache_dir,
            _strip_long_prefix(ai_cache_dir),
            _strip_long_prefix(figures_cache_dir),
        }
        if os.path.exists(win_safe_path(ai_cache_dir)):
            for root, dirs, files in os.walk(win_safe_path(ai_cache_dir), topdown=False):
                clean_r = _strip_long_prefix(os.path.abspath(root))
                if clean_r not in protected_dirs:
                    try:
                        if not os.listdir(win_safe_path(clean_r)):
                            os.rmdir(win_safe_path(clean_r))
                    except Exception:
                        pass

        return deleted_count

    def reconcile_deleted_files(
        self,
        watch_folder: str,
        ref_data: Optional[Dict[str, Any]] = None,
        discovered_rel_paths: Optional[Set[str]] = None
    ) -> int:
        """
        Public API to reconcile deletions for a watch folder.
        Removes orphan documents from SQLite and disk caches, pruning references.json.
        """
        watch_path = _strip_long_prefix(os.path.abspath(watch_folder))
        ref_json_path = os.path.join(watch_path, "references.json")
        save_ref = False
        if ref_data is None:
            save_ref = True
            ref_data = {}
            if os.path.exists(win_safe_path(ref_json_path)):
                try:
                    with open(win_safe_path(ref_json_path), "r", encoding="utf-8") as rf:
                        ref_data = json.load(rf)
                except Exception as e:
                    logger.warning(f"Failed to load references.json from {ref_json_path}: {e}")
                    ref_data = {}

        if discovered_rel_paths is None:
            discovered = self.scan_directory(watch_path)
            discovered_rel_paths = {item["rel_path"] for item in discovered}

        deleted = self._reconcile_deletions(watch_path, ref_data, discovered_rel_paths)

        if save_ref and deleted > 0:
            try:
                with open(win_safe_path(ref_json_path), "w", encoding="utf-8") as wf:
                    json.dump(ref_data, wf, indent=2, ensure_ascii=False)
            except Exception as e:
                logger.warning(f"Failed to write references.json at {ref_json_path}: {e}")
            self._write_references_markdown(watch_path, ref_data)

        return deleted

    def sync_watch_folder(
        self,
        watch_folder: str,
        progress_cb: Optional[Callable[[int, int, str, str], None]] = None,
        force_reindex: bool = False
    ) -> Dict[str, Any]:
        """
        Syncs a single watch folder:
        - Detects modified/new files.
        - Reconciles deleted/orphan files from DB and cache.
        - Writes markdown to `.ai_cache/`.
        - Updates SQLite DB and `references.json`.
        """
        watch_path = _strip_long_prefix(os.path.abspath(watch_folder))
        ai_cache_dir = os.path.join(watch_path, ".ai_cache")
        figures_cache_dir = os.path.join(ai_cache_dir, "figures")
        os.makedirs(win_safe_path(ai_cache_dir), exist_ok=True)
        os.makedirs(win_safe_path(figures_cache_dir), exist_ok=True)

        ref_json_path = os.path.join(watch_path, "references.json")
        ref_data: Dict[str, Any] = {}
        if os.path.exists(win_safe_path(ref_json_path)):
            try:
                with open(win_safe_path(ref_json_path), "r", encoding="utf-8") as rf:
                    ref_data = json.load(rf)
            except Exception as e:
                logger.warning(f"Failed to load references.json from {ref_json_path}: {e}")
                ref_data = {}

        discovered = self.scan_directory(watch_path)
        discovered_rel_paths = {item["rel_path"] for item in discovered}

        # Deletion reconciliation
        deleted = self._reconcile_deletions(watch_path, ref_data, discovered_rel_paths)

        total = len(discovered)
        processed = 0
        skipped = 0
        errors = 0

        for idx, item in enumerate(discovered):
            rel_path = item["rel_path"]
            full_path = item["full_path"]
            filename = item["filename"]
            ext = item["extension"]
            mtime = item["mtime"]
            fsize = item["file_size"]

            if progress_cb:
                progress_cb(idx + 1, total, filename, "Checking...")

            existing = ref_data.get(rel_path)
            # Incremental check
            if not force_reindex and existing:
                if existing.get("file_size") == fsize and abs(existing.get("mtime", 0) - mtime) < 0.1:
                    # If figure extraction is enabled, backfill figures if not already extracted
                    if self.extract_figures and ext in {".pdf", ".docx", ".pptx"} and not existing.get("figures_extracted"):
                        try:
                            conn = self.db.get_connection()
                            cur = conn.cursor()
                            cur.execute("SELECT id FROM documents WHERE watch_folder = ? AND rel_path = ?", (watch_path, rel_path))
                            doc_row = cur.fetchone()
                            if doc_row:
                                doc_id = doc_row["id"]
                                if progress_cb:
                                    progress_cb(idx + 1, total, filename, "Extracting figures...")
                                extracted_figs = []
                                if ext == ".pdf":
                                    extracted_figs = extract_figures_from_pdf(full_path, figures_cache_dir, max_figures=30)
                                elif ext == ".docx":
                                    extracted_figs = extract_figures_from_docx(full_path, figures_cache_dir)
                                elif ext == ".pptx":
                                    extracted_figs = extract_figures_from_pptx(full_path, figures_cache_dir)

                                if extracted_figs:
                                    for fig in extracted_figs:
                                        if "figure_index" not in fig and "fig_index" in fig:
                                            fig["figure_index"] = fig["fig_index"]
                                        if "page_number" not in fig and "page_num" in fig:
                                            fig["page_number"] = fig["page_num"]
                                        if not fig.get("rel_image_path") and fig.get("file_path"):
                                            try:
                                                fig["rel_image_path"] = os.path.relpath(
                                                    _strip_long_prefix(str(fig["file_path"])),
                                                    _strip_long_prefix(str(watch_path))
                                                ).replace("\\", "/")
                                            except Exception:
                                                fig["rel_image_path"] = fig.get("filename", "")
                                    self.db.index_figures(doc_id, extracted_figs)

                                existing["figures_extracted"] = True
                                existing["figures_count"] = len(extracted_figs)
                                ref_data[rel_path] = existing
                        except Exception as fe:
                            logger.warning(f"Figure backfill failed for '{rel_path}': {fe}")
                    skipped += 1
                    continue

            # Process file
            if progress_cb:
                progress_cb(idx + 1, total, filename, "Extracting text...")

            try:
                pages: List[Tuple[int, str]] = []
                structured_pages: List[Dict[str, Any]] = []
                extracted_meta: Dict[str, Any] = {}
                full_text = ""

                if ext == ".pdf":
                    full_text, structured_pages, extracted_meta = extract_pdf(full_path)
                    pages = [(p.get("page_num", 1), p.get("text", "")) for p in structured_pages]
                elif ext == ".docx":
                    full_text, structured_pages, extracted_meta = extract_docx(full_path)
                    pages = [(p.get("page_num", 1), p.get("text", "")) for p in structured_pages]
                elif ext == ".pptx":
                    full_text, structured_pages, extracted_meta = extract_pptx(full_path)
                    pages = [(p.get("page_num", 1), p.get("text", "")) for p in structured_pages]
                elif ext in {".txt", ".md"}:
                    with open(win_safe_path(full_path), "r", encoding="utf-8", errors="replace") as f_txt:
                        content = f_txt.read()
                        full_text = content
                        s_list = split_into_sentences(content)
                        structured_pages = [{
                            "page_num": 1,
                            "heading": "Document",
                            "text": content,
                            "sentences": s_list,
                        }]
                        pages = [(1, content)]
                        extracted_meta = {"pages": 1, "char_count": len(content)}

                # Extract DOI and Abstract from full_text
                meta_doi = extract_doi(full_text) or ""
                meta_abstract = extract_abstract(full_text) or ""

                fn_meta = parse_filename_metadata(filename)
                year_val = fn_meta.get("year")
                author_val = fn_meta.get("authors") or ""
                title_val = fn_meta.get("title") or os.path.splitext(filename)[0]

                # Bundle classification
                bundle_info = bundle_document_record(rel_path, filename, title_val)

                # Write markdown cache
                md_rel = os.path.join(".ai_cache", os.path.splitext(rel_path)[0] + ".md").replace("\\", "/")
                md_full = os.path.join(watch_path, md_rel)
                os.makedirs(win_safe_path(os.path.dirname(md_full)), exist_ok=True)

                self._write_markdown_cache(
                    md_full_path=md_full,
                    title=title_val,
                    category=item["category"],
                    authors=author_val or "Unknown",
                    year=year_val or "Unknown",
                    original_file=full_path,
                    doi=meta_doi,
                    abstract=meta_abstract,
                    format_name=ext[1:].upper(),
                    doc_role=bundle_info["role"],
                    bundle_id=bundle_info["bundle_id"],
                    pages=pages
                )

                # Upsert into SQLite
                doc_record = {
                    "watch_folder": watch_path,
                    "rel_path": rel_path,
                    "filename": filename,
                    "title": title_val,
                    "authors": author_val,
                    "year": year_val,
                    "doi": meta_doi,
                    "doc_role": bundle_info["role"],
                    "bundle_id": bundle_info["bundle_id"],
                    "page_count": len(pages),
                    "file_size": fsize,
                    "mtime": mtime,
                    "md_path": md_rel,
                    "abstract": meta_abstract
                }
                doc_id = self.db.upsert_document(doc_record)

                # Query previous figure paths to identify obsolete figures if modified
                old_figure_rels = set()
                try:
                    conn = self.db.get_connection()
                    cur = conn.cursor()
                    cur.execute("SELECT rel_image_path FROM figures WHERE doc_id = ?", (doc_id,))
                    for r in cur.fetchall():
                        if r["rel_image_path"]:
                            old_figure_rels.add(r["rel_image_path"])
                except Exception as e:
                    logger.debug(f"Could not query existing figures for doc_id {doc_id}: {e}")

                self.db.clear_document_content(doc_id)

                # Extract and index sentences directly consuming extractor pre-tokenized sentences
                all_sentences = []
                for p in structured_pages:
                    p_num = p.get("page_num", 1)
                    s_list = p.get("sentences")
                    if not s_list and p.get("text"):
                        s_list = split_into_sentences(p.get("text", ""))
                    if s_list:
                        for s_idx, s_str in enumerate(s_list):
                            if s_str and s_str.strip():
                                all_sentences.append((p_num, s_idx, s_str.strip()))

                if all_sentences:
                    self.db.index_sentences(doc_id, all_sentences)

                # Extract and index figures if enabled
                extracted_figs = []
                if self.extract_figures:
                    if progress_cb:
                        progress_cb(idx + 1, total, filename, "Extracting figures...")
                    try:
                        if ext == ".pdf":
                            extracted_figs = extract_figures_from_pdf(full_path, figures_cache_dir, max_figures=30)
                        elif ext == ".docx":
                            extracted_figs = extract_figures_from_docx(full_path, figures_cache_dir)
                        elif ext == ".pptx":
                            extracted_figs = extract_figures_from_pptx(full_path, figures_cache_dir)
                    except Exception as fe:
                        logger.warning(f"Figure extraction failed for '{rel_path}': {fe}")
                        extracted_figs = []

                    if extracted_figs:
                        # Harmonize figure dictionary keys for database insertion
                        for fig in extracted_figs:
                            if "figure_index" not in fig and "fig_index" in fig:
                                fig["figure_index"] = fig["fig_index"]
                            if "page_number" not in fig and "page_num" in fig:
                                fig["page_number"] = fig["page_num"]
                            if not fig.get("rel_image_path") and fig.get("file_path"):
                                try:
                                    fig["rel_image_path"] = os.path.relpath(
                                        _strip_long_prefix(str(fig["file_path"])),
                                        _strip_long_prefix(str(watch_path))
                                    ).replace("\\", "/")
                                except Exception:
                                    fig["rel_image_path"] = fig.get("filename", "")
                        self.db.index_figures(doc_id, extracted_figs)

                # Clean up any obsolete figure files on disk if the document lost figures
                new_figure_rels = {fig.get("rel_image_path") for fig in extracted_figs if fig.get("rel_image_path")}
                obsolete_figs = old_figure_rels - new_figure_rels
                for obs_rel in obsolete_figs:
                    try:
                        conn = self.db.get_connection()
                        cur = conn.cursor()
                        cur.execute(
                            "SELECT count(*) as cnt FROM figures WHERE rel_image_path = ? AND doc_id != ?",
                            (obs_rel, doc_id)
                        )
                        cnt_row = cur.fetchone()
                        if cnt_row and cnt_row["cnt"] == 0:
                            for cand in [
                                os.path.join(watch_path, obs_rel),
                                os.path.join(watch_path, ".ai_cache", obs_rel),
                                os.path.join(watch_path, ".ai_cache", "figures", os.path.basename(obs_rel))
                            ]:
                                if os.path.exists(win_safe_path(cand)):
                                    try:
                                        os.remove(win_safe_path(cand))
                                        logger.info(f"Removed obsolete figure cache: {cand}")
                                        break
                                    except Exception as e:
                                        logger.warning(f"Failed to remove obsolete figure '{cand}': {e}")
                    except Exception as e:
                        logger.warning(f"Error checking obsolete figure '{obs_rel}': {e}")

                # Update references dict
                ref_data[rel_path] = {
                    "rel_path": rel_path,
                    "category": item["category"],
                    "filename": filename,
                    "extension": ext,
                    "file_size": fsize,
                    "mtime": mtime,
                    "title": title_val,
                    "authors": author_val,
                    "year": year_val,
                    "doi": meta_doi or None,
                    "abstract": meta_abstract or None,
                    "doc_role": bundle_info["role"],
                    "bundle_id": bundle_info["bundle_id"],
                    "md_cache_rel": md_rel,
                    "original_file_uri": to_file_uri(full_path),
                    "format_meta": {
                        "pages": len(pages),
                        "format": ext[1:].upper(),
                        "char_count": sum(len(txt) for _, txt in pages)
                    },
                    "figures_extracted": bool(self.extract_figures),
                    "figures_count": len(extracted_figs),
                    "last_processed": datetime.datetime.now().isoformat()
                }
                processed += 1

            except Exception as e:
                errors += 1
                logger.error(f"Error processing file '{rel_path}' ({full_path}): {e}", exc_info=True)
                if progress_cb:
                    progress_cb(idx + 1, total, filename, f"Error: {e}")

        # Save references.json
        try:
            with open(win_safe_path(ref_json_path), "w", encoding="utf-8") as wf:
                json.dump(ref_data, wf, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.warning(f"Failed to write references.json at {ref_json_path}: {e}")

        # Save REFERENCES.md
        self._write_references_markdown(watch_path, ref_data)

        return {
            "total": total,
            "processed": processed,
            "skipped": skipped,
            "deleted": deleted,
            "errors": errors,
            "watch_folder": watch_folder
        }

    def _write_markdown_cache(
        self,
        md_full_path: str,
        title: str,
        category: str,
        authors: str,
        year: Any,
        original_file: str,
        doi: str,
        abstract: str,
        format_name: str,
        doc_role: str,
        bundle_id: str,
        pages: List[Tuple[int, str]]
    ):
        with open(win_safe_path(md_full_path), "w", encoding="utf-8", errors="replace") as mf:
            mf.write("---\n")
            mf.write(f'title: "{title}"\n')
            mf.write(f'category: "{category}"\n')
            mf.write(f'authors: "{authors}"\n')
            mf.write(f'year: {year}\n')
            mf.write(f'original_file: "{original_file}"\n')
            mf.write(f'doi: "{doi}"\n')
            mf.write(f'format: {format_name}\n')
            mf.write(f'doc_role: {doc_role}\n')
            mf.write(f'bundle_id: "{bundle_id}"\n')
            mf.write(f'extracted_at: {datetime.datetime.now().isoformat()}\n')
            mf.write("---\n\n")
            mf.write(f"# {title}\n\n")

            if abstract:
                mf.write("### Summary / Abstract\n")
                mf.write(f"> {abstract.strip()}\n\n")
                mf.write("---\n\n")

            for p_num, p_text in pages:
                clean_text = p_text.strip()
                if clean_text:
                    mf.write(f"## Page {p_num}\n\n")
                    mf.write(f"{clean_text}\n\n")

    def _write_references_markdown(self, watch_path: str, ref_data: Dict[str, Any]):
        out_md = os.path.join(watch_path, "REFERENCES.md")
        try:
            lines = [
                "# References Index\n\n",
                f"*Last synchronized: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n\n",
                "| Category | Role | Year | Title / File | Authors | DOI | MD Cache |\n",
                "|---|---|---|---|---|---|---|\n"
            ]

            sorted_items = sorted(
                ref_data.values(),
                key=lambda x: (
                    str(x.get("category") or ""),
                    str(x.get("bundle_id") or ""),
                    str(x.get("doc_role") or ""),
                    str(x.get("filename") or "")
                )
            )

            for item in sorted_items:
                cat = str(item.get("category") or "")
                role = str(item.get("doc_role") or "MAIN")
                year = str(item.get("year") or "")
                fn = str(item.get("filename") or "")
                title = str(item.get("title") or fn or "")
                auth = str(item.get("authors") or "")
                doi = str(item.get("doi") or "")
                md_rel = str(item.get("md_cache_rel") or "")
                uri = str(item.get("original_file_uri") or "")
                lines.append(f"| {cat} | {role} | {year} | [{title}]({uri}) | {auth} | {doi} | [{md_rel}]({md_rel}) |\n")

            with open(win_safe_path(out_md), "w", encoding="utf-8", errors="replace") as f:
                f.writelines(lines)
        except Exception as e:
            logger.warning(f"Failed to write REFERENCES.md at {out_md}: {e}")
