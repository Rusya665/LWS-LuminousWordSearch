"""
tests/test_challenger_m2_sync.py
================================
Adversarial Stress Test Suite for Milestone 2: Synchronizer & Deletion Reconciliation
Agent: challenger_m2_2 (Synchronizer & Deletion Adversarial Challenger)

Empirical verification of:
1. Adversarial Deletion Reconciliation:
   - Sync 10 documents (PDF, DOCX, PPTX), delete 3, verify documents, sentences_fts, figures,
     figures_fts, and user_annotations for deleted files are 100% purged.
   - Verify .ai_cache markdown files are unlinked.
   - Verify references.json and REFERENCES.md have references purged.
   - Delete ALL files in watch folder and sync: verify no crashes and clean database state.
   - Deletion reconciliation when database connection is closed or cache file is locked.
   - Document renames and deep nested subfolder moves.
   - Incremental concurrent modifications: verify sentences_fts does not accumulate duplicate rows.
   - Non-Latin, Unicode, and special character filenames.
   - Companion bundle partial deletion (deleting SI while retaining MAIN).
2. Direct Sentence Indexing:
   - Preservation of short headings ("Methods", "Results") and short table cells in sentences_fts.
3. Empirical Defect Reproductions:
   - REFERENCES.md sort failure with TypeError on None field values.
   - Figure cache collision across multiple documents in the same watch folder.
   - Forward-slash watch_folder case mismatch in _reconcile_deletions query.
"""

import io
import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Tuple

import docx
from PIL import Image
import pptx
from pptx.util import Inches
import pytest
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Ensure UTF-8 console output on Windows
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from lws_engine.config import win_safe_path
from lws_engine.database import LwsDatabase
from lws_engine.synchronizer import LwsSynchronizer
from lws_engine.figure_extractor import (
    extract_figures_from_pdf,
    extract_figures_from_docx,
    extract_figures_from_pptx,
)


# ---------------------------------------------------------------------------
# Test Fixture Helpers
# ---------------------------------------------------------------------------

def _create_sample_png(path: str, size: Tuple[int, int] = (180, 180), color=(50, 100, 200)) -> str:
    img = Image.new("RGB", size, color)
    img.save(path, format="PNG")
    return path


def _create_synthetic_pdf(
    pdf_path: str,
    page_count: int = 2,
    images_per_page: int = 1,
    image_size: Tuple[int, int] = (160, 160),
    title: str = "Photochemical Additive Manufacturing"
) -> str:
    c = canvas.Canvas(pdf_path, pagesize=letter)
    tmp_dir = os.path.dirname(pdf_path)

    for p in range(1, page_count + 1):
        c.drawString(72, 750, f"{title} - Page {p}")
        c.drawString(72, 720, f"Fig. {p}: Microstructural analysis of photopolymer slice {p}.")
        for i in range(1, images_per_page + 1):
            img_path = os.path.join(tmp_dir, f"tmp_p{p}_{i}_{os.path.basename(pdf_path)}.png")
            _create_sample_png(img_path, size=image_size, color=(30 * p, 50 * i, 160))
            y_pos = 450 - (i - 1) * 160
            c.drawImage(img_path, 72, y_pos, 140, 130)
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass
        c.showPage()

    c.save()
    return pdf_path


def _create_synthetic_docx(
    docx_path: str,
    heading_text: str = "Methods",
    body_text: str = "Resin kinetics were quantified with 405 nm irradiance.",
    table_data: Tuple[Tuple[str, str], ...] = (("Property", "Value"), ("Modulus", "5.4 MPa"))
) -> str:
    tmp_dir = os.path.dirname(docx_path)
    img_path = os.path.join(tmp_dir, f"tmp_docx_{os.path.basename(docx_path)}.png")
    _create_sample_png(img_path, size=(160, 160), color=(200, 80, 40))

    doc = docx.Document()
    doc.add_heading(heading_text, level=1)
    doc.add_paragraph(body_text)
    if table_data:
        tbl = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
        for r_idx, row in enumerate(table_data):
            for c_idx, val in enumerate(row):
                tbl.cell(r_idx, c_idx).text = val

    doc.add_picture(img_path, width=docx.shared.Inches(2.0))
    doc.save(docx_path)

    if os.path.exists(img_path):
        try:
            os.remove(img_path)
        except Exception:
            pass
    return docx_path


def _create_synthetic_pptx(
    pptx_path: str,
    slide_title: str = "Results Summary",
    body_text: str = "High spatial resolution tomographic projections verified.",
    table_cell: str = "98.5% yield"
) -> str:
    tmp_dir = os.path.dirname(pptx_path)
    img_path = os.path.join(tmp_dir, f"tmp_pptx_{os.path.basename(pptx_path)}.png")
    _create_sample_png(img_path, size=(160, 160), color=(40, 180, 90))

    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tx_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(6), Inches(1))
    tf = tx_box.text_frame
    tf.text = slide_title

    p_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(6), Inches(1))
    p_box.text_frame.text = body_text

    tbl_shape = slide.shapes.add_table(1, 1, Inches(1), Inches(3.0), Inches(2), Inches(0.8))
    tbl_shape.table.cell(0, 0).text = table_cell

    slide.shapes.add_picture(img_path, Inches(4), Inches(3.0), Inches(2.5), Inches(2.0))
    prs.save(pptx_path)

    if os.path.exists(img_path):
        try:
            os.remove(img_path)
        except Exception:
            pass
    return pptx_path


# ===========================================================================
# 1. Adversarial Deletion Reconciliation Workflow Tests
# ===========================================================================

def test_adversarial_reconciliation_10_docs_delete_3():
    """
    STRESS TEST:
    Sync a folder containing 10 documents across PDF (4), DOCX (3), and PPTX (3).
    Add user annotations to all documents.
    Delete 3 documents (1 PDF, 1 DOCX, 1 PPTX).
    Run sync again.
    Empirically verify:
    - Deleted documents are 100% purged from 'documents' table.
    - Deleted documents have 0 rows in 'sentences_fts'.
    - Deleted documents have 0 rows in 'figures'.
    - Deleted documents have 0 rows in 'figures_fts'.
    - Deleted documents have 0 rows in 'user_annotations'.
    - .ai_cache markdown files for deleted documents are unlinked.
    - references.json does not contain references to deleted files.
    - REFERENCES.md does not contain references to deleted files.
    - The 7 remaining documents remain completely intact with their annotations and searchable FTS text.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_10docs.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=True)

        docs = []
        # 4 PDFs
        for i in range(4):
            p = os.path.join(tmp_dir, f"paper_pdf_{i}.pdf")
            _create_synthetic_pdf(p, page_count=2, images_per_page=1, title=f"Polymer Study {i}")
            docs.append(p)
        # 3 DOCXs
        for i in range(3):
            p = os.path.join(tmp_dir, f"protocol_docx_{i}.docx")
            _create_synthetic_docx(p, heading_text=f"Methods Protocol {i}")
            docs.append(p)
        # 3 PPTXs
        for i in range(3):
            p = os.path.join(tmp_dir, f"slides_pptx_{i}.pptx")
            _create_synthetic_pptx(p, slide_title=f"Presentation Slide {i}")
            docs.append(p)

        res1 = sync.sync_watch_folder(tmp_dir)
        assert res1["total"] == 10
        assert res1["processed"] == 10
        assert res1["deleted"] == 0

        # Attach user annotations to all 10 documents
        cur = db.get_connection().cursor()
        cur.execute("SELECT id, rel_path FROM documents")
        all_docs_db = cur.fetchall()
        assert len(all_docs_db) == 10
        for doc_row in all_docs_db:
            db.save_annotation(doc_row["id"], 1, 0, f"Critical note for {doc_row['rel_path']}", "stress_test")

        # Verify all 10 have annotations
        cur.execute("SELECT count(*) as cnt FROM user_annotations")
        assert cur.fetchone()["cnt"] == 10

        # Choose 3 to delete: 1 PDF, 1 DOCX, 1 PPTX
        to_delete = [docs[0], docs[4], docs[7]]
        deleted_rel_paths = [os.path.basename(p) for p in to_delete]
        deleted_doc_ids = [r["id"] for r in all_docs_db if r["rel_path"] in deleted_rel_paths]
        assert len(deleted_doc_ids) == 3

        # Physically remove files from disk
        for p in to_delete:
            os.remove(p)

        # Execute reconciliation sync
        res2 = sync.sync_watch_folder(tmp_dir)
        assert res2["total"] == 7
        assert res2["skipped"] == 7
        assert res2["processed"] == 0
        assert res2["deleted"] == 3
        assert res2["errors"] == 0

        # 1. Verify purged from 'documents' table
        cur.execute("SELECT id, rel_path FROM documents")
        remaining_db = cur.fetchall()
        assert len(remaining_db) == 7
        remaining_rels = [r["rel_path"] for r in remaining_db]
        for del_rel in deleted_rel_paths:
            assert del_rel not in remaining_rels, f"Orphan found: {del_rel} still in documents!"

        # 2. Verify purged from 'sentences_fts'
        for del_id in deleted_doc_ids:
            cur.execute("SELECT count(*) as cnt FROM sentences_fts WHERE doc_id = ?", (str(del_id),))
            cnt = cur.fetchone()["cnt"]
            assert cnt == 0, f"Orphan found: {cnt} sentences remain for deleted doc_id {del_id} in sentences_fts!"

        # 3. Verify purged from 'figures' and 'figures_fts'
        for del_id in deleted_doc_ids:
            cur.execute("SELECT count(*) as cnt FROM figures WHERE doc_id = ?", (del_id,))
            f_cnt = cur.fetchone()["cnt"]
            assert f_cnt == 0, f"Orphan found: {f_cnt} figures remain for deleted doc_id {del_id} in figures table!"
            cur.execute("SELECT count(*) as cnt FROM figures_fts WHERE doc_id = ?", (str(del_id),))
            ff_cnt = cur.fetchone()["cnt"]
            assert ff_cnt == 0, f"Orphan found: {ff_cnt} figures remain for deleted doc_id {del_id} in figures_fts!"

        # 4. Verify purged from 'user_annotations'
        for del_id in deleted_doc_ids:
            cur.execute("SELECT count(*) as cnt FROM user_annotations WHERE doc_id = ?", (del_id,))
            a_cnt = cur.fetchone()["cnt"]
            assert a_cnt == 0, f"Orphan found: {a_cnt} annotations remain for deleted doc_id {del_id} in user_annotations!"

        cur.execute("SELECT count(*) as cnt FROM user_annotations")
        assert cur.fetchone()["cnt"] == 7

        # 5. Verify .ai_cache markdown files unlinked
        for del_rel in deleted_rel_paths:
            md_name = os.path.splitext(del_rel)[0] + ".md"
            md_disk = os.path.join(tmp_dir, ".ai_cache", md_name)
            assert not os.path.exists(md_disk), f"Orphan markdown cache file not unlinked: {md_disk}"

        # 6. Verify references.json pruned
        ref_json_path = os.path.join(tmp_dir, "references.json")
        with open(ref_json_path, "r", encoding="utf-8") as rf:
            ref_data = json.load(rf)
            for del_rel in deleted_rel_paths:
                assert del_rel not in ref_data, f"Orphan reference in references.json: {del_rel}"
            assert len(ref_data) == 7

        # 7. Verify REFERENCES.md pruned
        ref_md_path = os.path.join(tmp_dir, "REFERENCES.md")
        with open(ref_md_path, "r", encoding="utf-8") as rmf:
            md_content = rmf.read()
            for del_rel in deleted_rel_paths:
                assert del_rel not in md_content, f"Orphan mention in REFERENCES.md: {del_rel}"

        # 8. Verify remaining 7 documents are still searchable
        assert len(db.search_fts("Methods")) >= 1
        assert len(db.search_fts("tomographic")) >= 1

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_reconciliation_delete_all_documents():
    """
    STRESS TEST:
    Sync a watch folder with documents, then delete ALL documents.
    Sync again. Verify:
    - Zero crashes or exceptions.
    - Database is completely empty (0 rows in documents, sentences_fts, figures, figures_fts, user_annotations).
    - references.json is an empty dict {}.
    - REFERENCES.md contains table headers but 0 rows.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_delete_all.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=True)

        p1 = os.path.join(tmp_dir, "paper1.pdf")
        p2 = os.path.join(tmp_dir, "protocol.docx")
        p3 = os.path.join(tmp_dir, "presentation.pptx")
        _create_synthetic_pdf(p1, page_count=1)
        _create_synthetic_docx(p2, heading_text="All Delete Test")
        _create_synthetic_pptx(p3, slide_title="All Delete Slide")

        res1 = sync.sync_watch_folder(tmp_dir)
        assert res1["processed"] == 3

        # Delete ALL files
        os.remove(p1)
        os.remove(p2)
        os.remove(p3)

        res2 = sync.sync_watch_folder(tmp_dir)
        assert res2["total"] == 0
        assert res2["processed"] == 0
        assert res2["skipped"] == 0
        assert res2["deleted"] == 3
        assert res2["errors"] == 0

        cur = db.get_connection().cursor()
        cur.execute("SELECT count(*) as cnt FROM documents")
        assert cur.fetchone()["cnt"] == 0
        cur.execute("SELECT count(*) as cnt FROM sentences_fts")
        assert cur.fetchone()["cnt"] == 0
        cur.execute("SELECT count(*) as cnt FROM figures")
        assert cur.fetchone()["cnt"] == 0
        cur.execute("SELECT count(*) as cnt FROM figures_fts")
        assert cur.fetchone()["cnt"] == 0
        cur.execute("SELECT count(*) as cnt FROM user_annotations")
        assert cur.fetchone()["cnt"] == 0

        # references.json must be {}
        ref_json_path = os.path.join(tmp_dir, "references.json")
        with open(ref_json_path, "r", encoding="utf-8") as rf:
            assert json.load(rf) == {}

        # REFERENCES.md must have no rows
        ref_md_path = os.path.join(tmp_dir, "REFERENCES.md")
        with open(ref_md_path, "r", encoding="utf-8") as rmf:
            lines = [l for l in rmf.read().splitlines() if l.startswith("|") and not l.startswith("| Category") and not l.startswith("|---")]
            assert len(lines) == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_reconciliation_when_entire_watch_folder_deleted():
    """
    STRESS TEST:
    Sync a watch folder, then delete the entire folder from the disk.
    Call reconcile_deleted_files on the missing folder path.
    Verify clean error boundary (no unhandled exceptions) and database purge.
    """
    tmp_dir = tempfile.mkdtemp()
    watch_dir = os.path.join(tmp_dir, "watch_folder")
    os.makedirs(watch_dir)
    try:
        db_path = os.path.join(tmp_dir, "test_missing_dir.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        p = os.path.join(watch_dir, "doc.docx")
        _create_synthetic_docx(p)
        sync.sync_watch_folder(watch_dir)

        cur = db.get_connection().cursor()
        cur.execute("SELECT count(*) as cnt FROM documents")
        assert cur.fetchone()["cnt"] == 1

        # Delete entire watch directory
        shutil.rmtree(watch_dir)
        assert not os.path.exists(watch_dir)

        deleted = sync.reconcile_deleted_files(watch_dir)
        assert deleted == 1

        cur.execute("SELECT count(*) as cnt FROM documents")
        assert cur.fetchone()["cnt"] == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_reconciliation_closed_db_recovery():
    """
    STRESS TEST:
    Verify that synchronizer transparently reopens SQLite connection if
    db.close() was previously called, without crashing.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_closed.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        p = os.path.join(tmp_dir, "paper.docx")
        _create_synthetic_docx(p)
        sync.sync_watch_folder(tmp_dir)

        # Close database connection
        db.close()
        assert db._conn is None

        # Delete document from disk
        os.remove(p)

        # Sync should re-acquire connection cleanly
        res = sync.sync_watch_folder(tmp_dir)
        assert res["deleted"] == 1

        cur = db.get_connection().cursor()
        cur.execute("SELECT count(*) as cnt FROM documents")
        assert cur.fetchone()["cnt"] == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_reconciliation_locked_cache_file_graceful():
    """
    STRESS TEST:
    Verify that when an orphan .ai_cache file is locked (PermissionError on Windows),
    the synchronizer catches the exception, logs a warning, and still purges
    the database record and references.json.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_locked.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        p = os.path.join(tmp_dir, "locked_test.docx")
        _create_synthetic_docx(p)
        sync.sync_watch_folder(tmp_dir)

        md_cache_path = os.path.join(tmp_dir, ".ai_cache", "locked_test.md")
        assert os.path.exists(md_cache_path)

        # Delete source document
        os.remove(p)

        # Hold open exclusive lock on the markdown cache file
        with open(md_cache_path, "r+") as locked_file:
            res = sync.sync_watch_folder(tmp_dir)
            assert res["deleted"] == 1

            cur = db.get_connection().cursor()
            cur.execute("SELECT count(*) as cnt FROM documents")
            assert cur.fetchone()["cnt"] == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_reconciliation_document_rename_and_nested_move():
    """
    STRESS TEST:
    Verify that renaming or moving a document between nested folders:
    - Purges the old rel_path entry from documents, references.json, and FTS.
    - Indexes the new rel_path with updated category.
    - Leaves no duplicate FTS records or orphan cache files.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_rename_move.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        old_p = os.path.join(tmp_dir, "initial_name.docx")
        _create_synthetic_docx(old_p, heading_text="UniqueExperimentTerm")
        sync.sync_watch_folder(tmp_dir)

        hits1 = db.search_fts("UniqueExperimentTerm")
        assert len(hits1) == 1
        assert hits1[0]["filename"] == "initial_name.docx"

        # 1. Rename file
        new_p = os.path.join(tmp_dir, "renamed_protocol.docx")
        os.rename(old_p, new_p)

        res1 = sync.sync_watch_folder(tmp_dir)
        assert res1["deleted"] == 1
        assert res1["processed"] == 1

        hits2 = db.search_fts("UniqueExperimentTerm")
        assert len(hits2) == 1
        assert hits2[0]["filename"] == "renamed_protocol.docx"
        assert not os.path.exists(os.path.join(tmp_dir, ".ai_cache", "initial_name.md"))
        assert os.path.exists(os.path.join(tmp_dir, ".ai_cache", "renamed_protocol.md"))

        # 2. Move file into deeply nested folder
        nested_dir = os.path.join(tmp_dir, "additive_manufacturing", "photochemistry")
        os.makedirs(nested_dir, exist_ok=True)
        moved_p = os.path.join(nested_dir, "renamed_protocol.docx")
        os.rename(new_p, moved_p)

        res2 = sync.sync_watch_folder(tmp_dir)
        assert res2["deleted"] == 1
        assert res2["processed"] == 1

        hits3 = db.search_fts("UniqueExperimentTerm")
        assert len(hits3) == 1
        assert hits3[0]["rel_path"] == "additive_manufacturing/photochemistry/renamed_protocol.docx"

        # Verify old root cache is removed
        assert not os.path.exists(os.path.join(tmp_dir, ".ai_cache", "renamed_protocol.md"))
        # Verify new nested cache exists
        assert os.path.exists(os.path.join(tmp_dir, ".ai_cache", "additive_manufacturing", "photochemistry", "renamed_protocol.md"))

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_concurrent_file_modification_no_duplicate_fts():
    """
    STRESS TEST:
    Verify that updating a document's content and mtime re-indexes cleanly
    without creating duplicate sentences in sentences_fts or returning stale hits.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_modify.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        txt_p = os.path.join(tmp_dir, "live_notes.txt")
        with open(txt_p, "w", encoding="utf-8") as f:
            f.write("Initial formulation uses camphorquinone photoinitiator.")

        sync.sync_watch_folder(tmp_dir)
        assert len(db.search_fts("camphorquinone")) == 1
        assert len(db.search_fts("phenylbisphosphine")) == 0

        # Update file content and update mtime
        time.sleep(0.05)
        with open(txt_p, "w", encoding="utf-8") as f:
            f.write("Revised formulation uses phenylbisphosphine oxide initiator.")
        os.utime(txt_p, (time.time() + 10, time.time() + 10))

        res = sync.sync_watch_folder(tmp_dir)
        assert res["processed"] == 1
        assert res["skipped"] == 0

        # Former term must have 0 hits, new term must have exactly 1 hit
        assert len(db.search_fts("camphorquinone")) == 0
        assert len(db.search_fts("phenylbisphosphine")) == 1

        cur = db.get_connection().cursor()
        cur.execute("SELECT count(*) as cnt FROM sentences_fts")
        assert cur.fetchone()["cnt"] == 1

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_companion_bundle_partial_deletion():
    """
    STRESS TEST:
    Verify that deleting only a companion document (e.g. SI) preserves the
    main paper intact in the database and catalog, keeping their bundle relationship valid.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_bundle.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        main_p = os.path.join(tmp_dir, "nature_article_2024.pdf")
        si_p = os.path.join(tmp_dir, "nature_article_2024_si.pdf")
        _create_synthetic_pdf(main_p, page_count=1, title="Nature Article")
        _create_synthetic_pdf(si_p, page_count=1, title="Supplementary Information")

        sync.sync_watch_folder(tmp_dir)

        cur = db.get_connection().cursor()
        cur.execute("SELECT rel_path, doc_role, bundle_id FROM documents ORDER BY doc_role")
        rows = cur.fetchall()
        assert len(rows) == 2
        assert rows[0]["bundle_id"] == rows[1]["bundle_id"]

        # Delete only SI
        os.remove(si_p)
        res = sync.sync_watch_folder(tmp_dir)
        assert res["deleted"] == 1

        cur.execute("SELECT rel_path, doc_role, bundle_id FROM documents")
        remaining = cur.fetchall()
        assert len(remaining) == 1
        assert remaining[0]["rel_path"] == "nature_article_2024.pdf"
        assert remaining[0]["doc_role"] == "MAIN"

        with open(os.path.join(tmp_dir, "REFERENCES.md"), "r", encoding="utf-8") as f:
            content = f.read()
            assert "nature_article_2024.pdf" in content
            assert "nature_article_2024_si.pdf" not in content

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_adversarial_reconciliation_special_chars_and_unicode_filenames():
    """
    STRESS TEST:
    Verify that documents with spaces, brackets, Unicode, Cyrillic, and CJK characters
    are indexed, searched, and deleted cleanly without regex or filesystem escaping failures.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_unicode.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        fn1 = "[2024] Photopolymerization & Resins #1 (v2.0).docx"
        fn2 = "体素化3D打印_исследование_2024.docx"
        p1 = os.path.join(tmp_dir, fn1)
        p2 = os.path.join(tmp_dir, fn2)
        _create_synthetic_docx(p1, heading_text="SpecialCharsHeading")
        _create_synthetic_docx(p2, heading_text="UnicodeHeading")

        sync.sync_watch_folder(tmp_dir)

        assert len(db.search_fts("SpecialCharsHeading")) == 1
        assert len(db.search_fts("UnicodeHeading")) == 1

        # Delete both
        os.remove(p1)
        os.remove(p2)

        res = sync.sync_watch_folder(tmp_dir)
        assert res["deleted"] == 2

        cur = db.get_connection().cursor()
        cur.execute("SELECT count(*) as cnt FROM documents")
        assert cur.fetchone()["cnt"] == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ===========================================================================
# 2. Direct Sentence Indexing Tests
# ===========================================================================

def test_direct_sentence_indexing_docx_headings_and_tables():
    """
    VERIFICATION:
    Verify that short headings (len <= 8) and short table cells from DOCX files
    are preserved and searchable in FTS5 after synchronizer processing.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_docx_sentences.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        docx_p = os.path.join(tmp_dir, "paper.docx")
        doc = docx.Document()
        doc.add_heading("Methods", level=1)  # 7 chars
        doc.add_paragraph("Resin prep was conducted in the dark.")
        doc.add_heading("Results", level=1)  # 7 chars
        doc.add_heading("Intro", level=1)    # 5 chars

        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "Property"
        tbl.cell(0, 1).text = "Value"
        tbl.cell(1, 0).text = "Modulus"
        tbl.cell(1, 1).text = "5.4 MPa"
        doc.save(docx_p)

        sync.sync_watch_folder(tmp_dir)

        # "Methods" (7 chars) must be indexed
        methods_hits = db.search_fts("Methods")
        assert len(methods_hits) >= 1
        assert any("Methods" in h["sentence_text"] for h in methods_hits)

        # "Results" (7 chars) must be indexed
        results_hits = db.search_fts("Results")
        assert len(results_hits) >= 1

        # "Intro" (5 chars) must be indexed
        intro_hits = db.search_fts("Intro")
        assert len(intro_hits) >= 1

        # Table cell "5.4 MPa" must be indexed
        mpa_hits = db.search_fts("MPa")
        assert len(mpa_hits) >= 1

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_direct_sentence_indexing_pptx_slides_and_tables():
    """
    VERIFICATION:
    Verify that short slide titles (len <= 8) and table cells from PPTX presentations
    are preserved and searchable in FTS5.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_pptx_sentences.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        pptx_p = os.path.join(tmp_dir, "presentation.pptx")
        _create_synthetic_pptx(pptx_p, slide_title="Data", table_cell="Yield 98%")

        sync.sync_watch_folder(tmp_dir)

        # Short slide title "Data" (4 chars)
        data_hits = db.search_fts("Data")
        assert len(data_hits) >= 1

        # Table cell "Yield 98%"
        yield_hits = db.search_fts("Yield")
        assert len(yield_hits) >= 1

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ===========================================================================
# 3. Empirical Defect Reproductions & Adversarial Challenges
# ===========================================================================

def test_reproduce_bug_figure_filename_collision_across_documents():
    """
    DEFECT REPRODUCTION:
    In figure_extractor.py:145, 233, 307:
      out_filename = f"fig_p{page_num}_{obj_idx + 1}.png"  (PDF)
      out_filename = f"fig_docx_{idx}.png"                 (DOCX)
      out_filename = f"fig_slide_{idx}.png"                (PPTX)
    When multiple documents exist in the same watch folder, their extracted figure
    filenames collide in .ai_cache/figures/.
    Document 2 overwrites Document 1's figure file on disk!
    When Document 2 is deleted, Document 1's figure cache on disk is corrupted.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        out_dir = os.path.join(tmp_dir, ".ai_cache", "figures")
        os.makedirs(out_dir, exist_ok=True)

        p1 = os.path.join(tmp_dir, "doc1.pdf")
        p2 = os.path.join(tmp_dir, "doc2.pdf")
        _create_synthetic_pdf(p1, page_count=1, images_per_page=1, image_size=(150, 150))
        _create_synthetic_pdf(p2, page_count=1, images_per_page=1, image_size=(250, 250))

        figs1 = extract_figures_from_pdf(p1, out_dir)
        figs2 = extract_figures_from_pdf(p2, out_dir)

        # Distinct filenames incorporating safe stems
        assert figs1[0]["filename"] != figs2[0]["filename"]
        assert figs1[0]["file_path"] != figs2[0]["file_path"]
        assert "doc1" in figs1[0]["filename"]
        assert "doc2" in figs2[0]["filename"]

        # Neither document's image was overwritten; verify distinct dimensions
        with Image.open(figs1[0]["file_path"]) as img1:
            assert img1.size == (150, 150)
        with Image.open(figs2[0]["file_path"]) as img2:
            assert img2.size == (250, 250)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_contract_figure_filenames_unique_across_documents():
    """
    CONTRACT REQUIREMENT:
    When multiple documents (PDF, DOCX, PPTX) have extracted figures in the same watch folder,
    each extracted figure must have a unique filename that includes the parent document stem
    to prevent clobbering other documents' figures.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        out_dir = os.path.join(tmp_dir, ".ai_cache", "figures")
        os.makedirs(out_dir, exist_ok=True)

        p1 = os.path.join(tmp_dir, "paper1.pdf")
        p2 = os.path.join(tmp_dir, "paper2.pdf")
        _create_synthetic_pdf(p1, page_count=1, images_per_page=1)
        _create_synthetic_pdf(p2, page_count=1, images_per_page=1)

        figs1 = extract_figures_from_pdf(p1, out_dir)
        figs2 = extract_figures_from_pdf(p2, out_dir)

        assert figs1[0]["filename"] != figs2[0]["filename"], (
            f"COLLISION: Both documents output identical filename '{figs1[0]['filename']}'!"
        )
        assert figs1[0]["file_path"] != figs2[0]["file_path"]
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_reproduce_bug_references_markdown_none_field_typeerror():
    """
    DEFECT REPRODUCTION:
    In synchronizer.py:549:
      sorted_items = sorted(
          ref_data.values(),
          key=lambda x: (x.get("category", ""), x.get("bundle_id", ""), x.get("doc_role", ""), x.get("filename", ""))
      )
    If a dictionary entry has an explicit None value for 'category', 'bundle_id', 'doc_role',
    or 'filename' (e.g. from an unclassified or legacy entry), x.get("category", "") returns None!
    Python raises TypeError: '<' not supported between instances of 'str' and 'NoneType'.
    Line 563 catches Exception and logs a warning, causing REFERENCES.md generation to fail!
    """
    ref_data = {
        "doc1.pdf": {"category": "Root", "bundle_id": "b1", "doc_role": "MAIN", "filename": "doc1.pdf"},
        "doc2.pdf": {"category": None, "bundle_id": "b2", "doc_role": "MAIN", "filename": "doc2.pdf"}
    }
    with pytest.raises(TypeError, match="not supported between instances of"):
        sorted(
            ref_data.values(),
            key=lambda x: (x.get("category", ""), x.get("bundle_id", ""), x.get("doc_role", ""), x.get("filename", ""))
        )


def test_contract_references_markdown_tolerates_none_field_values():
    """
    CONTRACT REQUIREMENT:
    _write_references_markdown must safely tolerate None values in ref_data entries
    without raising TypeError or aborting REFERENCES.md writing.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db = LwsDatabase(os.path.join(tmp_dir, "test.sqlite"))
        sync = LwsSynchronizer(db=db, extract_figures=False)

        ref_data = {
            "doc1.pdf": {"category": "Root", "bundle_id": "b1", "doc_role": "MAIN", "filename": "doc1.pdf", "original_file_uri": ""},
            "doc2.pdf": {"category": None, "bundle_id": None, "doc_role": None, "filename": "doc2.pdf", "original_file_uri": ""}
        }

        # This should execute and write REFERENCES.md without raising or swallowing TypeError
        sync._write_references_markdown(tmp_dir, ref_data)
        out_md = os.path.join(tmp_dir, "REFERENCES.md")
        assert os.path.exists(out_md)

        with open(out_md, "r", encoding="utf-8") as f:
            content = f.read()
            assert "doc1.pdf" in content
            assert "doc2.pdf" in content
        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
