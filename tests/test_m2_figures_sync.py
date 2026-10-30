"""
tests/test_m2_figures_sync.py
=============================
Comprehensive unit and integration tests for Milestone 2:
- Figure extraction parameter signatures (max_figures, min_dimension).
- Path and string argument coercion across PDF, DOCX, and PPTX extractors.
- Figure dictionary key harmonization and SQLite schema compatibility.
- Context manager and CMYK image color mode conversion.
- Deletion reconciliation (purging documents, FTS records, .ai_cache, and references.json).
- Direct pre-tokenized sentence consumption preserving headings.
- End-to-end synchronization with extract_figures=True across PDF, DOCX, and PPTX.
- Defensive error logging and invalid input boundaries.
"""

import io
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Tuple

import pytest
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import docx
import pptx
from pptx.util import Inches

from lws_engine.figure_extractor import (
    extract_figures_from_pdf,
    extract_figures_from_docx,
    extract_figures_from_pptx,
    find_captions_in_page_text,
    _prepare_image_for_save,
    _compute_rel_image_path,
)
from lws_engine.database import LwsDatabase
from lws_engine.synchronizer import LwsSynchronizer


# ----------------------------------------------------------------------
# Synthetic Fixture Helpers
# ----------------------------------------------------------------------

def _create_sample_png(path: str, size: Tuple[int, int] = (200, 200), color=(0, 128, 255), mode="RGB") -> str:
    """Creates a sample test PNG image."""
    img = Image.new(mode, size, color)
    img.save(path, format="PNG")
    return path


def _create_synthetic_pdf(
    pdf_path: str,
    page_count: int = 2,
    images_per_page: int = 1,
    image_size: Tuple[int, int] = (200, 200),
    include_captions: bool = True
) -> str:
    """Generates a synthetic multi-page PDF with embedded images and captions using ReportLab."""
    c = canvas.Canvas(pdf_path, pagesize=letter)
    tmp_dir = os.path.dirname(pdf_path)

    for p in range(1, page_count + 1):
        if include_captions:
            c.drawString(72, 720, f"Fig. {p}: Microstructure analysis of sample slice {p}.")
            c.drawString(72, 700, f"Extended Data Fig. {p}.1: High-resolution tomographic detail.")

        for i in range(1, images_per_page + 1):
            img_path = os.path.join(tmp_dir, f"tmp_img_p{p}_{i}.png")
            _create_sample_png(img_path, size=image_size, color=(20 * p, 40 * i, 180))
            y_pos = 450 - (i - 1) * 160
            c.drawImage(img_path, 72, y_pos, 150, 140)
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
    image_size: Tuple[int, int] = (200, 200),
    heading_text: str = "Methods",
    image_mode: str = "RGB"
) -> str:
    """Generates a synthetic DOCX with heading and embedded image."""
    tmp_dir = os.path.dirname(docx_path)
    img_path = os.path.join(tmp_dir, "tmp_docx_img.png")

    img = Image.new(image_mode, image_size, (255, 100, 50))
    # If CMYK, Pillow cannot save directly to PNG, save as TIFF or JPEG first
    if image_mode == "CMYK":
        img_raw = io.BytesIO()
        img.save(img_raw, format="TIFF")
        # Direct zip creation for CMYK test fixture
        with zipfile.ZipFile(docx_path, "w") as z:
            z.writestr(
                "word/document.xml",
                f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body><w:p><w:pPr><w:pStyle w:val="heading1"/></w:pPr><w:r><w:t>{heading_text}</w:t></w:r></w:p></w:body></w:document>'
            )
            z.writestr("word/media/image1.tif", img_raw.getvalue())
        return docx_path

    img.save(img_path, format="PNG")
    doc = docx.Document()
    doc.add_heading(heading_text, level=1)
    doc.add_paragraph("This paragraph describes volumetric additive manufacturing protocol.")
    doc.add_picture(img_path, width=docx.shared.Inches(2.5))
    doc.save(docx_path)

    if os.path.exists(img_path):
        try:
            os.remove(img_path)
        except Exception:
            pass

    return docx_path


def _create_synthetic_pptx(
    pptx_path: str,
    image_size: Tuple[int, int] = (200, 200),
    slide_title: str = "Results Summary"
) -> str:
    """Generates a synthetic PPTX presentation with embedded image."""
    tmp_dir = os.path.dirname(pptx_path)
    img_path = os.path.join(tmp_dir, "tmp_pptx_img.png")
    _create_sample_png(img_path, size=image_size, color=(80, 200, 120))

    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    tx_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(6), Inches(1))
    tf = tx_box.text_frame
    tf.text = slide_title

    slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), Inches(3), Inches(2.5))
    prs.save(pptx_path)

    if os.path.exists(img_path):
        try:
            os.remove(img_path)
        except Exception:
            pass

    return pptx_path


# ----------------------------------------------------------------------
# 1. Figure Extractor Signature and Parameter Tests
# ----------------------------------------------------------------------

def test_extract_figures_from_pdf_max_figures():
    """Verifies extract_figures_from_pdf accepts and respects max_figures."""
    tmp_dir = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(tmp_dir, "multipage.pdf")
        _create_synthetic_pdf(pdf_path, page_count=3, images_per_page=1)
        out_dir = os.path.join(tmp_dir, "extracted")

        # Test max_figures=1
        figs_1 = extract_figures_from_pdf(pdf_path, out_dir, max_figures=1)
        assert len(figs_1) == 1
        assert figs_1[0]["figure_index"] == 1
        assert figs_1[0]["page_number"] == 1

        # Test max_figures=2
        figs_2 = extract_figures_from_pdf(pdf_path, out_dir, max_figures=2)
        assert len(figs_2) == 2

        # Test max_figures=0 or negative
        figs_0 = extract_figures_from_pdf(pdf_path, out_dir, max_figures=0)
        assert len(figs_0) == 0
        figs_neg = extract_figures_from_pdf(pdf_path, out_dir, max_figures=-5)
        assert len(figs_neg) == 0

        # Test max_figures=None (extract all)
        figs_all = extract_figures_from_pdf(pdf_path, out_dir, max_figures=None)
        assert len(figs_all) == 3
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_extract_figures_string_paths():
    """Verifies string path arguments (str) are coerced without AttributeError."""
    tmp_dir = tempfile.mkdtemp()
    try:
        pdf_path = str(os.path.join(tmp_dir, "doc.pdf"))
        docx_path = str(os.path.join(tmp_dir, "doc.docx"))
        pptx_path = str(os.path.join(tmp_dir, "doc.pptx"))
        out_dir = str(os.path.join(tmp_dir, "figures_out"))

        _create_synthetic_pdf(pdf_path, page_count=1)
        _create_synthetic_docx(docx_path)
        _create_synthetic_pptx(pptx_path)

        # Ensure strings do not trigger 'str' object has no attribute 'mkdir' or 'name'
        pdf_res = extract_figures_from_pdf(pdf_path, out_dir)
        assert len(pdf_res) >= 1
        assert isinstance(pdf_res[0]["file_path"], str)

        docx_res = extract_figures_from_docx(docx_path, out_dir)
        assert len(docx_res) >= 1
        assert isinstance(docx_res[0]["file_path"], str)

        pptx_res = extract_figures_from_pptx(pptx_path, out_dir)
        assert len(pptx_res) >= 1
        assert isinstance(pptx_res[0]["file_path"], str)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_figure_metadata_keys_and_schema_harmony():
    """Verifies output dictionaries emit both canonical keys and backward-compatible aliases."""
    tmp_dir = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(tmp_dir, "test.pdf")
        _create_synthetic_pdf(pdf_path, page_count=1)
        out_dir = os.path.join(tmp_dir, ".ai_cache", "figures")

        figs = extract_figures_from_pdf(pdf_path, out_dir, max_figures=10)
        assert len(figs) == 1
        fig = figs[0]

        # Canonical keys
        assert "figure_index" in fig
        assert "page_number" in fig
        assert "rel_image_path" in fig
        assert "file_path" in fig
        assert "width" in fig
        assert "height" in fig
        assert "caption" in fig
        assert "ocr_text" in fig

        # Backward compatibility aliases
        assert "fig_index" in fig
        assert "page_num" in fig
        assert fig["figure_index"] == fig["fig_index"]
        assert fig["page_number"] == fig["page_num"]

        # Test SQLite insertion directly
        db_path = os.path.join(tmp_dir, "schema_test.sqlite")
        db = LwsDatabase(db_path)
        doc_id = db.upsert_document({"watch_folder": tmp_dir, "rel_path": "test.pdf", "filename": "test.pdf"})
        db.index_figures(doc_id, figs)

        cur = db.get_connection().cursor()
        cur.execute("SELECT figure_index, page_number, rel_image_path, caption FROM figures WHERE doc_id = ?", (doc_id,))
        row = cur.fetchone()
        assert row is not None
        assert row["figure_index"] == 1
        assert row["page_number"] == 1
        assert ".ai_cache/figures" in row["rel_image_path"]
        assert "Microstructure" in row["caption"]

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_cmyk_image_conversion():
    """Verifies CMYK or non-PNG image modes are converted to RGB without crashing Pillow."""
    cmyk_img = Image.new("CMYK", (200, 200), (0, 100, 100, 0))
    prepared = _prepare_image_for_save(cmyk_img)
    assert prepared.mode == "RGB"

    tmp_dir = tempfile.mkdtemp()
    try:
        docx_path = os.path.join(tmp_dir, "cmyk_doc.docx")
        _create_synthetic_docx(docx_path, image_mode="CMYK")
        out_dir = os.path.join(tmp_dir, "out")

        figs = extract_figures_from_docx(docx_path, out_dir)
        assert len(figs) == 1
        saved_file = figs[0]["file_path"]
        assert os.path.exists(saved_file)

        # Verify saved image is valid and readable
        with Image.open(saved_file) as img:
            assert img.format == "PNG"
            assert img.mode in ("RGB", "RGBA")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_min_dimension_filter():
    """Verifies small UI icons or bullet graphics below min_dimension are filtered out."""
    tmp_dir = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(tmp_dir, "small_icon.pdf")
        _create_synthetic_pdf(pdf_path, page_count=1, image_size=(80, 80))
        out_dir = os.path.join(tmp_dir, "out")

        # Default min_dimension is 140; 80x80 should be filtered
        figs_filtered = extract_figures_from_pdf(pdf_path, out_dir, min_dimension=140)
        assert len(figs_filtered) == 0

        # With min_dimension=50, it should be extracted
        figs_included = extract_figures_from_pdf(pdf_path, out_dir, min_dimension=50)
        assert len(figs_included) == 1
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_caption_regex_patterns():
    """Verifies CAPTION_PATTERN matches standard, supplementary, and chapter figure captions."""
    test_cases = [
        ("Fig. 1: Schematic of the printing engine.", "Fig. 1"),
        ("Figure 2 - Calibration curve across 405 nm laser power.", "Figure 2"),
        ("Extended Data Fig. 3: Stress-strain characteristics.", "Extended Data Fig. 3"),
        ("Supplementary Figure S4. Mass spectrometry spectrum.", "Supplementary Figure S4"),
        ("Scheme 1: Synthesis of biocompatible macromers.", "Scheme 1"),
        ("Table 2: Comparison of mechanical properties.", "Table 2"),
    ]
    for text, expected_tag in test_cases:
        matches = find_captions_in_page_text(text)
        assert len(matches) == 1, f"Failed matching: {text}"
        assert matches[0][0] == expected_tag
        assert expected_tag in matches[0][1]

    # Non-caption / empty edge cases
    assert find_captions_in_page_text("") == []
    assert find_captions_in_page_text(None) == []
    assert find_captions_in_page_text("No figures or schemes are mentioned in this text.") == []


# ----------------------------------------------------------------------
# 2. Synchronizer Deletion Reconciliation Tests
# ----------------------------------------------------------------------

def test_deletion_reconciliation_workflow():
    """Verifies deleted files are purged from SQLite DB, FTS tables, .ai_cache, and references.json."""
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_sync.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        # 1. Create two files
        doc1 = os.path.join(tmp_dir, "paper1.txt")
        doc2 = os.path.join(tmp_dir, "paper2.txt")
        with open(doc1, "w", encoding="utf-8") as f:
            f.write("Volumetric 3D printing requires photochemical optimization and tomographic projections.")
        with open(doc2, "w", encoding="utf-8") as f:
            f.write("Two-photon polymerization utilizes femtosecond titanium-sapphire laser pulses.")

        res1 = sync.sync_watch_folder(tmp_dir)
        assert res1["total"] == 2
        assert res1["processed"] == 2
        assert res1["deleted"] == 0

        # Verify DB and cache state
        assert len(db.search_fts("photochemical")) == 1
        assert len(db.search_fts("femtosecond")) == 1
        assert os.path.exists(os.path.join(tmp_dir, ".ai_cache", "paper1.md"))
        assert os.path.exists(os.path.join(tmp_dir, ".ai_cache", "paper2.md"))

        ref_json_path = os.path.join(tmp_dir, "references.json")
        with open(ref_json_path, "r", encoding="utf-8") as rf:
            ref_data = json.load(rf)
            assert "paper1.txt" in ref_data
            assert "paper2.txt" in ref_data

        # 2. Delete paper1.txt from disk
        os.remove(doc1)

        # 3. Re-run sync
        res2 = sync.sync_watch_folder(tmp_dir)
        assert res2["total"] == 1
        assert res2["skipped"] == 1
        assert res2["processed"] == 0
        assert res2["deleted"] == 1

        # Verify paper1 is purged from SQLite documents & FTS
        conn = db.get_connection()
        cur = conn.cursor()
        cur.execute("SELECT rel_path FROM documents")
        remaining = [r["rel_path"] for r in cur.fetchall()]
        assert remaining == ["paper2.txt"]
        assert len(db.search_fts("photochemical")) == 0
        assert len(db.search_fts("femtosecond")) == 1

        # Verify paper1 cache markdown is deleted
        assert not os.path.exists(os.path.join(tmp_dir, ".ai_cache", "paper1.md"))
        assert os.path.exists(os.path.join(tmp_dir, ".ai_cache", "paper2.md"))

        # Verify references.json and REFERENCES.md pruned
        with open(ref_json_path, "r", encoding="utf-8") as rf:
            ref_data2 = json.load(rf)
            assert "paper1.txt" not in ref_data2
            assert "paper2.txt" in ref_data2

        with open(os.path.join(tmp_dir, "REFERENCES.md"), "r", encoding="utf-8") as mf:
            cat_content = mf.read()
            assert "paper1.txt" not in cat_content
            assert "paper2.txt" in cat_content

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_deletion_reconciliation_figures_cleanup():
    """Verifies deleting a document with extracted figures purges figures from DB and disk cache."""
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "figures_del.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=True)

        pdf_path = os.path.join(tmp_dir, "figure_paper.pdf")
        _create_synthetic_pdf(pdf_path, page_count=1)

        res1 = sync.sync_watch_folder(tmp_dir)
        assert res1["processed"] == 1
        assert res1["errors"] == 0

        # Confirm figure stored in DB and file exists
        cur = db.get_connection().cursor()
        cur.execute("SELECT id, rel_image_path FROM figures")
        fig_row = cur.fetchone()
        assert fig_row is not None
        fig_disk_path = os.path.join(tmp_dir, fig_row["rel_image_path"])
        assert os.path.exists(fig_disk_path)

        # Delete PDF and re-sync
        os.remove(pdf_path)
        res2 = sync.sync_watch_folder(tmp_dir)
        assert res2["deleted"] == 1

        # Confirm DB figures and FTS purged
        cur.execute("SELECT count(*) as cnt FROM figures")
        assert cur.fetchone()["cnt"] == 0
        cur.execute("SELECT count(*) as cnt FROM figures_fts")
        assert cur.fetchone()["cnt"] == 0

        # Confirm figure file removed from disk
        assert not os.path.exists(fig_disk_path)

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_standalone_reconcile_deleted_files_api():
    """Verifies public reconcile_deleted_files method operates correctly."""
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "api_test.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        f1 = os.path.join(tmp_dir, "file1.txt")
        with open(f1, "w") as f:
            f.write("Content 1.")
        sync.sync_watch_folder(tmp_dir)

        os.remove(f1)
        deleted = sync.reconcile_deleted_files(tmp_dir)
        assert deleted == 1

        cur = db.get_connection().cursor()
        cur.execute("SELECT count(*) as cnt FROM documents")
        assert cur.fetchone()["cnt"] == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ----------------------------------------------------------------------
# 3. Direct Sentence Consumption & Headings Preservation
# ----------------------------------------------------------------------

def test_direct_sentence_consumption_preserves_docx_headings():
    """Verifies synchronizer preserves short headings from structured_pages (length <= 8)."""
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "headings.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        docx_path = os.path.join(tmp_dir, "paper.docx")
        _create_synthetic_docx(docx_path, heading_text="Methods")

        res = sync.sync_watch_folder(tmp_dir)
        assert res["processed"] == 1

        # Heading "Methods" is 7 chars long. If re-split with len > 8 it would be discarded.
        # Direct consumption preserves it in FTS.
        hits = db.search_fts("Methods")
        assert len(hits) == 1
        assert "Methods" in hits[0]["highlighted_snippet"]

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ----------------------------------------------------------------------
# 4. End-to-End Synchronization with Figures (PDF, DOCX, PPTX)
# ----------------------------------------------------------------------

def test_end_to_end_sync_with_figures_all_formats():
    """Verifies end-to-end sync with extract_figures=True on PDF, DOCX, and PPTX files."""
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "all_formats.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=True)

        pdf_path = os.path.join(tmp_dir, "paper.pdf")
        docx_path = os.path.join(tmp_dir, "supplement.docx")
        pptx_path = os.path.join(tmp_dir, "presentation.pptx")

        _create_synthetic_pdf(pdf_path, page_count=1)
        _create_synthetic_docx(docx_path, heading_text="Supplementary Protocol")
        _create_synthetic_pptx(pptx_path, slide_title="Experimental Results")

        res = sync.sync_watch_folder(tmp_dir)
        assert res["total"] == 3
        assert res["processed"] == 3
        assert res["errors"] == 0
        assert res["deleted"] == 0

        # Check all 3 documents stored in DB
        cur = db.get_connection().cursor()
        cur.execute("SELECT rel_path, doc_role, page_count FROM documents ORDER BY rel_path")
        docs = cur.fetchall()
        assert len(docs) == 3

        # Check figures extracted and stored in figures table
        cur.execute("SELECT doc_id, figure_index, page_number, rel_image_path, caption FROM figures")
        figs = cur.fetchall()
        assert len(figs) == 3

        for f in figs:
            assert f["figure_index"] >= 1
            assert f["page_number"] >= 1
            assert f["rel_image_path"].startswith(".ai_cache/figures/")
            disk_p = os.path.join(tmp_dir, f["rel_image_path"])
            assert os.path.exists(disk_p)

        # Check figure FTS search
        f_hits = db.search_figures("Microstructure")
        assert len(f_hits) == 1
        assert "Microstructure" in f_hits[0]["caption"]

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ----------------------------------------------------------------------
# 5. Defensive Error Boundaries and Nonexistent / Corrupted Files
# ----------------------------------------------------------------------

def test_extract_figures_nonexistent_and_corrupt_files():
    """Verifies figure extractors safely return [] without unhandled exceptions on invalid inputs."""
    tmp_dir = tempfile.mkdtemp()
    try:
        out_dir = os.path.join(tmp_dir, "out")

        # Nonexistent files
        assert extract_figures_from_pdf("nonexistent.pdf", out_dir) == []
        assert extract_figures_from_docx("nonexistent.docx", out_dir) == []
        assert extract_figures_from_pptx("nonexistent.pptx", out_dir) == []

        # Zero-byte corrupted files
        corrupt_pdf = os.path.join(tmp_dir, "bad.pdf")
        corrupt_docx = os.path.join(tmp_dir, "bad.docx")
        corrupt_pptx = os.path.join(tmp_dir, "bad.pptx")
        for p in (corrupt_pdf, corrupt_docx, corrupt_pptx):
            with open(p, "wb") as f:
                f.write(b"not a valid archive")

        assert extract_figures_from_pdf(corrupt_pdf, out_dir) == []
        assert extract_figures_from_docx(corrupt_docx, out_dir) == []
        assert extract_figures_from_pptx(corrupt_pptx, out_dir) == []
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
