"""
tests/test_challenger_m2_figures.py
===================================
Adversarial Stress Test Suite for Milestone 2 (Figure Extractor & Synchronizer)
Agent: challenger_m2_1 (Figure Extraction Adversarial Challenger)

Empirical verification of:
1. Adversarial input boundaries:
   - Zero-byte files, non-existent paths, non-PDF files passed to extract_figures_from_pdf.
   - String vs Path instances across all functions.
   - PDFs with 0 images, 1 image, >50 images with max_figures limit (exact count, no resource leaks).
   - Image color modes: RGBA, CMYK, Palette (P), Grayscale (L), 1-bit monochrome (1), LA, HSV, LAB.
   - Corrupted image streams in DOCX/PPTX zip packages.
   - Non-ASCII, Unicode (Cyrillic, CJK), emoji, and spaces in filenames and output directory paths.
   - Deletion reconciliation when figures are extracted.
2. Caption Regex Matching:
   - Standard and edge-case caption headers: Fig. S1:, Figure 2A., FIG. 10 -, Scheme 1:, Chart 4:.
3. Empirical Defect Reproductions:
   - Figure filename collisions across multiple documents in the same watch folder.
   - Missing 'Chart' header in caption regex.
   - Multi-panel figure range caption truncation (e.g., Figure 1A-D:).
   - Unhandled TypeError on string max_figures.
"""

import io
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Tuple

import pytest
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Ensure UTF-8 console output on Windows
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from lws_engine.config import win_safe_path
from lws_engine.database import LwsDatabase
from lws_engine.figure_extractor import (
    CAPTION_PATTERN,
    _compute_rel_image_path,
    _prepare_image_for_save,
    extract_figures_from_docx,
    extract_figures_from_pdf,
    extract_figures_from_pptx,
    find_captions_in_page_text,
)
from lws_engine.synchronizer import LwsSynchronizer


# ---------------------------------------------------------------------------
# Test Helpers
# ---------------------------------------------------------------------------

def _create_test_image(
    path: str,
    size: Tuple[int, int] = (160, 160),
    color=(100, 150, 200),
    mode: str = "RGB"
) -> str:
    """Helper to generate a test image file."""
    im = Image.new(mode, size, color)
    im.save(path, format="PNG")
    return path


def _create_multipage_pdf_with_images(
    pdf_path: str,
    page_count: int = 10,
    images_per_page: int = 2,
    image_size: Tuple[int, int] = (150, 150)
) -> str:
    """Helper to generate a multipage PDF with multiple images per page."""
    c = canvas.Canvas(pdf_path, pagesize=letter)
    tmp_dir = os.path.dirname(pdf_path)
    img_sample = os.path.join(tmp_dir, "tmp_canvas_img.png")
    _create_test_image(img_sample, size=image_size)

    for p in range(1, page_count + 1):
        c.drawString(72, 720, f"Fig. {p}: Page {p} microstructure observation.")
        for i in range(1, images_per_page + 1):
            y_pos = 450 - (i - 1) * 160
            c.drawImage(img_sample, 72, y_pos, 140, 140)
        c.showPage()

    c.save()
    if os.path.exists(img_sample):
        try:
            os.remove(img_sample)
        except Exception:
            pass
    return pdf_path


# ===========================================================================
# 1. VERIFIED CONTRACTS: Input Boundaries & Type Resilience
# ===========================================================================

def test_contract_extract_figures_zero_byte_and_corrupt_files():
    """Verify zero-byte, truncated, and corrupt non-PDF/DOCX/PPTX files return [] gracefully."""
    tmp_dir = tempfile.mkdtemp()
    try:
        out_dir = os.path.join(tmp_dir, "out")

        # Zero-byte files
        for ext in [".pdf", ".docx", ".pptx"]:
            p = os.path.join(tmp_dir, f"zero{ext}")
            open(p, "wb").close()

        assert extract_figures_from_pdf(os.path.join(tmp_dir, "zero.pdf"), out_dir) == []
        assert extract_figures_from_docx(os.path.join(tmp_dir, "zero.docx"), out_dir) == []
        assert extract_figures_from_pptx(os.path.join(tmp_dir, "zero.pptx"), out_dir) == []

        # Corrupted / invalid binary files
        for ext in [".pdf", ".docx", ".pptx"]:
            p = os.path.join(tmp_dir, f"corrupt{ext}")
            with open(p, "wb") as f:
                f.write(b"\x00\xff\xfe\x01\x02\x03NotAValidBinaryStream\x00")

        assert extract_figures_from_pdf(os.path.join(tmp_dir, "corrupt.pdf"), out_dir) == []
        assert extract_figures_from_docx(os.path.join(tmp_dir, "corrupt.docx"), out_dir) == []
        assert extract_figures_from_pptx(os.path.join(tmp_dir, "corrupt.pptx"), out_dir) == []

        # Cross-extension mismatches (e.g., DOCX passed to PDF extractor)
        assert extract_figures_from_pdf(os.path.join(tmp_dir, "corrupt.docx"), out_dir) == []
        assert extract_figures_from_docx(os.path.join(tmp_dir, "corrupt.pdf"), out_dir) == []
        assert extract_figures_from_pptx(os.path.join(tmp_dir, "corrupt.pdf"), out_dir) == []
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_contract_extract_figures_string_and_path_instances():
    """Verify both str and Path instances are accepted across all figure extractors."""
    tmp_dir = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(tmp_dir, "sample.pdf")
        _create_multipage_pdf_with_images(pdf_path, page_count=1, images_per_page=1)

        out_str = os.path.join(tmp_dir, "out_str")
        out_path = Path(tmp_dir) / "out_path"

        # Str path & Str out_dir
        res1 = extract_figures_from_pdf(str(pdf_path), str(out_str))
        assert len(res1) == 1
        assert isinstance(res1[0]["file_path"], str)
        assert isinstance(res1[0]["rel_image_path"], str)

        # Path path & Path out_dir
        res2 = extract_figures_from_pdf(Path(pdf_path), out_path)
        assert len(res2) == 1
        assert isinstance(res2[0]["file_path"], str)
        assert isinstance(res2[0]["rel_image_path"], str)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_contract_extract_figures_nonexistent_and_invalid_argument_types():
    """Verify None, int, and non-existent paths return [] without crashing."""
    tmp_dir = tempfile.mkdtemp()
    try:
        out_dir = os.path.join(tmp_dir, "out")

        # Non-existent files
        assert extract_figures_from_pdf("does_not_exist_xyz.pdf", out_dir) == []
        assert extract_figures_from_docx("does_not_exist_xyz.docx", out_dir) == []
        assert extract_figures_from_pptx("does_not_exist_xyz.pptx", out_dir) == []

        # None or integer invalid argument types
        assert extract_figures_from_pdf(None, out_dir) == []
        assert extract_figures_from_docx(None, out_dir) == []
        assert extract_figures_from_pptx(None, out_dir) == []
        assert extract_figures_from_pdf(12345, out_dir) == []
        assert extract_figures_from_docx(12345, out_dir) == []
        assert extract_figures_from_pptx(12345, out_dir) == []

        # None output_dir
        assert extract_figures_from_pdf("dummy.pdf", None) == []
        assert extract_figures_from_docx("dummy.docx", None) == []
        assert extract_figures_from_pptx("dummy.pptx", None) == []
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ===========================================================================
# 2. VERIFIED CONTRACTS: Scale, Max Figures & Resource Leaks
# ===========================================================================

def test_contract_pdf_scale_and_max_figures_bounds():
    """
    Verify PDFs with 0 images, 1 image, and >50 images with max_figures limits
    enforce exact counts and release file handles on Windows.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        # 1. PDF with 0 images
        zero_pdf = os.path.join(tmp_dir, "zero_img.pdf")
        c = canvas.Canvas(zero_pdf, pagesize=letter)
        c.drawString(100, 700, "Text only. No embedded images anywhere.")
        c.save()

        res_zero = extract_figures_from_pdf(zero_pdf, os.path.join(tmp_dir, "out_zero"))
        assert len(res_zero) == 0

        # 2. PDF with 60 images (30 pages * 2 images per page)
        large_pdf = os.path.join(tmp_dir, "sixty_figs.pdf")
        _create_multipage_pdf_with_images(large_pdf, page_count=30, images_per_page=2)

        out_base = os.path.join(tmp_dir, "out_large")

        # Test exact boundary limits
        res_50 = extract_figures_from_pdf(large_pdf, out_base + "_50", max_figures=50)
        assert len(res_50) == 50

        res_30 = extract_figures_from_pdf(large_pdf, out_base + "_30", max_figures=30)
        assert len(res_30) == 30

        res_1 = extract_figures_from_pdf(large_pdf, out_base + "_1", max_figures=1)
        assert len(res_1) == 1

        res_0 = extract_figures_from_pdf(large_pdf, out_base + "_0", max_figures=0)
        assert len(res_0) == 0

        res_neg = extract_figures_from_pdf(large_pdf, out_base + "_neg", max_figures=-10)
        assert len(res_neg) == 0

        res_all = extract_figures_from_pdf(large_pdf, out_base + "_all", max_figures=None)
        assert len(res_all) == 60

        # Verify no file handle lock remains on Windows
        os.remove(large_pdf)
        assert not os.path.exists(large_pdf)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ===========================================================================
# 3. VERIFIED CONTRACTS: Color Modes & Corrupted Zip Media
# ===========================================================================

def test_contract_all_pillow_image_modes_converted_and_saved():
    """Verify RGBA, CMYK, Palette (P), Grayscale (L), 1-bit monochrome (1), and exotic modes."""
    modes = ["1", "L", "P", "RGB", "RGBA", "CMYK", "YCbCr", "LAB", "HSV", "LA"]
    for mode in modes:
        im = Image.new(mode, (80, 80))
        prepared = _prepare_image_for_save(im)
        buf = io.BytesIO()
        prepared.save(buf, format="PNG")
        assert len(buf.getvalue()) > 0
        with Image.open(buf) as reloaded:
            assert reloaded.format == "PNG"


def test_contract_docx_corrupted_image_streams_resilience():
    """Verify DOCX with corrupted or zero-byte media streams salvages valid images."""
    tmp_dir = tempfile.mkdtemp()
    try:
        docx_path = os.path.join(tmp_dir, "mixed_media.docx")
        out_dir = os.path.join(tmp_dir, "out_docx")

        im = Image.new("RGB", (180, 180), (255, 120, 0))
        valid_buf = io.BytesIO()
        im.save(valid_buf, format="PNG")

        with zipfile.ZipFile(docx_path, "w") as z:
            z.writestr("word/document.xml", "<w:document/>")
            z.writestr("word/media/image1.png", b"CORRUPTED_GARBAGE_PAYLOAD")
            z.writestr("word/media/image2.png", b"")  # Zero bytes
            z.writestr("word/media/image3.png", valid_buf.getvalue())

        figs = extract_figures_from_docx(docx_path, out_dir)
        assert len(figs) == 1
        assert figs[0]["figure_index"] == 3
        assert os.path.exists(figs[0]["file_path"])
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_contract_pptx_corrupted_image_streams_resilience():
    """Verify PPTX with corrupted media stream salvages valid images."""
    tmp_dir = tempfile.mkdtemp()
    try:
        pptx_path = os.path.join(tmp_dir, "mixed_slide_media.pptx")
        out_dir = os.path.join(tmp_dir, "out_pptx")

        im = Image.new("RGB", (180, 180), (0, 200, 100))
        valid_buf = io.BytesIO()
        im.save(valid_buf, format="PNG")

        with zipfile.ZipFile(pptx_path, "w") as z:
            z.writestr("ppt/presentation.xml", "<p:presentation/>")
            z.writestr("ppt/media/image1.png", b"TRUNCATED_GARBAGE")
            z.writestr("ppt/media/image2.png", valid_buf.getvalue())

        figs = extract_figures_from_pptx(pptx_path, out_dir)
        assert len(figs) == 1
        assert figs[0]["figure_index"] == 2
        assert os.path.exists(figs[0]["file_path"])
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ===========================================================================
# 4. VERIFIED CONTRACTS: Unicode, Emojis & Spaces in Paths
# ===========================================================================

def test_contract_unicode_emojis_and_spaces_in_paths():
    """Verify figure extraction across directory and file paths with emojis, spaces, and Unicode."""
    tmp_dir = tempfile.mkdtemp()
    try:
        special_dir = os.path.join(tmp_dir, "📁 Эксперименты & 测试 🔬 2026")
        os.makedirs(special_dir, exist_ok=True)
        pdf_path = os.path.join(special_dir, "图表 1 — 论文 (version #1) [draft] + notes.pdf")
        docx_path = os.path.join(special_dir, "报告 📄 #2.docx")
        out_dir = os.path.join(special_dir, ".ai_cache", "figures 🖼️")

        _create_multipage_pdf_with_images(pdf_path, page_count=1, images_per_page=1)

        im = Image.new("RGB", (160, 160), (50, 100, 150))
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        with zipfile.ZipFile(docx_path, "w") as z:
            z.writestr("word/document.xml", "<w:document/>")
            z.writestr("word/media/image1.png", buf.getvalue())

        res_pdf = extract_figures_from_pdf(pdf_path, out_dir)
        assert len(res_pdf) == 1
        assert os.path.exists(res_pdf[0]["file_path"])

        res_docx = extract_figures_from_docx(docx_path, out_dir)
        assert len(res_docx) == 1
        assert os.path.exists(res_docx[0]["file_path"])
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ===========================================================================
# 5. VERIFIED CONTRACTS: Caption Regex Standard Patterns
# ===========================================================================

def test_contract_caption_regex_standard_patterns():
    """Verify CAPTION_PATTERN matches standard and supplementary headers."""
    samples = [
        ("Fig. S1: Supplementary characterization curves.", "Fig. S1"),
        ("Figure 2A. Scanning electron microscopy image.", "Figure 2A"),
        ("FIG. 10 - High resolution tomography slice.", "FIG. 10"),
        ("Scheme 1: Chemical synthetic pathway.", "Scheme 1"),
        ("Extended Data Fig. 5: Cross sectional topography.", "Extended Data Fig. 5"),
        ("Supplementary Figure S3a: Gel electrophoresis.", "Supplementary Figure S3a"),
        ("Table 1: Physical constants.", "Table 1"),
    ]
    for text, expected_tag in samples:
        matches = find_captions_in_page_text(text)
        assert len(matches) == 1, f"Failed matching caption for: {text}"
        assert matches[0][0] == expected_tag


# ===========================================================================
# 6. EMPIRICAL DEFECT REPRODUCTIONS (BUGS FOUND BY CHALLENGER)
# ===========================================================================

def test_reproduce_bug_figure_filename_collision_multiple_documents():
    """
    CRITICAL DEFECT:
    In synchronizer.py:274, figures_cache_dir is '<watch_folder>/.ai_cache/figures'.
    In figure_extractor.py:145, out_filename is 'fig_p{page_num}_{obj_idx + 1}.png' (omitting file_path.stem).
    In figure_extractor.py:233, out_filename is 'fig_docx_{idx}.png'.
    In figure_extractor.py:307, out_filename is 'fig_slide_{idx}.png'.

    When multiple PDFs (e.g., doc1.pdf and doc2.pdf) are synced, both write to
    '.ai_cache/figures/fig_p1_1.png'. Document 2 overwrites and destroys Document 1's
    image file on disk, and both documents point to the same image in SQLite.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "collision_test.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=True)

        # Create doc1.pdf with a RED image
        img1_p = os.path.join(tmp_dir, "red.png")
        im1 = Image.new("RGB", (150, 150), (255, 0, 0))
        im1.save(img1_p)
        pdf1 = os.path.join(tmp_dir, "doc1.pdf")
        c1 = canvas.Canvas(pdf1)
        c1.drawImage(img1_p, 50, 50, 150, 150)
        c1.save()

        # Create doc2.pdf with a BLUE image
        img2_p = os.path.join(tmp_dir, "blue.png")
        im2 = Image.new("RGB", (150, 150), (0, 0, 255))
        im2.save(img2_p)
        pdf2 = os.path.join(tmp_dir, "doc2.pdf")
        c2 = canvas.Canvas(pdf2)
        c2.drawImage(img2_p, 50, 50, 150, 150)
        c2.save()

        os.remove(img1_p)
        os.remove(img2_p)

        res = sync.sync_watch_folder(tmp_dir)
        assert res["total"] == 2
        assert res["processed"] == 2

        cur = db.get_connection().cursor()
        cur.execute("SELECT doc_id, figure_index, rel_image_path FROM figures ORDER BY doc_id")
        rows = cur.fetchall()
        assert len(rows) == 2

        # Empirical assertion: Each document must have a UNIQUE image path on disk!
        # When the bug is present, rows[0]["rel_image_path"] == rows[1]["rel_image_path"] == ".ai_cache/figures/fig_p1_1.png"
        path1 = rows[0]["rel_image_path"]
        path2 = rows[1]["rel_image_path"]
        assert path1 != path2, (
            f"DEFECT REPRODUCED: Figure filename collision! Both documents share rel_image_path='{path1}'. "
            f"out_filename must include document identifier (e.g. file_path.stem) to avoid clobbering."
        )

        # Furthermore, verify that doc1's file on disk still contains the RED image
        full1 = os.path.join(tmp_dir, path1)
        with Image.open(full1) as im:
            pixel = im.getpixel((10, 10))
            assert pixel[0] > 200 and pixel[2] == 0, (
                f"DEFECT REPRODUCED: doc1 figure was overwritten by doc2! Pixel is {pixel}, expected Red."
            )

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_reproduce_bug_caption_chart_matching():
    r"""
    DEFECT:
    In figure_extractor.py:28, CAPTION_PATTERN only checks:
    (?:Figs?\.?|Figures?|Schemes?|Tables?)
    Scientific papers frequently use 'Chart 1:', 'Chart 4:', which are not matched.
    """
    chart_caption = "Chart 4: Power distribution profile across the optical field."
    matches = find_captions_in_page_text(chart_caption)
    assert len(matches) == 1, (
        f"DEFECT REPRODUCED: CAPTION_PATTERN omitted 'Chart' from caption keywords; "
        f"got {matches} for '{chart_caption}'"
    )
    assert matches[0][0] == "Chart 4"


def test_reproduce_bug_caption_multipanel_dash_truncation():
    """
    DEFECT:
    In figure_extractor.py:28, the word boundary \b after [a-zA-Z]? causes
    hyphenated multi-panel ranges like 'Figure 1A-D:' to be split:
    the tag becomes 'Figure 1A' and the description starts with 'D: ...'.
    """
    multipanel = "Figure 1A-D: Multi-panel experimental microscopy."
    matches = find_captions_in_page_text(multipanel)
    assert len(matches) == 1
    tag, desc = matches[0]
    assert tag in ("Figure 1A-D", "Figure 1A-D:"), (
        f"DEFECT REPRODUCED: Tag truncated to '{tag}' instead of 'Figure 1A-D'!"
    )
    assert not desc.startswith("Figure 1A: D:"), (
        f"DEFECT REPRODUCED: Description corrupted into '{desc}'!"
    )


def test_reproduce_bug_max_figures_string_typeerror():
    """
    DEFECT:
    In extract_figures_from_pdf, docx, pptx:
    'if max_figures is not None and max_figures <= 0:'
    raises TypeError when max_figures is passed as a string ('30').
    Extractors should defensively coerce max_figures = int(max_figures) if max_figures is not None.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(tmp_dir, "sample.pdf")
        _create_multipage_pdf_with_images(pdf_path, page_count=1)
        out_dir = os.path.join(tmp_dir, "out")

        # Passing max_figures="30" (e.g. from CLI or unparsed config)
        res = extract_figures_from_pdf(pdf_path, out_dir, max_figures="30")
        assert len(res) == 2

        # Verify max_figures="1" bounds to 1
        out_dir1 = os.path.join(tmp_dir, "out1")
        res1 = extract_figures_from_pdf(pdf_path, out_dir1, max_figures="1")
        assert len(res1) == 1

        # Verify max_figures="0" returns empty list
        out_dir0 = os.path.join(tmp_dir, "out0")
        res0 = extract_figures_from_pdf(pdf_path, out_dir0, max_figures="0")
        assert len(res0) == 0
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
