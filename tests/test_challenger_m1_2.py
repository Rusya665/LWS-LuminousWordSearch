"""
Adversarial Stress Test Suite for Milestone 1: Extractors & Thesaurus
====================================================================
Agent: challenger_m1_2 (Extractors Adversarial Verifier)
Target:
- lws_engine.extractors: extract_pptx, extract_docx, extract_pdf,
                         split_into_sentences, extract_abstract,
                         extract_doi, parse_filename_metadata
- lws_engine.thesaurus: ThesaurusManager
"""

import os
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

# Ensure UTF-8 output on Windows consoles
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import docx
import pptx
import pytest
from reportlab.pdfgen import canvas

from lws_engine.extractors import (
    extract_abstract,
    extract_docx,
    extract_doi,
    extract_pdf,
    extract_pptx,
    parse_filename_metadata,
    split_into_sentences,
)
from lws_engine.thesaurus import ThesaurusManager


# ===========================================================================
# 1. PPTX EXTRACTION: Deeply Nested Group Shapes
# ===========================================================================

def test_reproduce_bug_pptx_nested_group_shapes_omitted():
    """
    Verify extract_pptx extracts text from shapes inside GroupShapes recursively.
    """
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Level 1 Group Shape
    g1 = slide.shapes.add_group_shape()
    tb1 = g1.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(1), pptx.util.Inches(2), pptx.util.Inches(1))
    tb1.text_frame.text = "Primary group level 1 text."

    # Level 2 Group Shape (nested inside Level 1)
    g2 = g1.shapes.add_group_shape()
    tb2 = g2.shapes.add_textbox(pptx.util.Inches(2), pptx.util.Inches(2), pptx.util.Inches(2), pptx.util.Inches(1))
    tb2.text_frame.text = "Secondary group level 2 text."

    # Level 3 Group Shape (nested inside Level 2)
    g3 = g2.shapes.add_group_shape()
    tb3 = g3.shapes.add_textbox(pptx.util.Inches(3), pptx.util.Inches(3), pptx.util.Inches(2), pptx.util.Inches(1))
    tb3.text_frame.text = "Tertiary deeply nested group level 3 text."

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        pptx_path = tmp.name

    try:
        prs.save(pptx_path)
        full_text, structured_pages, meta = extract_pptx(pptx_path)

        assert "Primary group level 1 text." in full_text
        assert "Secondary group level 2 text." in full_text
        assert "Tertiary deeply nested group level 3 text." in full_text
        assert any("Primary group level 1 text." in s for s in structured_pages[0]["sentences"])
        assert any("Secondary group level 2 text." in s for s in structured_pages[0]["sentences"])
        assert any("Tertiary deeply nested group level 3 text." in s for s in structured_pages[0]["sentences"])
    finally:
        if os.path.exists(pptx_path):
            os.remove(pptx_path)


def test_contract_pptx_nested_group_shapes():
    """Contract: PPTX extractor must capture text from grouped and nested shapes."""
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    g1 = slide.shapes.add_group_shape()
    tb1 = g1.shapes.add_textbox(pptx.util.Inches(1), pptx.util.Inches(1), pptx.util.Inches(2), pptx.util.Inches(1))
    tb1.text_frame.text = "Critical grouped diagram note."

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        pptx_path = tmp.name

    try:
        prs.save(pptx_path)
        full_text, structured_pages, meta = extract_pptx(pptx_path)
        assert "Critical grouped diagram note." in full_text
        assert any("Critical grouped diagram note." in s for s in structured_pages[0]["sentences"])
    finally:
        if os.path.exists(pptx_path):
            os.remove(pptx_path)


# ===========================================================================
# 2. DOCX EXTRACTION: Tables, Complex Layouts, and Short Cell Values
# ===========================================================================

def test_reproduce_docx_tables_short_values_omitted_from_sentences():
    """
    Verify in DOCX tables that short cell contents are preserved in sentences via row joining.
    """
    doc = docx.Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Parameter"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Laser"
    table.cell(1, 1).text = "405 nm"
    table.cell(2, 0).text = "Acidity"
    table.cell(2, 1).text = "pH 7.4"

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        docx_path = tmp.name

    try:
        doc.save(docx_path)
        full_text, structured_pages, meta = extract_docx(docx_path)
        sentences = structured_pages[0]["sentences"]

        assert any("Parameter" in s for s in sentences)
        assert any("Value" in s for s in sentences)
        assert any("Laser" in s for s in sentences)
        assert any("405 nm" in s for s in sentences)
        assert any("Acidity" in s for s in sentences)
        assert any("pH 7.4" in s for s in sentences)
        assert "Parameter | Value" in sentences
        assert "Laser | 405 nm" in sentences
        assert "Acidity | pH 7.4" in sentences
    finally:
        if os.path.exists(docx_path):
            os.remove(docx_path)


def test_docx_complex_cell_layouts_stability():
    """Verify DOCX extractor does not crash on merged cells, empty cells, and nested tables."""
    doc = docx.Document()
    table = doc.add_table(rows=3, cols=3)

    # Merged cells
    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(0, 0).text = "Merged Header Across Two Columns"

    # Empty cell (cell with no paragraph text)
    _ = table.cell(1, 0)

    # Multi-paragraph cell
    p_cell = table.cell(1, 1)
    p_cell.text = "First paragraph in cell."
    p_cell.add_paragraph("Second paragraph in same cell.")

    # Nested table inside cell
    inner_table = table.cell(2, 2).add_table(rows=2, cols=2)
    inner_table.cell(0, 0).text = "Inner Table Data Alpha"
    inner_table.cell(1, 1).text = "Inner Table Data Beta"

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        docx_path = tmp.name

    try:
        doc.save(docx_path)
        full_text, structured_pages, meta = extract_docx(docx_path)
        assert meta["format"] == "DOCX"
        assert "Merged Header Across Two Columns" in full_text
        assert "Inner Table Data Alpha" in full_text
        assert "Inner Table Data Beta" in full_text
    finally:
        if os.path.exists(docx_path):
            os.remove(docx_path)


# ===========================================================================
# 3. MALFORMED, EMPTY, CORRUPTED FILES & INVALID ARGUMENTS
# ===========================================================================

def test_extractors_zero_byte_and_corrupted_files():
    """Verify appropriate exception types raised on 0-byte or corrupted files."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_0 = tmp.name
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        docx_0 = tmp.name
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        pptx_0 = tmp.name

    # Write truncated garbage to docx and pptx
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        docx_bad = tmp.name
        tmp.write(b"PK\x03\x04corrupted-archive")
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        pptx_bad = tmp.name
        tmp.write(b"PK\x03\x04corrupted-pptx")

    try:
        # PDF 0-byte raises PdfiumError
        import pypdfium2 as pdfium
        with pytest.raises(Exception):
            extract_pdf(pdf_0)

        # DOCX 0-byte raises BadZipFile
        with pytest.raises(zipfile.BadZipFile):
            extract_docx(docx_0)

        # DOCX corrupted raises BadZipFile
        with pytest.raises(zipfile.BadZipFile):
            extract_docx(docx_bad)

        # PPTX 0-byte raises PackageNotFoundError
        with pytest.raises(Exception):
            extract_pptx(pptx_0)

        # PPTX corrupted raises PackageNotFoundError
        with pytest.raises(Exception):
            extract_pptx(pptx_bad)
    finally:
        for p in [pdf_0, docx_0, pptx_0, docx_bad, pptx_bad]:
            if os.path.exists(p):
                os.remove(p)


def test_extractors_invalid_argument_types():
    """Verify behavior when None, integers, or empty strings are passed to extractors."""
    # None argument raises TypeError
    with pytest.raises(TypeError):
        extract_pdf(None)
    with pytest.raises(TypeError):
        extract_docx(None)
    with pytest.raises(TypeError):
        extract_pptx(None)

    # Int argument raises TypeError
    with pytest.raises(TypeError):
        extract_pdf(12345)
    with pytest.raises(TypeError):
        extract_docx(12345)
    with pytest.raises(TypeError):
        extract_pptx(12345)


# ===========================================================================
# 4. SENTENCE TOKENIZER: Typographic Quotes, Surrogates, and Huge Strings
# ===========================================================================

def test_reproduce_bug_sentence_split_typographic_unicode_quotes():
    """
    Verify split_into_sentences splits when a sentence begins with Unicode curly quote.
    """
    text = "The initial reaction was completed. \u201cWe observed no optical aberrations,\u201d the authors concluded."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "The initial reaction was completed."
    assert sents[1] == "\u201cWe observed no optical aberrations,\u201d the authors concluded."


def test_contract_sentence_split_typographic_unicode_quotes():
    """Contract: split_into_sentences must split on sentences beginning with Unicode curly quotes."""
    text = "The initial reaction was completed. \u201cWe observed no optical aberrations,\u201d the authors concluded."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "The initial reaction was completed."


def test_reproduce_bug_sentence_split_emoji_and_symbols():
    """
    Verify split_into_sentences splits sentences that begin with emojis, bullet symbols, or dashes.
    """
    text = "First sentence ends here. \U0001F680 Rocket launched into orbit. Next sentence starts."
    sents = split_into_sentences(text)
    assert len(sents) == 3
    assert sents[0] == "First sentence ends here."
    assert sents[1] == "\U0001F680 Rocket launched into orbit."
    assert sents[2] == "Next sentence starts."


def test_contract_sentence_split_emoji_and_symbols():
    """Contract: sentences starting with emojis or symbols must split properly."""
    text = "First sentence ends here. \U0001F680 Rocket launched into orbit."
    sents = split_into_sentences(text)
    assert len(sents) == 2


def test_sentence_tokenization_surrogate_and_control_chars_stability():
    """Verify split_into_sentences safely processes Unicode surrogates and control chars without crashing."""
    text = "Normal text before here. Text with \ud800 surrogate and \x00 null and \x01\x02 control. Final sentence follows."
    sents = split_into_sentences(text)
    assert len(sents) == 3
    assert sents[0] == "Normal text before here."
    assert "Final sentence follows." in sents[2]


def test_sentence_tokenization_large_string_performance():
    """Stress test split_into_sentences with a 2-million-character string to rule out catastrophic backtracking."""
    large_text = ("Volumetric additive manufacturing with photopolymers cures objects in seconds. " * 30000)
    t0 = time.time()
    sents = split_into_sentences(large_text)
    elapsed = time.time() - t0

    assert elapsed < 3.0, f"Sentence splitting exceeded 3.0s threshold ({elapsed:.2f}s)"
    assert len(sents) > 1000


# ===========================================================================
# 5. ABSTRACT & DOI EXTRACTION ADVERSARIAL STRESS
# ===========================================================================

def test_reproduce_bug_abstract_false_positive_mid_sentence():
    """
    Verify extract_abstract does not match mid-sentence mentions of 'abstract'.
    """
    prose = (
        "We discuss the abstract concept of volumetric additive manufacturing in great detail.\n\n"
        "1. Introduction\n"
        "Here begins the actual introduction section of the manuscript."
    )
    result = extract_abstract(prose)
    assert result != "concept of volumetric additive manufacturing in great detail."


def test_contract_abstract_no_mid_sentence_false_positive():
    """Contract: extract_abstract must not extract mid-sentence mentions of 'abstract'."""
    prose = (
        "We discuss the abstract concept of volumetric additive manufacturing in great detail.\n\n"
        "1. Introduction\n"
        "Here begins the actual introduction section of the manuscript."
    )
    result = extract_abstract(prose)
    # Should fall back to the first paragraph or None, not 'concept of...'
    assert result != "concept of volumetric additive manufacturing in great detail."


def test_doi_extraction_adversarial_patterns():
    """Verify extract_doi behavior across trailing punctuation, query strings, and brackets."""
    # Standard trailing periods and semicolons trimmed
    assert extract_doi("doi: 10.1038/s41586-020-03079-9.") == "10.1038/s41586-020-03079-9"
    assert extract_doi("10.1016/j.addma.2021.102145;") == "10.1016/j.addma.2021.102145"

    # URL query strings stopped at delimiter
    assert extract_doi("https://doi.org/10.1126/science.abc1234?view=full") == "10.1126/science.abc1234"

    # Slashes and parentheses inside valid DOIs preserved
    assert extract_doi("Published under 10.1002/(sici)1097-4628(1999)71:1<1::aid-app1>3.0.co;2-8") is not None

    # Invalid short DOIs (< 4 digits) rejected
    assert extract_doi("10.12/invalid-doi") is None
    assert extract_doi("10.123/invalid-doi") is None


# ===========================================================================
# 6. FILENAME METADATA: Multilingual, Special Characters, & Edge Cases
# ===========================================================================

def test_parse_filename_metadata_multilingual_and_edge_cases():
    """Verify parse_filename_metadata with non-ASCII characters, underscores, and multiple years."""
    # German umlauts
    meta_de = parse_filename_metadata("2022 - Müller & Gödel - Tomographie_Verfahren.pdf")
    assert meta_de["year"] == 2022
    assert meta_de["authors"] == "Müller & Gödel"
    assert meta_de["title"] == "Tomographie Verfahren"

    # Multiple years in stem (leading year takes precedence)
    meta_multi = parse_filename_metadata("2020 - Smith - Review of 2019 Developments.pdf")
    assert meta_multi["year"] == 2020
    assert meta_multi["authors"] == "Smith"
    assert meta_multi["title"] == "Review of 2019 Developments"

    # File with no year
    meta_noyear = parse_filename_metadata("Advanced_Materials_Report.docx")
    assert meta_noyear["year"] is None
    assert meta_noyear["authors"] is None
    assert meta_noyear["title"] == "Advanced Materials Report"


# ===========================================================================
# 7. THESAURUS MANAGER: Missing WordNet, max_count TypeError, & Large Inputs
# ===========================================================================

def test_reproduce_bug_thesaurus_max_count_none_typeerror():
    """
    Verify ThesaurusManager.get_wordnet_synonyms handles max_count=None gracefully.
    """
    tm = ThesaurusManager(enable_wordnet=False)
    assert tm.get_wordnet_synonyms("resin", max_count=None) == []


def test_contract_thesaurus_max_count_none_graceful():
    """Contract: get_wordnet_synonyms should handle max_count=None gracefully without raising TypeError."""
    tm = ThesaurusManager(enable_wordnet=False)
    assert tm.get_wordnet_synonyms("resin", max_count=None) == []


def test_thesaurus_missing_wordnet_graceful_fallbacks(caplog):
    """Verify ThesaurusManager operates gracefully when WordNet raises LookupError or ImportError."""
    tm = ThesaurusManager(enable_wordnet=True)
    tm._wordnet_attempted = False

    with patch("nltk.corpus.wordnet.synsets", side_effect=LookupError("WordNet corpus missing")):
        with caplog.at_level("WARNING", logger="lws_engine.thesaurus"):
            wn = tm._get_wordnet()
            assert wn is None

        # Synonyms still return scientific results
        syns = tm.expand_query("pdms resin", use_wordnet=True, use_scientific=True)
        assert "pdms" in syns
        assert "polydimethylsiloxane" in syns["pdms"]


def test_thesaurus_extremely_large_query_stability():
    """Verify ThesaurusManager does not degrade or hang on 10,000-word queries."""
    tm = ThesaurusManager(enable_wordnet=False)
    huge_query = "volumetric additive manufacturing computed axial lithography pdms silicone " * 2000
    t0 = time.time()
    res = tm.expand_query(huge_query, use_wordnet=False, use_scientific=True)
    elapsed = time.time() - t0

    assert elapsed < 1.0, f"Thesaurus expansion took too long ({elapsed:.2f}s)"
    assert "volumetric" in res or "pdms" in res
