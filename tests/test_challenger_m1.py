"""
Adversarial Stress Test Suite for Milestone 1 (M1)
Empirical verification of defects in:
- lws_engine.bundler: classify_document_role, compute_bundle_id
- lws_engine.extractors: split_into_sentences
"""

import os
import sys

# Ensure UTF-8 output on Windows consoles
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pytest
from lws_engine.bundler import (
    classify_document_role,
    compute_bundle_id,
    bundle_document_record,
)
from lws_engine.extractors import split_into_sentences


# ===========================================================================
# 1. VERIFIED CONTRACTS: Sentence Tokenizer Citations & Abbreviations
# ===========================================================================

def test_reproduce_bug_equation_parentheses_split():
    """
    Verify 'Eq. (1)' is not broken into two fragments.
    """
    text = "The formulation is given in Eq. (1) where sigma is stress. The parameters are shown in Table 1."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "The formulation is given in Eq. (1) where sigma is stress."
    assert sents[1] == "The parameters are shown in Table 1."


def test_reproduce_bug_reference_brackets_split():
    """
    Verify 'Ref. [14]' is not broken into two fragments.
    """
    text = "Prior methods were synthesized by Ref. [14] for organic crystals. Our work improves the throughput."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "Prior methods were synthesized by Ref. [14] for organic crystals."
    assert sents[1] == "Our work improves the throughput."


def test_reproduce_bug_supplementary_figure_split():
    """
    Verify 'Fig. S1' (Supplementary Figure) is not broken into two fragments.
    """
    text = "Refer to Fig. S1 for the schematic diagram. The experimental data matches Fig. S2 closely."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "Refer to Fig. S1 for the schematic diagram."
    assert sents[1] == "The experimental data matches Fig. S2 closely."


def test_reproduce_bug_roman_numeral_figure_split():
    """
    Verify 'Fig. IV' is not broken into two fragments.
    """
    text = "As seen in Fig. IV, the structure crystallized. Refer to Fig. II for magnification."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "As seen in Fig. IV, the structure crystallized."
    assert sents[1] == "Refer to Fig. II for magnification."


def test_reproduce_bug_honorific_titles_destruction():
    """
    Verify 'Mr.' and 'Mrs.' are preserved and not destroyed.
    """
    text = "Mr. Smith arrived at noon. Mrs. Smith joined him later."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "Mr. Smith arrived at noon."
    assert sents[1] == "Mrs. Smith joined him later."


def test_reproduce_bug_terminal_quote_sentence_merging():
    """
    Verify sentences ending with punctuation inside quotation marks split properly.
    """
    text1 = 'He said, "This is fantastic!" Then he left.'
    sents1 = split_into_sentences(text1)
    assert len(sents1) == 2
    assert sents1[0] == 'He said, "This is fantastic!"'
    assert sents1[1] == "Then he left."

    text2 = 'The author stated, "The yield reached 98%." We verified the claim.'
    sents2 = split_into_sentences(text2)
    assert len(sents2) == 2
    assert sents2[0] == 'The author stated, "The yield reached 98%."'
    assert sents2[1] == "We verified the claim."


def test_reproduce_bug_sentence_ending_with_no():
    """
    Verify sentences ending with 'no.' split properly.
    """
    text = "The committee voted no. The proposal was rejected."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "The committee voted no."
    assert sents[1] == "The proposal was rejected."


def test_reproduce_bug_mid_sentence_capitalized_citations():
    """
    Verify mid-sentence 'et al.' does not break when followed by capitalized words.
    """
    text = "In the work of Kelly et al. Bayesian optimization improved the resolution."
    sents = split_into_sentences(text)
    assert len(sents) == 2
    assert sents[0] == "In the work of Kelly et al."
    assert sents[1] == "Bayesian optimization improved the resolution."


# ===========================================================================
# 2. VERIFIED CONTRACTS: Bundler Silicon and Companion Defects
# ===========================================================================

def test_reproduce_bug_silicon_materials_classified_as_si():
    """
    Verify silicon materials filenames are correctly classified as MAIN.
    """
    silicon_files = [
        "si_nanowires.pdf",
        "si_photonics.pdf",
        "si_anode_batteries.pdf",
        "si-wafers.pdf",
        "si_solar_cells.pdf",
    ]
    for fn in silicon_files:
        role = classify_document_role(fn)
        assert role == "MAIN", f"Expected 'MAIN' for {fn}, got '{role}'"


def test_reproduce_bug_missing_companion_journal_suffixes():
    """
    Verify standard academic journal companion extensions
    '_suppl', '_sm', and '_esm' are recognized as SI and bundle with main paper.
    """
    folder = "papers/2021"
    main_name = "Smith2021.pdf"
    main_bid = compute_bundle_id(f"{folder}/{main_name}", main_name)

    # _suppl
    suppl_name = "Smith2021_suppl.pdf"
    assert classify_document_role(suppl_name) == "SI"
    assert compute_bundle_id(f"{folder}/{suppl_name}", suppl_name) == main_bid

    # _sm (RSC / ACS)
    sm_name = "Smith2021_sm.pdf"
    assert classify_document_role(sm_name) == "SI"
    assert compute_bundle_id(f"{folder}/{sm_name}", sm_name) == main_bid

    # _esm (Springer)
    esm_name = "Smith2021_esm.pdf"
    assert classify_document_role(esm_name) == "SI"
    assert compute_bundle_id(f"{folder}/{esm_name}", esm_name) == main_bid


def test_reproduce_bug_non_latin_filename_bundle_id_collision():
    """
    Verify multilingual filenames preserve non-Latin letters and avoid collision.
    """
    folder = "papers"
    fn1 = "论文_2021_介绍.pdf"
    fn2 = "论文_2021_方法.pdf"
    bid1 = compute_bundle_id(f"{folder}/{fn1}", fn1)
    bid2 = compute_bundle_id(f"{folder}/{fn2}", fn2)
    assert bid1 != bid2
    assert "论文" in bid1 and "论文" in bid2


# ===========================================================================
# 3. CONTRACT SPECIFICATIONS (Activated)
# ===========================================================================

def test_contract_eq_parentheses():
    text = "The formulation is given in Eq. (1) where sigma is stress. The parameters are shown in Table 1."
    sents = split_into_sentences(text)
    assert len(sents) == 2


def test_contract_ref_brackets():
    text = "Prior methods were synthesized by Ref. [14] for organic crystals. Our work improves the throughput."
    sents = split_into_sentences(text)
    assert len(sents) == 2


def test_contract_fig_supplementary():
    text = "Refer to Fig. S1 for the schematic diagram. The experimental data matches Fig. S2 closely."
    sents = split_into_sentences(text)
    assert len(sents) == 2


def test_contract_silicon_nanowires():
    assert classify_document_role("si_nanowires.pdf") == "MAIN"
    assert classify_document_role("si_photonics.pdf") == "MAIN"


def test_contract_terminal_quotes():
    text = 'He said, "This is fantastic!" Then he left.'
    sents = split_into_sentences(text)
    assert len(sents) == 2
