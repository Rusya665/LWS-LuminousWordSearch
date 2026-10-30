"""
Test suite for LWS 2.0 Engine
=============================
Validates sentence tokenization, bundling, thesaurus expansions,
FTS5 full-text search, figure indexing, and incremental synchronization.
"""

import os
import shutil
import tempfile
import pytest

from lws_engine.config import ConfigManager
from lws_engine.extractors import split_into_sentences, extract_doi
from lws_engine.bundler import classify_document_role, compute_bundle_id
from lws_engine.thesaurus import ThesaurusManager
from lws_engine.database import LwsDatabase
from lws_engine.synchronizer import LwsSynchronizer


def test_sentence_tokenization():
    text = (
        "Volumetric additive manufacturing enables fast fabrication. "
        "See Fig. 1 for details. "
        "Loterie et al. demonstrated high resolution at 405 nm. "
        "Is PDMS compatible? Yes, it is."
    )
    sentences = split_into_sentences(text)
    # Correctly splits into 5 sentences without breaking on "Fig. 1", "et al.", or "405 nm"
    assert len(sentences) == 5
    assert "Fig. 1" in sentences[1]
    assert "405 nm" in sentences[2]
    assert sentences[3] == "Is PDMS compatible?"
    assert sentences[4] == "Yes, it is."


def test_doi_extraction():
    sample = "Published in Nature. https://doi.org/10.1038/s41586-020-03079-9 received on Jan 2021."
    doi = extract_doi(sample)
    assert doi == "10.1038/s41586-020-03079-9"


def test_bundler_classification():
    # Roles
    assert classify_document_role("paper.pdf") == "MAIN"
    assert classify_document_role("2020 - Loterie - High-resolution SI.pdf") == "SI"
    assert classify_document_role("experiment_data_supp.pdf") == "SI"
    assert classify_document_role("US20210053284A1.pdf") == "PATENT"
    assert classify_document_role("peer_review_comments.pdf") == "PEER_REVIEW"
    assert classify_document_role("diagram.png") == "IMAGE"
    assert classify_document_role("presentation.pptx") == "SLIDES"

    # Bundle ID pairing
    b_main = compute_bundle_id("VAM/2020 - Loterie - High-res.pdf", "2020 - Loterie - High-res.pdf")
    b_si = compute_bundle_id("VAM/2020 - Loterie - High-res SI.pdf", "2020 - Loterie - High-res SI.pdf")
    assert b_main == b_si


def test_thesaurus_expansions():
    tm = ThesaurusManager(enable_wordnet=True)
    pdms_syns = tm.get_scientific_synonyms("pdms")
    assert "polydimethylsiloxane" in pdms_syns or "silicone" in pdms_syns

    vam_syns = tm.get_scientific_synonyms("vam")
    assert any("computed axial lithography" in s for s in vam_syns)

    expanded = tm.expand_query("pdms resin")
    assert "pdms" in expanded
    assert len(expanded["pdms"]) > 0


def test_database_and_fts():
    tmp_dir = tempfile.mkdtemp()
    try:
        db_path = os.path.join(tmp_dir, "test_lws.sqlite")
        db = LwsDatabase(db_path)

        # Upsert document
        doc_id = db.upsert_document({
            "watch_folder": tmp_dir,
            "rel_path": "test/paper.pdf",
            "filename": "paper.pdf",
            "title": "Tomographic Printing of PDMS",
            "authors": "Kelly et al.",
            "year": 2022,
            "doi": "10.1126/science.abc1234",
            "doc_role": "MAIN",
            "bundle_id": "test/paper",
            "page_count": 5,
            "file_size": 1024,
            "mtime": 123456.0,
            "md_path": ".ai_cache/test/paper.md",
            "abstract": "We report volumetric printing using a 405 nm DLP light engine."
        })
        assert doc_id > 0

        # Index sentences
        db.index_sentences(doc_id, [
            (1, 0, "Tomographic volumetric additive manufacturing requires index matching."),
            (2, 0, "We utilized PDMS elastomer formulations cured by TPO-L at 405 nm wavelength."),
            (3, 0, "Mechanical stiffness reached 1.2 MPa after UV post-curing.")
        ])

        # Test direct FTS search
        hits = db.search_fts("PDMS")
        assert len(hits) == 1
        assert hits[0]["doc_id"] == str(doc_id)
        assert "hit-direct" in hits[0]["highlighted_snippet"]

        # Test synonym expansion search
        hits_syn = db.search_fts("silicone", synonyms=["pdms"])
        assert len(hits_syn) == 1

        # Test annotation
        anno_id = db.save_annotation(doc_id, 2, 0, "Crucial formula note: check TPO-L wt%")
        assert anno_id > 0
        annos = db.get_annotations(doc_id)
        assert len(annos) == 1
        assert "TPO-L wt%" in annos[0]["note"]

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_synchronizer_workflow():
    tmp_dir = tempfile.mkdtemp()
    try:
        # Create nested folders
        sub1 = os.path.join(tmp_dir, "VAM")
        os.makedirs(sub1, exist_ok=True)

        doc1_path = os.path.join(sub1, "2021 - Test - Paper.txt")
        with open(doc1_path, "w", encoding="utf-8") as f:
            f.write("Computed axial lithography with 405 nm excitation produces high-resolution 3D parts.")

        db_path = os.path.join(tmp_dir, "test_lws.sqlite")
        db = LwsDatabase(db_path)
        sync = LwsSynchronizer(db=db, extract_figures=False)

        res = sync.sync_watch_folder(tmp_dir)
        assert res["total"] == 1
        assert res["processed"] == 1

        # Check .ai_cache generation
        md_cache_path = os.path.join(tmp_dir, ".ai_cache", "VAM", "2021 - Test - Paper.md")
        assert os.path.exists(md_cache_path)

        with open(md_cache_path, "r", encoding="utf-8") as mf:
            content = mf.read()
            assert "Computed axial lithography" in content
            assert "bundle_id:" in content

        # Check references.json
        ref_json = os.path.join(tmp_dir, "references.json")
        assert os.path.exists(ref_json)

        # Check incremental skip on second sync
        res2 = sync.sync_watch_folder(tmp_dir)
        assert res2["skipped"] == 1
        assert res2["processed"] == 0

        db.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
