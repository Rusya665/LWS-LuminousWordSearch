"""
tests.test_m3_m4_database_cli_bridge
====================================
Unit and integration tests for Milestone M3 (Database & FTS5 Query Hardening)
and Milestone M4 (CLI, Headless Sync, & Pywebview LwsApi Bridge).
"""

import os
import sys
import threading
import subprocess
import re
from pathlib import Path
import pytest

from lws_engine.config import ConfigManager
from lws_engine.database import LwsDatabase
from lws_engine.synchronizer import LwsSynchronizer
from lws_engine.thesaurus import ThesaurusManager
from LWS import LwsApi


# =========================================================================
# M3: Database Tests
# =========================================================================

def test_database_multithreaded_concurrency(tmp_path):
    """Verifies thread-safe concurrent reads/writes/searches across 8 threads."""
    db_file = tmp_path / "concurrent.sqlite"
    db = LwsDatabase(str(db_file))

    doc_id = db.upsert_document({
        "watch_folder": str(tmp_path),
        "rel_path": "paper.pdf",
        "filename": "paper.pdf",
        "title": "Concurrent Paper",
        "doc_role": "MAIN",
    })
    assert doc_id > 0

    errors = []

    def worker(worker_id):
        try:
            for i in range(10):
                db.index_sentences(doc_id, [(1, i, f"Worker {worker_id} sentence {i} mentioning PDMS and TPO-L")])
                hits = db.search_fts("PDMS")
                assert len(hits) >= 1
                db.save_annotation(doc_id, 1, i, f"Worker {worker_id} note {i}")
                annos = db.get_annotations(doc_id)
                assert len(annos) >= 1
        except Exception as e:
            errors.append(e)

    threads = [threading.Thread(target=worker, args=(t,)) for t in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    db.close()
    assert len(errors) == 0, f"Thread errors encountered: {errors}"


def test_database_in_memory():
    """Verifies :memory: database initialization, indexing, and querying."""
    db = LwsDatabase(":memory:")
    assert db.db_path == ":memory:"
    doc_id = db.upsert_document({
        "watch_folder": "mem",
        "rel_path": "mem_doc.pdf",
        "filename": "mem_doc.pdf",
        "title": "In Memory Doc",
    })
    assert doc_id > 0
    db.index_sentences(doc_id, [(1, 0, "Testing in memory FTS database")])
    hits = db.search_fts("memory")
    assert len(hits) == 1
    db.close()


def test_delete_document_by_rel_path_cascading(tmp_path):
    """Verifies atomic deletion of a document and its cascading dependencies."""
    db_file = tmp_path / "delete_test.sqlite"
    db = LwsDatabase(str(db_file))

    doc_id = db.upsert_document({
        "watch_folder": str(tmp_path),
        "rel_path": "nested/doc.pdf",
        "filename": "doc.pdf",
    })
    db.index_sentences(doc_id, [(1, 0, "Sentence to be deleted")])
    db.index_figures(doc_id, [{
        "figure_index": 0,
        "page_number": 1,
        "rel_image_path": "nested/fig.png",
        "caption": "Figure caption to be deleted"
    }])
    db.save_annotation(doc_id, 1, 0, "Annotation to be deleted")

    assert len(db.search_fts("deleted")) == 1
    assert len(db.search_figures("deleted")) == 1
    assert len(db.get_annotations(doc_id)) == 1

    # Normalization check: delete with backslash
    deleted = db.delete_document_by_rel_path("nested\\doc.pdf")
    assert deleted is True

    # Verify everything cascaded cleanly
    assert db.get_document_by_id(doc_id) is None
    assert len(db.search_fts("deleted")) == 0
    assert len(db.search_figures("deleted")) == 0
    assert len(db.get_annotations(doc_id)) == 0

    # Non-existent or empty path
    assert db.delete_document_by_rel_path("nonexistent.pdf") is False
    assert db.delete_document_by_rel_path("") is False
    db.close()


def test_index_figures_canonical_and_alias_keys(tmp_path):
    """Verifies that index_figures correctly handles both canonical and alias keys."""
    db_file = tmp_path / "figs_test.sqlite"
    db = LwsDatabase(str(db_file))

    doc_id = db.upsert_document({
        "watch_folder": str(tmp_path),
        "rel_path": "fig_doc.pdf",
        "filename": "fig_doc.pdf",
    })

    db.index_figures(doc_id, [
        {
            "fig_index": "2",
            "page_num": "5",
            "file_path": "sub\\image1.png",
            "caption": "Alias key figure",
            "width": "640",
            "height": "480",
            "ocr_text": "Sample OCR"
        },
        {
            "figure_index": 1,
            "page_number": 3,
            "rel_image_path": "sub/image2.png",
            "caption": "Canonical key figure"
        }
    ])

    figs = db.get_figures_for_doc(doc_id)
    assert len(figs) == 2
    f1 = [f for f in figs if "image1.png" in f["rel_image_path"]][0]
    assert f1["figure_index"] == 2
    assert f1["page_number"] == 5
    assert f1["rel_image_path"] == "sub/image1.png"
    assert f1["width"] == 640
    assert f1["height"] == 480

    db.close()


def test_boolean_fts_search_and_fallback(tmp_path):
    """Verifies boolean FTS5 searches (AND, OR, NOT, hyphenated terms) and syntax fallback."""
    db_file = tmp_path / "fts_bool.sqlite"
    db = LwsDatabase(str(db_file))

    doc_id = db.upsert_document({
        "watch_folder": str(tmp_path),
        "rel_path": "vam_paper.pdf",
        "filename": "vam_paper.pdf",
        "title": "VAM Paper",
    })

    db.index_sentences(doc_id, [
        (1, 0, "We demonstrated that PDMS elastomer formulations cured by TPO-L produce high fidelity parts."),
        (1, 1, "The hydrogel matrix was non-toxic and biocompatible."),
        (1, 2, "A toxic monomer caused cell death.")
    ])

    # 1. PDMS AND TPO-L (non-adjacent in sentence 0)
    hits = db.search_fts("PDMS AND TPO-L")
    assert len(hits) == 1
    assert int(hits[0]["sentence_index"]) == 0

    # 2. PDMS OR hydrogel (matches sentence 0 and sentence 1)
    hits_or = db.search_fts("PDMS OR hydrogel")
    assert len(hits_or) == 2

    # 3. PDMS NOT toxic (matches sentence 0, excludes toxic)
    hits_not = db.search_fts("PDMS NOT toxic")
    assert len(hits_not) == 1
    assert int(hits_not[0]["sentence_index"]) == 0

    # 4. Hyphenated term without boolean
    hits_hyphen = db.search_fts("TPO-L")
    assert len(hits_hyphen) == 1

    # 5. Malformed syntax fallback: unclosed quote
    hits_unclosed = db.search_fts('"PDMS elastomer')
    assert len(hits_unclosed) == 1

    # 6. Malformed syntax fallback: AND OR
    hits_and_or = db.search_fts("PDMS AND OR TPO-L")
    assert len(hits_and_or) == 1

    db.close()


# =========================================================================
# M4: CLI & Bridge Tests
# =========================================================================

def test_lws_api_add_watch_folder(tmp_path):
    """Verifies LwsApi validation on folder addition."""
    config_file = tmp_path / "config.json"
    cfg = ConfigManager(config_path=config_file)
    db = LwsDatabase(":memory:")
    thes = ThesaurusManager(enable_wordnet=False)
    sync = LwsSynchronizer(db=db, extract_figures=False)
    api = LwsApi(cfg, db, sync, thes)

    # Empty string
    res_empty = api.add_watch_folder("")
    assert res_empty["status"] == "error"

    # Non-existent path
    res_nonexist = api.add_watch_folder(str(tmp_path / "does_not_exist"))
    assert res_nonexist["status"] == "error"

    # Regular file path (not directory)
    test_file = tmp_path / "test.txt"
    test_file.write_text("hello")
    res_file = api.add_watch_folder(str(test_file))
    assert res_file["status"] == "error"

    # Valid directory
    sub_dir = tmp_path / "valid_watch"
    sub_dir.mkdir()
    res_valid = api.add_watch_folder(str(sub_dir))
    assert res_valid["status"] == "ok"
    assert str(sub_dir) in res_valid["watch_folders"]


def test_lws_api_choose_folder_dialog_cancel(tmp_path):
    """Verifies that cancelling pywebview folder dialog returns empty without tkinter fallback."""
    config_file = tmp_path / "config.json"
    cfg = ConfigManager(config_path=config_file)
    db = LwsDatabase(":memory:")
    thes = ThesaurusManager(enable_wordnet=False)
    sync = LwsSynchronizer(db=db, extract_figures=False)
    api = LwsApi(cfg, db, sync, thes)

    class MockWindow:
        def create_file_dialog(self, *args, **kwargs):
            return None

    api.set_window(MockWindow())
    res = api.choose_folder_dialog()
    assert res["status"] == "ok"
    assert res["added_folders"] == []


def test_lws_api_sync_folders_isolation(tmp_path):
    """Verifies that an exception in one watch folder does not crash other folder syncs."""
    config_file = tmp_path / "config.json"
    cfg = ConfigManager(config_path=config_file)
    db = LwsDatabase(":memory:")
    thes = ThesaurusManager(enable_wordnet=False)

    valid_folder = tmp_path / "valid"
    valid_folder.mkdir()
    failing_folder = tmp_path / "failing"
    failing_folder.mkdir()

    cfg.data["watched_roots"] = []
    cfg.add_watch_folder(str(failing_folder))
    cfg.add_watch_folder(str(valid_folder))

    class MockSynchronizer:
        def sync_watch_folder(self, folder):
            if "failing" in folder:
                raise RuntimeError("Simulated Disk I/O failure")
            return {"processed": 2, "skipped": 1, "errors": 0}

    api = LwsApi(cfg, db, MockSynchronizer(), thes)

    res = api.sync_folders()
    assert res["status"] == "ok"
    assert res["processed"] == 2
    assert res["errors"] == 1
    assert len(res["details"]) == 2


def test_lws_api_open_native_file_security(tmp_path):
    """Verifies path validation and executable extension blocking in open_native_file."""
    cfg = ConfigManager(config_path=tmp_path / "config.json")
    db = LwsDatabase(":memory:")
    thes = ThesaurusManager(enable_wordnet=False)
    sync = LwsSynchronizer(db=db, extract_figures=False)
    api = LwsApi(cfg, db, sync, thes)

    # Empty path
    assert api.open_native_file("")["status"] == "error"

    # Missing file
    assert api.open_native_file(str(tmp_path / "missing.pdf"))["status"] == "error"

    # Blocked dangerous extensions
    for ext in [".exe", ".bat", ".cmd", ".ps1", ".vbs"]:
        mal_file = tmp_path / f"test{ext}"
        mal_file.write_text("malicious")
        res = api.open_native_file(str(mal_file))
        assert res["status"] == "error"
        assert "Blocked executable" in res["message"]


def test_lws_api_search_regex_mode(tmp_path):
    """Verifies regex search mode and invalid regex handling in search_query."""
    db_file = tmp_path / "regex_test.sqlite"
    db = LwsDatabase(str(db_file))
    cfg = ConfigManager(config_path=tmp_path / "config.json")
    thes = ThesaurusManager(enable_wordnet=False)
    sync = LwsSynchronizer(db=db, extract_figures=False)
    api = LwsApi(cfg, db, sync, thes)

    doc_id = db.upsert_document({
        "watch_folder": str(tmp_path),
        "rel_path": "optics.pdf",
        "filename": "optics.pdf",
    })
    db.index_sentences(doc_id, [
        (1, 0, "Laser excitation was set at 405 nm with 15 mW power."),
        (1, 1, "Secondary measurement was at 380nm."),
        (1, 2, "No wavelength mentioned here.")
    ])

    # Valid regex matching 405 nm and 380nm
    res = api.search_query(r"\b\d+\s*nm\b", regex_mode=True)
    assert len(res["hits"]) == 2
    assert "hit-direct" in res["hits"][0]["highlighted_snippet"]

    # Invalid regex
    res_bad = api.search_query(r"[unclosed", regex_mode=True)
    assert "error" in res_bad
    assert res_bad["hits"] == []

    db.close()


def test_cli_headless_sync_exit_codes(tmp_path):
    """Verifies CLI exit codes for valid sync (0) and invalid directory arguments (1)."""
    python_exe = sys.executable
    lws_script = Path(__file__).resolve().parent.parent / "LWS.py"

    # 1. Valid folder -> exit 0
    valid_dir = tmp_path / "valid_cli_dir"
    valid_dir.mkdir()
    sample_file = valid_dir / "sample.txt"
    sample_file.write_text("Tomographic additive manufacturing paper.")

    try:
        proc_ok = subprocess.run(
            [python_exe, str(lws_script), "--dir", str(valid_dir), "--headless-sync", "--no-figures"],
            capture_output=True,
            text=True
        )
        assert proc_ok.returncode == 0, f"CLI stderr: {proc_ok.stderr}"
    finally:
        cfg = ConfigManager()
        cfg.remove_watch_folder(str(valid_dir))

    # 2. Invalid folder -> exit 1
    invalid_dir = tmp_path / "nonexistent_cli_dir_12345"
    proc_err = subprocess.run(
        [python_exe, str(lws_script), "--dir", str(invalid_dir), "--headless-sync", "--no-figures"],
        capture_output=True,
        text=True
    )
    assert proc_err.returncode == 1
    assert "Warning: Directory does not exist" in proc_err.stderr


def test_lws_api_get_figures_and_get_figure_data(tmp_path):
    """Verifies that LwsApi.get_figures and get_figure_data produce valid base64 data URIs."""
    from PIL import Image

    cfg_file = tmp_path / "config.json"
    cfg = ConfigManager(str(cfg_file))
    db_file = tmp_path / "test_fig_api.sqlite"
    db = LwsDatabase(str(db_file))
    thes = ThesaurusManager(enable_wordnet=False)
    sync = LwsSynchronizer(db=db, extract_figures=True)
    api = LwsApi(cfg, db, sync, thes)

    # Create dummy watch directory, document and figure image file
    watch_dir = tmp_path / "watch_papers"
    watch_dir.mkdir()
    figures_dir = watch_dir / ".ai_cache" / "figures"
    figures_dir.mkdir(parents=True)

    img_path = figures_dir / "sample_fig.png"
    im = Image.new("RGB", (100, 100), color=(255, 0, 0))
    im.save(str(img_path))

    doc_id = db.upsert_document({
        "watch_folder": str(watch_dir),
        "rel_path": "paper1.pdf",
        "filename": "paper1.pdf",
        "title": "VAM Optical Paper",
        "doc_role": "MAIN"
    })

    db.index_figures(doc_id, [{
        "fig_index": 1,
        "page_num": 3,
        "rel_image_path": ".ai_cache/figures/sample_fig.png",
        "caption": "Figure 1: High fidelity tomographic VAM reconstruction.",
        "width": 100,
        "height": 100
    }])

    # 1. Test get_figures
    figs = api.get_figures("tomographic")
    assert len(figs) == 1
    fig_item = figs[0]
    assert fig_item["page_number"] == 3
    assert fig_item["data_url"].startswith("data:image/png;base64,")
    assert os.path.exists(fig_item["full_image_path"])
    assert fig_item["full_doc_path"] == str(watch_dir / "paper1.pdf")

    # 2. Test get_figure_data
    fig_id = fig_item["id"]
    fig_data = api.get_figure_data(fig_id)
    assert fig_data["status"] == "ok"
    assert fig_data["data_url"].startswith("data:image/png;base64,")
    assert os.path.exists(fig_data["full_image_path"])

    db.close()


def test_synchronizer_protects_figures_directory(tmp_path):
    """Verifies that reconcile_deleted_files does not delete .ai_cache or .ai_cache/figures."""
    db = LwsDatabase(":memory:")
    sync = LwsSynchronizer(db=db, extract_figures=True)

    watch_dir = tmp_path / "watch_prot"
    watch_dir.mkdir()
    ai_cache = watch_dir / ".ai_cache"
    figures_cache = ai_cache / "figures"
    figures_cache.mkdir(parents=True)

    # Empty figures directory must NOT be deleted
    sync.reconcile_deleted_files(str(watch_dir), ref_data={}, discovered_rel_paths=set())
    assert figures_cache.exists(), "The .ai_cache/figures directory must be protected and preserved"
    assert ai_cache.exists(), "The .ai_cache directory must be protected and preserved"

    db.close()
