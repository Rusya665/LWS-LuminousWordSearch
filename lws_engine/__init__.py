"""
lws_engine
==========
The core backend engine for LWS (Luminous Word Search) 2.0.
"""

from .config import ConfigManager, win_safe_path, to_file_uri
from .database import LwsDatabase
from .synchronizer import LwsSynchronizer
from .thesaurus import ThesaurusManager, SCIENTIFIC_SYNONYMS
from .bundler import classify_document_role, compute_bundle_id, bundle_document_record
from .extractors import extract_pdf, extract_docx, extract_pptx, split_into_sentences, extract_doi
from .figure_extractor import extract_figures_from_pdf, extract_figures_from_docx, extract_figures_from_pptx

__all__ = [
    "ConfigManager",
    "win_safe_path",
    "to_file_uri",
    "LwsDatabase",
    "LwsSynchronizer",
    "ThesaurusManager",
    "SCIENTIFIC_SYNONYMS",
    "classify_document_role",
    "compute_bundle_id",
    "bundle_document_record",
    "extract_pdf",
    "extract_docx",
    "extract_pptx",
    "split_into_sentences",
    "extract_doi",
    "extract_figures_from_pdf",
    "extract_figures_from_docx",
    "extract_figures_from_pptx",
]
