"""
lws_engine.bundler
==================
Analyzes document metadata and file naming patterns to classify document role
(MAIN, SI, PEER_REVIEW, PATENT, SLIDES, IMAGE) and bundle companion documents
(such as main paper + supplementary information + peer review file) together.
"""

from typing import Dict, Any, Tuple
import os
import re

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
SLIDES_EXTENSIONS = {".pptx", ".ppt"}

# Regex for patent detection:
# Matches authority codes (US, WO, EP) followed by digits (e.g. US20210053284, US10123456, WO2020123456, EP2012345)
# or authority code followed by hyphen/underscore and 4+ digits (e.g. US-2021-0053284).
# Uses negative lookbehind (?<![a-z]) to prevent false positives like 'focus2020.pdf', 'status2021.pdf', 'step2020.pdf'.
# Matches whole-word 'patent' or 'patents' avoiding false positives on words like 'patently'.
PATENT_PATTERNS = [
    re.compile(r"(?<![a-z])(?:us|wo|ep)(?:\d{2}|[-_]\d{4,})", re.IGNORECASE),
    re.compile(r"(?<![a-z])patents?(?![a-z])", re.IGNORECASE),
]

# Delimiters supporting standard punctuation and Unicode en-dash / em-dash
_DELIM_CHARS = r"_\s\-\.\u2013\u2014"

# Regex for peer review detection:
# Matches phrases like peer_review, reviewer_comments, editorial_decision, peer-review
PEER_REVIEW_PATTERNS = [
    re.compile(
        rf"(?:^|[{_DELIM_CHARS}\[\(])(?:peer[{_DELIM_CHARS}]*review(?:[{_DELIM_CHARS}]*comments)?|reviewer[{_DELIM_CHARS}]*comments|editorial[{_DELIM_CHARS}]*decision)(?:$|[{_DELIM_CHARS}\]\)])",
        re.IGNORECASE,
    ),
]

# Regex for supplementary information (SI) detection:
# Hardened against false positives (e.g. silicon chemistry files like si_nanowires.pdf).
# 1. Unambiguous companion keywords (supplementary, supporting, supplement, suppl, esm).
# 2. 'supp' with descriptor (information, info, materials, data, notes), version/numbering, or trailing.
# 3. 'note' / 'notes' as trailing companion or in brackets (e.g. 'Smith2021_note_1', 'Smith2021 (Note 1)').
# 4. 'si' and 'sm' only when:
#    - trailing suffix before extension (e.g. Kelly2021_si.pdf, Kelly2021-si.pdf, Smith2021_sm.pdf, Kelly2021_si_v2.pdf)
#    - enclosed in brackets (e.g. [SI], (SI), [SM], (SM))
#    - explicitly followed by info/information/material/data/notes (e.g. si_info.pdf, si_data.pdf)
SI_PATTERNS = [
    re.compile(
        rf"(?:^|[{_DELIM_CHARS}\[\(])(?:supplementary|supporting|supplement|suppl|esm)(?:[{_DELIM_CHARS}]*(?:information|info|materials?|data|notes?))?(?:[{_DELIM_CHARS}]*(?:v\d+|\d+))?(?:$|[{_DELIM_CHARS}\]\)])",
        re.IGNORECASE,
    ),
    re.compile(
        rf"(?:^|[{_DELIM_CHARS}\[\(])supp(?:[{_DELIM_CHARS}]*(?:information|info|materials?|data|notes?|v\d+|\d+))?(?:$|[{_DELIM_CHARS}\]\)])",
        re.IGNORECASE,
    ),
    re.compile(
        rf"(?<=[{_DELIM_CHARS}\[\(])notes?(?:[{_DELIM_CHARS}]*(?:v\d+|\d+))?(?:$|[\)\]])",
        re.IGNORECASE,
    ),
    re.compile(
        rf"(?:^|[{_DELIM_CHARS}])(?:si|sm)(?:[{_DELIM_CHARS}]*(?:v\d+|\d+))?$",
        re.IGNORECASE,
    ),
    re.compile(
        rf"[\(\[]\s*(?:si|sm)(?:[{_DELIM_CHARS}]*(?:information|info|materials?|data|notes?|v\d+|\d+))?\s*[\)\]]",
        re.IGNORECASE,
    ),
    re.compile(
        rf"(?:^|[{_DELIM_CHARS}\[\(])(?:si|sm)[{_DELIM_CHARS}]+(?:information|info|materials?|data|notes?)(?:[{_DELIM_CHARS}]*(?:v\d+|\d+))?(?:$|[\)\]])",
        re.IGNORECASE,
    ),
]

# Suffix regexes for stripping companion tags in compute_bundle_id:
# Supports si, sm, esm, supp, suppl, supplement, supplementary, supporting, note, notes, peer review,
# and Unicode dashes (\u2013, \u2014).
_COMPANION_TAG_PATTERN = (
    rf"(?:si|sm|esm|supp|suppl|supplement|supplementary|supporting|notes?|peer[{_DELIM_CHARS}]*review(?:[{_DELIM_CHARS}]*comments)?|reviewer[{_DELIM_CHARS}]*comments|editorial[{_DELIM_CHARS}]*decision|review)"
    rf"(?:[{_DELIM_CHARS}]*(?:information|info|materials?|data|notes?))?"
    rf"(?:[{_DELIM_CHARS}]*(?:v\d+|\d+))?"
)

_RE_BRACKETED_COMPANION = re.compile(
    rf"[\s{_DELIM_CHARS}]*[\(\[]{_COMPANION_TAG_PATTERN}[\)\]]\s*$",
    re.IGNORECASE,
)

_RE_TRAILING_COMPANION = re.compile(
    rf"[{_DELIM_CHARS}]+{_COMPANION_TAG_PATTERN}[{_DELIM_CHARS}]*$",
    re.IGNORECASE,
)


def classify_document_role(filename: str, title: str = "") -> str:
    """
    Classifies a file's role based on filename patterns and extension.
    Returns: 'MAIN', 'SI', 'PEER_REVIEW', 'PATENT', 'SLIDES', or 'IMAGE'.
    """
    name_lower = filename.lower()
    ext = os.path.splitext(name_lower)[1]

    if ext in IMAGE_EXTENSIONS:
        return "IMAGE"

    if ext in SLIDES_EXTENSIONS:
        return "SLIDES"

    # Check for patent
    for pat in PATENT_PATTERNS:
        if pat.search(name_lower):
            return "PATENT"

    # Check for peer review
    for pat in PEER_REVIEW_PATTERNS:
        if pat.search(name_lower):
            return "PEER_REVIEW"

    # Check for SI (Supplementary Information)
    stem = os.path.splitext(name_lower)[0]
    for pat in SI_PATTERNS:
        if pat.search(stem):
            return "SI"

    return "MAIN"


def compute_bundle_id(rel_path: str, filename: str) -> str:
    """
    Computes a stable bundle_id that connects companion files (e.g., paper.pdf and paper_si.pdf).
    Files residing in the same relative folder whose filenames share the same base stem
    (after stripping SI/supp tags) share a bundle_id.
    """
    stem = os.path.splitext(filename)[0]

    # Strip bracketed companion tags first: e.g. " (si)", " [supplementary information]"
    clean_stem = _RE_BRACKETED_COMPANION.sub("", stem)

    # Strip trailing delimiter + companion tags: e.g. "_supplementary_information", "-si", "_supp_info"
    clean_stem = _RE_TRAILING_COMPANION.sub("", clean_stem)

    folder_dir = os.path.dirname(rel_path).replace("\\", "/")
    # Normalize clean stem: lowercase and replace non-alphanumeric with hyphen while preserving Unicode letters/digits
    normalized_stem = re.sub(r"[\W_]+", "-", clean_stem.lower(), flags=re.UNICODE).strip("-")

    if not normalized_stem:
        normalized_stem = "doc"

    if folder_dir:
        return f"{folder_dir}/{normalized_stem}"
    return normalized_stem


def bundle_document_record(rel_path: str, filename: str, title: str = "") -> Dict[str, str]:
    """
    Determines bundle classification information for a document.
    """
    role = classify_document_role(filename, title)
    bundle_id = compute_bundle_id(rel_path, filename)
    return {
        "role": role,
        "bundle_id": bundle_id,
    }
