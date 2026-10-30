"""
LWS 2.0 Configuration & Path Utilities
Manages multi-root workspaces, Windows MAX_PATH safety, and default references discovery.
"""

from __future__ import annotations

import json
import os
import sys
import urllib.parse
from pathlib import Path
from typing import Any, Dict, List, Optional

CACHE_DIR_NAME = ".ai_cache"
FIGURES_DIR_NAME = "figures"
INDEX_JSON_NAME = "references.json"
CATALOG_MD_NAME = "REFERENCES.md"
SQLITE_DB_NAME = "lws_index.sqlite"

SUPPORTED_DOC_EXTENSIONS = {".pdf", ".docx", ".pptx"}
SUPPORTED_IMG_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
ALL_SUPPORTED_EXTENSIONS = SUPPORTED_DOC_EXTENSIONS | SUPPORTED_IMG_EXTENSIONS


def win_safe_path(p: Path | str) -> Path:
    """Prepends Windows extended-length prefix (\\\\?\\) to bypass MAX_PATH (260 char) limit."""
    p_obj = Path(p).resolve()
    s = str(p_obj)
    if os.name == "nt" and not s.startswith("\\\\?\\") and not s.startswith("//?/"):
        return Path("\\\\?\\" + s)
    return p_obj


def to_file_uri(file_path: Path | str) -> str:
    """Converts a local file path into a standard clickable file:/// URI."""
    posix_path = Path(file_path).resolve().as_posix()
    return "file:///" + urllib.parse.quote(posix_path, safe="/:")


def find_default_references_dir() -> Path:
    """
    Dynamically finds the references directory without hardcoded paths.
    Searches OneDrive, Documents, and user profile paths.
    """
    if "REFERENCES_DIR" in os.environ:
        p = Path(os.environ["REFERENCES_DIR"]).resolve()
        if p.exists():
            return p

    user_profile = Path(os.environ.get("USERPROFILE", Path.home())).resolve()
    candidate_roots = []

    for var in ["OneDriveCommercial", "OneDrive", "OneDriveConsumer"]:
        val = os.environ.get(var)
        if val:
            p = Path(val).resolve()
            if p.exists() and p not in candidate_roots:
                candidate_roots.append(p)

    try:
        for child in user_profile.iterdir():
            if child.is_dir() and "onedrive" in child.name.lower() and child not in candidate_roots:
                candidate_roots.append(child)
    except Exception:
        pass

    candidate_roots.append(user_profile)

    subpaths = [
        Path("Documents") / "SOFTA" / "references",
        Path("SOFTA") / "references",
        Path("Documents") / "references",
        Path("references"),
    ]

    for root in candidate_roots:
        for sub in subpaths:
            candidate = root / sub
            if candidate.exists():
                return candidate.resolve()

    return user_profile / "Documents" / "SOFTA" / "references"


class ConfigManager:
    """Manages persistent LWS workspace configuration (~/.lws_config.json)."""

    def __init__(self, config_path: Optional[Union[Path, str]] = None) -> None:
        if config_path:
            self.config_file = Path(config_path).resolve()
        else:
            self.config_file = Path.home() / ".lws_config.json"
        self.data: Dict[str, Any] = self._load()

    def _load(self) -> Dict[str, Any]:
        default_ref = str(find_default_references_dir())
        default_data = {
            "watched_roots": [default_ref] if os.path.exists(default_ref) else [],
            "active_root": default_ref if os.path.exists(default_ref) else None,
            "theme": "dark",
            "extract_figures": True,
            "ocr_scanned_pages": True,
            "default_context_sentences": 1,
            "scientific_synonyms_enabled": True,
            "wordnet_enabled": True,
        }
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    default_data.update(saved)
            except Exception as e:
                print(f"[Warning] Failed to load {self.config_file}: {e}")
        return default_data

    def save(self) -> None:
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"[Error] Failed to write {self.config_file}: {e}")

    def get_watched_roots(self) -> List[str]:
        return list(self.data.get("watched_roots", []))

    def add_watched_root(self, folder_path: str | Path) -> bool:
        p = str(Path(folder_path).resolve())
        if not os.path.exists(p) or not os.path.isdir(p):
            return False
        roots = self.data.setdefault("watched_roots", [])
        if p not in roots:
            roots.append(p)
            if not self.data.get("active_root"):
                self.data["active_root"] = p
            self.save()
        return True

    def remove_watched_root(self, folder_path: str | Path) -> bool:
        p = str(Path(folder_path).resolve())
        roots = self.data.get("watched_roots", [])
        if p in roots:
            roots.remove(p)
            if self.data.get("active_root") == p:
                self.data["active_root"] = roots[0] if roots else None
            self.save()
            return True
        return False

    def get_active_root(self) -> Optional[str]:
        active = self.data.get("active_root")
        if active and os.path.exists(active):
            return active
        roots = self.get_watched_roots()
        return roots[0] if roots else None

    def set_active_root(self, folder_path: str | Path) -> bool:
        p = str(Path(folder_path).resolve())
        if os.path.exists(p) and os.path.isdir(p):
            self.data["active_root"] = p
            self.save()
            return True
        return False

    def get_watch_folders(self) -> List[str]:
        return self.get_watched_roots()

    def add_watch_folder(self, folder_path: str | Path) -> bool:
        return self.add_watched_root(folder_path)

    def remove_watch_folder(self, folder_path: str | Path) -> bool:
        return self.remove_watched_root(folder_path)

    def get_database_path(self) -> str:
        active = self.get_active_root()
        if active:
            return str(Path(active) / ".ai_cache" / SQLITE_DB_NAME)
        return str(Path.home() / ".lws" / SQLITE_DB_NAME)
