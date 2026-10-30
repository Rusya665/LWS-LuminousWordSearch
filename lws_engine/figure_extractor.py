"""
LWS 2.0 Figure & Image Extractor
Extracts embedded figures and plots from PDFs, DOCX, and PPTX with caption matching.
"""

from __future__ import annotations

import io
import hashlib
import logging
import os
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from PIL import Image

from lws_engine.config import win_safe_path

logger = logging.getLogger(__name__)

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None


CAPTION_PATTERN = re.compile(
    r"(?im)(?:^|(?<=[.!?\r\n]))[ \t]*(?:[•*–—\-]\s+)?\b("
    r"(?:(?:Supplementary|Extended\s+Data|Suppl?\.?)\s+)?"
    r"(?:Figs?|Figures?|Schemes?|Tables?|Charts?)\.?"
    r"\s*"
    r"S?\d+(?:[.\-]\d+)?"
    r"(?:\s*\([a-zA-Z](?:(?:[\-–—]|,\s*)[a-zA-Z])*\)|[-–—]?[a-zA-Z](?:(?:[\-–—]|,\s*)[a-zA-Z])*)?"
    r")"
    r"(?:"
    r"\s*[:|][\-–—]?\s*"
    r"|"
    r"\.\s+"
    r"|"
    r"\s+[\-–—]\s*"
    r")"
    r"([^\n\r]{5,500})"
)

IN_TEXT_PREPOSITIONS = re.compile(
    r"(?i)\b(?:as\s+shown\s+in|shown\s+in|seen\s+in|illustrated\s+in|given\s+in|depicted\s+in|refer\s+to|see|in|from|of|and|with|at)\s*$"
)


def find_captions_in_page_text(page_text: Optional[str]) -> List[Tuple[str, str]]:
    """Extracts figure references and captions from page text, filtering in-text citations."""
    if not page_text or not isinstance(page_text, str):
        return []
    results: List[Tuple[str, str]] = []
    for match in CAPTION_PATTERN.finditer(page_text):
        tag = match.group(1).strip()
        desc = match.group(2).strip()

        start_idx = match.start()
        line_start = page_text.rfind("\n", 0, start_idx)
        preceding = page_text[0 if line_start == -1 else line_start + 1:start_idx].strip()

        if preceding and IN_TEXT_PREPOSITIONS.search(preceding):
            continue

        full_match = match.group(0)
        tag_match = match.group(1)
        delimiter_part = full_match[len(tag_match):-len(match.group(2))].strip()
        if preceding and delimiter_part == ".":
            continue

        results.append((tag, f"{tag}: {desc}"))
    return results


def _compute_rel_image_path(out_path: Path, output_dir: Path) -> str:
    """Computes a normalized relative path for the extracted image."""
    parts = out_path.parts
    if ".ai_cache" in parts:
        idx = parts.index(".ai_cache")
        return "/".join(parts[idx:])
    elif output_dir.name and output_dir.name not in (".", "/"):
        return f"{output_dir.name}/{out_path.name}"
    else:
        return out_path.name


def _prepare_image_for_save(img: Image.Image) -> Image.Image:
    """Ensures image is in a PNG-compatible mode before saving."""
    if img.mode not in ("RGB", "RGBA", "L", "1", "P"):
        return img.convert("RGB")
    return img


def _compute_safe_stem(file_path: Union[str, Path]) -> str:
    """
    Computes a filesystem-safe, collision-resilient stem for document namespacing.
    Replaces non-alphanumeric/non-hyphen characters with underscores.
    Falls back to a deterministic MD5 hash for pure non-ASCII or symbol filenames.
    Truncates exceptionally long stems to prevent MAX_PATH filesystem errors.
    """
    p = Path(file_path)
    safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", p.stem)
    if not safe.strip("_"):
        h = hashlib.md5(p.stem.encode("utf-8", errors="ignore")).hexdigest()[:8]
        return f"doc_{h}" if h else "doc"
    if len(safe) > 80:
        h = hashlib.md5(p.stem.encode("utf-8", errors="ignore")).hexdigest()[:8]
        trimmed = safe[:64].rstrip("_")
        safe = f"{trimmed}_{h}"
    return safe


def _coerce_max_figures(max_figures: Any) -> Optional[int]:
    """
    Defensively coerces max_figures parameter into Optional[int].
    Handles strings ('30', '0', '-5', 'none', 'unlimited'), floats (30.0), and ints.
    Logs a warning and returns None on unparseable inputs.
    """
    if max_figures is None:
        return None
    if isinstance(max_figures, str):
        cleaned = max_figures.strip().lower()
        if cleaned in ("", "none", "null", "unlimited", "all"):
            return None
        try:
            return int(cleaned)
        except ValueError:
            try:
                return int(float(cleaned))
            except (ValueError, OverflowError):
                logger.warning("Invalid max_figures string '%s'; defaulting to None", max_figures)
                return None
    try:
        return int(max_figures)
    except (TypeError, ValueError, OverflowError):
        logger.warning("Unrecognized max_figures value %r; defaulting to None", max_figures)
        return None



def extract_figures_from_pdf(
    file_path: Union[str, Path],
    output_dir: Union[str, Path],
    min_dimension: int = 140,
    max_figures: Optional[int] = 30,
) -> List[Dict[str, Any]]:
    """
    Extracts embedded raster images from PDF pages using pypdfium2.
    Saves images to output_dir and returns metadata records.
    """
    if pdfium is None:
        logger.warning("pypdfium2 is not available; cannot extract figures from PDF: %s", file_path)
        return []

    max_figures = _coerce_max_figures(max_figures)
    if max_figures is not None and max_figures <= 0:
        return []

    try:
        file_path = Path(file_path)
        output_dir = Path(output_dir)
    except Exception as e:
        logger.warning("Invalid file_path or output_dir argument: %s, %s (error: %s)", file_path, output_dir, e)
        return []

    safe_stem = _compute_safe_stem(file_path)

    safe_fp = win_safe_path(file_path)
    if not safe_fp.exists():
        logger.warning("PDF file does not exist for figure extraction: %s", file_path)
        return []

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        logger.warning("Could not create output directory %s: %s", output_dir, e)
        return []

    figures: List[Dict[str, Any]] = []

    try:
        with pdfium.PdfDocument(str(safe_fp)) as doc:
            fig_counter = 0

            for page_idx in range(len(doc)):
                if max_figures is not None and len(figures) >= max_figures:
                    break

                page_num = page_idx + 1
                try:
                    page = doc[page_idx]

                    # Extract captions present on this page
                    try:
                        textpage = page.get_textpage()
                        page_text = textpage.get_text_range() or ""
                    except Exception as tp_err:
                        logger.debug("Failed extracting textpage on page %d of %s: %s", page_num, file_path.name, tp_err)
                        page_text = ""

                    page_captions = find_captions_in_page_text(page_text)

                    # Extract embedded image objects
                    try:
                        image_objs = [obj for obj in page.get_objects() if obj.type == pdfium.raw.FPDF_PAGEOBJ_IMAGE]
                    except Exception as obj_err:
                        logger.debug("Failed extracting image objects on page %d of %s: %s", page_num, file_path.name, obj_err)
                        image_objs = []

                    caption_idx = 0
                    for obj_idx, img_obj in enumerate(image_objs):
                        if max_figures is not None and len(figures) >= max_figures:
                            break

                        try:
                            bitmap = img_obj.get_bitmap()
                            pil_image = bitmap.to_pil()
                            w, h = pil_image.size

                            # Filter out small UI icons, bullets, and thin lines
                            if min_dimension > 0 and (w < min_dimension or h < min_dimension):
                                continue

                            fig_counter += 1
                            fig_id = f"{safe_stem}_p{page_num}_{obj_idx + 1}"
                            out_filename = f"{fig_id}.png"
                            out_path = output_dir / out_filename

                            pil_image = _prepare_image_for_save(pil_image)
                            pil_image.save(win_safe_path(out_path), format="PNG")

                            # Match caption if available on the same page
                            caption = ""
                            if caption_idx < len(page_captions):
                                caption = page_captions[caption_idx][1]
                                caption_idx += 1
                            elif page_captions:
                                caption = page_captions[-1][1]

                            rel_img_path = _compute_rel_image_path(out_path, output_dir)

                            figures.append({
                                "fig_id": fig_id,
                                "figure_index": obj_idx + 1,
                                "fig_index": obj_idx + 1,
                                "page_number": page_num,
                                "page_num": page_num,
                                "filename": out_filename,
                                "rel_image_path": rel_img_path,
                                "file_path": str(out_path.resolve()),
                                "width": w,
                                "height": h,
                                "caption": caption,
                                "ocr_text": "",
                            })
                        except Exception as img_err:
                            logger.debug("Failed processing image obj %d on page %d of %s: %s", obj_idx + 1, page_num, file_path.name, img_err)
                            continue

                except Exception as page_err:
                    logger.debug("Failed processing page %d of %s: %s", page_num, file_path.name, page_err)
                    continue

    except Exception as e:
        logger.warning("Could not open %s for figure extraction: %s", file_path.name, e)
        return []

    return figures


def extract_figures_from_docx(
    file_path: Union[str, Path],
    output_dir: Union[str, Path],
    min_dimension: int = 140,
    max_figures: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Extracts embedded images from a DOCX zip archive."""
    max_figures = _coerce_max_figures(max_figures)
    if max_figures is not None and max_figures <= 0:
        return []

    try:
        file_path = Path(file_path)
        output_dir = Path(output_dir)
    except Exception as e:
        logger.warning("Invalid file_path or output_dir argument: %s, %s (error: %s)", file_path, output_dir, e)
        return []

    safe_stem = _compute_safe_stem(file_path)

    safe_fp = win_safe_path(file_path)
    if not safe_fp.exists():
        logger.warning("DOCX file does not exist for figure extraction: %s", file_path)
        return []

    figures: List[Dict[str, Any]] = []

    try:
        with zipfile.ZipFile(safe_fp) as z:
            media_files = [n for n in z.namelist() if n.startswith("word/media/")]
            if not media_files:
                return []

            output_dir.mkdir(parents=True, exist_ok=True)

            for idx, media_name in enumerate(media_files, 1):
                if max_figures is not None and len(figures) >= max_figures:
                    break

                try:
                    data = z.read(media_name)
                    img = Image.open(io.BytesIO(data))
                    w, h = img.size
                    if min_dimension > 0 and (w < min_dimension or h < min_dimension):
                        continue

                    fig_id = f"{safe_stem}_docx_{idx}"
                    out_filename = f"{fig_id}.png"
                    out_path = output_dir / out_filename
                    img = _prepare_image_for_save(img)
                    img.save(win_safe_path(out_path), format="PNG")

                    rel_img_path = _compute_rel_image_path(out_path, output_dir)

                    figures.append({
                        "fig_id": fig_id,
                        "figure_index": idx,
                        "fig_index": idx,
                        "page_number": 1,
                        "page_num": 1,
                        "filename": out_filename,
                        "rel_image_path": rel_img_path,
                        "file_path": str(out_path.resolve()),
                        "width": w,
                        "height": h,
                        "caption": f"Embedded figure #{idx} from {file_path.name}",
                        "ocr_text": "",
                    })
                except Exception as img_err:
                    logger.debug("Failed extracting docx media %s from %s: %s", media_name, file_path.name, img_err)
                    continue
    except Exception as e:
        logger.warning("Failed extracting docx media from %s: %s", file_path.name, e)

    return figures


def extract_figures_from_pptx(
    file_path: Union[str, Path],
    output_dir: Union[str, Path],
    min_dimension: int = 140,
    max_figures: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Extracts embedded images from a PPTX file."""
    max_figures = _coerce_max_figures(max_figures)
    if max_figures is not None and max_figures <= 0:
        return []

    try:
        file_path = Path(file_path)
        output_dir = Path(output_dir)
    except Exception as e:
        logger.warning("Invalid file_path or output_dir argument: %s, %s (error: %s)", file_path, output_dir, e)
        return []

    safe_stem = _compute_safe_stem(file_path)

    safe_fp = win_safe_path(file_path)
    if not safe_fp.exists():
        logger.warning("PPTX file does not exist for figure extraction: %s", file_path)
        return []

    figures: List[Dict[str, Any]] = []

    try:
        with zipfile.ZipFile(safe_fp) as z:
            media_files = [n for n in z.namelist() if n.startswith("ppt/media/")]
            if not media_files:
                return []

            output_dir.mkdir(parents=True, exist_ok=True)

            for idx, media_name in enumerate(media_files, 1):
                if max_figures is not None and len(figures) >= max_figures:
                    break

                try:
                    data = z.read(media_name)
                    img = Image.open(io.BytesIO(data))
                    w, h = img.size
                    if min_dimension > 0 and (w < min_dimension or h < min_dimension):
                        continue

                    fig_id = f"{safe_stem}_slide_{idx}"
                    out_filename = f"{fig_id}.png"
                    out_path = output_dir / out_filename
                    img = _prepare_image_for_save(img)
                    img.save(win_safe_path(out_path), format="PNG")

                    rel_img_path = _compute_rel_image_path(out_path, output_dir)

                    figures.append({
                        "fig_id": fig_id,
                        "figure_index": idx,
                        "fig_index": idx,
                        "page_number": idx,
                        "page_num": idx,
                        "filename": out_filename,
                        "rel_image_path": rel_img_path,
                        "file_path": str(out_path.resolve()),
                        "width": w,
                        "height": h,
                        "caption": f"Slide figure #{idx} from {file_path.name}",
                        "ocr_text": "",
                    })
                except Exception as img_err:
                    logger.debug("Failed extracting pptx media %s from %s: %s", media_name, file_path.name, img_err)
                    continue
    except Exception as e:
        logger.warning("Failed extracting pptx media from %s: %s", file_path.name, e)

    return figures
