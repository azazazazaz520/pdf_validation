from __future__ import annotations

import unicodedata
from typing import Any

from pdf_routing import count_text_characters, normalize_page_text


DEFAULT_LOW_CONFIDENCE_THRESHOLD = 0.75
DEFAULT_LOW_CONFIDENCE_RATIO = 0.25
DEFAULT_GARBLED_CHAR_RATIO = 0.05


def _field(value: Any, name: str, default: Any = None) -> Any:
    if isinstance(value, dict):
        return value.get(name, default)
    return getattr(value, name, default)


def _as_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    try:
        return list(value)
    except TypeError:
        return []


def _ocr_result(result: Any) -> Any:
    overall = _field(result, "overall_ocr_res")
    return overall if overall is not None else result


def _is_suspicious_character(value: str) -> bool:
    if value in {"\ufffd", "\u0000"}:
        return True
    return unicodedata.category(value) == "Co"


def summarize_ocr_page(
    result: Any,
    *,
    page_number: int,
    low_confidence_threshold: float = DEFAULT_LOW_CONFIDENCE_THRESHOLD,
    low_confidence_ratio_threshold: float = DEFAULT_LOW_CONFIDENCE_RATIO,
    garbled_char_ratio_threshold: float = DEFAULT_GARBLED_CHAR_RATIO,
) -> dict[str, Any]:
    """提取单页 OCR 结果的置信度和异常信号。"""
    ocr_result = _ocr_result(result)
    texts = [
        normalize_page_text(str(text or ""))
        for text in _as_list(_field(ocr_result, "rec_texts", []))
    ]
    scores: list[float] = []
    for score in _as_list(_field(ocr_result, "rec_scores", [])):
        try:
            scores.append(float(score))
        except (TypeError, ValueError):
            continue

    text_char_count = sum(count_text_characters(text) for text in texts)
    suspicious_char_count = sum(
        _is_suspicious_character(character)
        for text in texts
        for character in text
    )
    low_confidence_count = sum(score < low_confidence_threshold for score in scores)
    low_confidence_ratio = low_confidence_count / len(scores) if scores else 1.0
    mean_confidence = sum(scores) / len(scores) if scores else None
    min_confidence = min(scores) if scores else None
    garbled_char_ratio = suspicious_char_count / max(text_char_count, 1)
    needs_review = (
        not texts
        or not any(text.strip() for text in texts)
        or not scores
        or low_confidence_ratio > low_confidence_ratio_threshold
        or (mean_confidence is not None and mean_confidence < low_confidence_threshold)
        or garbled_char_ratio > garbled_char_ratio_threshold
    )

    return {
        "page_number": page_number,
        "ocr_line_count": len(texts),
        "ocr_nonempty_line_count": sum(bool(text.strip()) for text in texts),
        "ocr_text_char_count": text_char_count,
        "ocr_mean_confidence": (
            round(mean_confidence, 4) if mean_confidence is not None else None
        ),
        "ocr_min_confidence": (
            round(min_confidence, 4) if min_confidence is not None else None
        ),
        "ocr_low_confidence_count": low_confidence_count,
        "ocr_low_confidence_ratio": round(low_confidence_ratio, 4),
        "ocr_garbled_char_count": suspicious_char_count,
        "ocr_garbled_char_ratio": round(garbled_char_ratio, 4),
        "ocr_needs_review": needs_review,
    }
