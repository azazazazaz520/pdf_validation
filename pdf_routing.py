from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pypdf import PdfReader


@dataclass(frozen=True)
class PdfTextAnalysis:
    """保存 PDF 文本层检测结果和逐页提取文本。"""

    page_texts: list[str]
    usable_page_count: int
    text_char_count: int

    @property
    def page_count(self) -> int:
        return len(self.page_texts)

    @property
    def usable_page_ratio(self) -> float:
        return self.usable_page_count / max(self.page_count, 1)

    def has_usable_text_layer(
        self,
        *,
        min_page_chars: int,
        min_page_ratio: float,
    ) -> bool:
        return (
            self.usable_page_count > 0
            and self.text_char_count >= min_page_chars
            and self.usable_page_ratio >= min_page_ratio
        )


def normalize_page_text(value: str | None) -> str:
    """清理文本层中的空字符和行内多余空格，同时保留换行。"""
    if not value:
        return ""
    lines = []
    for line in value.replace("\x00", "").splitlines():
        normalized = re.sub(r"[ \t]+", " ", line).strip()
        if normalized:
            lines.append(normalized)
    return "\n".join(lines)


def count_text_characters(value: str) -> int:
    """统计去除空白后的文本字符数量。"""
    return len(re.sub(r"\s+", "", value))


def analyze_pdf_text(path: Path, *, min_page_chars: int) -> PdfTextAnalysis:
    """提取 PDF 各页文本并统计可用于快速路线判断的页面数量。"""
    if min_page_chars < 1:
        raise ValueError("文本层最小字符数必须大于 0")

    reader = PdfReader(str(path))
    page_texts: list[str] = []
    usable_page_count = 0
    text_char_count = 0
    for page in reader.pages:
        text = normalize_page_text(page.extract_text())
        page_texts.append(text)
        page_chars = count_text_characters(text)
        text_char_count += page_chars
        if page_chars >= min_page_chars:
            usable_page_count += 1

    return PdfTextAnalysis(
        page_texts=page_texts,
        usable_page_count=usable_page_count,
        text_char_count=text_char_count,
    )
