from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pypdfium2 as pdfium

from .pdf_routing import normalize_page_text


_PAGEOBJ_PATH = 2
_MAX_LINE_THICKNESS = 2.0
_MIN_HORIZONTAL_LINE_LENGTH = 20.0
_MIN_VERTICAL_LINE_LENGTH = 10.0
_COORDINATE_TOLERANCE = 2.5
_MAX_TABLE_ROW_GAP = 60.0
_MIN_TABLE_WIDTH = 60.0


@dataclass(frozen=True)
class PdfTextLine:
    """保存可用于阅读顺序和列表识别的 PDF 文本行。"""

    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    font_size: float

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2


@dataclass(frozen=True)
class PdfTable:
    """保存 PDF 几何表格的边界、列边界和单元格文本。"""

    bbox: tuple[float, float, float, float]
    column_boundaries: tuple[float, ...]
    row_boundaries: tuple[float, ...]
    rows: tuple[tuple[str, ...], ...]

    @property
    def column_count(self) -> int:
        return max(len(self.column_boundaries) - 1, 0)

    @property
    def row_count(self) -> int:
        return max(len(self.row_boundaries) - 1, 0)


@dataclass(frozen=True)
class PdfPageLayout:
    """保存单页的尺寸、文本行和几何表格。"""

    width: float
    height: float
    lines: tuple[PdfTextLine, ...]
    tables: tuple[PdfTable, ...]


@dataclass(frozen=True)
class PdfDocumentLayout:
    """保存 PDF 全文的布局中间模型。"""

    pages: tuple[PdfPageLayout, ...]

    @property
    def page_count(self) -> int:
        return len(self.pages)

    @property
    def table_count(self) -> int:
        return sum(len(page.tables) for page in self.pages)


@dataclass(frozen=True)
class _HorizontalLine:
    top: float
    x0: float
    x1: float


@dataclass(frozen=True)
class _VerticalLine:
    x: float
    top: float
    bottom: float


@dataclass
class _TextCharacter:
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    font_size: float

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2


def _compact_text(value: str) -> str:
    return normalize_page_text(value.replace("\r", "\n")).replace("\n", " ")


def _extract_text_characters(
    text_page: Any,
    page_height: float,
) -> list[_TextCharacter]:
    characters: list[_TextCharacter] = []
    text = text_page.get_text_range()
    for index, raw_text in enumerate(text):
        if raw_text in {"", "\r", "\n", "\t", "\ufffe"}:
            continue
        if raw_text == " ":
            text = raw_text
        else:
            text = _compact_text(raw_text)
        if not text:
            continue
        x0, y0, x1, y1 = (float(value) for value in text_page.get_charbox(index))
        height = max(abs(y1 - y0), 1.0)
        characters.append(
            _TextCharacter(
                text=text,
                x0=x0,
                top=page_height - max(y0, y1),
                x1=x1,
                bottom=page_height - min(y0, y1),
                font_size=height,
            )
        )
    return characters


def _same_text_band(left: _TextCharacter, right: _TextCharacter) -> bool:
    overlap = min(left.bottom, right.bottom) - max(left.top, right.top)
    minimum_height = min(left.bottom - left.top, right.bottom - right.top)
    if overlap >= minimum_height * 0.35:
        return True
    return abs(left.center_y - right.center_y) <= max(
        6.0,
        max(left.font_size, right.font_size) * 0.6,
    )


def _split_text_band(characters: list[_TextCharacter]) -> list[PdfTextLine]:
    characters.sort(key=lambda item: item.x0)
    runs: list[list[_TextCharacter]] = []
    current: list[_TextCharacter] = []
    for character in characters:
        obj = character
        if current:
            previous = current[-1]
            gap = character.x0 - previous.x1
            split_gap = max(
                14.0,
                max(previous.font_size, character.font_size) * 1.3,
            )
            if gap > split_gap:
                runs.append(current)
                current = []
        current.append(character)
    if current:
        runs.append(current)

    lines: list[PdfTextLine] = []
    for run in runs:
        text = _compact_text("".join(item.text for item in run))
        if not text:
            continue
        lines.append(
            PdfTextLine(
                text=text,
                x0=min(item.x0 for item in run),
                top=min(item.top for item in run),
                x1=max(item.x1 for item in run),
                bottom=max(item.bottom for item in run),
                font_size=max(item.font_size for item in run),
            )
        )
    return lines


def _extract_text_lines(
    text_page: Any,
    page_height: float,
) -> list[PdfTextLine]:
    characters = _extract_text_characters(text_page, page_height)
    bands: list[list[_TextCharacter]] = []
    for character in sorted(characters, key=lambda item: (item.top, item.x0)):
        matching_band = next(
            (
                band
                for band in reversed(bands)
                if _same_text_band(band[-1], character)
            ),
            None,
        )
        if matching_band is None:
            bands.append([character])
        else:
            matching_band.append(character)

    lines: list[PdfTextLine] = []
    for band in bands:
        lines.extend(_split_text_band(band))
    return sorted(lines, key=lambda line: (line.top, line.x0))


def _merge_intervals(
    intervals: list[tuple[float, float]],
) -> list[tuple[float, float]]:
    if not intervals:
        return []
    merged: list[list[float]] = []
    for start, end in sorted(intervals):
        if not merged or start > merged[-1][1] + _COORDINATE_TOLERANCE:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def _cluster_horizontal_lines(
    segments: list[_HorizontalLine],
) -> list[_HorizontalLine]:
    clusters: list[list[_HorizontalLine]] = []
    for segment in sorted(segments, key=lambda item: item.top):
        if (
            clusters
            and abs(clusters[-1][-1].top - segment.top)
            <= _COORDINATE_TOLERANCE
        ):
            clusters[-1].append(segment)
        else:
            clusters.append([segment])
    return [
        _HorizontalLine(
            top=sum(item.top for item in cluster) / len(cluster),
            x0=min(item.x0 for item in cluster),
            x1=max(item.x1 for item in cluster),
        )
        for cluster in clusters
    ]


def _cluster_vertical_lines(
    segments: list[_VerticalLine],
) -> list[_VerticalLine]:
    clusters: list[list[_VerticalLine]] = []
    for segment in sorted(segments, key=lambda item: item.x):
        if (
            clusters
            and abs(clusters[-1][-1].x - segment.x)
            <= _COORDINATE_TOLERANCE
        ):
            clusters[-1].append(segment)
        else:
            clusters.append([segment])
    merged: list[_VerticalLine] = []
    for cluster in clusters:
        intervals = _merge_intervals(
            [(item.top, item.bottom) for item in cluster]
        )
        for index, (top, bottom) in enumerate(intervals):
            merged.append(
                _VerticalLine(
                    x=sum(item.x for item in cluster) / len(cluster),
                    top=top,
                    bottom=bottom,
                )
            )
    return merged


def _extract_vector_lines(
    page: Any,
    page_height: float,
) -> tuple[list[_HorizontalLine], list[_VerticalLine]]:
    horizontal: list[_HorizontalLine] = []
    vertical: list[_VerticalLine] = []
    for obj in page.get_objects(filter=[_PAGEOBJ_PATH]):
        x0, y0, x1, y1 = (float(value) for value in obj.get_bounds())
        width = abs(x1 - x0)
        height = abs(y1 - y0)
        if height <= _MAX_LINE_THICKNESS and width >= _MIN_HORIZONTAL_LINE_LENGTH:
            horizontal.append(
                _HorizontalLine(
                    top=page_height - (y0 + y1) / 2,
                    x0=min(x0, x1),
                    x1=max(x0, x1),
                )
            )
        elif width <= _MAX_LINE_THICKNESS and height >= _MIN_VERTICAL_LINE_LENGTH:
            vertical.append(
                _VerticalLine(
                    x=(x0 + x1) / 2,
                    top=page_height - max(y0, y1),
                    bottom=page_height - min(y0, y1),
                )
            )
    return _cluster_horizontal_lines(horizontal), _cluster_vertical_lines(vertical)


def _vertical_coverage(
    lines: list[_VerticalLine],
    *,
    top: float,
    bottom: float,
) -> float:
    if bottom <= top:
        return 0.0
    intervals = _merge_intervals(
        [
            (max(line.top, top), min(line.bottom, bottom))
            for line in lines
            if line.bottom > top and line.top < bottom
        ]
    )
    covered = sum(max(end - start, 0.0) for start, end in intervals)
    return covered / (bottom - top)


def _table_boundary_lines(
    vertical: list[_VerticalLine],
    *,
    x0: float,
    x1: float,
    top: float,
    bottom: float,
) -> list[_VerticalLine]:
    return [
        line
        for line in vertical
        if x0 - _COORDINATE_TOLERANCE <= line.x <= x1 + _COORDINATE_TOLERANCE
        and _vertical_coverage([line], top=top, bottom=bottom) >= 0.9
    ]


def _horizontal_run_tables(
    horizontal: list[_HorizontalLine],
    vertical: list[_VerticalLine],
) -> list[tuple[float, float, float, float, tuple[float, ...], tuple[float, ...]]]:
    candidates: list[
        tuple[float, float, float, float, tuple[float, ...], tuple[float, ...]]
    ] = []
    index = 0
    while index < len(horizontal):
        run = [horizontal[index]]
        index += 1
        while index < len(horizontal):
            previous = run[-1]
            current = horizontal[index]
            shared_width = min(previous.x1, current.x1) - max(previous.x0, current.x0)
            if (
                current.top - previous.top > _MAX_TABLE_ROW_GAP
                or shared_width < _MIN_TABLE_WIDTH
            ):
                break
            prospective_run = [*run, current]
            prospective_x0 = min(item.x0 for item in prospective_run)
            prospective_x1 = max(item.x1 for item in prospective_run)
            if len(run) >= 1 and len(
                _table_boundary_lines(
                    vertical,
                    x0=prospective_x0,
                    x1=prospective_x1,
                    top=prospective_run[0].top,
                    bottom=prospective_run[-1].top,
                )
            ) < 2:
                break
            run.append(current)
            index += 1
        if len(run) < 3:
            continue

        table_top = run[0].top
        table_bottom = run[-1].top
        x0 = min(item.x0 for item in run)
        x1 = max(item.x1 for item in run)
        if x1 - x0 < _MIN_TABLE_WIDTH:
            continue

        boundary_lines = _table_boundary_lines(
            vertical,
            x0=x0,
            x1=x1,
            top=table_top,
            bottom=table_bottom,
        )
        boundary_lines.sort(key=lambda line: line.x)
        column_boundaries: list[float] = []
        for line in boundary_lines:
            if (
                not column_boundaries
                or abs(column_boundaries[-1] - line.x) > _COORDINATE_TOLERANCE
            ):
                column_boundaries.append(line.x)
        if len(column_boundaries) < 2:
            continue
        if column_boundaries[0] > x0 + 4.0:
            x0 = column_boundaries[0]
        if column_boundaries[-1] < x1 - 4.0:
            x1 = column_boundaries[-1]
        row_boundaries = tuple(item.top for item in run)
        candidates.append(
            (
                x0,
                table_top,
                x1,
                table_bottom,
                tuple(column_boundaries),
                row_boundaries,
            )
        )
    return candidates


def _line_is_in_table(line: PdfTextLine, bbox: tuple[float, float, float, float]) -> bool:
    x0, top, x1, bottom = bbox
    return (
        x0 - 1.0 <= line.center_x <= x1 + 1.0
        and top - 1.0 <= line.center_y <= bottom + 1.0
    )


def _extract_table(
    lines: list[PdfTextLine],
    candidate: tuple[
        float,
        float,
        float,
        float,
        tuple[float, ...],
        tuple[float, ...],
    ],
) -> PdfTable:
    x0, top, x1, bottom, column_boundaries, row_boundaries = candidate
    cell_lines: dict[tuple[int, int], list[PdfTextLine]] = {}
    bbox = (x0, top, x1, bottom)
    for line in lines:
        if not _line_is_in_table(line, bbox):
            continue
        row_index = next(
            (
                index
                for index in range(len(row_boundaries) - 1)
                if row_boundaries[index] <= line.center_y <= row_boundaries[index + 1]
            ),
            None,
        )
        column_index = next(
            (
                index
                for index in range(len(column_boundaries) - 1)
                if column_boundaries[index] <= line.center_x <= column_boundaries[index + 1]
            ),
            None,
        )
        if row_index is None or column_index is None:
            continue
        cell_lines.setdefault((row_index, column_index), []).append(line)

    rows: list[tuple[str, ...]] = []
    for row_index in range(len(row_boundaries) - 1):
        row: list[str] = []
        for column_index in range(len(column_boundaries) - 1):
            values = sorted(
                cell_lines.get((row_index, column_index), []),
                key=lambda line: (line.top, line.x0),
            )
            row.append("\n".join(line.text for line in values).strip())
        rows.append(tuple(row))

    return PdfTable(
        bbox=bbox,
        column_boundaries=column_boundaries,
        row_boundaries=row_boundaries,
        rows=tuple(rows),
    )


def _find_tables(
    page: Any,
    lines: list[PdfTextLine],
    page_height: float,
) -> tuple[PdfTable, ...]:
    horizontal, vertical = _extract_vector_lines(page, page_height)
    candidates = _horizontal_run_tables(horizontal, vertical)
    tables = [
        _extract_table(lines, candidate)
        for candidate in candidates
    ]
    return tuple(
        table
        for table in tables
        if any(value for row in table.rows for value in row)
    )


def extract_pdf_layout(source_pdf: Path) -> PdfDocumentLayout:
    """提取 PDF 文本行、坐标和几何表格，形成可编辑导出的布局模型。"""
    if not source_pdf.is_file():
        raise FileNotFoundError(f"布局解析所需的 PDF 不存在：{source_pdf}")

    document = pdfium.PdfDocument(str(source_pdf))
    pages: list[PdfPageLayout] = []
    try:
        for page in document:
            width, height = (float(value) for value in page.get_size())
            text_page = page.get_textpage()
            try:
                lines = _extract_text_lines(text_page, height)
                tables = _find_tables(page, lines, height)
            finally:
                text_page.close()
                page.close()
            pages.append(
                PdfPageLayout(
                    width=width,
                    height=height,
                    lines=tuple(lines),
                    tables=tables,
                )
            )
    finally:
        document.close()
    return PdfDocumentLayout(pages=tuple(pages))
