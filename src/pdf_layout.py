from __future__ import annotations

from dataclasses import dataclass, replace
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
_MIN_CELL_LINE_LENGTH = 8.0
_MIN_BOUNDARY_COVERAGE = 0.6


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
class PdfTableCell:
    """保存表格单元格的网格位置、跨行跨列信息和文本。"""

    row_index: int
    column_index: int
    row_span: int
    column_span: int
    bbox: tuple[float, float, float, float]
    text: str


@dataclass(frozen=True)
class PdfTable:
    """保存 PDF 几何表格的边界、列边界、行边界和单元格。"""

    bbox: tuple[float, float, float, float]
    column_boundaries: tuple[float, ...]
    row_boundaries: tuple[float, ...]
    rows: tuple[tuple[str, ...], ...]
    cells: tuple[PdfTableCell, ...] = ()
    header_row_count: int = 1
    continued_from_previous_page: bool = False
    continuation_header_rows: tuple[tuple[str, ...], ...] = ()

    @property
    def column_count(self) -> int:
        return max(len(self.column_boundaries) - 1, 0)

    @property
    def row_count(self) -> int:
        return max(len(self.row_boundaries) - 1, 0)

    @property
    def row_heights(self) -> tuple[float, ...]:
        """返回每一行在 PDF 中的高度，单位为 point。"""
        return tuple(
            max(self.row_boundaries[index + 1] - self.row_boundaries[index], 0.0)
            for index in range(self.row_count)
        )


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
    segments: tuple[tuple[float, float], ...] = ()


@dataclass(frozen=True)
class _VerticalLine:
    x: float
    top: float
    bottom: float


@dataclass(frozen=True)
class _TableCandidate:
    """保存表格候选区域及其可用于识别合并单元格的线段。"""

    bbox: tuple[float, float, float, float]
    column_boundaries: tuple[float, ...]
    row_boundaries: tuple[float, ...]
    horizontal_lines: tuple[_HorizontalLine, ...]
    vertical_lines: tuple[_VerticalLine, ...]


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
            segments=tuple((item.x0, item.x1) for item in cluster),
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


def _table_column_boundaries(
    vertical: list[_VerticalLine],
    *,
    x0: float,
    x1: float,
    top: float,
    bottom: float,
) -> tuple[float, ...]:
    """从候选区域内的垂直线段提取列边界，保留局部边界以支持合并单元格。"""
    positions = [x0, x1]
    for line in vertical:
        if line.x < x0 - _COORDINATE_TOLERANCE or line.x > x1 + _COORDINATE_TOLERANCE:
            continue
        overlap = min(line.bottom, bottom) - max(line.top, top)
        if overlap >= _MIN_CELL_LINE_LENGTH:
            positions.append(line.x)

    clusters: list[list[float]] = []
    for position in sorted(positions):
        if (
            clusters
            and position - clusters[-1][-1] <= _COORDINATE_TOLERANCE
        ):
            clusters[-1].append(position)
        else:
            clusters.append([position])
    return tuple(sum(cluster) / len(cluster) for cluster in clusters)


def _horizontal_coverage(
    lines: list[_HorizontalLine],
    *,
    top: float,
    x0: float,
    x1: float,
) -> float:
    if x1 <= x0:
        return 0.0
    intervals: list[tuple[float, float]] = []
    for line in lines:
        if abs(line.top - top) > _COORDINATE_TOLERANCE:
            continue
        segments = line.segments or ((line.x0, line.x1),)
        intervals.extend(
            (
                max(segment_x0, x0),
                min(segment_x1, x1),
            )
            for segment_x0, segment_x1 in segments
            if segment_x1 > x0 and segment_x0 < x1
        )
    merged = _merge_intervals(intervals)
    covered = sum(max(end - start, 0.0) for start, end in merged)
    return covered / (x1 - x0)


def _has_vertical_boundary(
    lines: list[_VerticalLine],
    *,
    x: float,
    top: float,
    bottom: float,
) -> bool:
    return _vertical_coverage(
        [line for line in lines if abs(line.x - x) <= _COORDINATE_TOLERANCE],
        top=top,
        bottom=bottom,
    ) >= _MIN_BOUNDARY_COVERAGE


def _has_horizontal_boundary(
    lines: list[_HorizontalLine],
    *,
    top: float,
    x0: float,
    x1: float,
) -> bool:
    return (
        _horizontal_coverage(lines, top=top, x0=x0, x1=x1)
        >= _MIN_BOUNDARY_COVERAGE
    )


def _extract_cell_spans(candidate: _TableCandidate) -> list[tuple[int, int, int, int]]:
    """根据局部边界缺失情况推断表格中的跨行和跨列单元格。"""
    column_boundaries = candidate.column_boundaries
    row_boundaries = candidate.row_boundaries
    row_count = max(len(row_boundaries) - 1, 0)
    column_count = max(len(column_boundaries) - 1, 0)
    occupied: set[tuple[int, int]] = set()
    spans: list[tuple[int, int, int, int]] = []

    for row_index in range(row_count):
        for column_index in range(column_count):
            if (row_index, column_index) in occupied:
                continue
            column_span = 1
            while column_index + column_span < column_count:
                divider_x = column_boundaries[column_index + column_span]
                if _has_vertical_boundary(
                    list(candidate.vertical_lines),
                    x=divider_x,
                    top=row_boundaries[row_index],
                    bottom=row_boundaries[row_index + 1],
                ):
                    break
                column_span += 1

            row_span = 1
            while row_index + row_span < row_count:
                next_row = row_index + row_span
                if any(
                    (next_row, column) in occupied
                    for column in range(column_index, column_index + column_span)
                ):
                    break
                if _has_horizontal_boundary(
                    list(candidate.horizontal_lines),
                    top=row_boundaries[next_row],
                    x0=column_boundaries[column_index],
                    x1=column_boundaries[column_index + column_span],
                ):
                    break
                row_span += 1

            span = (row_index, column_index, row_span, column_span)
            spans.append(span)
            for row in range(row_index, row_index + row_span):
                for column in range(column_index, column_index + column_span):
                    occupied.add((row, column))
    return spans


def _horizontal_run_tables(
    horizontal: list[_HorizontalLine],
    vertical: list[_VerticalLine],
) -> list[_TableCandidate]:
    candidates: list[_TableCandidate] = []
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

        column_boundaries = _table_column_boundaries(
            vertical,
            x0=x0,
            x1=x1,
            top=table_top,
            bottom=table_bottom,
        )
        if len(column_boundaries) < 2:
            continue
        x0 = column_boundaries[0]
        x1 = column_boundaries[-1]
        row_boundaries = tuple(item.top for item in run)
        candidates.append(
            _TableCandidate(
                bbox=(x0, table_top, x1, table_bottom),
                column_boundaries=column_boundaries,
                row_boundaries=row_boundaries,
                horizontal_lines=tuple(run),
                vertical_lines=tuple(vertical),
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
    candidate: _TableCandidate,
) -> PdfTable:
    x0, top, x1, bottom = candidate.bbox
    column_boundaries = candidate.column_boundaries
    row_boundaries = candidate.row_boundaries
    cell_spans = _extract_cell_spans(candidate)
    cell_lines: dict[tuple[int, int], list[PdfTextLine]] = {
        (row_index, column_index): []
        for row_index, column_index, _, _ in cell_spans
    }
    for line in lines:
        if not _line_is_in_table(line, candidate.bbox):
            continue
        matching_cells: list[tuple[float, tuple[int, int]]] = []
        for row_index, column_index, row_span, column_span in cell_spans:
            cell_top = row_boundaries[row_index]
            cell_bottom = row_boundaries[row_index + row_span]
            cell_x0 = column_boundaries[column_index]
            cell_x1 = column_boundaries[column_index + column_span]
            overlap_x = min(line.x1, cell_x1) - max(line.x0, cell_x0)
            if overlap_x <= 0 and not cell_x0 <= line.center_x <= cell_x1:
                continue
            if not cell_top <= line.center_y <= cell_bottom:
                continue
            overlap_y = min(line.bottom, cell_bottom) - max(line.top, cell_top)
            score = max(overlap_x, 0.1) * max(overlap_y, 0.1)
            matching_cells.append((score, (row_index, column_index)))
        if not matching_cells:
            continue
        _, cell_key = max(matching_cells, key=lambda item: item[0])
        cell_lines[cell_key].append(line)

    rows = [
        ["" for _ in range(max(len(column_boundaries) - 1, 0))]
        for _ in range(max(len(row_boundaries) - 1, 0))
    ]
    cells: list[PdfTableCell] = []
    for row_index, column_index, row_span, column_span in cell_spans:
        values = sorted(
            cell_lines[(row_index, column_index)],
            key=lambda line: (line.top, line.x0),
        )
        text = "\n".join(line.text for line in values).strip()
        rows[row_index][column_index] = text
        cells.append(
            PdfTableCell(
                row_index=row_index,
                column_index=column_index,
                row_span=row_span,
                column_span=column_span,
                bbox=(
                    column_boundaries[column_index],
                    row_boundaries[row_index],
                    column_boundaries[column_index + column_span],
                    row_boundaries[row_index + row_span],
                ),
                text=text,
            )
        )

    return PdfTable(
        bbox=candidate.bbox,
        column_boundaries=column_boundaries,
        row_boundaries=row_boundaries,
        rows=tuple(tuple(row) for row in rows),
        cells=tuple(cells),
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


def _column_width_ratios(table: PdfTable) -> tuple[float, ...]:
    widths = [
        table.column_boundaries[index + 1] - table.column_boundaries[index]
        for index in range(table.column_count)
    ]
    total = sum(widths)
    if total <= 0:
        return ()
    return tuple(width / total for width in widths)


def _tables_can_continue(
    previous_page: PdfPageLayout,
    previous_table: PdfTable,
    current_page: PdfPageLayout,
    current_table: PdfTable,
) -> bool:
    """判断相邻页面的表格是否具有续表的几何特征。"""
    if previous_table.column_count != current_table.column_count:
        return False
    if not previous_table.rows or not current_table.rows:
        return False
    if previous_table.bbox[3] < previous_page.height * 0.7:
        return False
    if current_table.bbox[1] > current_page.height * 0.3:
        return False
    previous_x0, _, previous_x1, _ = previous_table.bbox
    current_x0, _, current_x1, _ = current_table.bbox
    previous_width = max(previous_x1 - previous_x0, 1.0)
    current_width = max(current_x1 - current_x0, 1.0)
    if abs(previous_x0 / previous_page.width - current_x0 / current_page.width) > 0.04:
        return False
    if abs(previous_x1 / previous_page.width - current_x1 / current_page.width) > 0.04:
        return False
    if abs(previous_width / previous_page.width - current_width / current_page.width) > 0.04:
        return False
    previous_ratios = _column_width_ratios(previous_table)
    current_ratios = _column_width_ratios(current_table)
    return bool(
        previous_ratios
        and len(previous_ratios) == len(current_ratios)
        and max(
            abs(previous - current)
            for previous, current in zip(previous_ratios, current_ratios)
        )
        <= 0.06
    )


def _mark_table_continuations(
    pages: tuple[PdfPageLayout, ...],
) -> tuple[PdfPageLayout, ...]:
    """为相邻页面的续表补充表头元数据。"""
    updated_pages = list(pages)
    for page_index in range(1, len(updated_pages)):
        previous_page = updated_pages[page_index - 1]
        current_page = updated_pages[page_index]
        if not previous_page.tables or not current_page.tables:
            continue
        previous_table = previous_page.tables[-1]
        current_table = current_page.tables[0]
        if not _tables_can_continue(
            previous_page,
            previous_table,
            current_page,
            current_table,
        ):
            continue
        header = (previous_table.rows[0],)
        current_header_is_present = current_table.rows[0] == previous_table.rows[0]
        updated_table = replace(
            current_table,
            continued_from_previous_page=True,
            header_row_count=1 if current_header_is_present else 0,
            continuation_header_rows=() if current_header_is_present else header,
        )
        updated_pages[page_index] = replace(
            current_page,
            tables=(updated_table, *current_page.tables[1:]),
        )
    return tuple(updated_pages)


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
    return PdfDocumentLayout(pages=_mark_table_continuations(tuple(pages)))
