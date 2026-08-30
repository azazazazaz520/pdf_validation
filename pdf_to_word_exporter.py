from __future__ import annotations

import re
from html.parser import HTMLParser
from pathlib import Path
from collections.abc import Mapping
from typing import Any, Iterable

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt


_ORDERED_ITEM = re.compile(r"^\s*\d+[.．、)]\s*(.+)$")
_BULLET_ITEM = re.compile(r"^\s*[•·●▪◦]\s*(.+)$")
_PLUGIN_ITEM = re.compile(r"^\s*(plugin_[a-zA-Z0-9_]+)\s*[:：]\s*(.+)$")
_HTML_TAG = re.compile(r"<[^>]+>")


class _TableParser(HTMLParser):
    """读取 PaddleOCR 表格 HTML 中的行和单元格文本。"""

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._cell = []

    def handle_endtag(self, tag: str) -> None:
        if tag in ("td", "th") and self._row is not None and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)


def _plain_text(value: str) -> str:
    value = _HTML_TAG.sub("", value)
    return re.sub(r"[ \t]+", " ", value).strip()


def _table_rows(value: str) -> list[list[str]]:
    parser = _TableParser()
    parser.feed(value)
    return parser.rows


def _field(value: Any, name: str, default: Any = None) -> Any:
    if isinstance(value, Mapping):
        return value.get(name, default)
    attribute = getattr(value, name, None)
    if attribute is not None:
        return attribute
    aliases = {
        "block_label": "label",
        "block_content": "content",
        "block_bbox": "bbox",
        "block_order": "order_index",
    }
    return getattr(value, aliases.get(name, name), default)


def _ordered_blocks(result: Any) -> Iterable[Any]:
    blocks = list(_field(result, "parsing_res_list", []) or [])
    indexed = list(enumerate(blocks))

    def sort_key(item: tuple[int, Any]) -> tuple[int, int]:
        index, block = item
        order = _field(block, "block_order")
        if isinstance(order, (int, float)):
            return (0, int(order))
        bbox = _field(block, "block_bbox") or [0, index]
        return (1, int(bbox[1]) if len(bbox) > 1 else index)

    for _, block in sorted(indexed, key=sort_key):
        yield block


def _set_document_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(10.5)
    for style_name in ("Heading 1", "Heading 2", "Heading 3"):
        style = document.styles[style_name]
        style.font.name = "Microsoft YaHei"
        style.font.bold = True


def _add_table(document: Document, html: str) -> None:
    rows = _table_rows(html)
    if not rows:
        return
    column_count = max(len(row) for row in rows)
    table = document.add_table(rows=0, cols=column_count)
    table.style = "Table Grid"
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            if index < len(cells):
                cells[index].text = value


def _add_content_lines(document: Document, content: str, block_label: str) -> None:
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    if not lines:
        return

    if block_label == "abstract" and len(lines[0]) <= 24 and not re.search(r"[：:。；;]", lines[0]):
        document.add_heading(_plain_text(lines.pop(0)), level=1)

    previous_list_paragraph = None
    for line in lines:
        ordered_match = _ORDERED_ITEM.match(line)
        bullet_match = _BULLET_ITEM.match(line)
        plugin_match = _PLUGIN_ITEM.match(line)
        if ordered_match:
            paragraph = document.add_paragraph(style="List Number")
            paragraph.add_run(_plain_text(ordered_match.group(1)))
            previous_list_paragraph = paragraph
        elif bullet_match:
            paragraph = document.add_paragraph(style="List Bullet")
            paragraph.add_run(_plain_text(bullet_match.group(1)))
            previous_list_paragraph = paragraph
        elif plugin_match:
            paragraph = document.add_paragraph(style="List Bullet")
            paragraph.add_run(
                f"{plugin_match.group(1)}：{_plain_text(plugin_match.group(2))}"
            )
            previous_list_paragraph = paragraph
        elif previous_list_paragraph is not None and len(line) < 120:
            previous_list_paragraph.add_run(f" {_plain_text(line)}")
        else:
            document.add_paragraph(_plain_text(line))
            previous_list_paragraph = None


def _add_block(document: Document, block: Any) -> None:
    label = str(_field(block, "block_label") or "text")
    content = str(_field(block, "block_content") or "").strip()
    if not content:
        return
    if label == "table" or "<table" in content.lower():
        _add_table(document, content)
    elif label in {"doc_title", "paragraph_title", "title", "figure_title"}:
        document.add_heading(_plain_text(content), level=1)
    else:
        _add_content_lines(document, content, label)


def export_results_to_docx(
    results: Iterable[Any], output_path: Path, title: str
) -> None:
    """将多页 PaddleOCR 版面结果汇总为一个可编辑 DOCX。"""
    document = Document()
    _set_document_styles(document)
    document.core_properties.title = title
    result_list = list(results)
    for page_index, result in enumerate(result_list):
        if page_index > 0:
            document.add_page_break()
        for block in _ordered_blocks(result):
            _add_block(document, block)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
