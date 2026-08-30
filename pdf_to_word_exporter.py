from __future__ import annotations

import math
import re
from io import BytesIO
from html.parser import HTMLParser
from pathlib import Path
from collections.abc import Callable, Iterable, Mapping
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from PIL import Image
from pypdf import PdfReader


_ORDERED_ITEM = re.compile(r"^\s*\d+[.．、)]\s*(.+)$")
_BULLET_ITEM = re.compile(r"^\s*[•·●▪◦]\s*(.+)$")
_PLUGIN_ITEM = re.compile(r"^\s*(plugin_[a-zA-Z0-9_]+)\s*[:：]\s*(.+)$")
_HTML_TAG = re.compile(r"<[^>]+>")
EXPORT_MODES = frozenset({"hybrid", "text"})
DEFAULT_PAGE_IMAGE_MAX_PIXELS = 4 * 1024 * 1024
DEFAULT_PAGE_IMAGE_JPEG_QUALITY = 88
MAX_PAGE_DIMENSION_INCHES = 11.0

StageCallback = Callable[[str, dict[str, Any]], None]


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


def _notify_stage(
    callback: StageCallback | None, stage: str, **details: Any
) -> None:
    if callback is not None:
        callback(stage, details)


def _set_section_page(
    section: Any,
    width_inches: float,
    height_inches: float,
    margin_inches: float,
) -> None:
    from docx.shared import Inches

    section.page_width = Inches(width_inches)
    section.page_height = Inches(height_inches)
    section.top_margin = Inches(margin_inches)
    section.bottom_margin = Inches(margin_inches)
    section.left_margin = Inches(margin_inches)
    section.right_margin = Inches(margin_inches)


def _scaled_page_size(
    width_points: float, height_points: float
) -> tuple[float, float]:
    width_inches = max(width_points / 72.0, 0.01)
    height_inches = max(height_points / 72.0, 0.01)
    scale = min(1.0, MAX_PAGE_DIMENSION_INCHES / max(width_inches, height_inches))
    return width_inches * scale, height_inches * scale


def _encode_page_image(
    image: Image.Image,
    *,
    max_pixels: int,
    jpeg_quality: int,
) -> bytes:
    image = image.convert("RGB")
    pixel_count = image.width * image.height
    if pixel_count > max_pixels:
        scale = math.sqrt(max_pixels / pixel_count)
        image = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    output = BytesIO()
    image.save(
        output,
        format="JPEG",
        quality=max(1, min(100, jpeg_quality)),
        optimize=True,
    )
    return output.getvalue()


def _extract_full_page_image(
    pdf_page: Any,
    *,
    page_width: float,
    page_height: float,
    max_pixels: int,
    jpeg_quality: int,
) -> bytes | None:
    try:
        images = list(pdf_page.images)
    except Exception:
        return None
    if len(images) != 1:
        return None

    try:
        with Image.open(BytesIO(images[0].data)) as image:
            image.load()
            page_ratio = page_width / max(page_height, 0.01)
            image_ratio = image.width / max(image.height, 1)
            if abs(image_ratio - page_ratio) / max(page_ratio, 0.01) > 0.03:
                return None
            return _encode_page_image(
                image,
                max_pixels=max_pixels,
                jpeg_quality=jpeg_quality,
            )
    except Exception:
        return None


def _render_page_image(
    pdf_document: Any,
    page_index: int,
    *,
    max_pixels: int,
    jpeg_quality: int,
) -> bytes:
    import pypdfium2 as pdfium

    pdf_page = pdf_document[page_index]
    try:
        width, height = pdf_page.get_size()
        scale = min(2.0, math.sqrt(max_pixels / max(width * height, 1.0)))
        bitmap = pdf_page.render(scale=max(scale, 0.1))
        try:
            image = bitmap.to_pil()
            return _encode_page_image(
                image,
                max_pixels=max_pixels,
                jpeg_quality=jpeg_quality,
            )
        finally:
            close_bitmap = getattr(bitmap, "close", None)
            if close_bitmap is not None:
                close_bitmap()
    finally:
        pdf_page.close()


def _iter_page_images(
    source_pdf: Path,
    *,
    max_pixels: int,
    jpeg_quality: int,
) -> Iterable[tuple[int, bytes, float, float]]:
    import pypdfium2 as pdfium

    reader = PdfReader(str(source_pdf))
    pdf_document = pdfium.PdfDocument(str(source_pdf))
    try:
        for page_index, pdf_page in enumerate(reader.pages):
            page_width = float(pdf_page.mediabox.width)
            page_height = float(pdf_page.mediabox.height)
            image_bytes = _extract_full_page_image(
                pdf_page,
                page_width=page_width,
                page_height=page_height,
                max_pixels=max_pixels,
                jpeg_quality=jpeg_quality,
            )
            if image_bytes is None:
                image_bytes = _render_page_image(
                    pdf_document,
                    page_index,
                    max_pixels=max_pixels,
                    jpeg_quality=jpeg_quality,
                )
            yield page_index, image_bytes, page_width, page_height
    finally:
        pdf_document.close()


def _add_page_image(
    document: Document,
    section: Any,
    image_bytes: bytes,
) -> None:
    paragraph = (
        document.paragraphs[0]
        if len(document.paragraphs) == 1
        else document.add_paragraph()
    )
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.add_run().add_picture(BytesIO(image_bytes), width=section.page_width)


def _add_source_pages(
    document: Document,
    source_pdf: Path,
    *,
    max_pixels: int,
    jpeg_quality: int,
    stage_callback: StageCallback | None,
) -> int:
    page_count = 0
    for page_index, image_bytes, page_width, page_height in _iter_page_images(
        source_pdf,
        max_pixels=max_pixels,
        jpeg_quality=jpeg_quality,
    ):
        if page_index > 0:
            section = document.add_section(WD_SECTION.NEW_PAGE)
        else:
            section = document.sections[0]
        width_inches, height_inches = _scaled_page_size(page_width, page_height)
        _set_section_page(section, width_inches, height_inches, 0)
        _add_page_image(
            document,
            section,
            image_bytes,
        )
        page_count += 1
        _notify_stage(
            stage_callback,
            "page_image_completed",
            page=page_index + 1,
            page_count=page_count,
        )
    return page_count


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


def _add_ocr_pages(document: Document, result_list: list[Any]) -> None:
    for page_index, result in enumerate(result_list):
        if page_index > 0:
            document.add_page_break()
        document.add_heading(f"第 {page_index + 1} 页", level=2)
        for block in _ordered_blocks(result):
            _add_block(document, block)


def _add_text_pages(document: Document, page_texts: Iterable[str]) -> int:
    page_count = 0
    for page_index, text in enumerate(page_texts):
        if page_index > 0:
            document.add_page_break()
        _add_content_lines(document, text, "text")
        page_count += 1
    return page_count


def export_text_pages_to_docx(
    page_texts: Iterable[str],
    output_path: Path,
    title: str,
    *,
    stage_callback: StageCallback | None = None,
) -> None:
    """将 PDF 文本层逐页导出为可编辑 DOCX。"""
    document = Document()
    _set_document_styles(document)
    document.core_properties.title = title
    text_pages = list(page_texts)
    _notify_stage(stage_callback, "text_export_started", page_count=len(text_pages))
    page_count = _add_text_pages(document, text_pages)
    _notify_stage(stage_callback, "text_export_pages_completed", page_count=page_count)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    _notify_stage(stage_callback, "text_export_completed", page_count=page_count)


def export_results_to_docx(
    results: Iterable[Any],
    output_path: Path,
    title: str,
    *,
    source_pdf: Path | None = None,
    mode: str = "text",
    image_max_pixels: int = DEFAULT_PAGE_IMAGE_MAX_PIXELS,
    image_jpeg_quality: int = DEFAULT_PAGE_IMAGE_JPEG_QUALITY,
    stage_callback: StageCallback | None = None,
) -> None:
    """将多页 PaddleOCR 版面结果导出为文字版或混合版 DOCX。"""
    normalized_mode = mode.strip().lower()
    if normalized_mode not in EXPORT_MODES:
        raise ValueError(f"不支持的 DOCX 导出模式：{mode}")
    if image_max_pixels < 1:
        raise ValueError("页面图像像素上限必须大于 0")
    if source_pdf is not None:
        source_pdf = Path(source_pdf)
    document = Document()
    _set_document_styles(document)
    document.core_properties.title = title
    result_list = list(results)

    if normalized_mode == "hybrid":
        if source_pdf is None or not source_pdf.is_file():
            raise FileNotFoundError(f"混合导出所需的 PDF 不存在：{source_pdf}")
        _notify_stage(
            stage_callback,
            "page_images_started",
            page_count=len(result_list),
        )
        image_page_count = _add_source_pages(
            document,
            source_pdf,
            max_pixels=image_max_pixels,
            jpeg_quality=image_jpeg_quality,
            stage_callback=stage_callback,
        )
        _notify_stage(
            stage_callback,
            "page_images_completed",
            page_count=image_page_count,
        )
        ocr_section = document.add_section(WD_SECTION.NEW_PAGE)
        _set_section_page(ocr_section, 8.5, 11.0, 0.75)
        _notify_stage(
            stage_callback,
            "ocr_text_started",
            page_count=len(result_list),
        )
        _add_ocr_pages(document, result_list)
        _notify_stage(
            stage_callback,
            "ocr_text_completed",
            page_count=len(result_list),
        )
    else:
        _add_ocr_pages(document, result_list)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
