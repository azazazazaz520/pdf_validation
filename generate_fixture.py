from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "fixtures" / "synthetic_text_table.pdf"


def register_font() -> str:
    candidates = [
        Path(r"C:\Windows\Fonts\msyh.ttf"),
        Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
    ]
    for font_path in candidates:
        if not font_path.is_file():
            continue
        try:
            pdfmetrics.registerFont(TTFont("ValidationFont", str(font_path)))
            return "ValidationFont"
        except Exception:
            continue
    return "Helvetica"


def build_pdf() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    font_name = register_font()
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "ValidationTitle",
        parent=styles["Title"],
        fontName=font_name,
        alignment=TA_CENTER,
        fontSize=18,
        leading=24,
    )
    heading = ParagraphStyle(
        "ValidationHeading",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=13,
        leading=18,
    )
    body = ParagraphStyle(
        "ValidationBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=10,
        leading=16,
    )

    document = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Prism PDF 转 Word 可行性验证样本",
        author="Prism validation",
    )
    story = [
        Paragraph("Prism PDF 转 Word 可行性验证样本", title),
        Spacer(1, 8 * mm),
        Paragraph("第一部分：文本、标题和段落顺序", heading),
        Paragraph(
            "这是用于接口冒烟测试的合成 PDF。The quick brown fox jumps over the lazy dog. "
            "数字、日期和标点：2026-08-30、1,234.56、A-102。",
            body,
        ),
        Spacer(1, 5 * mm),
        Paragraph("本页用于检查 PDF 输入、文本识别、标题层级、段落顺序和 DOCX 输出。", body),
        Spacer(1, 8 * mm),
        Paragraph("第二部分：表格", heading),
    ]

    table_data = [
        [Paragraph("编号", body), Paragraph("项目", body), Paragraph("数量", body), Paragraph("备注", body)],
        [Paragraph("A-01", body), Paragraph("文本识别", body), Paragraph("12", body), Paragraph("中文与 English", body)],
        [Paragraph("A-02", body), Paragraph("表格识别", body), Paragraph("8", body), Paragraph("需要检查列顺序", body)],
        [Paragraph("A-03", body), Paragraph("DOCX 导出", body), Paragraph("1", body), Paragraph("检查文件可打开性", body)],
    ]
    table = Table(table_data, colWidths=[24 * mm, 42 * mm, 24 * mm, 72 * mm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#6B7280")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(table)
    story.extend(
        [
            PageBreak(),
            Paragraph("第三部分：第二页连续内容", heading),
            Paragraph(
                "第二页用于检查多页输入和页面重组。此样本不代表真实扫描件质量，不能替代中文扫描 PDF、复杂表格或多栏报告的回归测试。",
                body,
            ),
        ]
    )
    document.build(story)
    print(OUTPUT)


if __name__ == "__main__":
    build_pdf()
