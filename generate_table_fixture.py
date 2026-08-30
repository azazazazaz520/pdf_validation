from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "fixtures" / "synthetic_table_one_page.pdf"


def build_pdf() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    styles = getSampleStyleSheet()
    body = styles["BodyText"]
    body.fontSize = 10
    body.leading = 15
    document = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title="Prism 表格验证样本",
    )
    rows = [
        ["ID", "Product", "Qty", "Amount"],
        ["T-001", "OCR text", "12", "1,234.56"],
        ["T-002", "Table structure", "8", "800.00"],
        ["T-003", "DOCX export", "1", "99.00"],
    ]
    story = [
        Paragraph("表格识别验证样本", styles["Title"]),
        Spacer(1, 6 * mm),
        Paragraph("用于检查单页表格边界、单元格内容和 Word 导出。", body),
        Spacer(1, 6 * mm),
        Table(
            [[Paragraph(str(value), body) for value in row] for row in rows],
            colWidths=[30 * mm, 65 * mm, 25 * mm, 40 * mm],
            repeatRows=1,
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#6B7280")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            ),
        ),
    ]
    document.build(story)
    print(OUTPUT)


if __name__ == "__main__":
    build_pdf()
