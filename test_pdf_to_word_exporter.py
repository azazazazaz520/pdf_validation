from __future__ import annotations

import zipfile
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from reportlab.pdfgen.canvas import Canvas

from pdf_to_word_exporter import export_results_to_docx


class HybridExportTest(unittest.TestCase):
    def test_hybrid_export_keeps_page_image_and_ocr_text(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "input.pdf"
            docx_path = root / "output.docx"

            canvas = Canvas(str(pdf_path), pagesize=(360, 240))
            canvas.drawString(36, 200, "Rendered PDF page")
            canvas.rect(36, 36, 288, 120)
            canvas.save()

            results = [
                {
                    "parsing_res_list": [
                        {
                            "block_label": "title",
                            "block_content": "示例 OCR 文本",
                            "block_bbox": [36, 36, 240, 80],
                            "block_order": 0,
                        }
                    ]
                }
            ]

            export_results_to_docx(
                results,
                docx_path,
                title="hybrid-test",
                source_pdf=pdf_path,
                mode="hybrid",
            )

            document = Document(str(docx_path))
            paragraphs = "\n".join(p.text for p in document.paragraphs)
            self.assertIn("示例 OCR 文本", paragraphs)
            self.assertNotIn("OCR 可编辑文本", paragraphs)
            self.assertNotIn("以下文字由 OCR 识别", paragraphs)
            self.assertGreaterEqual(len(document.sections), 2)
            with zipfile.ZipFile(docx_path) as archive:
                media = [name for name in archive.namelist() if name.startswith("word/media/")]
            self.assertEqual(len(media), 1)


if __name__ == "__main__":
    unittest.main()
