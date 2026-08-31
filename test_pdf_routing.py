from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas

from pdf_routing import analyze_pdf_text, normalize_page_text


def _create_text_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "Fast route page one with selectable text")
    canvas.showPage()
    canvas.drawString(36, 200, "Fast route page two with selectable text")
    canvas.save()


def _create_image_backed_text_pdf(root: Path, path: Path) -> None:
    background_path = root / "background.png"
    Image.new("RGB", (720, 480), color=(220, 230, 240)).save(background_path)
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawImage(ImageReader(str(background_path)), 0, 0, width=360, height=240)
    canvas.drawString(36, 200, "Only a partial selectable text layer")
    canvas.save()


class PdfRoutingTest(unittest.TestCase):
    def test_special_font_mapping_is_normalized_to_editable_text(self) -> None:
        value = "𝑥\u0b35 + 𝑦\u0b36 + 𝛼 + ቐ + \uf0b7 项目"

        self.assertEqual(normalize_page_text(value), "x₁ + y₂ + α + ⎧ + • 项目")

    def test_analyze_pdf_text_detects_text_layer(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "text.pdf"
            _create_text_pdf(pdf_path)

            analysis = analyze_pdf_text(pdf_path, min_page_chars=20)

            self.assertEqual(analysis.page_count, 2)
            self.assertEqual(analysis.usable_page_count, 2)
            self.assertEqual(analysis.usable_page_ratio, 1.0)
            self.assertIn("Fast route page one", analysis.page_texts[0])
            self.assertTrue(
                analysis.has_usable_text_layer(
                    min_page_chars=20,
                    min_page_ratio=0.6,
                )
            )
            self.assertEqual(analysis.high_quality_page_count, 2)
            self.assertTrue(
                analysis.has_complete_text_layer(
                    min_page_chars=20,
                    min_high_quality_ratio=0.8,
                )
            )

    def test_image_backed_text_layer_is_not_complete(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "image-backed.pdf"
            _create_image_backed_text_pdf(root, pdf_path)

            analysis = analyze_pdf_text(pdf_path, min_page_chars=20)

            self.assertEqual(analysis.page_count, 1)
            self.assertEqual(analysis.usable_page_count, 1)
            self.assertEqual(analysis.full_page_image_page_count, 1)
            self.assertEqual(analysis.high_quality_page_count, 0)
            self.assertTrue(
                analysis.has_usable_text_layer(
                    min_page_chars=20,
                    min_page_ratio=0.6,
                )
            )
            self.assertFalse(
                analysis.has_complete_text_layer(
                    min_page_chars=20,
                    min_high_quality_ratio=0.8,
                )
            )

if __name__ == "__main__":
    unittest.main()
