from __future__ import annotations

import threading
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from docx import Document
from reportlab.pdfgen.canvas import Canvas

from pdf_routing import analyze_pdf_text
from pdf_to_word_service import CONFIG, Job, JobManager


def _create_text_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "Fast route page one with selectable text")
    canvas.showPage()
    canvas.drawString(36, 200, "Fast route page two with selectable text")
    canvas.save()


class PdfRoutingTest(unittest.TestCase):
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

    def test_service_uses_text_route_without_loading_model(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_path = root / "input.pdf"
            output_path = root / "result.docx"
            _create_text_pdf(input_path)
            job = Job(
                job_id="text-route-job",
                filename="input.pdf",
                workspace=root,
                input_path=input_path,
                output_path=output_path,
                page_count=2,
            )
            manager = JobManager()
            route_mode_patch = patch.object(CONFIG, "route_mode", "auto")
            export_mode_patch = patch.object(CONFIG, "export_mode", "text")
            worker = threading.Thread(target=manager._process, args=(job,))

            try:
                route_mode_patch.start()
                export_mode_patch.start()
                worker.start()
                worker.join(timeout=5)
            finally:
                manager.close()
                export_mode_patch.stop()
                route_mode_patch.stop()

            self.assertFalse(worker.is_alive())
            self.assertEqual(job.status, "succeeded")
            self.assertEqual(job.route, "text")
            self.assertEqual(job.route_reason, "text_layer_detected")
            self.assertIsNone(manager._pipeline)
            self.assertTrue(output_path.is_file())
            paragraphs = "\n".join(paragraph.text for paragraph in Document(output_path).paragraphs)
            self.assertIn("Fast route page one", paragraphs)
            self.assertIn("Fast route page two", paragraphs)


if __name__ == "__main__":
    unittest.main()
