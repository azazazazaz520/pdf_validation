from __future__ import annotations

import time
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from docx import Document
from reportlab.pdfgen.canvas import Canvas

import pdf_to_word_service as service


def _create_text_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "Worker process page one with selectable text")
    canvas.showPage()
    canvas.drawString(36, 200, "Worker process page two with selectable text")
    canvas.save()


class JobProcessingTest(unittest.TestCase):
    def test_worker_process_completes_text_route_and_writes_stage_log(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source_path = root / "input.pdf"
            _create_text_pdf(source_path)
            jobs_root = root / "jobs"
            jobs_root.mkdir()
            original_job_root = service.JOB_ROOT
            service.JOB_ROOT = jobs_root
            manager = None

            try:
                with patch.object(service.CONFIG, "route_mode", "auto"), patch.object(
                    service.CONFIG, "export_mode", "text"
                ):
                    manager = service.JobManager()
                    job = manager.create("input.pdf", source_path, page_count=2)
                    deadline = time.monotonic() + 10
                    while time.monotonic() < deadline:
                        current = manager.get(job.job_id)
                        if current is not None and current.status in {
                            "succeeded",
                            "failed",
                            "cancelled",
                        }:
                            break
                        time.sleep(0.05)

                    current = manager.get(job.job_id)
                    self.assertIsNotNone(current)
                    assert current is not None
                    self.assertEqual(current.status, "succeeded")
                    self.assertEqual(current.route, "text")
                    self.assertEqual(current.route_reason, "text_layer_complete")
                    self.assertIsNotNone(current.worker_pid)
                    assert current.worker_pid is not None
                    self.assertGreater(current.worker_pid, 0)
                    self.assertTrue(current.output_path.is_file())
                    stage_log = current.stage_log_path.read_text(encoding="utf-8")
                    self.assertIn('"stage": "route_selected"', stage_log)
                    self.assertIn('"stage": "finished"', stage_log)
                    paragraphs = "\n".join(
                        paragraph.text
                        for paragraph in Document(current.output_path).paragraphs
                    )
                    self.assertIn("Worker process page one", paragraphs)
                    self.assertIn("Worker process page two", paragraphs)
            finally:
                if manager is not None:
                    manager.close()
                service.JOB_ROOT = original_job_root

    def test_queue_capacity_rejects_new_job(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source_path = root / "input.pdf"
            source_path.write_bytes(b"%PDF-test")
            manager = service.JobManager()
            placeholder = service.Job(
                job_id="queued-job",
                filename="queued.pdf",
                workspace=root,
                input_path=root / "queued-input.pdf",
                output_path=root / "queued-result.docx",
                progress_path=root / "queued-progress.json",
                stage_log_path=root / "queued-stages.jsonl",
                cancel_path=root / "queued.cancel",
                page_count=1,
            )
            with manager._lock:
                manager._jobs[placeholder.job_id] = placeholder

            try:
                with patch.object(service.CONFIG, "max_pending_jobs", 1):
                    with self.assertRaises(service.QueueFullError):
                        manager.create("input.pdf", source_path, page_count=1)
                self.assertTrue(source_path.is_file())
            finally:
                manager.close()


if __name__ == "__main__":
    unittest.main()
