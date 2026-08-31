from __future__ import annotations

import json
import time
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import src.pdf_worker as pdf_worker


class _SlowPipeline:
    def predict_iter(self, _: str):
        time.sleep(0.02)
        yield {
            "overall_ocr_res": {
                "rec_texts": ["测试文字"],
                "rec_scores": [0.98],
            }
        }


class PdfWorkerTest(unittest.TestCase):
    def test_ocr_time_budget_returns_terminal_timeout_status(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            input_path = root / "input.pdf"
            input_path.write_bytes(b"%PDF-test")
            progress_path = root / "progress.json"
            stage_log_path = root / "stages.jsonl"
            cancel_path = root / "cancel.requested"
            payload = {
                "job_id": "timeout-test",
                "filename": "input.pdf",
                "input_path": str(input_path),
                "output_path": str(root / "result.docx"),
                "progress_path": str(progress_path),
                "stage_log_path": str(stage_log_path),
                "cancel_path": str(cancel_path),
                "page_count": 1,
                "engine": "structure-lite",
                "route_mode": "ocr",
                "task_timeout_seconds": 10.0,
                "ocr_time_budget_seconds": 0.001,
            }

            with patch.object(pdf_worker, "_get_pipeline", return_value=_SlowPipeline()):
                result = pdf_worker.process_job(payload)

            self.assertEqual(result["status"], "timed_out")
            events = [
                json.loads(line)
                for line in stage_log_path.read_text(encoding="utf-8").splitlines()
            ]
            self.assertEqual(events[-1]["stage"], "timed_out")
            self.assertEqual(events[-1]["status"], "timed_out")
            self.assertEqual(events[-1]["timeout_stage"], "inference_page_completed")
            self.assertFalse((root / "result.docx").exists())


if __name__ == "__main__":
    unittest.main()
