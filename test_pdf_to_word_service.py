from __future__ import annotations

import threading
import time
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from pdf_to_word_service import CONFIG, Job, JobManager


class _ControlledPipeline:
    """控制推理入口返回时机，用于验证服务的进度更新时机。"""

    def __init__(self) -> None:
        self.first_result_started = threading.Event()
        self.release_second_result = threading.Event()

    def predict(self, _: str) -> list[dict[str, list[object]]]:
        self.first_result_started.set()
        self.release_second_result.wait(timeout=2)
        return [{"parsing_res_list": []}, {"parsing_res_list": []}]

    def predict_iter(self, _: str):
        self.first_result_started.set()
        yield {"parsing_res_list": []}
        self.release_second_result.wait(timeout=2)
        yield {"parsing_res_list": []}

    def close(self) -> None:
        pass


class JobProcessingTest(unittest.TestCase):
    def test_progress_updates_after_first_prediction_batch(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            workspace = Path(temporary_directory)
            input_path = workspace / "input.pdf"
            output_path = workspace / "result.docx"
            input_path.write_bytes(b"%PDF-test")
            job = Job(
                job_id="test-job",
                filename="input.pdf",
                workspace=workspace,
                input_path=input_path,
                output_path=output_path,
                page_count=2,
            )
            manager = JobManager()
            pipeline = _ControlledPipeline()
            manager._pipeline = pipeline
            worker = threading.Thread(target=manager._process, args=(job,))
            export_mode_patch = patch.object(CONFIG, "export_mode", "text")
            route_mode_patch = patch.object(CONFIG, "route_mode", "ocr")

            try:
                export_mode_patch.start()
                route_mode_patch.start()
                worker.start()
                self.assertTrue(pipeline.first_result_started.wait(timeout=1))

                deadline = time.monotonic() + 1
                while job.progress <= 5 and time.monotonic() < deadline:
                    time.sleep(0.01)

                self.assertGreater(job.progress, 5)
            finally:
                pipeline.release_second_result.set()
                worker.join(timeout=3)
                export_mode_patch.stop()
                route_mode_patch.stop()
                manager.close()

            self.assertFalse(worker.is_alive())
            self.assertEqual(job.status, "succeeded")
            self.assertEqual(job.progress, 100)
            self.assertTrue(output_path.is_file())


if __name__ == "__main__":
    unittest.main()
