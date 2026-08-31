from __future__ import annotations

import unittest

from src.ocr_quality import summarize_ocr_page


class OcrQualityTest(unittest.TestCase):
    def test_page_quality_reports_confidence_and_garbled_signal(self) -> None:
        quality = summarize_ocr_page(
            {
                "overall_ocr_res": {
                    "rec_texts": ["清晰文字", "�异常"],
                    "rec_scores": [0.96, 0.42],
                }
            },
            page_number=3,
        )

        self.assertEqual(quality["page_number"], 3)
        self.assertEqual(quality["ocr_line_count"], 2)
        self.assertEqual(quality["ocr_low_confidence_count"], 1)
        self.assertEqual(quality["ocr_low_confidence_ratio"], 0.5)
        self.assertEqual(quality["ocr_min_confidence"], 0.42)
        self.assertGreaterEqual(quality["ocr_garbled_char_count"], 1)
        self.assertTrue(quality["ocr_needs_review"])

    def test_empty_ocr_page_needs_review(self) -> None:
        quality = summarize_ocr_page({"rec_texts": [], "rec_scores": []}, page_number=1)

        self.assertEqual(quality["ocr_line_count"], 0)
        self.assertTrue(quality["ocr_needs_review"])


if __name__ == "__main__":
    unittest.main()
