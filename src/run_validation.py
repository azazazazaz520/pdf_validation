from __future__ import annotations

import argparse
import json
import os
import platform
import sys
import time
import traceback
import zipfile
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
os.environ.setdefault("PADDLE_PDX_CACHE_HOME", str(ROOT / "model_cache"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="运行 PaddleOCR PDF 转 Word 独立验证")
    parser.add_argument(
        "--engine",
        choices=("vl", "structure", "structure-lite", "structure-table-lite"),
        required=True,
    )
    parser.add_argument("--input", type=Path, default=ROOT / "fixtures" / "synthetic_text_table.pdf")
    parser.add_argument("--output", type=Path, default=ROOT / "artifacts" / "outputs")
    return parser.parse_args()


def build_pipeline(engine: str) -> Any:
    if engine == "vl":
        from paddleocr import PaddleOCRVL

        return PaddleOCRVL(
            pipeline_version="v1.6",
            device="cpu",
            enable_mkldnn=False,
            cpu_threads=4,
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_chart_recognition=False,
            use_seal_recognition=False,
        )

    from paddleocr import PPStructureV3

    lite = engine in ("structure-lite", "structure-table-lite")
    table_enabled = engine in ("structure", "structure-table-lite")

    return PPStructureV3(
        device="cpu",
        enable_mkldnn=False,
        cpu_threads=4,
        layout_detection_model_name="PP-DocLayout-S" if lite else None,
        text_detection_model_name="PP-OCRv5_mobile_det" if lite else None,
        text_recognition_model_name="PP-OCRv5_mobile_rec" if lite else None,
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        use_table_recognition=table_enabled,
        use_chart_recognition=False,
        use_seal_recognition=False,
        use_formula_recognition=False,
        use_region_detection=False,
    )


def get_page_count(path: Path) -> int | None:
    try:
        from pypdfium2 import PdfDocument

        document = PdfDocument(str(path))
        count = len(document)
        document.close()
        return count
    except Exception:
        return None


def summarize_result(result: Any) -> dict[str, Any]:
    summary: dict[str, Any] = {"type": type(result).__name__}
    if isinstance(result, dict):
        for key in ("input_path", "page_index", "page_count", "input_img", "layout_det_res"):
            if key in result:
                value = result[key]
                if key == "input_img" and value is not None:
                    summary[key] = "present"
                elif key == "layout_det_res" and isinstance(value, dict):
                    summary[key] = {"keys": sorted(value.keys())[:20]}
                elif key == "layout_det_res" and isinstance(value, (list, tuple)):
                    summary[key] = {"type": type(value).__name__, "length": len(value)}
                else:
                    summary[key] = summarize_value(value)
        summary["keys"] = sorted(str(key) for key in result.keys())
        for key in ("markdown", "pruned_result", "parsing_res_list"):
            value = result.get(key)
            if isinstance(value, str):
                summary[f"{key}_chars"] = len(value)
            elif isinstance(value, dict):
                summary[f"{key}_keys"] = sorted(str(item) for item in value.keys())
    return summary


def summarize_value(value: Any) -> Any:
    """将摘要中的标量和数组转换为可序列化的简洁值。"""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    shape = getattr(value, "shape", None)
    if shape is not None:
        return {
            "type": type(value).__name__,
            "shape": list(shape),
            "dtype": str(getattr(value, "dtype", "unknown")),
        }
    return str(value)


def validate_docx(path: Path) -> dict[str, Any]:
    result: dict[str, Any] = {"path": str(path), "exists": path.is_file()}
    if not path.is_file():
        return result
    result["bytes"] = path.stat().st_size
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
            result["zip_test"] = archive.testzip() is None
            result["required_entries"] = sorted(required & names)
            result["required_entries_complete"] = required <= names
            result["media_count"] = sum(name.startswith("word/media/") for name in names)
            document_xml = archive.read("word/document.xml")
            result["document_xml_bytes"] = len(document_xml)
            result["text_node_count"] = document_xml.count(b"<w:t")
    except (OSError, zipfile.BadZipFile, KeyError) as error:
        result["zip_error"] = f"{type(error).__name__}: {error}"
    return result


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    if not input_path.is_file():
        print(json.dumps({"status": "failed", "error": f"输入文件不存在：{input_path}"}, ensure_ascii=False))
        return 2

    run_dir = args.output.resolve() / f"{args.engine}_{time.strftime('%Y%m%d_%H%M%S')}"
    json_dir = run_dir / "json"
    markdown_dir = run_dir / "markdown"
    word_dir = run_dir / "word"
    for directory in (json_dir, markdown_dir, word_dir):
        directory.mkdir(parents=True, exist_ok=True)

    report: dict[str, Any] = {
        "status": "started",
        "engine": args.engine,
        "input": str(input_path),
        "input_bytes": input_path.stat().st_size,
        "input_page_count": get_page_count(input_path),
        "python": sys.version,
        "platform": platform.platform(),
        "cache_home": os.environ["PADDLE_PDX_CACHE_HOME"],
        "run_dir": str(run_dir),
    }
    started = time.perf_counter()
    pipeline = None
    raw_results: list[Any] = []
    final_results: list[Any] = []
    try:
        print(json.dumps({"event": "started", "engine": args.engine}, ensure_ascii=False), flush=True)
        init_started = time.perf_counter()
        pipeline = build_pipeline(args.engine)
        report["init_seconds"] = round(time.perf_counter() - init_started, 3)
        print(json.dumps({"event": "pipeline_ready", "seconds": report["init_seconds"]}), flush=True)

        predict_started = time.perf_counter()
        raw_results = list(pipeline.predict(str(input_path)))
        report["predict_seconds"] = round(time.perf_counter() - predict_started, 3)
        report["raw_result_count"] = len(raw_results)
        report["raw_results"] = [summarize_result(result) for result in raw_results]
        print(json.dumps({"event": "predicted", "result_count": len(raw_results)}), flush=True)

        restructure_started = time.perf_counter()
        if args.engine == "vl":
            final_results = list(
                pipeline.restructure_pages(
                    raw_results,
                    merge_tables=True,
                    relevel_titles=True,
                    concatenate_pages=True,
                )
            )
        else:
            final_results = raw_results
        report["restructure_seconds"] = round(time.perf_counter() - restructure_started, 3)
        report["final_result_count"] = len(final_results)
        report["final_results"] = [summarize_result(result) for result in final_results]
        print(json.dumps({"event": "restructured", "result_count": len(final_results)}), flush=True)

        save_errors: list[str] = []
        for result in final_results:
            for directory, method_name in (
                (json_dir, "save_to_json"),
                (markdown_dir, "save_to_markdown"),
                (word_dir, "save_to_word"),
            ):
                try:
                    getattr(result, method_name)(save_path=str(directory))
                except Exception as error:
                    save_errors.append(f"{method_name}: {type(error).__name__}: {error}")
        report["save_errors"] = save_errors
        report["docx_files"] = [
            validate_docx(path) for path in sorted(word_dir.glob("*.docx"))
        ]
        report["status"] = "success" if not save_errors else "partial"
    except Exception as error:
        report["status"] = "failed"
        report["error"] = f"{type(error).__name__}: {error}"
        traceback.print_exc()
    finally:
        if pipeline is not None:
            try:
                pipeline.close()
            except Exception as error:
                report["close_error"] = f"{type(error).__name__}: {error}"
        report["total_seconds"] = round(time.perf_counter() - started, 3)

    report_path = run_dir / "report.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=summarize_value),
        encoding="utf-8",
    )
    print(json.dumps({"event": "completed", "status": report["status"], "report": str(report_path)}, ensure_ascii=False))
    return 0 if report["status"] == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
