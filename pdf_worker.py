from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pdf_routing import analyze_pdf_text
from ocr_quality import summarize_ocr_page
from pdf_to_word_exporter import (
    export_results_to_docx,
    export_source_pages_to_docx,
    export_text_pages_to_docx,
)
from run_validation import build_pipeline


_PIPELINES: dict[str, Any] = {}


class _Cancelled(Exception):
    pass


class _TimedOut(Exception):
    def __init__(self, stage: str, elapsed_seconds: float, budget_seconds: float):
        self.stage = stage
        self.elapsed_seconds = elapsed_seconds
        self.budget_seconds = budget_seconds
        super().__init__(
            f"任务超过时间预算：stage={stage}, "
            f"elapsed={elapsed_seconds:.3f}s, budget={budget_seconds:.3f}s"
        )


class _ProgressWriter:
    """以原子方式写入任务最新状态，并追加阶段日志。"""

    def __init__(self, progress_path: Path, stage_log_path: Path) -> None:
        self.progress_path = progress_path
        self.stage_log_path = stage_log_path
        self.progress_path.parent.mkdir(parents=True, exist_ok=True)
        self.stage_log_path.parent.mkdir(parents=True, exist_ok=True)

    def emit(
        self,
        stage: str,
        *,
        status: str = "processing",
        progress: int | None = None,
        route: str | None = None,
        route_reason: str | None = None,
        error: str | None = None,
        **details: Any,
    ) -> None:
        event = {
            "stage": stage,
            "status": status,
            "progress": progress,
            "route": route,
            "route_reason": route_reason,
            "error": error,
            "worker_pid": os.getpid(),
            "updated_at": datetime.now(timezone.utc).isoformat(),
            **details,
        }
        temporary_path = self.progress_path.with_name(
            f".{self.progress_path.name}.tmp"
        )
        temporary_path.write_text(
            json.dumps(event, ensure_ascii=False), encoding="utf-8"
        )
        for attempt in range(10):
            try:
                os.replace(temporary_path, self.progress_path)
                break
            except PermissionError:
                if attempt == 9:
                    raise
                time.sleep(0.02 * (attempt + 1))
        with self.stage_log_path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(event, ensure_ascii=False) + "\n")


def _get_pipeline(engine: str) -> Any:
    pipeline = _PIPELINES.get(engine)
    if pipeline is None:
        pipeline = build_pipeline(engine)
        _PIPELINES[engine] = pipeline
    return pipeline


def _is_cancelled(cancel_path: Path) -> bool:
    return cancel_path.exists()


def _raise_if_cancelled(cancel_path: Path) -> None:
    if _is_cancelled(cancel_path):
        raise _Cancelled()


def process_job(payload: dict[str, Any]) -> dict[str, Any]:
    """在独立 worker 进程中执行一个 PDF 转 Word 任务。"""
    job_id = str(payload["job_id"])
    input_path = Path(payload["input_path"])
    output_path = Path(payload["output_path"])
    progress_path = Path(payload["progress_path"])
    stage_log_path = Path(payload["stage_log_path"])
    cancel_path = Path(payload["cancel_path"])
    page_count = int(payload["page_count"])
    route_mode = str(payload["route_mode"])
    route = "ocr"
    route_reason = "forced_ocr"
    task_started = time.perf_counter()
    task_timeout_seconds = float(payload.get("task_timeout_seconds", 300.0))
    ocr_time_budget_seconds = float(payload.get("ocr_time_budget_seconds", 60.0))
    writer = _ProgressWriter(progress_path, stage_log_path)

    def raise_if_timed_out(stage: str) -> None:
        budget_seconds = (
            ocr_time_budget_seconds if route == "ocr" else task_timeout_seconds
        )
        elapsed_seconds = time.perf_counter() - task_started
        if elapsed_seconds > budget_seconds:
            raise _TimedOut(stage, elapsed_seconds, budget_seconds)

    def emit_export_stage(stage: str, details: dict[str, Any]) -> None:
        _raise_if_cancelled(cancel_path)
        raise_if_timed_out(stage)
        progress = None
        if stage in {"text_export_started", "page_images_started"}:
            progress = 60
        elif stage in {
            "text_export_pages_completed",
            "page_images_completed",
        }:
            progress = 90
        elif stage == "ocr_text_started":
            progress = 90
        elif stage == "ocr_text_completed":
            progress = 95
        elif stage == "source_pages_completed":
            progress = 95
        writer.emit(
            stage,
            progress=progress,
            route=route,
            route_reason=route_reason,
            **details,
        )

    try:
        writer.emit(
            "started",
            progress=5,
            job_id=job_id,
            page_count=page_count,
            filename=str(payload["filename"]),
            task_timeout_seconds=task_timeout_seconds,
            ocr_time_budget_seconds=ocr_time_budget_seconds,
        )

        analysis = None
        if route_mode != "ocr":
            analysis_started = time.perf_counter()
            writer.emit("text_analysis_started", progress=8, job_id=job_id)
            analysis = analyze_pdf_text(
                input_path,
                min_page_chars=int(payload["text_min_page_chars"]),
                full_page_image_min_pixels=int(
                    payload["text_full_page_image_min_pixels"]
                ),
                garbled_char_ratio_threshold=float(
                    payload["text_garbled_char_ratio"]
                ),
            )
            has_text_layer = analysis.has_usable_text_layer(
                min_page_chars=int(payload["text_min_page_chars"]),
                min_page_ratio=float(payload["text_min_page_ratio"]),
            )
            has_complete_text_layer = analysis.has_complete_text_layer(
                min_page_chars=int(payload["text_min_page_chars"]),
                min_high_quality_ratio=float(payload["text_high_quality_ratio"]),
            )
            writer.emit(
                "text_analysis_completed",
                progress=12,
                job_id=job_id,
                elapsed_sec=round(time.perf_counter() - analysis_started, 3),
                usable_page_count=analysis.usable_page_count,
                page_count=analysis.page_count,
                text_char_count=analysis.text_char_count,
                usable_page_ratio=round(analysis.usable_page_ratio, 3),
                high_quality_page_count=analysis.high_quality_page_count,
                high_quality_page_ratio=round(analysis.high_quality_page_ratio, 3),
                full_page_image_page_count=analysis.full_page_image_page_count,
                garbled_char_count=analysis.garbled_char_count,
                has_text_layer=has_text_layer,
                has_complete_text_layer=has_complete_text_layer,
            )
            if route_mode == "text":
                if not has_text_layer:
                    raise RuntimeError("PDF 不满足文本层快速路线的检测阈值")
                route = "text"
                route_reason = "forced_text"
            elif has_complete_text_layer:
                route = "text"
                route_reason = "text_layer_complete"
            elif has_text_layer:
                route = "page_image"
                route_reason = "text_layer_incomplete"
            else:
                route_reason = "text_layer_not_usable"

            _raise_if_cancelled(cancel_path)
            raise_if_timed_out("route_selected")

        writer.emit(
            "route_selected",
            progress=15,
            route=route,
            route_reason=route_reason,
        )

        if route in {"text", "page_image"}:
            if analysis is None:
                raise RuntimeError("文本层快速路线缺少分析结果")
            _raise_if_cancelled(cancel_path)
            export_started = time.perf_counter()
            writer.emit(
                "docx_export_started",
                progress=60,
                route=route,
                route_reason=route_reason,
                page_count=analysis.page_count,
            )
            if route == "text":
                export_text_pages_to_docx(
                    analysis.page_texts,
                    output_path,
                    title=Path(str(payload["filename"])).stem,
                    source_pdf=input_path,
                    stage_callback=emit_export_stage,
                )
            else:
                export_source_pages_to_docx(
                    output_path,
                    title=Path(str(payload["filename"])).stem,
                    source_pdf=input_path,
                    image_max_pixels=int(payload["page_image_max_pixels"]),
                    image_jpeg_quality=int(payload["page_image_jpeg_quality"]),
                    stage_callback=emit_export_stage,
                )
            export_elapsed = time.perf_counter() - export_started
        else:
            model_started = time.perf_counter()
            writer.emit(
                "model_loading_started",
                progress=15,
                route=route,
                route_reason=route_reason,
                engine=str(payload["engine"]),
            )
            pipeline = _get_pipeline(str(payload["engine"]))
            writer.emit(
                "model_loading_completed",
                progress=20,
                route=route,
                route_reason=route_reason,
                elapsed_sec=round(time.perf_counter() - model_started, 3),
            )
            inference_started = time.perf_counter()
            writer.emit(
                "inference_started",
                progress=20,
                route=route,
                route_reason=route_reason,
                input=input_path.name,
            )
            results: list[dict[str, Any]] = []
            for result in pipeline.predict_iter(str(input_path)):
                _raise_if_cancelled(cancel_path)
                raise_if_timed_out("inference_page_completed")
                results.append(result)
                progress = min(80, 20 + int(60 * len(results) / max(page_count, 1)))
                page_quality = summarize_ocr_page(
                    result,
                    page_number=len(results),
                )
                writer.emit(
                    "inference_page_completed",
                    progress=progress,
                    route=route,
                    route_reason=route_reason,
                    **page_quality,
                )
                writer.emit(
                    "inference_batch_completed",
                    progress=progress,
                    route=route,
                    route_reason=route_reason,
                    result_count=len(results),
                )
            _raise_if_cancelled(cancel_path)
            writer.emit(
                "inference_completed",
                progress=85,
                route=route,
                route_reason=route_reason,
                result_count=len(results),
                elapsed_sec=round(time.perf_counter() - inference_started, 3),
            )
            export_started = time.perf_counter()
            writer.emit(
                "docx_export_started",
                progress=90,
                route=route,
                route_reason=route_reason,
                result_count=len(results),
            )
            export_results_to_docx(
                results,
                output_path,
                title=Path(str(payload["filename"])).stem,
                source_pdf=input_path,
                mode=str(payload["export_mode"]),
                image_max_pixels=int(payload["page_image_max_pixels"]),
                image_jpeg_quality=int(payload["page_image_jpeg_quality"]),
                stage_callback=emit_export_stage,
            )
            export_elapsed = time.perf_counter() - export_started

        output_bytes = output_path.stat().st_size
        writer.emit(
            "docx_export_completed",
            progress=95,
            route=route,
            route_reason=route_reason,
            elapsed_sec=round(export_elapsed, 3),
            output_bytes=output_bytes,
        )
        writer.emit(
            "finished",
            status="succeeded",
            progress=100,
            route=route,
            route_reason=route_reason,
            elapsed_sec=round(time.perf_counter() - task_started, 3),
        )
        return {
            "status": "succeeded",
            "progress": 100,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
        }
    except _TimedOut as error:
        writer.emit(
            "timed_out",
            status="timed_out",
            progress=0,
            route=route,
            route_reason=route_reason,
            error=str(error),
            timeout_stage=error.stage,
            elapsed_sec=round(error.elapsed_seconds, 3),
            timeout_budget_sec=error.budget_seconds,
        )
        return {
            "status": "timed_out",
            "progress": 0,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
            "error": str(error),
        }
    except _Cancelled:
        writer.emit(
            "cancelled",
            status="cancelled",
            progress=0,
            route=route,
            route_reason=route_reason,
            elapsed_sec=round(time.perf_counter() - task_started, 3),
        )
        return {
            "status": "cancelled",
            "progress": 0,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
        }
    except Exception as error:
        writer.emit(
            "failed",
            status="failed",
            progress=0,
            route=route,
            route_reason=route_reason,
            error=f"{type(error).__name__}: {error}",
            elapsed_sec=round(time.perf_counter() - task_started, 3),
        )
        return {
            "status": "failed",
            "progress": 0,
            "route": route,
            "route_reason": route_reason,
            "worker_pid": os.getpid(),
            "error": f"{type(error).__name__}: {error}",
        }
