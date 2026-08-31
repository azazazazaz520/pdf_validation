from __future__ import annotations

import asyncio
import hmac
import json
import logging
import multiprocessing
import os
import re
import shutil
import threading
import time
import uuid
from contextlib import asynccontextmanager
from concurrent.futures import CancelledError, Future, ProcessPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import Depends, FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pypdf import PdfReader

from pdf_to_word_exporter import (
    DEFAULT_PAGE_IMAGE_JPEG_QUALITY,
    DEFAULT_PAGE_IMAGE_MAX_PIXELS,
    EXPORT_MODES,
)
from pdf_worker import process_job


ROOT = Path(__file__).resolve().parent
LOGGER = logging.getLogger("pdf_to_word_service")
ROUTE_MODES = frozenset({"auto", "text", "ocr"})
TERMINAL_STATUSES = frozenset({"succeeded", "failed", "cancelled", "timed_out"})


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError as error:
        raise RuntimeError(f"环境变量 {name} 必须是整数") from error


def _env_float(name: str, default: float) -> float:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except ValueError as error:
        raise RuntimeError(f"环境变量 {name} 必须是数字") from error


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _iso(value: datetime | None) -> str | None:
    return value.isoformat() if value else None


@dataclass
class ServiceConfig:
    engine: str = os.getenv("PDF_SERVICE_ENGINE", "structure-lite")
    worker_processes: int = _env_int("PDF_SERVICE_WORKER_PROCESSES", 1)
    max_pending_jobs: int = _env_int("PDF_SERVICE_MAX_PENDING_JOBS", 4)
    route_mode: str = os.getenv("PDF_SERVICE_ROUTE_MODE", "auto").strip().lower()
    text_min_page_chars: int = _env_int("PDF_SERVICE_TEXT_MIN_PAGE_CHARS", 20)
    text_min_page_ratio: float = _env_float("PDF_SERVICE_TEXT_MIN_PAGE_RATIO", 0.6)
    text_high_quality_ratio: float = _env_float(
        "PDF_SERVICE_TEXT_HIGH_QUALITY_RATIO", 0.8
    )
    text_full_page_image_min_pixels: int = _env_int(
        "PDF_SERVICE_TEXT_FULL_PAGE_IMAGE_MIN_PIXELS", 300_000
    )
    text_garbled_char_ratio: float = _env_float(
        "PDF_SERVICE_TEXT_GARBLED_CHAR_RATIO", 0.05
    )
    export_mode: str = os.getenv("PDF_SERVICE_EXPORT_MODE", "hybrid").strip().lower()
    page_image_max_pixels: int = _env_int(
        "PDF_SERVICE_PAGE_IMAGE_MAX_PIXELS", DEFAULT_PAGE_IMAGE_MAX_PIXELS
    )
    page_image_jpeg_quality: int = _env_int(
        "PDF_SERVICE_PAGE_IMAGE_JPEG_QUALITY", DEFAULT_PAGE_IMAGE_JPEG_QUALITY
    )
    data_root: Path = Path(
        os.getenv("PDF_SERVICE_DATA_ROOT", str(ROOT / "service_data"))
    ).resolve()
    max_upload_bytes: int = _env_int("PDF_SERVICE_MAX_UPLOAD_BYTES", 50 * 1024 * 1024)
    max_pages: int = _env_int("PDF_SERVICE_MAX_PAGES", 100)
    job_ttl_seconds: int = _env_int("PDF_SERVICE_JOB_TTL_SECONDS", 3600)
    cleanup_interval_seconds: int = _env_int("PDF_SERVICE_CLEANUP_INTERVAL_SECONDS", 60)
    task_timeout_seconds: float = _env_float(
        "PDF_SERVICE_TASK_TIMEOUT_SECONDS", 300.0
    )
    ocr_time_budget_seconds: float = _env_float(
        "PDF_SERVICE_OCR_TIME_BUDGET_SECONDS", 60.0
    )
    auth_token: str = os.getenv("PDF_SERVICE_TOKEN", "")

    def __post_init__(self) -> None:
        if not 1 <= self.worker_processes <= 4:
            raise RuntimeError("环境变量 PDF_SERVICE_WORKER_PROCESSES 必须在 1 到 4 之间")
        if self.max_pending_jobs < self.worker_processes:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_MAX_PENDING_JOBS 不得小于 worker 进程数"
            )
        if self.route_mode not in ROUTE_MODES:
            raise RuntimeError("环境变量 PDF_SERVICE_ROUTE_MODE 必须是 auto、text 或 ocr")
        if self.text_min_page_chars < 1:
            raise RuntimeError("环境变量 PDF_SERVICE_TEXT_MIN_PAGE_CHARS 必须大于 0")
        if not 0 < self.text_min_page_ratio <= 1:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_TEXT_MIN_PAGE_RATIO 必须大于 0 且不超过 1"
            )
        if not 0 < self.text_high_quality_ratio <= 1:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_TEXT_HIGH_QUALITY_RATIO 必须大于 0 且不超过 1"
            )
        if self.text_full_page_image_min_pixels < 1:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_TEXT_FULL_PAGE_IMAGE_MIN_PIXELS 必须大于 0"
            )
        if not 0 <= self.text_garbled_char_ratio <= 1:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_TEXT_GARBLED_CHAR_RATIO 必须在 0 到 1 之间"
            )
        if self.export_mode not in EXPORT_MODES:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_EXPORT_MODE 必须是 hybrid 或 text"
            )
        if self.page_image_max_pixels < 1:
            raise RuntimeError("环境变量 PDF_SERVICE_PAGE_IMAGE_MAX_PIXELS 必须大于 0")
        if not 1 <= self.page_image_jpeg_quality <= 100:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_PAGE_IMAGE_JPEG_QUALITY 必须在 1 到 100 之间"
            )
        if self.task_timeout_seconds <= 0:
            raise RuntimeError("环境变量 PDF_SERVICE_TASK_TIMEOUT_SECONDS 必须大于 0")
        if self.ocr_time_budget_seconds <= 0:
            raise RuntimeError(
                "环境变量 PDF_SERVICE_OCR_TIME_BUDGET_SECONDS 必须大于 0"
            )


CONFIG = ServiceConfig()
JOB_ROOT = CONFIG.data_root / "jobs"
JOB_ROOT.mkdir(parents=True, exist_ok=True)


class QueueFullError(RuntimeError):
    """任务队列达到容量上限。"""


@dataclass
class Job:
    job_id: str
    filename: str
    workspace: Path
    input_path: Path
    output_path: Path
    progress_path: Path
    stage_log_path: Path
    cancel_path: Path
    page_count: int
    created_at: datetime = field(default_factory=_utc_now)
    started_at: datetime | None = None
    finished_at: datetime | None = None
    status: str = "queued"
    progress: int = 0
    route: str = "pending"
    route_reason: str | None = None
    worker_pid: int | None = None
    error: str | None = None
    future: Future[dict[str, Any]] | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id,
            "filename": self.filename,
            "status": self.status,
            "progress": self.progress,
            "route": self.route,
            "route_reason": self.route_reason,
            "worker_pid": self.worker_pid,
            "page_count": self.page_count,
            "created_at": _iso(self.created_at),
            "started_at": _iso(self.started_at),
            "finished_at": _iso(self.finished_at),
            "error": self.error,
            "download_url": (
                f"/api/pdf-to-word/jobs/{self.job_id}/result"
                if self.status == "succeeded"
                else None
            ),
        }


class JobManager:
    """管理 API 进程中的任务状态，并将转换交给独立 worker 进程。"""

    def __init__(self) -> None:
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()
        context_name = "spawn" if os.name == "nt" else "fork"
        self._executor = ProcessPoolExecutor(
            max_workers=CONFIG.worker_processes,
            mp_context=multiprocessing.get_context(context_name),
        )

    def _active_job_count_locked(self) -> int:
        return sum(
            job.status in {"queued", "processing"} for job in self._jobs.values()
        )

    def has_capacity(self) -> bool:
        with self._lock:
            return self._active_job_count_locked() < CONFIG.max_pending_jobs

    def create(self, filename: str, source_path: Path, page_count: int) -> Job:
        with self._lock:
            if self._active_job_count_locked() >= CONFIG.max_pending_jobs:
                raise QueueFullError("任务队列已满，请稍后重试")

        job_id = uuid.uuid4().hex
        workspace = JOB_ROOT / job_id
        workspace.mkdir(parents=True, exist_ok=False)
        input_path = workspace / "input.pdf"
        output_path = workspace / "result.docx"
        progress_path = workspace / "progress.json"
        stage_log_path = workspace / "stages.jsonl"
        cancel_path = workspace / "cancel.requested"
        try:
            shutil.move(str(source_path), input_path)
            job = Job(
                job_id=job_id,
                filename=filename,
                workspace=workspace,
                input_path=input_path,
                output_path=output_path,
                progress_path=progress_path,
                stage_log_path=stage_log_path,
                cancel_path=cancel_path,
                page_count=page_count,
            )
            with self._lock:
                self._jobs[job_id] = job
            job.future = self._executor.submit(process_job, self._worker_payload(job))
            job.future.add_done_callback(
                lambda future, current_job=job: self._on_worker_done(current_job, future)
            )
            return job
        except Exception:
            with self._lock:
                self._jobs.pop(job_id, None)
            shutil.rmtree(workspace, ignore_errors=True)
            raise

    def get(self, job_id: str) -> Job | None:
        with self._lock:
            job = self._jobs.get(job_id)
        if job is not None:
            self._refresh_from_worker(job)
        return job

    def cancel(self, job_id: str) -> Job:
        job = self.get(job_id)
        if job is None:
            raise KeyError(job_id)
        with self._lock:
            if job.status in {"queued", "processing"}:
                job.cancel_path.touch(exist_ok=True)
                if job.status == "queued" and job.future and job.future.cancel():
                    job.status = "cancelled"
                    job.progress = 0
                    job.finished_at = _utc_now()
        return job

    def cleanup_expired(self) -> int:
        cutoff = time.time() - CONFIG.job_ttl_seconds
        expired: list[Job] = []
        with self._lock:
            for job in self._jobs.values():
                if job.finished_at and job.finished_at.timestamp() < cutoff:
                    expired.append(job)
            for job in expired:
                self._jobs.pop(job.job_id, None)
        for job in expired:
            shutil.rmtree(job.workspace, ignore_errors=True)
        return len(expired)

    def model_loaded(self) -> bool:
        """返回 API 进程是否持有模型；模型实际驻留在 worker 进程中。"""
        return False

    def close(self) -> None:
        self._executor.shutdown(wait=True, cancel_futures=True)

    def _worker_payload(self, job: Job) -> dict[str, Any]:
        return {
            "job_id": job.job_id,
            "filename": job.filename,
            "input_path": str(job.input_path),
            "output_path": str(job.output_path),
            "progress_path": str(job.progress_path),
            "stage_log_path": str(job.stage_log_path),
            "cancel_path": str(job.cancel_path),
            "page_count": job.page_count,
            "engine": CONFIG.engine,
            "route_mode": CONFIG.route_mode,
            "text_min_page_chars": CONFIG.text_min_page_chars,
            "text_min_page_ratio": CONFIG.text_min_page_ratio,
            "text_high_quality_ratio": CONFIG.text_high_quality_ratio,
            "text_full_page_image_min_pixels": CONFIG.text_full_page_image_min_pixels,
            "text_garbled_char_ratio": CONFIG.text_garbled_char_ratio,
            "export_mode": CONFIG.export_mode,
            "page_image_max_pixels": CONFIG.page_image_max_pixels,
            "page_image_jpeg_quality": CONFIG.page_image_jpeg_quality,
            "task_timeout_seconds": CONFIG.task_timeout_seconds,
            "ocr_time_budget_seconds": CONFIG.ocr_time_budget_seconds,
        }

    def _refresh_from_worker(self, job: Job) -> None:
        try:
            payload = json.loads(job.progress_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        with self._lock:
            if job.status in TERMINAL_STATUSES:
                return
            status = payload.get("status")
            if isinstance(status, str):
                job.status = status
                if status == "processing" and job.started_at is None:
                    job.started_at = _utc_now()
            progress = payload.get("progress")
            if isinstance(progress, int):
                job.progress = progress
            route = payload.get("route")
            if isinstance(route, str):
                job.route = route
            route_reason = payload.get("route_reason")
            if isinstance(route_reason, str):
                job.route_reason = route_reason
            worker_pid = payload.get("worker_pid")
            if isinstance(worker_pid, int):
                job.worker_pid = worker_pid
            error = payload.get("error")
            if isinstance(error, str):
                job.error = error

    def _set_state(
        self,
        job: Job,
        *,
        status: str | None = None,
        progress: int | None = None,
        route: str | None = None,
        route_reason: str | None = None,
        worker_pid: int | None = None,
        error: str | None = None,
    ) -> None:
        with self._lock:
            if status is not None:
                job.status = status
            if progress is not None:
                job.progress = progress
            if route is not None:
                job.route = route
            if route_reason is not None:
                job.route_reason = route_reason
            if worker_pid is not None:
                job.worker_pid = worker_pid
            if error is not None:
                job.error = error

    def _on_worker_done(
        self, job: Job, future: Future[dict[str, Any]]
    ) -> None:
        with self._lock:
            if job.started_at is None:
                job.started_at = _utc_now()
        if future.cancelled():
            with self._lock:
                if job.status not in TERMINAL_STATUSES:
                    job.status = "cancelled"
                    job.progress = 0
            job.finished_at = _utc_now()
            return

        try:
            result = future.result()
        except Exception as error:
            LOGGER.exception("worker 进程未正常返回任务 %s", job.job_id)
            self._set_state(
                job,
                status="failed",
                progress=0,
                error=f"{type(error).__name__}: {error}",
            )
            job.finished_at = _utc_now()
            return

        self._set_state(
            job,
            status=str(result.get("status", "failed")),
            progress=int(result.get("progress", 0)),
            route=result.get("route"),
            route_reason=result.get("route_reason"),
            worker_pid=result.get("worker_pid"),
            error=result.get("error"),
        )
        job.finished_at = _utc_now()


MANAGER = JobManager()


def _require_auth(request: Request) -> None:
    if not CONFIG.auth_token:
        return
    authorization = request.headers.get("authorization", "")
    api_key = request.headers.get("x-api-key", "")
    bearer = authorization.removeprefix("Bearer ").strip()
    if not hmac.compare_digest(bearer or api_key, CONFIG.auth_token):
        raise HTTPException(status_code=401, detail="未提供有效的服务访问凭证")


def _safe_filename(filename: str | None) -> str:
    name = Path(filename or "input.pdf").name
    name = re.sub(r"[^\w.\-一-龥]", "_", name)
    return name or "input.pdf"


async def _save_upload(upload: UploadFile, target: Path) -> int:
    total = 0
    first_chunk = b""
    with target.open("wb") as stream:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            if not first_chunk:
                first_chunk = chunk[:16]
            total += len(chunk)
            if total > CONFIG.max_upload_bytes:
                raise HTTPException(status_code=413, detail="PDF 文件超过大小限制")
            stream.write(chunk)
    if not first_chunk.lstrip().startswith(b"%PDF-"):
        raise HTTPException(status_code=415, detail="上传文件不是有效的 PDF")
    return total


def _read_page_count(path: Path) -> int:
    try:
        reader = PdfReader(str(path))
        if reader.is_encrypted:
            raise HTTPException(status_code=400, detail="暂不处理加密 PDF")
        page_count = len(reader.pages)
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=400, detail=f"PDF 解析失败：{error}") from error
    if page_count < 1:
        raise HTTPException(status_code=400, detail="PDF 不包含页面")
    if page_count > CONFIG.max_pages:
        raise HTTPException(status_code=413, detail="PDF 页数超过限制")
    return page_count


async def _cleanup_loop() -> None:
    while True:
        await asyncio.sleep(CONFIG.cleanup_interval_seconds)
        removed = MANAGER.cleanup_expired()
        if removed:
            LOGGER.info("清理过期任务：%s", removed)


@asynccontextmanager
async def lifespan(_: FastAPI):
    cleanup_task = asyncio.create_task(_cleanup_loop())
    try:
        yield
    finally:
        cleanup_task.cancel()
        try:
            await cleanup_task
        except asyncio.CancelledError:
            pass
        MANAGER.close()


app = FastAPI(
    title="PDF 转 Word 服务原型",
    version="0.1.0",
    lifespan=lifespan,
)

allowed_origins = [
    origin.strip()
    for origin in os.getenv("PDF_SERVICE_ALLOWED_ORIGINS", "").split(",")
    if origin.strip()
]
if allowed_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=allowed_origins,
        allow_credentials=False,
        allow_methods=["GET", "POST", "DELETE"],
        allow_headers=["Authorization", "Content-Type", "X-API-Key"],
    )


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "engine": CONFIG.engine,
        "worker_processes": CONFIG.worker_processes,
        "max_pending_jobs": CONFIG.max_pending_jobs,
        "route_mode": CONFIG.route_mode,
        "text_min_page_chars": CONFIG.text_min_page_chars,
        "text_min_page_ratio": CONFIG.text_min_page_ratio,
        "text_high_quality_ratio": CONFIG.text_high_quality_ratio,
        "text_full_page_image_min_pixels": CONFIG.text_full_page_image_min_pixels,
        "text_garbled_char_ratio": CONFIG.text_garbled_char_ratio,
        "export_mode": CONFIG.export_mode,
        "model_loaded": MANAGER.model_loaded(),
        "max_upload_bytes": CONFIG.max_upload_bytes,
        "max_pages": CONFIG.max_pages,
        "page_image_max_pixels": CONFIG.page_image_max_pixels,
        "task_timeout_seconds": CONFIG.task_timeout_seconds,
        "ocr_time_budget_seconds": CONFIG.ocr_time_budget_seconds,
    }


@app.post("/api/pdf-to-word/jobs", dependencies=[Depends(_require_auth)])
async def create_job(file: UploadFile = File(...)) -> dict[str, Any]:
    MANAGER.cleanup_expired()
    if not MANAGER.has_capacity():
        raise HTTPException(status_code=429, detail="任务队列已满，请稍后重试")
    staging_root = CONFIG.data_root / "staging"
    staging_root.mkdir(parents=True, exist_ok=True)
    staging_path = staging_root / f"{uuid.uuid4().hex}.upload"
    try:
        await _save_upload(file, staging_path)
        page_count = _read_page_count(staging_path)
        job = MANAGER.create(_safe_filename(file.filename), staging_path, page_count)
        return job.as_dict()
    except HTTPException:
        staging_path.unlink(missing_ok=True)
        raise
    except QueueFullError as error:
        staging_path.unlink(missing_ok=True)
        raise HTTPException(status_code=429, detail=str(error)) from error
    except Exception as error:
        staging_path.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"创建任务失败：{error}") from error
    finally:
        await file.close()


@app.get("/api/pdf-to-word/jobs/{job_id}", dependencies=[Depends(_require_auth)])
def get_job(job_id: str) -> dict[str, Any]:
    job = MANAGER.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    return job.as_dict()


@app.get("/api/pdf-to-word/jobs/{job_id}/result", dependencies=[Depends(_require_auth)])
def download_result(job_id: str) -> FileResponse:
    job = MANAGER.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    if job.status != "succeeded" or not job.output_path.is_file():
        raise HTTPException(status_code=409, detail=f"任务当前状态为 {job.status}")
    return FileResponse(
        job.output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{Path(job.filename).stem}.docx",
    )


@app.delete("/api/pdf-to-word/jobs/{job_id}", dependencies=[Depends(_require_auth)])
def cancel_job(job_id: str) -> dict[str, Any]:
    try:
        job = MANAGER.cancel(job_id)
    except KeyError as error:
        raise HTTPException(status_code=404, detail="任务不存在或已过期") from error
    return job.as_dict()


if __name__ == "__main__":
    import uvicorn

    logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"))
    uvicorn.run(
        app,
        host=os.getenv("PDF_SERVICE_HOST", "127.0.0.1"),
        port=_env_int("PDF_SERVICE_PORT", 8765),
    )
