from __future__ import annotations

import asyncio
import hmac
import logging
import os
import re
import shutil
import threading
import time
import uuid
from contextlib import asynccontextmanager
from concurrent.futures import Future, ThreadPoolExecutor
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
    export_results_to_docx,
)
from run_validation import build_pipeline


ROOT = Path(__file__).resolve().parent
LOGGER = logging.getLogger("pdf_to_word_service")


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError as error:
        raise RuntimeError(f"环境变量 {name} 必须是整数") from error


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _iso(value: datetime | None) -> str | None:
    return value.isoformat() if value else None


@dataclass
class ServiceConfig:
    engine: str = os.getenv("PDF_SERVICE_ENGINE", "structure-table-lite")
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
    auth_token: str = os.getenv("PDF_SERVICE_TOKEN", "")

    def __post_init__(self) -> None:
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


CONFIG = ServiceConfig()
JOB_ROOT = CONFIG.data_root / "jobs"
JOB_ROOT.mkdir(parents=True, exist_ok=True)


@dataclass
class Job:
    job_id: str
    filename: str
    workspace: Path
    input_path: Path
    output_path: Path
    page_count: int
    created_at: datetime = field(default_factory=_utc_now)
    started_at: datetime | None = None
    finished_at: datetime | None = None
    status: str = "queued"
    progress: int = 0
    error: str | None = None
    cancel_event: threading.Event = field(default_factory=threading.Event)
    future: Future[None] | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id,
            "filename": self.filename,
            "status": self.status,
            "progress": self.progress,
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
    """管理单进程内存任务，并以单工作线程保护模型资源。"""

    def __init__(self) -> None:
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="pdf-converter"
        )
        self._pipeline: Any | None = None
        self._pipeline_lock = threading.Lock()

    def create(self, filename: str, source_path: Path, page_count: int) -> Job:
        job_id = uuid.uuid4().hex
        workspace = JOB_ROOT / job_id
        workspace.mkdir(parents=True, exist_ok=False)
        input_path = workspace / "input.pdf"
        output_path = workspace / "result.docx"
        shutil.move(str(source_path), input_path)
        job = Job(
            job_id=job_id,
            filename=filename,
            workspace=workspace,
            input_path=input_path,
            output_path=output_path,
            page_count=page_count,
        )
        with self._lock:
            self._jobs[job_id] = job
        job.future = self._executor.submit(self._process, job)
        return job

    def get(self, job_id: str) -> Job | None:
        with self._lock:
            return self._jobs.get(job_id)

    def cancel(self, job_id: str) -> Job:
        job = self.get(job_id)
        if job is None:
            raise KeyError(job_id)
        with self._lock:
            if job.status in {"queued", "processing"}:
                job.cancel_event.set()
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
        return self._pipeline is not None

    def close(self) -> None:
        self._executor.shutdown(wait=True, cancel_futures=True)
        if self._pipeline is not None:
            try:
                self._pipeline.close()
            except Exception:
                LOGGER.exception("关闭 PaddleOCR 模型失败")

    def _get_pipeline(self) -> Any:
        with self._pipeline_lock:
            if self._pipeline is None:
                LOGGER.info("加载模型：%s", CONFIG.engine)
                self._pipeline = build_pipeline(CONFIG.engine)
            return self._pipeline

    def _set_state(
        self,
        job: Job,
        *,
        status: str | None = None,
        progress: int | None = None,
        error: str | None = None,
    ) -> None:
        with self._lock:
            if status is not None:
                job.status = status
            if progress is not None:
                job.progress = progress
            if error is not None:
                job.error = error

    def _process(self, job: Job) -> None:
        task_started = time.perf_counter()
        LOGGER.info(
            "pdf_to_word_stage job_id=%s stage=started filename=%s page_count=%s",
            job.job_id,
            job.filename,
            job.page_count,
        )
        try:
            with self._lock:
                job.status = "processing"
                job.started_at = _utc_now()
                job.progress = 5

            model_started = time.perf_counter()
            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=model_loading_started engine=%s",
                job.job_id,
                CONFIG.engine,
            )
            pipeline = self._get_pipeline()
            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=model_loading_completed elapsed_sec=%.3f",
                job.job_id,
                time.perf_counter() - model_started,
            )

            results: list[dict[str, Any]] = []
            inference_started = time.perf_counter()
            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=inference_started input=%s",
                job.job_id,
                job.input_path.name,
            )
            for result in pipeline.predict_iter(str(job.input_path)):
                if job.cancel_event.is_set():
                    LOGGER.info(
                        "pdf_to_word_stage job_id=%s stage=cancelled result_count=%s",
                        job.job_id,
                        len(results),
                    )
                    self._set_state(job, status="cancelled", progress=0)
                    return
                results.append(result)
                progress = min(80, 10 + int(70 * len(results) / max(job.page_count, 1)))
                self._set_state(job, progress=progress)
                LOGGER.info(
                    "pdf_to_word_stage job_id=%s stage=inference_batch_completed result_count=%s progress=%s",
                    job.job_id,
                    len(results),
                    progress,
                )

            if job.cancel_event.is_set():
                LOGGER.info(
                    "pdf_to_word_stage job_id=%s stage=cancelled result_count=%s",
                    job.job_id,
                    len(results),
                )
                self._set_state(job, status="cancelled", progress=0)
                return

            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=inference_completed result_count=%s elapsed_sec=%.3f",
                job.job_id,
                len(results),
                time.perf_counter() - inference_started,
            )
            self._set_state(job, progress=90)
            export_started = time.perf_counter()
            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=docx_export_started result_count=%s",
                job.job_id,
                len(results),
            )
            def export_stage(stage: str, details: dict[str, Any]) -> None:
                detail_text = " ".join(
                    f"{key}={value}" for key, value in details.items()
                )
                LOGGER.info(
                    "pdf_to_word_stage job_id=%s stage=%s%s",
                    job.job_id,
                    stage,
                    f" {detail_text}" if detail_text else "",
                )

            export_results_to_docx(
                results,
                job.output_path,
                title=Path(job.filename).stem,
                source_pdf=job.input_path,
                mode=CONFIG.export_mode,
                image_max_pixels=CONFIG.page_image_max_pixels,
                image_jpeg_quality=CONFIG.page_image_jpeg_quality,
                stage_callback=export_stage,
            )
            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=docx_export_completed elapsed_sec=%.3f output_bytes=%s",
                job.job_id,
                time.perf_counter() - export_started,
                job.output_path.stat().st_size,
            )
            self._set_state(job, status="succeeded", progress=100)
        except Exception as error:
            LOGGER.exception("任务 %s 处理失败", job.job_id)
            LOGGER.error(
                "pdf_to_word_stage job_id=%s stage=failed elapsed_sec=%.3f error_type=%s",
                job.job_id,
                time.perf_counter() - task_started,
                type(error).__name__,
            )
            self._set_state(job, status="failed", progress=0, error=f"{type(error).__name__}: {error}")
        finally:
            with self._lock:
                job.finished_at = _utc_now()
            LOGGER.info(
                "pdf_to_word_stage job_id=%s stage=finished status=%s elapsed_sec=%.3f",
                job.job_id,
                job.status,
                time.perf_counter() - task_started,
            )


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
        "export_mode": CONFIG.export_mode,
        "model_loaded": MANAGER.model_loaded(),
        "max_upload_bytes": CONFIG.max_upload_bytes,
        "max_pages": CONFIG.max_pages,
        "page_image_max_pixels": CONFIG.page_image_max_pixels,
    }


@app.post("/api/pdf-to-word/jobs", dependencies=[Depends(_require_auth)])
async def create_job(file: UploadFile = File(...)) -> dict[str, Any]:
    MANAGER.cleanup_expired()
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
