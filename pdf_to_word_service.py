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
from multiprocessing import Process
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
from job_store import JobStore


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
    max_retries: int = _env_int("PDF_SERVICE_MAX_RETRIES", 1)
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
        if not 0 <= self.max_retries <= 3:
            raise RuntimeError("环境变量 PDF_SERVICE_MAX_RETRIES 必须在 0 到 3 之间")


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
    attempt: int = 0
    max_retries: int = 0
    process: Process | None = field(default=None, repr=False, compare=False)
    worker_reported_status: str | None = field(
        default=None, repr=False, compare=False
    )

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
            "attempt": self.attempt,
            "max_retries": self.max_retries,
            "download_url": (
                f"/api/pdf-to-word/jobs/{self.job_id}/result"
                if self.status == "succeeded"
                else None
            ),
        }

    def as_record(self) -> dict[str, Any]:
        """返回可写入任务存储的标量字段。"""
        return {
            "job_id": self.job_id,
            "filename": self.filename,
            "workspace": str(self.workspace),
            "input_path": str(self.input_path),
            "output_path": str(self.output_path),
            "progress_path": str(self.progress_path),
            "stage_log_path": str(self.stage_log_path),
            "cancel_path": str(self.cancel_path),
            "page_count": self.page_count,
            "created_at": _iso(self.created_at),
            "started_at": _iso(self.started_at),
            "finished_at": _iso(self.finished_at),
            "status": self.status,
            "progress": self.progress,
            "route": self.route,
            "route_reason": self.route_reason,
            "worker_pid": self.worker_pid,
            "error": self.error,
            "attempt": self.attempt,
            "max_retries": self.max_retries,
        }


def _parse_datetime(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


class JobManager:
    """持久化任务状态，并监管可独立终止的转换 worker 进程。"""

    def __init__(self, store_path: Path | None = None) -> None:
        self._jobs: dict[str, Job] = {}
        self._processes: dict[str, Process] = {}
        self._lock = threading.Lock()
        context_name = "spawn" if os.name == "nt" else "fork"
        self._mp_context = multiprocessing.get_context(context_name)
        self._store = JobStore(store_path or CONFIG.data_root / "jobs.sqlite3")
        self._stop_event = threading.Event()
        self._closed = False
        self._load_persisted_jobs()
        self._scheduler_thread = threading.Thread(
            target=self._scheduler_loop,
            name="pdf-worker-supervisor",
            daemon=True,
        )
        self._scheduler_thread.start()

    def _load_persisted_jobs(self) -> None:
        for record in self._store.load_all():
            job = self._job_from_record(record)
            if job is None:
                continue
            if job.status in {"queued", "processing"}:
                if not job.input_path.is_file():
                    job.status = "failed"
                    job.progress = 0
                    job.finished_at = _utc_now()
                    job.error = "服务重启后找不到任务输入文件"
                elif job.attempt > job.max_retries:
                    job.status = "failed"
                    job.progress = 0
                    job.finished_at = _utc_now()
                    job.error = "服务重启时任务已耗尽重试次数"
                else:
                    job.status = "queued"
                    job.progress = 0
                    job.started_at = None
                    job.finished_at = None
                    job.worker_pid = None
                    job.error = "服务重启后任务重新排队"
                    job.cancel_path.unlink(missing_ok=True)
            self._jobs[job.job_id] = job
            self._store.save(job.as_record())

    @staticmethod
    def _job_from_record(record: dict[str, Any]) -> Job | None:
        try:
            return Job(
                job_id=str(record["job_id"]),
                filename=str(record["filename"]),
                workspace=Path(str(record["workspace"])),
                input_path=Path(str(record["input_path"])),
                output_path=Path(str(record["output_path"])),
                progress_path=Path(str(record["progress_path"])),
                stage_log_path=Path(str(record["stage_log_path"])),
                cancel_path=Path(str(record["cancel_path"])),
                page_count=int(record["page_count"]),
                created_at=_parse_datetime(record.get("created_at")) or _utc_now(),
                started_at=_parse_datetime(record.get("started_at")),
                finished_at=_parse_datetime(record.get("finished_at")),
                status=str(record.get("status") or "queued"),
                progress=int(record.get("progress") or 0),
                route=str(record.get("route") or "pending"),
                route_reason=record.get("route_reason"),
                worker_pid=(
                    int(record["worker_pid"])
                    if record.get("worker_pid") is not None
                    else None
                ),
                error=record.get("error"),
                attempt=int(record.get("attempt") or 0),
                max_retries=int(record.get("max_retries") or 0),
            )
        except (KeyError, TypeError, ValueError):
            LOGGER.exception("跳过无法读取的持久化任务记录")
            return None

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
        moved_input = False
        try:
            shutil.move(str(source_path), input_path)
            moved_input = True
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
                max_retries=CONFIG.max_retries,
            )
            with self._lock:
                if self._active_job_count_locked() >= CONFIG.max_pending_jobs:
                    raise QueueFullError("任务队列已满，请稍后重试")
                self._jobs[job_id] = job
                self._store.save(job.as_record())
            return job
        except QueueFullError:
            if moved_input and input_path.is_file() and not source_path.exists():
                shutil.move(str(input_path), source_path)
            with self._lock:
                self._jobs.pop(job_id, None)
            shutil.rmtree(workspace, ignore_errors=True)
            raise
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
            status = job.status
            if status == "queued":
                job.status = "cancelled"
                job.progress = 0
                job.finished_at = _utc_now()
                job.error = "任务已取消"
                self._store.save(job.as_record())
            elif status != "processing":
                return job
        if status == "queued":
            self._write_manager_state(job, "cancelled", "cancelled")
            return job

        job.cancel_path.touch(exist_ok=True)
        self._stop_process(job.job_id)
        with self._lock:
            job.status = "cancelled"
            job.progress = 0
            job.finished_at = _utc_now()
            job.worker_pid = None
            job.error = "任务已取消"
            self._store.save(job.as_record())
        self._write_manager_state(job, "cancelled", "cancelled")
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
        self._store.delete_many(job.job_id for job in expired)
        for job in expired:
            shutil.rmtree(job.workspace, ignore_errors=True)
        return len(expired)

    def model_loaded(self) -> bool:
        """返回 API 进程是否持有模型；模型实际驻留在 worker 进程中。"""
        return False

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._stop_event.set()
        self._scheduler_thread.join(timeout=5)
        with self._lock:
            jobs = [
                job for job in self._jobs.values() if job.status == "processing"
            ]
        for job in jobs:
            self._stop_process(job.job_id)
            with self._lock:
                if job.status == "processing":
                    job.status = "queued"
                    job.progress = 0
                    job.started_at = None
                    job.finished_at = None
                    job.worker_pid = None
                    job.error = "服务停止后任务等待恢复"
                    self._store.save(job.as_record())
            self._write_manager_state(job, "service_stopped", "queued")

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
            "attempt": job.attempt,
        }

    def _scheduler_loop(self) -> None:
        while not self._stop_event.is_set():
            with self._lock:
                jobs = list(self._jobs.values())
            for job in jobs:
                self._refresh_from_worker(job)
                self._monitor_process(job)

            with self._lock:
                available = CONFIG.worker_processes - len(self._processes)
                queued = sorted(
                    (
                        job
                        for job in self._jobs.values()
                        if job.status == "queued"
                    ),
                    key=lambda item: item.created_at,
                )
            for job in queued[: max(available, 0)]:
                self._start_process(job)
            self._stop_event.wait(0.1)

    def _start_process(self, job: Job) -> None:
        with self._lock:
            if job.status != "queued" or len(self._processes) >= CONFIG.worker_processes:
                return
            job.attempt += 1
            job.status = "processing"
            job.progress = 5
            job.started_at = _utc_now()
            job.finished_at = None
            job.worker_pid = None
            job.error = None
            job.route = "pending"
            job.route_reason = None
            job.worker_reported_status = None
            job.cancel_path.unlink(missing_ok=True)
            job.output_path.unlink(missing_ok=True)
            process = self._mp_context.Process(
                target=process_job,
                args=(self._worker_payload(job),),
                name=f"pdf-worker-{job.job_id[:8]}",
            )
            self._processes[job.job_id] = process
            self._store.save(job.as_record())
        try:
            process.start()
        except Exception as error:
            with self._lock:
                self._processes.pop(job.job_id, None)
            self._handle_attempt_failure(
                job,
                "failed",
                f"worker 启动失败：{type(error).__name__}: {error}",
            )
            return
        with self._lock:
            job.worker_pid = process.pid
            self._store.save(job.as_record())

    def _monitor_process(self, job: Job) -> None:
        with self._lock:
            process = self._processes.get(job.job_id)
            status = job.status
        if process is None:
            return
        if process.is_alive():
            if status in {"queued", "processing"} and self._deadline_exceeded(job):
                self._stop_process(job.job_id)
                self._handle_attempt_failure(
                    job,
                    "timed_out",
                    f"任务超过时间预算（attempt={job.attempt}）",
                )
            return

        process.join(timeout=0)
        with self._lock:
            self._processes.pop(job.job_id, None)
            job.process = None
            status = job.status
            error = job.error
            worker_reported_status = job.worker_reported_status
        if status == "succeeded" and job.output_path.is_file():
            with self._lock:
                job.finished_at = job.finished_at or _utc_now()
                self._store.save(job.as_record())
            return
        if status == "cancelled":
            return
        if worker_reported_status == "timed_out" or status == "timed_out":
            self._handle_attempt_failure(
                job,
                "timed_out",
                error or "worker 达到软时间预算",
            )
            return
        exit_code = process.exitcode
        self._handle_attempt_failure(
            job,
            "failed",
            error or f"worker 异常退出，exitcode={exit_code}",
        )

    def _deadline_exceeded(self, job: Job) -> bool:
        if job.started_at is None:
            return False
        budget = (
            CONFIG.ocr_time_budget_seconds
            if job.route == "ocr"
            else CONFIG.task_timeout_seconds
        )
        return (_utc_now() - job.started_at).total_seconds() > budget

    def _stop_process(self, job_id: str) -> None:
        with self._lock:
            process = self._processes.pop(job_id, None)
        if process is None:
            return
        if process.is_alive():
            process.terminate()
            process.join(timeout=5)
        if process.is_alive():
            kill = getattr(process, "kill", None)
            if kill is not None:
                kill()
                process.join(timeout=2)
        with self._lock:
            job = self._jobs.get(job_id)
            if job is not None:
                job.process = None

    def _handle_attempt_failure(
        self,
        job: Job,
        status: str,
        error: str,
    ) -> None:
        with self._lock:
            if job.status == "cancelled":
                return
            if job.attempt <= job.max_retries:
                job.status = "queued"
                job.progress = 0
                job.started_at = None
                job.finished_at = None
                job.worker_pid = None
                job.error = (
                    f"{error}；将在重试次数 {job.attempt}/{job.max_retries} 后重新执行"
                )
                job.route = "pending"
                job.route_reason = None
                job.output_path.unlink(missing_ok=True)
                self._store.save(job.as_record())
                retry = True
            else:
                job.status = status
                job.progress = 0
                job.finished_at = _utc_now()
                job.worker_pid = None
                job.error = error
                self._store.save(job.as_record())
                retry = False
        if retry:
            self._write_manager_state(job, "retry_scheduled", "queued")
        else:
            self._write_manager_state(job, "manager_finished", status)

    def _refresh_from_worker(self, job: Job) -> None:
        with self._lock:
            if job.status in TERMINAL_STATUSES:
                return
            if job.status == "queued" and job.job_id not in self._processes:
                return
        try:
            payload = json.loads(job.progress_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        with self._lock:
            if job.status in TERMINAL_STATUSES:
                return
            status = payload.get("status")
            if status in {"failed", "timed_out"}:
                job.worker_reported_status = status
            elif isinstance(status, str):
                job.status = status
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
            self._store.save(job.as_record())

    def _write_manager_state(self, job: Job, stage: str, status: str) -> None:
        event = {
            "stage": stage,
            "status": status,
            "progress": job.progress,
            "route": job.route,
            "route_reason": job.route_reason,
            "error": job.error,
            "worker_pid": job.worker_pid,
            "updated_at": _utc_now().isoformat(),
            "job_id": job.job_id,
            "attempt": job.attempt,
        }
        temporary_path = job.progress_path.with_name(
            f".{job.progress_path.name}.manager.tmp"
        )
        try:
            job.progress_path.parent.mkdir(parents=True, exist_ok=True)
            temporary_path.write_text(
                json.dumps(event, ensure_ascii=False), encoding="utf-8"
            )
            for attempt in range(10):
                try:
                    os.replace(temporary_path, job.progress_path)
                    break
                except PermissionError:
                    if attempt == 9:
                        raise
                    time.sleep(0.02 * (attempt + 1))
            with job.stage_log_path.open("a", encoding="utf-8") as stream:
                stream.write(json.dumps(event, ensure_ascii=False) + "\n")
        except OSError:
            LOGGER.exception("写入任务管理状态失败：%s", job.job_id)


MANAGER: JobManager | None = None


def _get_manager() -> JobManager:
    """返回当前服务实例的任务管理器。"""
    global MANAGER
    if MANAGER is None:
        MANAGER = JobManager()
    return MANAGER


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


async def _cleanup_loop(manager: JobManager) -> None:
    while True:
        await asyncio.sleep(CONFIG.cleanup_interval_seconds)
        removed = manager.cleanup_expired()
        if removed:
            LOGGER.info("清理过期任务：%s", removed)


@asynccontextmanager
async def lifespan(_: FastAPI):
    global MANAGER
    manager = _get_manager()
    cleanup_task = asyncio.create_task(_cleanup_loop(manager))
    try:
        yield
    finally:
        cleanup_task.cancel()
        try:
            await cleanup_task
        except asyncio.CancelledError:
            pass
        manager.close()
        if MANAGER is manager:
            MANAGER = None


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
    manager = _get_manager()
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
        "model_loaded": manager.model_loaded(),
        "max_upload_bytes": CONFIG.max_upload_bytes,
        "max_pages": CONFIG.max_pages,
        "page_image_max_pixels": CONFIG.page_image_max_pixels,
        "task_timeout_seconds": CONFIG.task_timeout_seconds,
        "ocr_time_budget_seconds": CONFIG.ocr_time_budget_seconds,
    }


@app.post("/api/pdf-to-word/jobs", dependencies=[Depends(_require_auth)])
async def create_job(file: UploadFile = File(...)) -> dict[str, Any]:
    manager = _get_manager()
    manager.cleanup_expired()
    if not manager.has_capacity():
        raise HTTPException(status_code=429, detail="任务队列已满，请稍后重试")
    staging_root = CONFIG.data_root / "staging"
    staging_root.mkdir(parents=True, exist_ok=True)
    staging_path = staging_root / f"{uuid.uuid4().hex}.upload"
    try:
        await _save_upload(file, staging_path)
        page_count = _read_page_count(staging_path)
        job = manager.create(_safe_filename(file.filename), staging_path, page_count)
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
    job = _get_manager().get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    return job.as_dict()


@app.get("/api/pdf-to-word/jobs/{job_id}/result", dependencies=[Depends(_require_auth)])
def download_result(job_id: str) -> FileResponse:
    job = _get_manager().get(job_id)
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
        job = _get_manager().cancel(job_id)
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
