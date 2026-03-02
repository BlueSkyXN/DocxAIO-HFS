#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DocxAIO-HFS Web 服务入口。

提供:
1. WebUI 上传 DOCX 并选择处理模式
2. 调用 docx_allinone 核心逻辑处理文档
3. 自动打包所有输出文件为 ZIP 下载
"""

from __future__ import annotations

import asyncio
import contextlib
import fcntl
import io
import logging
import os
import shlex
import shutil
import subprocess
import sys
import tempfile
import uuid
import zipfile
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import docx_allinone as aio


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


APP_TITLE = "DocxAIO-HFS"
MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "120"))
REQUEST_TIMEOUT_SECONDS = int(os.getenv("REQUEST_TIMEOUT_SECONDS", "1200"))
MAX_CONCURRENT_TASKS = int(os.getenv("MAX_CONCURRENT_TASKS", "2"))


def _parse_workers_value(raw: str, source: str) -> int:
    try:
        workers = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"{source} must be a positive integer.") from exc
    if workers < 1:
        raise RuntimeError(f"{source} must be a positive integer.")
    return workers


def _extract_workers_from_tokens(tokens: list[str]) -> Optional[int]:
    for index, token in enumerate(tokens):
        if token == "--workers" and index + 1 < len(tokens):
            return _parse_workers_value(tokens[index + 1], "workers flag")
        if token.startswith("--workers="):
            return _parse_workers_value(token.split("=", 1)[1], "workers flag")
    return None


def _read_non_empty_env(name: str) -> Optional[str]:
    raw = os.getenv(name)
    if raw is None:
        return None
    value = raw.strip()
    if not value:
        return None
    return value


def detect_configured_workers() -> tuple[int, str]:
    argv_workers = _extract_workers_from_tokens(sys.argv)
    if argv_workers is not None:
        return argv_workers, "argv:--workers"

    try:
        parent_cmdline = subprocess.check_output(
            ["ps", "-o", "command=", "-p", str(os.getppid())],
            text=True,
        ).strip()
    except Exception:
        parent_cmdline = ""

    if parent_cmdline:
        try:
            parent_tokens = shlex.split(parent_cmdline)
        except ValueError:
            parent_tokens = []
        parent_workers = _extract_workers_from_tokens(parent_tokens)
        if parent_workers is not None:
            return parent_workers, "parent_cmd:--workers"

    workers_env = _read_non_empty_env("WORKERS")
    if workers_env is not None:
        return _parse_workers_value(workers_env, "WORKERS"), "env:WORKERS"

    web_concurrency = _read_non_empty_env("WEB_CONCURRENCY")
    if web_concurrency is not None:
        return _parse_workers_value(web_concurrency, "WEB_CONCURRENCY"), "env:WEB_CONCURRENCY"

    return 1, "default"


CONFIGURED_WORKERS, CONFIGURED_WORKERS_SOURCE = detect_configured_workers()
if CONFIGURED_WORKERS > 1:
    raise RuntimeError(
        f"Detected workers={CONFIGURED_WORKERS} ({CONFIGURED_WORKERS_SOURCE}), "
        f"which is not supported with local semaphore limiting. "
        f"Effective max concurrency would be {CONFIGURED_WORKERS} x {MAX_CONCURRENT_TASKS} = {CONFIGURED_WORKERS * MAX_CONCURRENT_TASKS}. "
        "Please run with a single worker."
    )
BASE_DIR = Path(__file__).resolve().parent
TEMP_ROOT = Path(os.getenv("TEMP_DIR", tempfile.gettempdir())) / "docxaio-hfs"
PROCESS_LOCK_FILE = Path(os.getenv("PROCESS_LOCK_FILE", str(TEMP_ROOT / "process.lock")))
process_lock_handle = None


def enforce_single_process() -> None:
    """通过文件锁确保仅有一个应用进程运行。"""
    global process_lock_handle
    PROCESS_LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
    lock_flags = os.O_RDWR | os.O_CREAT
    if hasattr(os, "O_NOFOLLOW"):
        lock_flags |= os.O_NOFOLLOW
    try:
        lock_fd = os.open(PROCESS_LOCK_FILE, lock_flags, 0o600)
    except OSError as exc:
        raise RuntimeError(f"Failed to open process lock file: {PROCESS_LOCK_FILE}") from exc
    lock_handle = os.fdopen(lock_fd, "r+", encoding="utf-8")
    try:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError as exc:
        lock_handle.close()
        raise RuntimeError(
            "Detected multiple app processes. "
            "This service requires a single worker process to keep concurrency limits safe."
        ) from exc
    process_lock_handle = lock_handle

processing_semaphore = asyncio.Semaphore(MAX_CONCURRENT_TASKS)
enforce_single_process()

app = FastAPI(
    title=APP_TITLE,
    description="DOCX All-in-One Web API for Hugging Face Spaces",
    version="1.0.0",
)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

TEMP_ROOT.mkdir(parents=True, exist_ok=True)


def parse_checkbox(value: Optional[str]) -> bool:
    """将 HTML checkbox 值转换为布尔值。"""
    if value is None:
        return False
    return value.lower() in {"on", "true", "1", "yes"}


def cleanup_temp_dir(temp_dir: Path) -> None:
    """删除会话临时目录。"""
    try:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
    except Exception as exc:
        logger.error("清理临时目录失败: %s", exc, exc_info=True)


async def get_upload_file_size(upload_file: UploadFile) -> int:
    """获取上传文件大小（字节）。"""
    current_position = upload_file.file.tell()
    upload_file.file.seek(0, 2)
    size = upload_file.file.tell()
    upload_file.file.seek(current_position)
    return size


@dataclass
class ProcessRequest:
    """docx_allinone 参数对象。"""

    input_path: str = ""
    word_table: bool = False
    extract_excel: bool = False
    image: bool = False
    keep_attachment: bool = False
    remove_watermark: bool = False
    a3: bool = False
    table_extract: bool = False
    split_images: bool = False
    split_remove_images: bool = False
    split_output_dir: Optional[str] = None
    split_no_optimize: bool = False
    split_jpeg_quality: int = 85
    workers: int = 1


def build_args(
    *,
    word_table: bool,
    extract_excel: bool,
    image: bool,
    keep_attachment: bool,
    remove_watermark: bool,
    a3: bool,
    table_extract: bool,
    split_images: bool,
    split_remove_images: bool,
    split_no_optimize: bool,
    split_jpeg_quality: int,
) -> ProcessRequest:
    """构造 docx_allinone 处理参数对象。"""
    args = ProcessRequest(
        input_path="",  # 占位字段，核心逻辑不依赖该值
        word_table=word_table,
        extract_excel=extract_excel,
        image=image,
        keep_attachment=keep_attachment,
        remove_watermark=remove_watermark,
        a3=a3,
        table_extract=table_extract,
        split_images=split_images,
        split_remove_images=split_remove_images,
        split_output_dir=None,
        split_no_optimize=split_no_optimize,
        split_jpeg_quality=split_jpeg_quality,
        workers=1,
    )

    # 与 CLI 行为保持一致：未选择任何模式时默认 --word-table
    has_any_mode = any(
        [
            args.word_table,
            args.extract_excel,
            args.image,
            args.remove_watermark,
            args.a3,
            args.table_extract,
            args.split_images,
        ]
    )
    if not has_any_mode:
        args.word_table = True
    return args


def run_processing(input_path: Path, args: ProcessRequest) -> str:
    """执行 docx_allinone 核心处理并返回日志。"""
    output_buffer = io.StringIO()
    with contextlib.redirect_stdout(output_buffer), contextlib.redirect_stderr(output_buffer):
        aio.process_document_with_extensions(str(input_path), args)
    return output_buffer.getvalue()


def build_result_zip(
    *,
    session_dir: Path,
    input_path: Path,
    output_files: list[Path],
    log_path: Path,
) -> Path:
    """将输出文件打包为 ZIP。"""
    zip_name = f"{input_path.stem}-outputs.zip"
    zip_path = session_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in output_files:
            zf.write(file_path, arcname=file_path.name)
        zf.write(log_path, arcname="process.log")
    return zip_path


@app.get("/", summary="Web UI")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/health", summary="Health Check")
async def health_check():
    """健康检查: 依赖导入 + 临时目录空间。"""
    try:
        import matplotlib  # noqa: F401
        import openpyxl  # noqa: F401
        import reportlab  # noqa: F401
        from PIL import Image  # noqa: F401

        disk_info = os.statvfs(str(TEMP_ROOT))
        free_space_mb = (disk_info.f_bavail * disk_info.f_frsize) / (1024 * 1024)
        return {
            "status": "healthy",
            "service": APP_TITLE,
            "temp_root": str(TEMP_ROOT),
            "free_space_mb": round(free_space_mb, 2),
            "limits": {
                "max_file_size_mb": MAX_FILE_SIZE_MB,
                "timeout_seconds": REQUEST_TIMEOUT_SECONDS,
                "max_concurrent_tasks": MAX_CONCURRENT_TASKS,
            },
            "concurrency": {
                "semaphore_limit": MAX_CONCURRENT_TASKS,
                "configured_workers": CONFIGURED_WORKERS,
                "configured_workers_source": CONFIGURED_WORKERS_SOURCE,
                "effective_max": MAX_CONCURRENT_TASKS * CONFIGURED_WORKERS,
                "warning": "using local semaphore, safe only with a single worker process",
                "single_process_lock_file": str(PROCESS_LOCK_FILE),
            },
        }
    except Exception as exc:
        return JSONResponse(
            status_code=500,
            content={"status": "unhealthy", "error": str(exc)},
        )


@app.post("/process", response_class=FileResponse, summary="处理 DOCX 并下载 ZIP")
async def process_docx(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="待处理 DOCX 文件"),
    word_table: Optional[str] = Form(None),
    extract_excel: Optional[str] = Form(None),
    image: Optional[str] = Form(None),
    keep_attachment: Optional[str] = Form(None),
    remove_watermark: Optional[str] = Form(None),
    a3: Optional[str] = Form(None),
    table_extract: Optional[str] = Form(None),
    split_images: Optional[str] = Form(None),
    split_remove_images: Optional[str] = Form(None),
    split_no_optimize: Optional[str] = Form(None),
    split_jpeg_quality: int = Form(85),
):
    """上传 DOCX，调用 AIO 处理逻辑并返回 ZIP。"""
    session_dir = TEMP_ROOT / str(uuid.uuid4())
    session_dir.mkdir(parents=True, exist_ok=True)
    cleanup_by_finally = True

    try:
        if not file.filename:
            raise HTTPException(status_code=400, detail="Filename is required.")
        if not file.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="Only .docx files are supported.")
        if not (1 <= split_jpeg_quality <= 100):
            raise HTTPException(status_code=422, detail="split_jpeg_quality must be between 1 and 100.")

        file_size_mb = await get_upload_file_size(file) / (1024 * 1024)
        if file_size_mb > MAX_FILE_SIZE_MB:
            raise HTTPException(
                status_code=400,
                detail=f"File too large. Max size is {MAX_FILE_SIZE_MB}MB, got {file_size_mb:.2f}MB.",
            )

        input_path = session_dir / file.filename
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        args = build_args(
            word_table=parse_checkbox(word_table),
            extract_excel=parse_checkbox(extract_excel),
            image=parse_checkbox(image),
            keep_attachment=parse_checkbox(keep_attachment),
            remove_watermark=parse_checkbox(remove_watermark),
            a3=parse_checkbox(a3),
            table_extract=parse_checkbox(table_extract),
            split_images=parse_checkbox(split_images),
            split_remove_images=parse_checkbox(split_remove_images),
            split_no_optimize=parse_checkbox(split_no_optimize),
            split_jpeg_quality=split_jpeg_quality,
        )

        before_files = {p.resolve() for p in session_dir.iterdir() if p.is_file()}

        try:
            async with processing_semaphore:
                logger.info("开始处理: %s", input_path.name)
                logs = await asyncio.wait_for(
                    asyncio.to_thread(run_processing, input_path, args),
                    timeout=REQUEST_TIMEOUT_SECONDS,
                )
        except asyncio.TimeoutError as exc:
            raise HTTPException(
                status_code=504,
                detail=f"Processing timed out after {REQUEST_TIMEOUT_SECONDS} seconds.",
            ) from exc
        except aio.DocumentSkipError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        log_path = session_dir / "process.log"
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"timestamp={datetime.utcnow().isoformat()}Z\n")
            f.write(f"input={input_path.name}\n")
            f.write("----- logs -----\n")
            f.write(logs)
            if not logs.endswith("\n"):
                f.write("\n")

        after_files = {p.resolve() for p in session_dir.iterdir() if p.is_file()}
        output_files = sorted(
            [
                p
                for p in after_files - before_files
                if p.name not in {input_path.name, "process.log"} and not p.name.endswith(".zip")
            ],
            key=lambda x: x.name,
        )

        if not output_files:
            # 为了便于排查，直接返回日志中前几行
            short_log = logs.strip().splitlines()[-20:]
            raise HTTPException(
                status_code=400,
                detail="No output files generated.\n" + "\n".join(short_log),
            )

        zip_path = build_result_zip(
            session_dir=session_dir,
            input_path=input_path,
            output_files=output_files,
            log_path=log_path,
        )

        background_tasks.add_task(cleanup_temp_dir, session_dir)
        cleanup_by_finally = False

        return FileResponse(
            path=zip_path,
            media_type="application/zip",
            filename=zip_path.name,
            background=background_tasks,
        )

    except HTTPException:
        raise
    except Exception as exc:
        logger.error("处理失败: %s", exc, exc_info=True)
        raise HTTPException(status_code=500, detail=f"Unexpected server error: {exc}") from exc
    finally:
        await file.close()
        if cleanup_by_finally:
            cleanup_temp_dir(session_dir)
