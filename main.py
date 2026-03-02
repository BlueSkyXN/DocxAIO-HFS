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
import io
import logging
import os
import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
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
BASE_DIR = Path(__file__).resolve().parent
TEMP_ROOT = Path(os.getenv("TEMP_DIR", tempfile.gettempdir())) / "docxaio-hfs"

processing_semaphore = asyncio.Semaphore(MAX_CONCURRENT_TASKS)

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
) -> SimpleNamespace:
    """构造 docx_allinone 处理参数对象。"""
    args = SimpleNamespace(
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


def run_processing(input_path: Path, args: SimpleNamespace) -> str:
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
