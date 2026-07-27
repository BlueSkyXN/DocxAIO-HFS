# DocxAIO-HFS

基于 FastAPI + Docker 的 DOCX 处理服务，适用于 Hugging Face Spaces。

核心能力由 `docx_allinone.py` 提供，包括：

1. 嵌入 Excel 转 Word 表格 / 图片 / 提取 `.xlsx`
2. 文档水印移除（文本/图片/背景）
3. A3 横向页面布局
4. 表格提取（TXT/XLSX/PDF + 标记文档）
5. 图片分离（附图 PDF + 标记文档）

## 本地运行

```bash
docker build -t docxaio-hfs .
docker run --rm -p 8000:8000 docxaio-hfs
```

启动后访问：

- WebUI: `http://localhost:8000/`
- 健康检查: `http://localhost:8000/health`

## Hugging Face Spaces 部署

GitHub 产品仓是完整源代码和发布过程的事实源；Hugging Face Space 只接收由 `cloud/hfs/` 导出的五文件 flat wrapper，不保存 `main.py`、`docx_allinone.py`、模板或静态资源。Space 的 Docker 多阶段构建会从 GitHub 拉取并校验一个明确的完整 commit SHA。

candidate 与 production 使用 `hfs-dev.candidate.toml`、`hfs-dev.toml` 两个固定 profile。
Settings 从忽略的本地 `.env` 事实源执行 `diff → push → readback`，不在网页维护最终值：

```bash
python3 scripts/hf_space_sync.py diff --manifest hfs-dev.candidate.toml --env-file .env
python3 scripts/hf_space_sync.py push --manifest hfs-dev.candidate.toml --env-file .env
python3 scripts/hf_space_sync.py diff --manifest hfs-dev.candidate.toml --env-file .env
```

本项目没有 Secret；Variable 必须按值读回。清理授权前不得使用 `--prune --yes`。

本地导出 wrapper：

```bash
cloud/hfs/export_space_bundle.sh /tmp/docxaio-hfs-space
```

导出目录严格只有 `README.md`、`Dockerfile`、`hfs-dev.toml`、`.dockerignore` 和 `BUILD_SOURCE.txt`。`BUILD_SOURCE.txt` 中的 SHA 会同时写入 Dockerfile 的 `DOCXAIO_SOURCE_COMMIT`，构建阶段 checkout 后断言 `HEAD` 完全一致。

- 本地验证：`scripts/static-check.sh`
- 容器 smoke：导出后 `docker build -t docxaio-hfs-space /tmp/docxaio-hfs-space`，启动容器并运行 `cloud/hfs/smoke-test.sh`
- 发布：仅使用 `.github/workflows/sync-to-hf-space.yml` 的手动 `workflow_dispatch`，并明确输入 `confirm=yes`。该 workflow 只上传导出的 wrapper，随后以 CLI 下载 `Dockerfile` 和 `BUILD_SOURCE.txt` 逐字节读回核对。

## 环境变量

- `PORT`：服务端口（默认 `8000`）
- `TEMP_DIR`：临时目录根（默认 `/app/temp`）
- `MAX_FILE_SIZE_MB`：上传大小上限（默认 `120`）
- `REQUEST_TIMEOUT_SECONDS`：单请求超时秒数（默认 `1200`）
- `MAX_CONCURRENT_TASKS`：并发处理任务数（默认 `2`）
- `WORKERS`：Uvicorn worker 数（**必须为 `1`**，默认 `1`，否则服务将启动失败以避免并发失控）
- `PROCESS_LOCK_FILE`：单进程锁文件路径（默认 `${TEMP_DIR}/docxaio-hfs/process.lock`）

> 并发说明：当前使用进程内 semaphore 限流，`/health` 的 `concurrency` 字段会返回 `semaphore_limit`、`configured_workers`、`configured_workers_source`、`effective_max`、`single_process_lock_file` 便于排查配置。  
> 若手动启动 uvicorn，请显式使用 `--workers 1`（本项目仅支持单 worker + 本地 semaphore 组合）。
> worker 检测优先级为：显式 `--workers`（当前进程/父进程） > `WORKERS` > `WEB_CONCURRENCY`；空值环境变量会按“未设置”处理。

## 输出说明

上传一个 DOCX 后，服务会执行所选模式，并将所有输出文件打包为一个 ZIP 返回下载，ZIP 中附带 `process.log` 便于排查。
