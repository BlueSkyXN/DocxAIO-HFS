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

本项目按 HFS v2.1 分类为 Preview。`hfs-dev.toml` 是 canonical primary profile，允许日常
Preview 变更直接更新当前 Space；`hfs-dev.candidate.toml` 仅用于高风险变更的可选隔离验证，
不是常规前置门禁。workflow 中保留的 `production` target 名称仅是兼容输入，实际选择的是
canonical preview profile。

candidate 与 production 使用 `hfs-dev.candidate.toml`、`hfs-dev.toml` 两个固定 profile。production profile 的 target 必须显式等于 canonical `BlueSkyXN/DocxAIO-HFS`；production 发布还会在上传紧前重新 fetch `origin/main`，并要求 workflow ref 为 `refs/heads/main`，checkout `HEAD`、`GITHUB_SHA`、导出使用的 source commit 与最新 `origin/main` 完全相等。candidate 保留从手动触发所选 Git ref 发布的既有语义。

Settings 必须从 manifest 声明的、Git ignored 的本地明文事实源执行 `diff → push → readback`，不能只在 Space 网页维护最终值。canonical 使用 `.env`：

```bash
python3 scripts/hf_space_sync.py diff --manifest hfs-dev.toml
python3 scripts/hf_space_sync.py push --manifest hfs-dev.toml
python3 scripts/hf_space_sync.py diff --manifest hfs-dev.toml
```

如确需 candidate，使用其独立的 `local/hfs-targets/candidate.env`，脚本会从 manifest 读取：

```bash
python3 scripts/hf_space_sync.py diff --manifest hfs-dev.candidate.toml
python3 scripts/hf_space_sync.py push --manifest hfs-dev.candidate.toml
python3 scripts/hf_space_sync.py diff --manifest hfs-dev.candidate.toml
```

本项目没有 Secret；Variable 必须按值读回。清理授权前不得使用 `--prune --yes`。

本地导出 wrapper：

```bash
cloud/hfs/export_space_bundle.sh /tmp/docxaio-hfs-space
```

导出目录严格只有 `README.md`、`Dockerfile`、`hfs-dev.toml`、`.dockerignore` 和 `BUILD_SOURCE.txt`。`BUILD_SOURCE.txt` 中的 SHA 会同时写入 Dockerfile 的 `DOCXAIO_SOURCE_COMMIT`，构建阶段 checkout 后断言 `HEAD` 完全一致。

- 本地验证：`scripts/static-check.sh`
- 容器 smoke：导出后 `docker build -t docxaio-hfs-space /tmp/docxaio-hfs-space`，启动容器并运行 `cloud/hfs/smoke-test.sh`
- 发布：仅使用 `.github/workflows/sync-to-hf-space.yml` 的手动 `workflow_dispatch`，并明确输入 `confirm=PUBLISH_WRAPPER`。canonical target、private visibility、thin-wrapper tree 和 production main provenance 等 gate 全部在首次 HF upload 前执行；该 workflow 只上传导出的 wrapper，随后以 CLI 下载全部五个文件逐字节读回核对。

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
