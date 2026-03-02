---
title: DocxAIO-HFS
emoji: 📄
colorFrom: blue
colorTo: green
sdk: docker
app_port: 8000
pinned: false
---

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

1. 新建 Space，选择 `Docker` SDK。
2. 推送本目录文件到 Space 仓库。
3. Space 自动构建镜像并启动服务。

## 环境变量

- `PORT`：服务端口（默认 `8000`）
- `TEMP_DIR`：临时目录根（默认 `/app/temp`）
- `MAX_FILE_SIZE_MB`：上传大小上限（默认 `120`）
- `REQUEST_TIMEOUT_SECONDS`：单请求超时秒数（默认 `1200`）
- `MAX_CONCURRENT_TASKS`：并发处理任务数（默认 `2`）

## 输出说明

上传一个 DOCX 后，服务会执行所选模式，并将所有输出文件打包为一个 ZIP 返回下载，ZIP 中附带 `process.log` 便于排查。
