---
title: DocxAIO-HFS
emoji: 📄
colorFrom: blue
colorTo: green
sdk: docker
app_port: 8000
---

# DocxAIO-HFS Space wrapper

This Space is a five-file deployment wrapper, not a copy of the product repository. Its Docker build clones the public product source from `BlueSkyXN/DocxAIO-HFS`, checks out the exact full commit recorded in `BUILD_SOURCE.txt`, and verifies that checkout before copying only the runtime files into the final image.

## Source provenance

- Product source of record: `https://github.com/BlueSkyXN/DocxAIO-HFS.git`
- Build source: the 40-character Git commit in `BUILD_SOURCE.txt`
- Docker build argument: `DOCXAIO_SOURCE_COMMIT`
- Runtime: Python 3.11 slim, CJK fonts, the product requirements, and the existing single-worker entrypoint

The wrapper intentionally does not contain `main.py`, `docx_allinone.py`, templates, static assets, or any local configuration and credentials.

## Health and smoke checks

The service exposes `GET /health`. A healthy response has `status: "healthy"` and `concurrency.configured_workers: 1`; the latter is required because the application uses an in-process semaphore and a local process lock.

`cloud/hfs/smoke-test.sh` waits for `/health`, creates a minimal table-only DOCX with Python standard library code, posts it to `/process`, and checks that the ZIP response contains `process.log` and a generated DOCX. It requires only Bash, curl, and Python 3.

From a product-repository checkout, test a locally exported wrapper bundle:

```bash
REPO_ROOT=/path/to/DocxAIO-HFS
"$REPO_ROOT/cloud/hfs/export_space_bundle.sh" /tmp/docxaio-hfs-space
docker build -t docxaio-hfs-space /tmp/docxaio-hfs-space
docker run --rm -d --name docxaio-hfs-space -p 8000:8000 docxaio-hfs-space
"$REPO_ROOT/cloud/hfs/smoke-test.sh" http://127.0.0.1:8000
```

The GitHub product repository owns source changes. `cloud/hfs/export_space_bundle.sh` is the only supported way to create the flat Space upload context.
