#!/usr/bin/env bash
# Smoke test a running DocxAIO-HFS container without test fixtures or third-party Python packages.
set -euo pipefail

base_url=${1:-http://127.0.0.1:8000}
base_url=${base_url%/}
work_dir=$(mktemp -d "${TMPDIR:-/tmp}/docxaio-hfs-smoke.XXXXXX")
trap 'rm -rf "$work_dir"' EXIT

input_docx="$work_dir/minimal-table.docx"
result_zip="$work_dir/result.zip"
concurrent_zip="$work_dir/result-concurrent.zip"
health_json="$work_dir/health.json"
smoke_timeout=${SMOKE_TIMEOUT_SECONDS:-120}
max_output_bytes=${SMOKE_MAX_OUTPUT_BYTES:-104857600}

python3 - "$input_docx" <<'PY'
from pathlib import Path
import sys
import zipfile

output = Path(sys.argv[1])
contents = {
    "[Content_Types].xml": """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>
</Types>""",
    "_rels/.rels": """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>
</Relationships>""",
    "word/document.xml": """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    <w:p><w:r><w:rPr><w:rFonts w:ascii="Noto Sans CJK SC" w:eastAsia="Noto Sans CJK SC"/></w:rPr><w:t>DocxAIO-HFS 中文字体回归</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>
      <w:tr><w:tc><w:p><w:r><w:t>Column A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Column B</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>
  </w:body>
</w:document>""",
}
with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
    for name, content in contents.items():
        archive.writestr(name, content)
PY

for ((attempt = 0; attempt < 60; attempt++)); do
    if curl --fail --silent --show-error "$base_url/health" > "$health_json"; then
        break
    fi
    sleep 1
done

python3 - "$health_json" <<'PY'
import json
from pathlib import Path
import sys

payload = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
if payload.get("status") != "healthy":
    raise SystemExit(f"Unhealthy response: {payload!r}")
if payload.get("concurrency", {}).get("configured_workers") != 1:
    raise SystemExit(f"Expected one configured worker: {payload!r}")
PY

run_conversion() {
    local output=$1
    local content_type_file=$2
    curl --fail --silent --show-error \
        --max-time "$smoke_timeout" \
        --output "$output" \
        --write-out '%{content_type}' \
        --form "file=@$input_docx;type=application/vnd.openxmlformats-officedocument.wordprocessingml.document" \
        --form 'a3=on' \
        "$base_url/process" > "$content_type_file"
}

run_conversion "$result_zip" "$work_dir/content-type-1" &
first_pid=$!
run_conversion "$concurrent_zip" "$work_dir/content-type-2" &
second_pid=$!
wait "$first_pid"
wait "$second_pid"

for content_type_file in "$work_dir/content-type-1" "$work_dir/content-type-2"; do
    content_type=$(<"$content_type_file")
    case "$content_type" in
        application/zip*) ;;
        *)
            printf 'Expected application/zip, got %s\n' "$content_type" >&2
            exit 1
            ;;
    esac
done

python3 - "$result_zip" "$concurrent_zip" "$max_output_bytes" <<'PY'
from pathlib import Path
import sys
import zipfile

limit = int(sys.argv[3])
for raw_path in sys.argv[1:3]:
    result = Path(raw_path)
    if not 0 < result.stat().st_size <= limit:
        raise SystemExit(f"ZIP response size is outside 1..{limit} bytes: {result.stat().st_size}")
    with zipfile.ZipFile(result) as archive:
        names = archive.namelist()
        if "process.log" not in names:
            raise SystemExit("ZIP response does not contain process.log")
        docx_names = [name for name in names if name.lower().endswith(".docx")]
        if not docx_names:
            raise SystemExit("ZIP response does not contain a DOCX output")
        docx_bytes = archive.read(docx_names[0])
    with zipfile.ZipFile(__import__("io").BytesIO(docx_bytes)) as document:
        xml = document.read("word/document.xml").decode("utf-8")
    if "中文字体回归" not in xml or "Noto Sans CJK SC" not in xml:
        raise SystemExit("DOCX output did not preserve the Chinese font regression fixture")
PY

printf 'Smoke test passed for %s\n' "$base_url"
