#!/usr/bin/env bash
# Run repository-local static checks with no network, credentials, or absolute paths.
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
root_dir=$(CDPATH= cd -- "$script_dir/.." && pwd)

for script in \
    "$root_dir/scripts/validate-hfs-contract.sh" \
    "$root_dir/scripts/static-check.sh" \
    "$root_dir/cloud/hfs/export_space_bundle.sh" \
    "$root_dir/cloud/hfs/smoke-test.sh"; do
    bash -n "$script"
done

python3 - "$root_dir" <<'PY'
import ast
import sys
from pathlib import Path

root = Path(sys.argv[1])
for path in (root / "main.py", root / "docx_allinone.py", root / "scripts/hf_space_sync.py"):
    ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
PY

PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s "$root_dir/scripts" -p 'test_*.py'

exec "$root_dir/scripts/validate-hfs-contract.sh"
