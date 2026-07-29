#!/usr/bin/env bash
# Validate the source-wrapper contract without network access or local credentials.
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
root_dir=$(CDPATH= cd -- "$script_dir/.." && pwd)
failures=0

fail() {
    printf 'FAIL: %s\n' "$*" >&2
    failures=$((failures + 1))
}

require_file() {
    [ -f "$1" ] || fail "missing ${1#$root_dir/}"
}

require_grep() {
    local pattern=$1
    local file=$2
    grep -Eq -- "$pattern" "$file" || fail "${file#$root_dir/} does not match ${pattern}"
}

for file in \
    "$root_dir/hfs-dev.toml" \
    "$root_dir/hfs-dev.candidate.toml" \
    "$root_dir/.env.example" \
    "$root_dir/cloud/hfs/README.md" \
    "$root_dir/cloud/hfs/Dockerfile.template" \
    "$root_dir/cloud/hfs/.dockerignore" \
    "$root_dir/cloud/hfs/export_space_bundle.sh" \
    "$root_dir/cloud/hfs/smoke-test.sh" \
    "$root_dir/.github/workflows/sync-to-hf-space.yml" \
    "$root_dir/.github/workflows/hfs-verify.yml"; do
    require_file "$file"
done
require_file "$root_dir/scripts/hf_space_sync.py"

python3 - "$root_dir/hfs-dev.toml" "$root_dir/.env.example" <<'PY' || failures=$((failures + 1))
import re
import sys
import tomllib
from pathlib import Path

manifest_path = Path(sys.argv[1])
env_example_path = Path(sys.argv[2])
manifest = tomllib.loads(manifest_path.read_text(encoding="utf-8"))
expected = {
    "standard": "2.0",
    "project": "docxaio-hfs",
    "space": "BlueSkyXN/DocxAIO-HFS",
    "sovereignty": "sovereign",
    "lane": "source",
    "version_source": "commit",
}
for key, value in expected.items():
    if manifest.get(key) != value:
        raise SystemExit(f"hfs-dev.toml {key} must be {value!r}")
if manifest.get("local_only") != ["HF_TOKEN", "GH_TOKEN"]:
    raise SystemExit("hfs-dev.toml local_only must contain only HF_TOKEN and GH_TOKEN")
if manifest.get("secrets") != []:
    raise SystemExit("hfs-dev.toml secrets must be an empty list")
variables = {
    "PORT", "TEMP_DIR", "MAX_FILE_SIZE_MB", "REQUEST_TIMEOUT_SECONDS",
    "MAX_CONCURRENT_TASKS", "WORKERS", "PROCESS_LOCK_FILE",
}
if set(manifest.get("variables", [])) != variables:
    raise SystemExit("hfs-dev.toml variables do not match the public runtime contract")

values = {}
for line in env_example_path.read_text(encoding="utf-8").splitlines():
    line = line.strip()
    if not line or line.startswith("#"):
        continue
    key, separator, value = line.partition("=")
    if not separator or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key):
        raise SystemExit(".env.example must contain only KEY=VALUE entries")
    values[key] = value
if not {"HF_TOKEN", "GH_TOKEN"}.issubset(values):
    raise SystemExit(".env.example must include local control-plane placeholders")
if any(re.search(r"(?:hf_[A-Za-z0-9]{20,}|ghp_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9]{20,})", value) for value in values.values()):
    raise SystemExit(".env.example appears to contain a credential")
if not variables.issubset(values):
    raise SystemExit(".env.example must document every public runtime variable")
PY

python3 - "$root_dir/hfs-dev.toml" "$root_dir/hfs-dev.candidate.toml" <<'PY' || failures=$((failures + 1))
import sys
import tomllib
from pathlib import Path

production = tomllib.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
candidate = tomllib.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
if candidate.get("space") != "BlueSkyXN/DocxAIO-HFS-v2-candidate":
    raise SystemExit("candidate manifest has the wrong fixed Space id")
for key in sorted(set(production) | set(candidate)):
    if key != "space" and production.get(key) != candidate.get(key):
        raise SystemExit(f"candidate manifest differs from production at {key}")
PY

require_grep '^ARG DOCXAIO_SOURCE_COMMIT=__DOCXAIO_SOURCE_COMMIT__$' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep 'DOCXAIO_SOURCE_COMMIT must equal the checked-out wrapper commit' "$root_dir/cloud/hfs/export_space_bundle.sh"
require_grep 'Refusing to export uncommitted or untracked wrapper inputs' "$root_dir/cloud/hfs/export_space_bundle.sh"
require_grep '^FROM python:3\.11-slim AS source$' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep 'ca-certificates' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep 'git clone https://github\.com/BlueSkyXN/DocxAIO-HFS\.git' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep 'git fetch --depth=1 origin' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep 'git checkout --detach' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep 'git rev-parse HEAD' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep '^COPY --from=source /src/docxaio-hfs/main\.py \./$' "$root_dir/cloud/hfs/Dockerfile.template"
require_grep '^EXPOSE 8000$' "$root_dir/cloud/hfs/Dockerfile.template"
if grep -Eq '^COPY (main\.py|docx_allinone\.py|entrypoint\.sh|templates|static)' "$root_dir/cloud/hfs/Dockerfile.template"; then
    fail 'wrapper Dockerfile must not COPY product source from the Space context'
fi

require_grep '^ENV PORT=8000' "$root_dir/Dockerfile"
require_grep '^EXPOSE 8000$' "$root_dir/Dockerfile"
require_grep '--workers "\$\{WORKERS:-1\}"' "$root_dir/entrypoint.sh"
require_grep 'CONFIGURED_WORKERS > 1' "$root_dir/main.py"
require_grep 'enforce_single_process\(\)' "$root_dir/main.py"
require_grep '@app\.get\("/health"' "$root_dir/main.py"
require_grep 'configured_workers' "$root_dir/main.py"
require_grep '/health' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep '/process' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep 'application/zip' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep 'process\.log' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep '中文字体回归' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep 'SMOKE_TIMEOUT_SECONDS' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep 'SMOKE_MAX_OUTPUT_BYTES' "$root_dir/cloud/hfs/smoke-test.sh"
require_grep 'second_pid' "$root_dir/cloud/hfs/smoke-test.sh"

sync_workflow="$root_dir/.github/workflows/sync-to-hf-space.yml"
verify_workflow="$root_dir/.github/workflows/hfs-verify.yml"
require_grep 'workflow_dispatch:' "$sync_workflow"
require_grep 'target:' "$sync_workflow"
require_grep 'HFS_MANIFEST:' "$sync_workflow"
require_grep 'confirm:' "$sync_workflow"
require_grep "if: inputs\.confirm == 'PUBLISH_WRAPPER'" "$sync_workflow"
require_grep 'DOCXAIO_SOURCE_COMMIT: \$\{\{ github\.sha \}\}' "$sync_workflow"
require_grep 'huggingface_hub\.cli\.hf upload' "$sync_workflow"
require_grep 'huggingface_hub\.cli\.hf download' "$sync_workflow"
require_grep 'huggingface_hub==1\.5\.0' "$sync_workflow"
require_grep 'click==8\.3\.3' "$sync_workflow"
require_grep 'python -m huggingface_hub\.cli\.hf version' "$sync_workflow"
require_grep 'python -m huggingface_hub\.cli\.hf --help' "$sync_workflow"
require_grep 'python -m huggingface_hub\.cli\.hf upload --help' "$sync_workflow"
require_grep 'python -m huggingface_hub\.cli\.hf download --help' "$sync_workflow"
require_grep 'cmp "\$BUNDLE_DIR/\$file" "\$READBACK_DIR/\$file"' "$sync_workflow"
require_grep 'candidate Space must be private' "$sync_workflow"
require_grep 'Space readback contains extra product files' "$sync_workflow"
if grep -Eq 'git remote|git push|--force|--delete' "$sync_workflow"; then
    fail 'sync workflow must not use Git remotes, pushes, or force-pushes'
fi
require_grep 'hf_space_sync\.py diff' "$root_dir/README.md"
require_grep 'hf_space_sync\.py push' "$root_dir/README.md"
require_grep 'pull_request:' "$verify_workflow"
require_grep 'docker build' "$verify_workflow"
require_grep 'cloud/hfs/smoke-test\.sh' "$verify_workflow"
if grep -q 'secrets\.' "$verify_workflow"; then
    fail 'verification workflow must not require secrets'
fi

for ignored in '.env' '.env.*' '!.env.example' 'config.toml' 'local/'; do
    grep -Fxq -- "$ignored" "$root_dir/.gitignore" || fail ".gitignore misses $ignored"
    grep -Fxq -- "$ignored" "$root_dir/.dockerignore" || fail ".dockerignore misses $ignored"
done

source_commit=$(git -C "$root_dir" rev-parse --verify HEAD^{commit}) || fail 'cannot resolve HEAD commit'
case "$source_commit" in
    [0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f]) ;;
    *) fail 'HEAD is not a full lowercase commit SHA' ;;
esac

bundle_root=$(mktemp -d "${TMPDIR:-/tmp}/docxaio-hfs-contract.XXXXXX")
trap 'rm -rf "$bundle_root"' EXIT
if DOCXAIO_SOURCE_COMMIT=HEAD "$root_dir/cloud/hfs/export_space_bundle.sh" "$bundle_root/invalid" >/dev/null 2>&1; then
    fail 'exporter accepted a non-full source commit'
fi
[ ! -e "$bundle_root/invalid" ] || fail 'exporter wrote a bundle after rejecting a non-full source commit'
"$root_dir/cloud/hfs/export_space_bundle.sh" "$bundle_root/space"

expected_names='.dockerignore BUILD_SOURCE.txt Dockerfile README.md hfs-dev.toml'
actual_names=$(find "$bundle_root/space" -mindepth 1 -maxdepth 1 -type f -exec basename {} \; | LC_ALL=C sort | tr '\n' ' ' | sed 's/ $//')
[ "$actual_names" = "$expected_names" ] || fail "bundle files are not the five-file allowlist: $actual_names"
[ "$(<"$bundle_root/space/BUILD_SOURCE.txt")" = "$source_commit" ] || fail 'BUILD_SOURCE.txt does not match HEAD'
require_grep "^ARG DOCXAIO_SOURCE_COMMIT=$source_commit$" "$bundle_root/space/Dockerfile"
HFS_MANIFEST=hfs-dev.candidate.toml "$root_dir/cloud/hfs/export_space_bundle.sh" "$bundle_root/candidate-space"
cmp "$root_dir/hfs-dev.candidate.toml" "$bundle_root/candidate-space/hfs-dev.toml" || fail 'candidate export used the wrong manifest'
if find "$bundle_root/space" -mindepth 1 -maxdepth 1 \( -name 'main.py' -o -name 'docx_allinone.py' -o -name 'entrypoint.sh' -o -name 'templates' -o -name 'static' \) -print -quit | grep -q .; then
    fail 'bundle contains product source outside the wrapper boundary'
fi

if [ "$failures" -ne 0 ]; then
    printf 'HFS contract validation failed with %s issue(s).\n' "$failures" >&2
    exit 1
fi
printf 'HFS contract validation passed for %s.\n' "$source_commit"
