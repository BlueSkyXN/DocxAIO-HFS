#!/usr/bin/env bash
# Export the only files that may be uploaded to the Hugging Face Space.
set -euo pipefail

usage() {
    printf 'Usage: %s OUTPUT_DIR\n' "${0##*/}" >&2
}

if [ "$#" -ne 1 ]; then
    usage
    exit 2
fi

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
root_dir=$(CDPATH= cd -- "$script_dir/../.." && pwd)
output_dir=$1
manifest_file=${HFS_MANIFEST:-$root_dir/hfs-dev.toml}
case "$manifest_file" in
    /*) ;;
    *) manifest_file="$root_dir/$manifest_file" ;;
esac

source_candidate=${DOCXAIO_SOURCE_COMMIT:-}
if [ -z "$source_candidate" ]; then
    source_candidate=$(git -C "$root_dir" rev-parse --verify HEAD^{commit}) || {
        printf '%s\n' 'Unable to resolve the current Git commit.' >&2
        exit 1
    }
elif [[ ! "$source_candidate" =~ ^[0-9a-fA-F]{40}$ ]]; then
    printf '%s\n' 'DOCXAIO_SOURCE_COMMIT must be a full 40-character commit SHA.' >&2
    exit 1
fi

source_commit=$(git -C "$root_dir" rev-parse --verify "${source_candidate}^{commit}") || {
    printf '%s\n' 'DOCXAIO_SOURCE_COMMIT must resolve to a Git commit.' >&2
    exit 1
}
if [[ ! "$source_commit" =~ ^[0-9a-f]{40}$ ]]; then
    printf '%s\n' 'The resolved source commit must be a full 40-character lowercase SHA.' >&2
    exit 1
fi

for required in README.md .dockerignore Dockerfile.template; do
    if [ ! -f "$script_dir/$required" ]; then
        printf 'Missing wrapper input: %s\n' "$required" >&2
        exit 1
    fi
done
if [ ! -f "$manifest_file" ]; then
    printf 'Missing selected HFS manifest: %s\n' "$manifest_file" >&2
    exit 1
fi

if [ -e "$output_dir" ] && [ ! -d "$output_dir" ]; then
    printf 'Output path is not a directory: %s\n' "$output_dir" >&2
    exit 1
fi
mkdir -p "$output_dir"
if [ -n "$(find "$output_dir" -mindepth 1 -maxdepth 1 -print -quit)" ]; then
    printf 'Output directory must be empty: %s\n' "$output_dir" >&2
    exit 1
fi

python3 - "$script_dir/Dockerfile.template" "$output_dir/Dockerfile" "$source_commit" <<'PY'
from pathlib import Path
import sys

template_path = Path(sys.argv[1])
output_path = Path(sys.argv[2])
source_commit = sys.argv[3]
text = template_path.read_text(encoding="utf-8")
placeholder = "__DOCXAIO_SOURCE_COMMIT__"
if text.count(placeholder) != 1:
    raise SystemExit("Dockerfile template must contain the source SHA placeholder exactly once.")
Path(output_path).write_text(text.replace(placeholder, source_commit), encoding="utf-8")
PY

cp "$script_dir/README.md" "$output_dir/README.md"
cp "$script_dir/.dockerignore" "$output_dir/.dockerignore"
cp "$manifest_file" "$output_dir/hfs-dev.toml"
printf '%s\n' "$source_commit" > "$output_dir/BUILD_SOURCE.txt"

expected_names='.dockerignore BUILD_SOURCE.txt Dockerfile README.md hfs-dev.toml'
actual_names=$(find "$output_dir" -mindepth 1 -maxdepth 1 -type f -exec basename {} \; | LC_ALL=C sort | tr '\n' ' ' | sed 's/ $//')
if [ "$actual_names" != "$expected_names" ]; then
    printf '%s\n' 'Bundle allowlist check failed.' >&2
    exit 1
fi

printf 'Exported HFS wrapper for %s to %s\n' "$source_commit" "$output_dir"
