#!/usr/bin/env python3
"""HFS v2 示例同步脚本。

命令：
  diff   比较本地登记、Space 设置、种子和实例配置；有差异返回 1
  push   从本地 env 文件推送已登记设置，并更新和读回种子
  pull   将实例配置回收到 local/hfs-sync-pulled/，绝不覆盖根种子

依赖：Python 3.11+、huggingface_hub==1.5.0、click==8.3.3
（后者是本脚本调用的 module HF CLI 的直接运行依赖）；
仅处理 YAML seed 时需要 PyYAML>=6.0。
脚本不会打印 secret 值；HF_TOKEN/GH_TOKEN 只作为本地控制凭据，不推 Space。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import stat
import subprocess
import sys
import tempfile
import time
import tomllib
import urllib.parse
import urllib.request
from pathlib import Path, PurePosixPath
from typing import Any

from huggingface_hub import HfApi
from huggingface_hub.utils import build_hf_headers, validate_repo_id


STANDARD = "2.0"
DEFAULT_DIST_BUCKET = "hfs-dist"
DEFAULT_LOCAL_ONLY = {"HF_TOKEN", "GH_TOKEN"}
SOVEREIGNTIES = {"sovereign", "fork", "port"}
LANES = {"source", "artifact"}
VERSION_SOURCES = {"latest", "tag", "commit"}
ENV_KEY = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
SAFE_SLUG = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,95}$")
TEXT_KEY_VALUE = re.compile(r"^\s*([A-Za-z0-9_.-]+)\s*[:=]\s*(.*?)\s*$")
URL_VALUE = re.compile(r"[A-Za-z][A-Za-z0-9+.-]*://[^\s\"']+")
DSN_FIELD = re.compile(
    r"(?:^|[\s;])(?P<key>[A-Za-z][A-Za-z0-9_.-]*)\s*=\s*(?P<value>[^\s;]+)"
)
SENSITIVE_KEY_SUFFIXES = {
    "pass",
    "password",
    "passwd",
    "passphrase",
    "pwd",
    "token",
    "secret",
    "api_key",
    "apikey",
    "access_key",
    "secret_key",
    "private_key",
    "client_secret",
    "credential",
    "credentials",
}
SAFE_PLACEHOLDERS = {"__FROM_ENV__", "CHANGE_ME", "REPLACE_ME", "REDACTED", "<SECRET>"}
ANGLE_PLACEHOLDER = re.compile(
    r"^<[^<>]*(?:placeholder|secret|token|password|passphrase|api[-_ ]?key|"
    r"replace|change|example|sample|todo|value)[^<>]*>$",
    flags=re.IGNORECASE,
)
STRUCTURED_SEED_SUFFIXES = {".toml", ".json", ".yaml", ".yml"}
TEXT_SEED_SUFFIXES = {"", ".cfg", ".conf", ".env", ".ini", ".properties", ".txt"}
LITERAL_SECRET = re.compile(
    r"(hf_[A-Za-z0-9]{20,}|ghp_[A-Za-z0-9]{20,}"
    r"|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9_-]{20,})"
)


class SyncError(RuntimeError):
    """可安全展示给用户的同步错误。"""


def resolve_local_file(root: Path, value: Path, field: str) -> Path:
    root = root.resolve()
    candidate = value if value.is_absolute() else root / value
    candidate = candidate.resolve()
    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise SyncError(f"{field} 不能指向项目根以外：{value}") from exc
    return candidate


def load_manifest(root: Path, manifest_file: Path = Path("hfs-dev.toml")) -> dict[str, Any]:
    path = resolve_local_file(root, manifest_file, "manifest")
    if not path.exists():
        raise SyncError(f"缺少 {path}；见规范第 7 节")
    try:
        with path.open("rb") as file:
            manifest = tomllib.load(file)
    except tomllib.TOMLDecodeError as exc:
        raise SyncError(f"{path.name} 解析失败：{exc}") from exc
    return manifest


def strip_env_inline_comment(value: str) -> str:
    quote: str | None = None
    escaped = False
    first_value = len(value) - len(value.lstrip())
    outer_quote = (
        value[first_value]
        if first_value < len(value) and value[first_value] in {'"', "'"}
        else None
    )
    outer_closed: int | None = None

    for index, character in enumerate(value):
        if escaped:
            escaped = False
            continue
        if character == "\\" and quote is not None:
            escaped = True
            continue
        if character in {'"', "'"}:
            if quote is None:
                quote = character
            elif quote == character:
                quote = None
                if outer_quote == character and outer_closed is None:
                    outer_closed = index
            continue
        if character != "#" or quote is not None:
            continue
        follows_whitespace = index > 0 and value[index - 1].isspace()
        follows_outer_quote = (
            outer_closed is not None
            and not value[outer_closed + 1 : index].strip()
        )
        if follows_whitespace or follows_outer_quote:
            return value[:index].rstrip()
    return value


def load_env(root: Path, env_file: Path = Path(".env")) -> dict[str, str]:
    path = resolve_local_file(root, env_file, "env-file")
    if not path.exists():
        raise SyncError(f"缺少 {path}；本地值必须先登记再同步")
    values: dict[str, str] = {}
    for number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            raise SyncError(f"{path.name} 第 {number} 行不是 KEY=VALUE")
        key, _, value = line.partition("=")
        key = key.strip()
        if not ENV_KEY.fullmatch(key):
            raise SyncError(f"{path.name} 第 {number} 行键名非法：{key!r}")
        if key in values:
            raise SyncError(f"{path.name} 中键名重复：{key}")
        value = strip_env_inline_comment(value).strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        values[key] = value
    return values


def string_list(manifest: dict[str, Any], field: str) -> list[str]:
    value = manifest.get(field, [])
    if not isinstance(value, list) or not all(isinstance(item, str) and item for item in value):
        raise SyncError(f"hfs-dev manifest 的 {field} 必须是非空字符串组成的数组")
    if len(value) != len(set(value)):
        raise SyncError(f"hfs-dev manifest 的 {field} 存在重复项")
    return value


def validate_slug(value: Any, field: str) -> str:
    if not isinstance(value, str) or not SAFE_SLUG.fullmatch(value):
        raise SyncError(f"{field} 必须是安全的单段名称：{value!r}")
    return value


def validate_object_path(value: Any, field: str) -> str:
    if not isinstance(value, str) or not value:
        raise SyncError(f"{field} 必须是非空相对路径")
    path = PurePosixPath(value)
    if path.is_absolute() or ".." in path.parts or "." in path.parts:
        raise SyncError(f"{field} 不能是绝对路径或包含 . / ..：{value!r}")
    return str(path)


def validate_setting_names(names: set[str], field: str) -> None:
    invalid = sorted(name for name in names if not ENV_KEY.fullmatch(name))
    if invalid:
        raise SyncError(f"{field} 包含非法环境变量键名：{invalid}")


def validate_manifest(manifest: dict[str, Any]) -> None:
    if str(manifest.get("standard", "")) != STANDARD:
        raise SyncError(f'standard 必须为 "{STANDARD}"')
    validate_slug(manifest.get("project"), "project")

    space = manifest.get("space")
    if not isinstance(space, str) or not space:
        raise SyncError("space 必须是非空字符串")
    try:
        validate_repo_id(space)
    except Exception as exc:
        raise SyncError(f"space ID 非法：{space!r}") from exc

    sovereignty = manifest.get("sovereignty")
    if not isinstance(sovereignty, str) or sovereignty not in SOVEREIGNTIES:
        raise SyncError("sovereignty 必须是 sovereign、fork 或 port")
    lane = manifest.get("lane")
    if not isinstance(lane, str) or lane not in LANES:
        raise SyncError("lane 必须是 source 或 artifact")
    version_source = manifest.get("version_source")
    if not isinstance(version_source, str) or version_source not in VERSION_SOURCES:
        raise SyncError("version_source 必须是 latest、tag 或 commit")

    if "bucket_namespace" in manifest:
        validate_slug(manifest["bucket_namespace"], "bucket_namespace")
    validate_slug(manifest.get("dist_bucket", DEFAULT_DIST_BUCKET), "dist_bucket")
    if "mount_config_bucket" in manifest and manifest["mount_config_bucket"] not in (None, ""):
        validate_slug(manifest["mount_config_bucket"], "mount_config_bucket")
        validate_object_path(manifest.get("mount_config_object", "config/config.toml"), "mount_config_object")

    secrets = set(string_list(manifest, "secrets"))
    optional_secrets = set(string_list(manifest, "optional_secrets"))
    variables = set(string_list(manifest, "variables"))
    local_only = DEFAULT_LOCAL_ONLY | set(string_list(manifest, "local_only"))
    validate_setting_names(secrets, "secrets")
    validate_setting_names(optional_secrets, "optional_secrets")
    validate_setting_names(variables, "variables")
    validate_setting_names(local_only, "local_only")
    if secrets & optional_secrets:
        raise SyncError(
            "同一键不能同时登记为 required secret 和 optional secret："
            f"{sorted(secrets & optional_secrets)}"
        )
    secret_variable_overlap = (secrets | optional_secrets) & variables
    if secret_variable_overlap:
        raise SyncError(
            "同一键不能同时登记为 secret/optional secret 和 variable："
            f"{sorted(secret_variable_overlap)}"
        )
    local_overlap = (secrets | optional_secrets | variables) & local_only
    if local_overlap:
        raise SyncError(
            "本地控制凭据不能登记为 Space secret/optional secret/variable："
            f"{sorted(local_overlap)}"
        )

    seed_value = manifest.get("seed_file")
    other_objects = string_list(manifest, "other_objects")
    if seed_value not in (None, ""):
        seed_relative = validate_object_path(seed_value, "seed_file")
        if Path(seed_relative).name not in other_objects:
            raise SyncError("seed_file 的文件名必须登记在 other_objects")
    for item in other_objects:
        validate_object_path(item, "other_objects")


def local_only_names(manifest: dict[str, Any]) -> set[str]:
    return DEFAULT_LOCAL_ONLY | set(string_list(manifest, "local_only"))


def registered_names(manifest: dict[str, Any]) -> tuple[set[str], set[str], set[str]]:
    secrets = set(string_list(manifest, "secrets"))
    optional_secrets = set(string_list(manifest, "optional_secrets"))
    variables = set(string_list(manifest, "variables"))
    return secrets, optional_secrets, variables


def hf_token(env_values: dict[str, str]) -> str:
    token = env_values.get("HF_TOKEN", "").strip()
    if not token:
        raise SyncError("env 文件缺少 HF_TOKEN；脚本不会隐式改用机器上的其他账号")
    return token


def api_client(token: str) -> HfApi:
    return HfApi(token=token)


def token_namespace(api: HfApi, token: str) -> str:
    name = api.whoami(token=token).get("name")
    if not isinstance(name, str) or not name:
        raise SyncError("HF_TOKEN 无法解析当前账号")
    return name


def space_id(manifest: dict[str, Any], token_owner: str) -> str:
    slug = manifest["space"].strip()
    return slug if "/" in slug else f"{token_owner}/{slug}"


def bucket_namespace(manifest: dict[str, Any], resolved_space: str) -> str:
    explicit = manifest.get("bucket_namespace")
    return explicit if isinstance(explicit, str) and explicit else resolved_space.split("/", 1)[0]


def seed_path(root: Path, manifest: dict[str, Any]) -> Path | None:
    value = manifest.get("seed_file")
    if value in (None, ""):
        return None
    return resolve_local_file(root, Path(value), "seed_file")


def seed_uri(storage_owner: str, manifest: dict[str, Any], local_seed: Path) -> str:
    bucket = manifest.get("dist_bucket", DEFAULT_DIST_BUCKET)
    project = manifest["project"]
    return f"hf://buckets/{storage_owner}/{bucket}/{project}/other/{local_seed.name}"


def instance_uri(manifest: dict[str, Any], storage_owner: str) -> str | None:
    bucket = manifest.get("mount_config_bucket")
    if bucket in (None, ""):
        return None
    object_name = manifest.get("mount_config_object", "config/config.toml")
    return f"hf://buckets/{storage_owner}/{bucket}/{str(object_name).lstrip('/')}"


def space_secret_names(space: str, token: str) -> set[str]:
    request = urllib.request.Request(
        f"https://huggingface.co/api/spaces/{space}/secrets",
        headers=build_hf_headers(token=token),
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        data = json.load(response)
    if not isinstance(data, dict):
        raise SyncError("HF Secret 名单接口返回了无法识别的数据")
    return set(data)


def bucket_cp(source: str, destination: str, token: str) -> tuple[bool, str]:
    process_env = os.environ.copy()
    process_env["HF_TOKEN"] = token
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "huggingface_hub.cli.hf",
            "buckets",
            "cp",
            source,
            destination,
            "--quiet",
        ],
        capture_output=True,
        text=True,
        check=False,
        env=process_env,
    )
    message = (result.stderr or result.stdout).strip().splitlines()
    return result.returncode == 0, message[-1] if message else ""


def bucket_read_bytes(source: str, token: str) -> tuple[bool, bytes]:
    process_env = os.environ.copy()
    process_env["HF_TOKEN"] = token
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "huggingface_hub.cli.hf",
            "buckets",
            "cp",
            source,
            "-",
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        env=process_env,
    )
    return result.returncode == 0, result.stdout


def normalized_sensitive_key(key: str) -> bool:
    with_boundaries = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", key)
    normalized = with_boundaries.lower().replace("-", "_").replace(".", "_")
    return any(normalized == term or normalized.endswith(f"_{term}") for term in SENSITIVE_KEY_SUFFIXES)


def placeholder_value(value: Any) -> bool:
    if value is None or value is False:
        return True
    if not isinstance(value, str):
        return False
    stripped = value.strip().strip('"').strip("'")
    if not stripped or stripped.upper() in SAFE_PLACEHOLDERS:
        return True
    if re.fullmatch(r"\$\{[A-Za-z_][A-Za-z0-9_]*(?::-[^}]*)?\}", stripped):
        return True
    if re.fullmatch(r"env:[A-Za-z_][A-Za-z0-9_]*", stripped):
        return True
    if ANGLE_PLACEHOLDER.fullmatch(stripped):
        return True
    return False


def embedded_credential(value: str) -> bool:
    for candidate in URL_VALUE.findall(value):
        candidate = candidate.rstrip(",);]")
        try:
            parsed = urllib.parse.urlsplit(candidate)
            password = parsed.password
        except ValueError:
            parsed = None
            password = None
        if password is not None and not placeholder_value(urllib.parse.unquote(password)):
            return True
        if parsed is not None:
            for key, query_value in urllib.parse.parse_qsl(
                parsed.query,
                keep_blank_values=True,
            ):
                if normalized_sensitive_key(key) and not placeholder_value(query_value):
                    return True
    for match in DSN_FIELD.finditer(value):
        if normalized_sensitive_key(match.group("key")) and not placeholder_value(
            match.group("value")
        ):
            return True
    return False


def matching_protected_names(
    value: str,
    protected_values: dict[str, tuple[str, str]],
) -> list[tuple[str, str]]:
    matches: list[tuple[str, str]] = []
    for name, (category, secret) in protected_values.items():
        if value == secret or (len(secret) >= 8 and secret in value):
            matches.append((category, name))
    return matches


def scalar_sensitive_fields(
    value: Any,
    location: str,
    protected_values: dict[str, tuple[str, str]],
) -> list[str]:
    if not isinstance(value, str):
        return []
    findings: list[str] = []
    if embedded_credential(value):
        findings.append(f"{location}:embedded-credential")
    for category, name in matching_protected_names(value, protected_values):
        findings.append(f"{location}:{category}:{name}")
    return findings


def structured_sensitive_fields(
    value: Any,
    protected_values: dict[str, tuple[str, str]],
    prefix: str = "",
) -> list[str]:
    findings: list[str] = []
    if isinstance(value, dict):
        for key, child in value.items():
            key_text = str(key)
            path = f"{prefix}.{key_text}" if prefix else key_text
            if normalized_sensitive_key(key_text) and not placeholder_value(child):
                findings.append(path)
            findings.extend(scalar_sensitive_fields(child, path, protected_values))
            findings.extend(structured_sensitive_fields(child, protected_values, path))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            path = f"{prefix}[{index}]"
            findings.extend(scalar_sensitive_fields(child, path, protected_values))
            findings.extend(structured_sensitive_fields(child, protected_values, path))
    else:
        findings.extend(scalar_sensitive_fields(value, prefix or "<root>", protected_values))
    return findings


def deduplicated(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))


def strip_unquoted_comment(line: str) -> str:
    quote: str | None = None
    escaped = False
    for index, character in enumerate(line):
        if escaped:
            escaped = False
            continue
        if character == "\\" and quote is not None:
            escaped = True
            continue
        if character in {'"', "'"}:
            if quote is None:
                quote = character
            elif quote == character:
                quote = None
            continue
        if character == "#" and quote is None:
            return line[:index]
    return line


def sensitive_seed_text(
    raw: str,
    suffix: str,
    protected_values: dict[str, tuple[str, str]] | None = None,
    *,
    strict_format: bool = False,
) -> list[str]:
    protected_values = protected_values or {}
    suffix = suffix.lower()
    literal_findings = ["literal-token-pattern"] if LITERAL_SECRET.search(raw) else []
    if suffix == ".toml":
        try:
            parsed = tomllib.loads(raw)
        except tomllib.TOMLDecodeError:
            raise SyncError("种子 TOML 解析失败（不输出配置原文）") from None
        return deduplicated(literal_findings + structured_sensitive_fields(parsed, protected_values))
    if suffix == ".json":
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise SyncError(f"种子 JSON 解析失败：第 {exc.lineno} 行第 {exc.colno} 列") from None
        return deduplicated(literal_findings + structured_sensitive_fields(parsed, protected_values))
    if suffix in {".yaml", ".yml"}:
        try:
            import yaml
        except ImportError as exc:
            raise SyncError("YAML seed 解析需要安装 PyYAML>=6.0") from exc
        try:
            documents = list(yaml.safe_load_all(raw))
        except yaml.YAMLError as exc:
            mark = getattr(exc, "problem_mark", None)
            if mark is not None and isinstance(mark.line, int) and isinstance(mark.column, int):
                location = f"第 {mark.line + 1} 行第 {mark.column + 1} 列"
            else:
                location = "位置未知"
            raise SyncError(f"种子 YAML 解析失败：{location}（不输出配置原文）") from None
        findings = list(literal_findings)
        for index, document in enumerate(documents):
            findings.extend(structured_sensitive_fields(document, protected_values, f"document[{index}]"))
        return deduplicated(findings)

    if strict_format and suffix not in TEXT_SEED_SUFFIXES:
        supported = ", ".join(sorted(item or "无扩展名" for item in STRUCTURED_SEED_SUFFIXES | TEXT_SEED_SUFFIXES))
        raise SyncError(
            f"种子格式 {suffix or '无扩展名'} 无法可靠检查；"
            f"示例脚本仅允许 {supported}，请为上游格式提供项目级安全扫描后再同步"
        )

    findings: list[str] = list(literal_findings)
    for number, raw_line in enumerate(raw.splitlines(), start=1):
        findings.extend(
            scalar_sensitive_fields(raw_line, f"line:{number}", protected_values)
        )
        line = strip_unquoted_comment(raw_line).rstrip()
        match = TEXT_KEY_VALUE.match(line)
        if match:
            location = f"line:{number}:{match.group(1)}"
            if normalized_sensitive_key(match.group(1)) and not placeholder_value(match.group(2)):
                findings.append(location)
            findings.extend(scalar_sensitive_fields(match.group(2), location, protected_values))
        else:
            findings.extend(scalar_sensitive_fields(line, f"line:{number}", protected_values))
    return deduplicated(findings)


def sensitive_seed_fields(
    path: Path,
    protected_values: dict[str, tuple[str, str]] | None = None,
    *,
    strict_format: bool = False,
) -> list[str]:
    return sensitive_seed_text(
        path.read_text(encoding="utf-8", errors="replace"),
        path.suffix,
        protected_values,
        strict_format=strict_format,
    )


def configured_optional_secrets(
    env_values: dict[str, str], optional_secrets: set[str]
) -> set[str]:
    return {name for name in optional_secrets if env_values.get(name, "")}


def protected_secret_values(
    env_values: dict[str, str], manifest: dict[str, Any]
) -> dict[str, tuple[str, str]]:
    secrets, optional_secrets, _ = registered_names(manifest)
    protected: dict[str, tuple[str, str]] = {}
    for name in secrets | optional_secrets:
        if env_values.get(name) and not placeholder_value(env_values[name]):
            protected[name] = ("env-secret", env_values[name])
    for name in local_only_names(manifest):
        if env_values.get(name) and not placeholder_value(env_values[name]):
            protected[name] = ("local-only", env_values[name])
    return protected


def unsafe_variable_reasons(
    name: str,
    value: str,
    protected_values: dict[str, tuple[str, str]],
) -> list[str]:
    reasons: list[str] = []
    if embedded_credential(value):
        reasons.append(f"{name}:embedded-credential")
    if LITERAL_SECRET.search(value):
        reasons.append(f"{name}:literal-token-pattern")
    for category, protected_name in matching_protected_names(value, protected_values):
        reasons.append(f"{name}:{category}:{protected_name}")
    return reasons


def unsafe_local_only_aliases(
    env_values: dict[str, str],
    manifest: dict[str, Any],
    setting_names: set[str],
) -> list[str]:
    protected_values = {
        name: ("local-only", env_values[name])
        for name in local_only_names(manifest)
        if env_values.get(name) and not placeholder_value(env_values[name])
    }
    findings: list[str] = []
    for name in sorted(setting_names):
        for _, protected_name in matching_protected_names(env_values[name], protected_values):
            findings.append(f"{name}:local-only:{protected_name}")
    return findings


def report(title: str, items: list[str]) -> int:
    print(f"[{title}] {len(items)} 项" if items else f"[{title}] 无差异")
    for item in items:
        print(f"  - {item}")
    return len(items)


def preflight(
    root: Path,
    manifest_file: Path = Path("hfs-dev.toml"),
    env_file: Path = Path(".env"),
    *,
    for_push: bool = False,
) -> tuple[
    dict[str, Any],
    dict[str, str],
    str,
    set[str],
    set[str],
    set[str],
    Path | None,
]:
    root = root.resolve()
    manifest = load_manifest(root, manifest_file)
    validate_manifest(manifest)
    env_values = load_env(root, env_file)
    token = hf_token(env_values)
    secrets, optional_secrets, variables = registered_names(manifest)
    missing = sorted(name for name in secrets | variables if not env_values.get(name, ""))
    if missing:
        raise SyncError(f"env 文件缺少已登记值，未执行任何写操作：{missing}")
    present_optional_secrets = configured_optional_secrets(env_values, optional_secrets)
    placeholder_settings = sorted(
        name
        for name in secrets | variables | present_optional_secrets
        if placeholder_value(env_values[name])
    )
    if placeholder_settings:
        raise SyncError(
            f"env 文件中的 Secret/Variable 仍是占位符，未执行任何写操作：{placeholder_settings}"
        )
    protected_values = protected_secret_values(env_values, manifest)
    variable_findings: list[str] = []
    for name in sorted(variables):
        variable_findings.extend(
            unsafe_variable_reasons(name, env_values[name], protected_values)
        )
    if variable_findings:
        raise SyncError(
            "Space Variable 值疑似含凭据，未执行任何写操作；"
            f"仅列出键名和原因：{deduplicated(variable_findings)}"
        )
    local_only_aliases = unsafe_local_only_aliases(
        env_values,
        manifest,
        secrets | present_optional_secrets,
    )
    if local_only_aliases:
        raise SyncError(
            "Space Secret 值与本地控制凭据重合，未执行任何写操作；"
            f"仅列出键名和来源：{local_only_aliases}"
        )
    local_seed = seed_path(root, manifest)
    if local_seed is not None and not local_seed.is_file():
        raise SyncError(f"seed_file 不存在，未执行任何写操作：{local_seed}")
    if for_push and local_seed is not None:
        sensitive = sensitive_seed_fields(
            local_seed,
            protected_values,
            strict_format=True,
        )
        if sensitive:
            raise SyncError(f"种子疑似包含实际 secret，禁止上传；字段位置：{sensitive}")
    return (
        manifest,
        env_values,
        token,
        secrets,
        optional_secrets,
        variables,
        local_seed,
    )


def resolve_targets(api: HfApi, manifest: dict[str, Any], token: str) -> tuple[str, str]:
    user = token_namespace(api, token)
    space = space_id(manifest, user)
    storage_owner = bucket_namespace(manifest, space)
    print(f"目标 Space：{space}")
    print(f"bucket namespace：{storage_owner}")
    print(f"下载 bucket：hf://buckets/{storage_owner}/{manifest.get('dist_bucket', DEFAULT_DIST_BUCKET)}")
    runtime = instance_uri(manifest, storage_owner)
    print(f"实例配置：{runtime or '未登记'}")
    return space, storage_owner


def cmd_diff(
    root: Path,
    manifest_file: Path = Path("hfs-dev.toml"),
    env_file: Path = Path(".env"),
) -> int:
    root = root.resolve()
    (
        manifest,
        env_values,
        token,
        secrets,
        optional_secrets,
        variables,
        local_seed,
    ) = preflight(root, manifest_file, env_file)
    api = api_client(token)
    space, storage_owner = resolve_targets(api, manifest, token)
    api.space_info(space, token=token)

    differences = 0
    present_optional_secrets = configured_optional_secrets(env_values, optional_secrets)
    managed_secrets = secrets | optional_secrets
    expected_secrets = secrets | present_optional_secrets
    registered = managed_secrets | variables
    deployable_env = set(env_values) - local_only_names(manifest)
    differences += report("env 有但未登记", sorted(deployable_env - registered))

    remote_secrets = space_secret_names(space, token)
    remote_variables = api.get_space_variables(space, token=token)
    differences += report("远端多出 secret", sorted(remote_secrets - managed_secrets))
    differences += report("远端缺 secret", sorted(expected_secrets - remote_secrets))
    differences += report("远端多出 variable", sorted(set(remote_variables) - variables))
    differences += report("远端缺 variable", sorted(variables - set(remote_variables)))
    differences += report(
        "variable 值不一致",
        sorted(
            name
            for name in variables & set(remote_variables)
            if remote_variables[name].value != env_values[name]
        ),
    )

    with tempfile.TemporaryDirectory(prefix="hfs-sync-diff-") as temp_dir:
        temp = Path(temp_dir)
        seed_copy: Path | None = None
        seed_ok = False

        if local_seed is None:
            print("[种子] 未配置 seed_file，跳过")
        else:
            sensitive = sensitive_seed_fields(
                local_seed,
                protected_secret_values(env_values, manifest),
            )
            differences += report("种子含受保护值或敏感字段", sensitive)

            seed_copy = temp / f"seed-{local_seed.name}"
            seed_ok, _ = bucket_cp(
                seed_uri(storage_owner, manifest, local_seed),
                str(seed_copy),
                token,
            )
            if not seed_ok:
                print("[种子] 远端不存在或不可读")
                differences += 1
            elif seed_copy.read_bytes() == local_seed.read_bytes():
                print("[种子 vs 本地] 一致")
            else:
                print("[种子 vs 本地] 不一致（push 会更新种子）")
                differences += 1

        runtime_uri = instance_uri(manifest, storage_owner)
        if runtime_uri is None:
            print("[实例配置] 未登记 mount_config_bucket，跳过")
        else:
            instance_name = PurePosixPath(
                str(manifest.get("mount_config_object", "config/config.toml"))
            ).name
            instance_copy = temp / f"instance-{instance_name}"
            instance_ok, _ = bucket_cp(runtime_uri, str(instance_copy), token)
            if not instance_ok:
                print("[实例配置] 挂载位置不存在或不可读")
                differences += 1
            elif local_seed is None:
                print("[实例配置] 存在且可读（未配置 seed_file，跳过与种子比较）")
            elif not seed_ok or seed_copy is None:
                print("[实例 vs 种子] 未比较：远端种子不可读")
            elif instance_copy.read_bytes() == seed_copy.read_bytes():
                print("[实例 vs 种子] 一致")
            else:
                print("[实例 vs 种子] 不一致：refresh 前先 pull 到 local/，检查后再决定")
                differences += 1

    return 1 if differences else 0


def verify_seed_round_trip(local_seed: Path, remote_uri: str, token: str) -> None:
    with tempfile.TemporaryDirectory(prefix="hfs-sync-seed-readback-") as temp_dir:
        downloaded = Path(temp_dir) / local_seed.name
        success, _ = bucket_cp(remote_uri, str(downloaded), token)
        if not success:
            raise SyncError(f"种子读回失败：{remote_uri}")
        if downloaded.read_bytes() != local_seed.read_bytes():
            raise SyncError("种子读回内容与本地不一致")


def cmd_push(
    root: Path,
    prune: bool,
    yes: bool,
    manifest_file: Path = Path("hfs-dev.toml"),
    env_file: Path = Path(".env"),
) -> int:
    root = root.resolve()
    if prune and not yes:
        raise SyncError("--prune 会删除远端设置，必须同时传 --yes")
    (
        manifest,
        env_values,
        token,
        secrets,
        optional_secrets,
        variables,
        local_seed,
    ) = preflight(root, manifest_file, env_file, for_push=True)
    api = api_client(token)
    space, storage_owner = resolve_targets(api, manifest, token)
    api.space_info(space, token=token)

    present_optional_secrets = configured_optional_secrets(env_values, optional_secrets)
    pushed_secrets = secrets | present_optional_secrets
    managed_secrets = secrets | optional_secrets
    for name in sorted(pushed_secrets):
        api.add_space_secret(space, name, env_values[name], token=token)
        print(f"secret 已推送：{name}")
    for name in sorted(variables):
        api.add_space_variable(space, name, env_values[name], token=token)
        print(f"variable 已推送：{name}")

    if local_seed is not None:
        destination = seed_uri(storage_owner, manifest, local_seed)
        uploaded, _ = bucket_cp(str(local_seed), destination, token)
        if not uploaded:
            raise SyncError(f"种子上传失败：{destination}")
        verify_seed_round_trip(local_seed, destination, token)
        print(f"种子上传和读回通过：{destination}")
    else:
        print("未配置 seed_file，不推送种子")

    if prune:
        remote_secrets = space_secret_names(space, token)
        remote_variables = api.get_space_variables(space, token=token)
        for name in sorted(remote_secrets - managed_secrets):
            api.delete_space_secret(space, name, token=token)
            print(f"secret 已删除：{name}")
        for name in sorted(set(remote_variables) - variables):
            api.delete_space_variable(space, name, token=token)
            print(f"variable 已删除：{name}")

    remote_secrets = space_secret_names(space, token)
    remote_variables = api.get_space_variables(space, token=token)
    missing_names = (pushed_secrets - remote_secrets) | (variables - set(remote_variables))
    value_drift = {
        name
        for name in variables & set(remote_variables)
        if remote_variables[name].value != env_values[name]
    }
    extras = set()
    if prune:
        extras = (remote_secrets - managed_secrets) | (set(remote_variables) - variables)
    if missing_names or value_drift or extras:
        raise SyncError(
            "读回校验失败："
            f"缺少名称 {sorted(missing_names)}；"
            f"variable 值不一致 {sorted(value_drift)}；"
            f"prune 后多余项 {sorted(extras)}"
        )
    print("读回校验通过（secret 比名称，variable 比值，种子比内容）")
    return 0


def ensure_contained(root: Path, path: Path) -> None:
    try:
        path.resolve().relative_to(root.resolve())
    except ValueError as exc:
        raise SyncError(f"pull 目录不能逃逸项目根：{path.name}") from exc


def directory_open_flags() -> int:
    flags = os.O_RDONLY
    flags |= getattr(os, "O_DIRECTORY", 0)
    flags |= getattr(os, "O_NOFOLLOW", 0)
    return flags


def validate_open_directory(root: Path, path: Path, descriptor: int) -> None:
    try:
        path_mode = path.lstat().st_mode
        path_stat = path.stat(follow_symlinks=False)
    except FileNotFoundError as exc:
        raise SyncError(f"pull 路径组件在校验期间消失：{path.name}") from exc
    if stat.S_ISLNK(path_mode):
        raise SyncError(f"pull 路径组件不能是符号链接：{path.name}")
    if not stat.S_ISDIR(path_mode):
        raise SyncError(f"pull 路径组件必须是目录：{path.name}")
    descriptor_stat = os.fstat(descriptor)
    if not stat.S_ISDIR(descriptor_stat.st_mode):
        raise SyncError(f"pull 路径组件必须是目录：{path.name}")
    if (path_stat.st_dev, path_stat.st_ino) != (
        descriptor_stat.st_dev,
        descriptor_stat.st_ino,
    ):
        raise SyncError(f"pull 路径组件在校验期间被替换：{path.name}")
    ensure_contained(root, path)
    os.fchmod(descriptor, 0o700)


def open_private_directory(root: Path, path: Path) -> int:
    try:
        descriptor = os.open(path, directory_open_flags())
    except OSError as exc:
        raise SyncError(f"pull 路径组件无法安全打开为目录：{path.name}") from exc
    try:
        validate_open_directory(root, path, descriptor)
    except Exception:
        os.close(descriptor)
        raise
    return descriptor


def ensure_private_directory(root: Path, path: Path) -> None:
    try:
        mode = path.lstat().st_mode
    except FileNotFoundError:
        ensure_contained(root, path)
        try:
            path.mkdir(mode=0o700, parents=False, exist_ok=False)
        except FileExistsError:
            pass
        mode = path.lstat().st_mode
    if stat.S_ISLNK(mode):
        raise SyncError(f"pull 路径组件不能是符号链接：{path.name}")
    if not stat.S_ISDIR(mode):
        raise SyncError(f"pull 路径组件必须是目录：{path.name}")

    descriptor = open_private_directory(root, path)
    try:
        validate_open_directory(root, path, descriptor)
    finally:
        os.close(descriptor)


def unique_pull_dir(root: Path, space: str) -> Path:
    root = root.resolve()
    if not root.is_dir():
        raise SyncError(f"项目根必须是目录：{root}")
    space_slug = validate_slug(space.split("/", 1)[-1], "Space slug")
    local_root = root / "local"
    archive_root = local_root / "hfs-sync-pulled"
    base = archive_root / space_slug
    for path in (local_root, archive_root, base):
        ensure_private_directory(root, path)
    timestamp = time.strftime("%Y%m%d%H%M%S")
    base_descriptor = open_private_directory(root, base)
    try:
        for attempt in range(256):
            name = timestamp if attempt == 0 else f"{timestamp}-{time.time_ns()}-{attempt - 1}"
            try:
                os.mkdir(name, mode=0o700, dir_fd=base_descriptor)
            except FileExistsError:
                try:
                    existing_mode = os.stat(
                        name,
                        dir_fd=base_descriptor,
                        follow_symlinks=False,
                    ).st_mode
                except OSError as exc:
                    raise SyncError(f"pull 唯一目录碰撞后无法安全检查：{name}") from exc
                if stat.S_ISLNK(existing_mode):
                    raise SyncError(f"pull 路径组件不能是符号链接：{name}")
                if not stat.S_ISDIR(existing_mode):
                    raise SyncError(f"pull 路径组件必须是目录：{name}")
                continue
            except OSError as exc:
                raise SyncError(f"无法创建私有 pull 目录：{name}") from exc

            candidate = base / name
            try:
                candidate_descriptor = os.open(
                    name,
                    directory_open_flags(),
                    dir_fd=base_descriptor,
                )
            except OSError as exc:
                raise SyncError(f"新建 pull 目录无法安全打开：{name}") from exc
            try:
                validate_open_directory(root, candidate, candidate_descriptor)
            except Exception:
                try:
                    os.rmdir(name, dir_fd=base_descriptor)
                except OSError:
                    pass
                raise
            finally:
                os.close(candidate_descriptor)
            return candidate
    finally:
        os.close(base_descriptor)
    raise SyncError("无法创建唯一的 pull 目录")


def stat_identity(value: os.stat_result) -> tuple[int, int]:
    return value.st_dev, value.st_ino


def validate_open_directory_at(
    parent_descriptor: int,
    name: str,
    descriptor: int,
) -> None:
    try:
        entry_stat = os.stat(
            name,
            dir_fd=parent_descriptor,
            follow_symlinks=False,
        )
    except FileNotFoundError as exc:
        raise SyncError("pull staging 在校验期间消失") from exc
    descriptor_stat = os.fstat(descriptor)
    if stat.S_ISLNK(entry_stat.st_mode) or not stat.S_ISDIR(entry_stat.st_mode):
        raise SyncError("pull staging 必须是普通目录")
    if not stat.S_ISDIR(descriptor_stat.st_mode):
        raise SyncError("pull staging 打开结果不是目录")
    if stat_identity(entry_stat) != stat_identity(descriptor_stat):
        raise SyncError("pull staging 在校验期间被替换")
    os.fchmod(descriptor, 0o700)


def create_private_staging_at(parent_descriptor: int) -> tuple[str, int]:
    for _ in range(256):
        name = f".staging-{os.urandom(12).hex()}"
        try:
            os.mkdir(name, mode=0o700, dir_fd=parent_descriptor)
        except FileExistsError:
            continue
        except OSError as exc:
            raise SyncError("无法在受信 pull 目录中创建私有 staging") from exc

        try:
            descriptor = os.open(
                name,
                directory_open_flags(),
                dir_fd=parent_descriptor,
            )
        except OSError as exc:
            try:
                os.rmdir(name, dir_fd=parent_descriptor)
            except OSError:
                pass
            raise SyncError("pull staging 无法安全打开") from exc
        try:
            validate_open_directory_at(parent_descriptor, name, descriptor)
        except Exception:
            os.close(descriptor)
            try:
                os.rmdir(name, dir_fd=parent_descriptor)
            except OSError:
                pass
            raise
        return name, descriptor
    raise SyncError("无法创建唯一的 pull staging 目录")


def validate_open_regular_file_at(
    directory_descriptor: int,
    filename: str,
    descriptor: int,
) -> None:
    try:
        entry_stat = os.stat(
            filename,
            dir_fd=directory_descriptor,
            follow_symlinks=False,
        )
    except FileNotFoundError as exc:
        raise SyncError("实例配置 staging 文件在校验期间消失") from exc
    descriptor_stat = os.fstat(descriptor)
    if stat.S_ISLNK(entry_stat.st_mode) or not stat.S_ISREG(entry_stat.st_mode):
        raise SyncError("实例配置 staging 结果不是普通文件")
    if not stat.S_ISREG(descriptor_stat.st_mode):
        raise SyncError("实例配置 staging 打开结果不是普通文件")
    if stat_identity(entry_stat) != stat_identity(descriptor_stat):
        raise SyncError("实例配置 staging 文件在校验期间被替换")
    os.fchmod(descriptor, 0o600)


def create_private_file_at(
    directory_descriptor: int,
    filename: str,
    content: bytes,
) -> int:
    flags = os.O_RDWR | os.O_CREAT | os.O_EXCL
    flags |= getattr(os, "O_NOFOLLOW", 0)
    try:
        descriptor = os.open(
            filename,
            flags,
            0o600,
            dir_fd=directory_descriptor,
        )
    except OSError as exc:
        raise SyncError("无法在 pull staging 中创建私有配置文件") from exc
    try:
        validate_open_regular_file_at(directory_descriptor, filename, descriptor)
        remaining = memoryview(content)
        while remaining:
            try:
                written = os.write(descriptor, remaining)
            except InterruptedError:
                continue
            if written <= 0:
                raise SyncError("写入 pull staging 文件失败")
            remaining = remaining[written:]
        os.lseek(descriptor, 0, os.SEEK_SET)
    except Exception:
        os.close(descriptor)
        try:
            os.unlink(filename, dir_fd=directory_descriptor)
        except OSError:
            pass
        raise
    return descriptor


def read_open_file(descriptor: int) -> bytes:
    os.lseek(descriptor, 0, os.SEEK_SET)
    chunks: list[bytes] = []
    while True:
        try:
            chunk = os.read(descriptor, 1024 * 1024)
        except InterruptedError:
            continue
        if not chunk:
            break
        chunks.append(chunk)
    os.lseek(descriptor, 0, os.SEEK_SET)
    return b"".join(chunks)


def validate_published_file_at(
    directory_descriptor: int,
    filename: str,
    source_descriptor: int,
) -> None:
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
    try:
        published_descriptor = os.open(
            filename,
            flags,
            dir_fd=directory_descriptor,
        )
    except OSError as exc:
        raise SyncError("实例配置回收结果无法安全打开") from exc
    try:
        published_stat = os.fstat(published_descriptor)
        source_stat = os.fstat(source_descriptor)
        if not stat.S_ISREG(published_stat.st_mode):
            raise SyncError("实例配置回收结果不是普通文件")
        if stat_identity(published_stat) != stat_identity(source_stat):
            raise SyncError("实例配置发布结果与受信 staging 文件不一致")
        os.fchmod(published_descriptor, 0o600)
    finally:
        os.close(published_descriptor)


def remove_entry_at(directory_descriptor: int, name: str) -> None:
    try:
        entry_stat = os.stat(
            name,
            dir_fd=directory_descriptor,
            follow_symlinks=False,
        )
    except FileNotFoundError:
        return
    if stat.S_ISDIR(entry_stat.st_mode) and not stat.S_ISLNK(entry_stat.st_mode):
        os.rmdir(name, dir_fd=directory_descriptor)
    else:
        os.unlink(name, dir_fd=directory_descriptor)


def cleanup_staging_at(
    parent_descriptor: int,
    staging_name: str,
    staging_descriptor: int,
    filename: str,
) -> None:
    try:
        os.unlink(filename, dir_fd=staging_descriptor)
    except FileNotFoundError:
        pass

    staging_identity = stat_identity(os.fstat(staging_descriptor))
    matching_names: list[str] = []
    # The owned directory may have been renamed; find it by inode without
    # resolving any attacker-controlled path.
    for entry in os.listdir(parent_descriptor):
        try:
            entry_stat = os.stat(
                entry,
                dir_fd=parent_descriptor,
                follow_symlinks=False,
            )
        except FileNotFoundError:
            continue
        if stat_identity(entry_stat) == staging_identity:
            matching_names.append(entry)

    for entry in matching_names:
        os.rmdir(entry, dir_fd=parent_descriptor)
    if staging_name not in matching_names:
        remove_entry_at(parent_descriptor, staging_name)
    if not matching_names:
        raise SyncError("pull staging 已被移出受信父目录，无法安全清理")


def cmd_pull(
    root: Path,
    manifest_file: Path = Path("hfs-dev.toml"),
    env_file: Path = Path(".env"),
) -> int:
    root = root.resolve()
    manifest = load_manifest(root, manifest_file)
    validate_manifest(manifest)
    env_values = load_env(root, env_file)
    token = hf_token(env_values)
    api = api_client(token)
    space, storage_owner = resolve_targets(api, manifest, token)
    api.space_info(space, token=token)
    runtime_uri = instance_uri(manifest, storage_owner)
    if runtime_uri is None:
        raise SyncError("未登记 mount_config_bucket，无法定位实例配置")

    filename = PurePosixPath(str(manifest.get("mount_config_object", "config/config.toml"))).name
    pull_dir = unique_pull_dir(root, space)
    downloaded = pull_dir / filename
    pull_descriptor = open_private_directory(root, pull_dir)
    staging_name: str | None = None
    staging_descriptor: int | None = None
    file_descriptor: int | None = None
    published = False
    sensitive: list[str] = []
    try:
        success, content = bucket_read_bytes(runtime_uri, token)
        if not success:
            raise SyncError(f"实例配置回收失败：{runtime_uri}")
        validate_open_directory(root, pull_dir, pull_descriptor)
        staging_name, staging_descriptor = create_private_staging_at(pull_descriptor)
        file_descriptor = create_private_file_at(
            staging_descriptor,
            filename,
            content,
        )
        staged_content = read_open_file(file_descriptor)
        if staged_content != content:
            raise SyncError("pull staging 内容与 HF CLI 内存输出不一致")
        sensitive = sensitive_seed_text(
            staged_content.decode("utf-8", errors="replace"),
            PurePosixPath(filename).suffix,
            protected_secret_values(env_values, manifest),
        )
        validate_open_directory_at(
            pull_descriptor,
            staging_name,
            staging_descriptor,
        )
        validate_open_regular_file_at(
            staging_descriptor,
            filename,
            file_descriptor,
        )
        if read_open_file(file_descriptor) != content:
            raise SyncError("pull staging 内容在安全检查期间被修改")
        try:
            os.link(
                filename,
                filename,
                src_dir_fd=staging_descriptor,
                dst_dir_fd=pull_descriptor,
                follow_symlinks=False,
            )
        except FileExistsError as exc:
            raise SyncError("实例配置回收目标已存在，拒绝覆盖") from exc
        except OSError as exc:
            raise SyncError("实例配置无法安全发布到最终 pull 目录") from exc
        published = True
        validate_published_file_at(
            pull_descriptor,
            filename,
            file_descriptor,
        )
        validate_open_directory(root, pull_dir, pull_descriptor)
    except Exception:
        if published:
            try:
                os.unlink(filename, dir_fd=pull_descriptor)
            except OSError:
                pass
        raise
    finally:
        # A cleanup failure must not replace the security/config error already
        # in flight from validation or parsing.
        primary_exception_active = sys.exc_info()[0] is not None
        cleanup_error: Exception | None = None
        try:
            if file_descriptor is not None:
                try:
                    os.close(file_descriptor)
                except OSError as exc:
                    cleanup_error = exc
            if staging_name is not None and staging_descriptor is not None:
                try:
                    cleanup_staging_at(
                        pull_descriptor,
                        staging_name,
                        staging_descriptor,
                        filename,
                    )
                except Exception as exc:
                    cleanup_error = cleanup_error or exc
            if cleanup_error is not None and not primary_exception_active and published:
                try:
                    os.unlink(filename, dir_fd=pull_descriptor)
                    published = False
                except OSError as exc:
                    cleanup_error = cleanup_error or exc
        finally:
            if staging_descriptor is not None:
                try:
                    os.close(staging_descriptor)
                except OSError as exc:
                    cleanup_error = cleanup_error or exc
            try:
                os.close(pull_descriptor)
            except OSError as exc:
                cleanup_error = cleanup_error or exc
        if cleanup_error is not None and not primary_exception_active:
            raise SyncError("pull staging 清理失败") from cleanup_error

    relative = downloaded.relative_to(root)
    print(f"实例配置已回收：{relative}（来源 {runtime_uri}）")
    if sensitive:
        print(f"WARN：回收配置包含受保护值或敏感字段：{sensitive}")
    print("根种子未修改；请人工 diff、脱密后再手工合并，随后单独执行 push")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["diff", "push", "pull"])
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="项目根目录（默认当前目录）")
    parser.add_argument("--manifest", type=Path, default=Path("hfs-dev.toml"), help="manifest 路径")
    parser.add_argument("--env-file", type=Path, default=Path(".env"), help="本地 env 文件路径")
    parser.add_argument("--prune", action="store_true", help="push 时删除远端多余设置；默认不删除")
    parser.add_argument("--yes", action="store_true", help="确认执行 --prune 的远端删除")
    args = parser.parse_args()
    root = args.root.resolve()

    try:
        if args.command == "diff":
            if args.prune or args.yes:
                raise SyncError("--prune/--yes 只适用于 push")
            return cmd_diff(root, args.manifest, args.env_file)
        if args.command == "push":
            return cmd_push(root, args.prune, args.yes, args.manifest, args.env_file)
        if args.prune or args.yes:
            raise SyncError("--prune/--yes 只适用于 push")
        return cmd_pull(root, args.manifest, args.env_file)
    except (SyncError, OSError, KeyError, subprocess.SubprocessError) as exc:
        print(f"ERROR：{exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(
            f"ERROR：外部 HF/API 调用失败（{type(exc).__name__}）；响应细节未输出，以免泄露凭据",
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
