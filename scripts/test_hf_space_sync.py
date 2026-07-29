from __future__ import annotations

import importlib.util
import json
import stat
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]


def load_sync_module():
    fake_huggingface_hub = types.ModuleType("huggingface_hub")
    fake_huggingface_utils = types.ModuleType("huggingface_hub.utils")
    fake_huggingface_hub.HfApi = type("HfApi", (), {})
    fake_huggingface_utils.build_hf_headers = lambda **_kwargs: {}

    def validate_repo_id(repo_id: str) -> None:
        if not isinstance(repo_id, str) or "/" not in repo_id:
            raise ValueError("invalid repo id")

    fake_huggingface_utils.validate_repo_id = validate_repo_id
    spec = importlib.util.spec_from_file_location(
        "repo_local_hf_space_sync",
        ROOT / "scripts" / "hf_space_sync.py",
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("unable to load hf_space_sync.py")
    module = importlib.util.module_from_spec(spec)
    with mock.patch.dict(
        sys.modules,
        {
            "huggingface_hub": fake_huggingface_hub,
            "huggingface_hub.utils": fake_huggingface_utils,
        },
    ):
        spec.loader.exec_module(module)
    return module


sync = load_sync_module()


class FakeApi:
    def space_info(self, *_args: object, **_kwargs: object) -> None:
        return None


def write_fixture(
    root: Path,
    *,
    app_mode: str = "public",
    secrets: list[str] | None = None,
    extra_env: list[str] | None = None,
) -> None:
    secret_names = secrets or ["APP_SECRET"]
    manifest = [
        'standard = "2.1"',
        'project = "sync-security-test"',
        'space = "example/sync-security-test"',
        'project_class = "preview"',
        'target_role = "primary"',
        'sovereignty = "sovereign"',
        'lane = "source"',
        'version_source = "commit"',
        'env_file = ".env"',
        'secret_files = []',
        'mount_config_bucket = "sync-security-data"',
        'mount_config_object = "config/config.toml"',
        'local_only = ["HF_TOKEN", "GH_TOKEN", "PROJECT_CONTROL"]',
        f"secrets = {json.dumps(secret_names)}",
        'variables = ["APP_MODE"]',
        "other_objects = []",
        "",
    ]
    (root / "hfs-dev.toml").write_text("\n".join(manifest), encoding="utf-8")
    env_lines = [
        "HF_TOKEN=test-control-token",
        "GH_TOKEN=test-github-control",
        "PROJECT_CONTROL=local-control-value",
        "APP_SECRET=registered-secret-value",
        f"APP_MODE={app_mode}",
    ]
    env_lines.extend(extra_env or [])
    (root / ".env").write_text("\n".join(env_lines) + "\n", encoding="utf-8")


class SyncSecurityTests(unittest.TestCase):
    def test_push_rejects_sensitive_variables_before_remote_calls(self) -> None:
        cases = [
            "postgresql://app:" + "LEAKME-url-password" + "@db.example/app",
            "https://api.example/v1?client_" + "secret=LEAKME-query-secret",
            "Server=db.example;Access" + "Token=LEAKME-dsn-token;Database=app",
            "hf_" + ("1" * 20),
            "prefix-registered-secret-value-suffix",
            "prefix-local-control-value-suffix",
        ]
        for value in cases:
            with self.subTest(value=value[:24]):
                with tempfile.TemporaryDirectory() as temporary:
                    root = Path(temporary)
                    write_fixture(root, app_mode=value)
                    with (
                        mock.patch.object(sync, "api_client") as api_client,
                        mock.patch.object(sync, "bucket_cp") as bucket_copy,
                        self.assertRaises(sync.SyncError) as caught,
                    ):
                        sync.cmd_push(root, False, False)
                    api_client.assert_not_called()
                    bucket_copy.assert_not_called()
                    message = str(caught.exception)
                    self.assertIn("APP_MODE", message)
                    self.assertNotIn("LEAKME", message)
                    self.assertNotIn(value, message)

    def test_secret_cannot_alias_custom_local_only_value(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            write_fixture(
                root,
                secrets=["APP_SECRET", "CONTROL_ALIAS"],
                extra_env=["CONTROL_ALIAS=local-control-value"],
            )
            with (
                mock.patch.object(sync, "api_client") as api_client,
                self.assertRaises(sync.SyncError) as caught,
            ):
                sync.cmd_push(root, False, False)
            api_client.assert_not_called()
            message = str(caught.exception)
            self.assertIn("CONTROL_ALIAS", message)
            self.assertIn("PROJECT_CONTROL", message)
            self.assertNotIn("local-control-value", message)

    def test_custom_local_only_value_is_protected_in_seed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            write_fixture(root)
            manifest_path = root / "hfs-dev.toml"
            manifest_path.write_text(
                manifest_path.read_text(encoding="utf-8").replace(
                    "other_objects = []",
                    'seed_file = "config.toml"\nother_objects = ["config.toml"]',
                ),
                encoding="utf-8",
            )
            (root / "config.toml").write_text(
                'service_value = "local-control-value"\n',
                encoding="utf-8",
            )
            with self.assertRaises(sync.SyncError) as caught:
                sync.preflight(root, for_push=True)
            message = str(caught.exception)
            self.assertIn("local-only:PROJECT_CONTROL", message)
            self.assertNotIn("local-control-value", message)

    def test_public_and_placeholder_urls_remain_valid_variables(self) -> None:
        cases = [
            "https://api.example/v1?format=json&mode=public",
            "https://api.example/v1?api_" + "key=%3CSECRET%3E",
            "postgresql://readonly@db.example/app?sslmode=require",
        ]
        for value in cases:
            with self.subTest(value=value):
                with tempfile.TemporaryDirectory() as temporary:
                    root = Path(temporary)
                    write_fixture(root, app_mode=value)
                    sync.preflight(root, for_push=True)

    def test_pull_rejects_symlink_and_non_directory_components_before_read(self) -> None:
        for kind in ("symlink", "file"):
            for component in ("local", "hfs-sync-pulled", "sync-security-test"):
                with self.subTest(kind=kind, component=component):
                    with (
                        tempfile.TemporaryDirectory() as temporary,
                        tempfile.TemporaryDirectory() as outside,
                    ):
                        root = Path(temporary)
                        write_fixture(root)
                        parent = root
                        for name in ("local", "hfs-sync-pulled", "sync-security-test"):
                            path = parent / name
                            if name == component:
                                if kind == "symlink":
                                    path.symlink_to(Path(outside), target_is_directory=True)
                                else:
                                    path.write_text("not a directory\n", encoding="utf-8")
                                break
                            path.mkdir(mode=0o700)
                            parent = path
                        with (
                            mock.patch.object(sync, "api_client", return_value=FakeApi()),
                            mock.patch.object(
                                sync,
                                "resolve_targets",
                                return_value=("example/sync-security-test", "example"),
                            ),
                            mock.patch.object(sync, "bucket_read_bytes") as bucket_read,
                            self.assertRaisesRegex(sync.SyncError, "符号链接|目录"),
                        ):
                            sync.cmd_pull(root)
                        bucket_read.assert_not_called()
                        self.assertEqual(list(Path(outside).iterdir()), [])

    def test_pull_rejects_unsafe_final_parent_before_read(self) -> None:
        for kind in ("symlink", "file"):
            with self.subTest(kind=kind):
                with (
                    tempfile.TemporaryDirectory() as temporary,
                    tempfile.TemporaryDirectory() as outside,
                ):
                    root = Path(temporary)
                    write_fixture(root)
                    base = root / "local" / "hfs-sync-pulled" / "sync-security-test"
                    base.mkdir(parents=True, mode=0o700)
                    target = base / "20260730010101"
                    if kind == "symlink":
                        target.symlink_to(Path(outside), target_is_directory=True)
                    else:
                        target.write_text("not a directory\n", encoding="utf-8")
                    with (
                        mock.patch.object(sync, "api_client", return_value=FakeApi()),
                        mock.patch.object(
                            sync,
                            "resolve_targets",
                            return_value=("example/sync-security-test", "example"),
                        ),
                        mock.patch.object(
                            sync.time,
                            "strftime",
                            return_value="20260730010101",
                        ),
                        mock.patch.object(sync, "bucket_read_bytes") as bucket_read,
                        self.assertRaisesRegex(sync.SyncError, "符号链接|目录"),
                    ):
                        sync.cmd_pull(root)
                    bucket_read.assert_not_called()
                    self.assertEqual(list(Path(outside).iterdir()), [])

    def test_pull_rejects_final_parent_symlink_swap_after_read(self) -> None:
        with (
            tempfile.TemporaryDirectory() as temporary,
            tempfile.TemporaryDirectory() as outside,
        ):
            root = Path(temporary)
            outside_root = Path(outside)
            write_fixture(root)

            def replace_pull_dir(
                _source: str,
                _token: str,
            ) -> tuple[bool, bytes]:
                base = root / "local" / "hfs-sync-pulled" / "sync-security-test"
                pull_dir = next(base.iterdir())
                pull_dir.rmdir()
                pull_dir.symlink_to(outside_root, target_is_directory=True)
                return True, b"enabled = true\n"

            with (
                mock.patch.object(sync, "api_client", return_value=FakeApi()),
                mock.patch.object(
                    sync,
                    "resolve_targets",
                    return_value=("example/sync-security-test", "example"),
                ),
                mock.patch.object(
                    sync,
                    "bucket_read_bytes",
                    side_effect=replace_pull_dir,
                ),
                self.assertRaisesRegex(
                    sync.SyncError,
                    "符号链接|安全发布|校验期间被替换",
                ),
            ):
                sync.cmd_pull(root)
            self.assertFalse((outside_root / "config.toml").exists())
            self.assertFalse(any(root.glob(".staging-*")))

    def test_pull_rejects_space_slug_containment_escape(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            with self.assertRaisesRegex(sync.SyncError, "安全的单段名称|项目根"):
                sync.unique_pull_dir(root, "example/../../outside")
            self.assertFalse((root / "outside").exists())

    def test_pull_writes_private_regular_file_without_staging_leftovers(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            write_fixture(root)
            with (
                mock.patch.object(sync, "api_client", return_value=FakeApi()),
                mock.patch.object(
                    sync,
                    "resolve_targets",
                    return_value=("example/sync-security-test", "example"),
                ),
                mock.patch.object(
                    sync,
                    "bucket_read_bytes",
                    return_value=(True, b"enabled = true\n"),
                ),
            ):
                self.assertEqual(sync.cmd_pull(root), 0)
            pulled = list(
                (root / "local" / "hfs-sync-pulled" / "sync-security-test").glob(
                    "*/config.toml"
                )
            )
            self.assertEqual(len(pulled), 1)
            self.assertEqual(pulled[0].read_bytes(), b"enabled = true\n")
            self.assertEqual(stat.S_IMODE(pulled[0].stat().st_mode), 0o600)
            self.assertEqual(stat.S_IMODE(pulled[0].parent.stat().st_mode), 0o700)
            self.assertFalse(any(root.rglob(".staging-*")))


if __name__ == "__main__":
    unittest.main()
