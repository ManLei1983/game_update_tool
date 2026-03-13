import argparse
import hashlib
import json
import shlex
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional


SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "game_tool_config.json"
EXAMPLE_CONFIG_FILE = SCRIPT_DIR / "game_tool_config.example.json"

DEFAULT_CONFIG: Dict[str, Any] = {
    "server": {
        "base_url": "http://127.0.0.1:18080",
        "agent_id": "VM-3-1",
        "auth_token": "",
        "use_query_token": False,
        "timeout_seconds": 15,
    },
    "paths": {
        "cache_dir": "cache",
        "downloads_dir": "downloads",
        "backups_dir": "backups",
        "runtime_dir": "runtime",
        "bootstrap_file": "runtime/bootstrap.json",
        "payload_json_file": "runtime/task_payload.json",
        "payload_text_file": "runtime/task_payload.txt",
        "launch_file": "runtime/launch.json",
        "manifest_file": "runtime/manifest.json",
        "state_file": "cache/state.json",
        "exe_path": "QianNian.exe",
    },
    "behavior": {
        "download_resources": True,
        "download_exe": False,
        "launch_after_sync": False,
        "fail_on_missing_manifest": False,
    },
}


def print_line(message: str) -> None:
    print(message, flush=True)


def load_json_file(path: Path, default: Optional[Any] = None) -> Any:
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def save_json_file(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def merge_dict(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    merged: Dict[str, Any] = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = merge_dict(merged[key], value)
        else:
            merged[key] = value
    return merged


def sha256_of_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def backup_file(path: Path, backups_dir: Path) -> Optional[Path]:
    if not path.exists():
        return None
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    backup_path = backups_dir / f"{path.name}.{timestamp}.bak"
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup_path)
    return backup_path


class GameTool:
    def __init__(self, config: Dict[str, Any]) -> None:
        self.config = config
        self.server = config["server"]
        self.paths = config["paths"]
        self.behavior = config["behavior"]

        self.base_url = str(self.server["base_url"]).rstrip("/")
        self.agent_id = str(self.server["agent_id"]).strip()
        self.auth_token = str(self.server.get("auth_token", "")).strip()
        self.use_query_token = bool(self.server.get("use_query_token", False))
        self.timeout_seconds = int(self.server.get("timeout_seconds", 15) or 15)

        self.cache_dir = self.resolve_path(self.paths["cache_dir"])
        self.downloads_dir = self.resolve_path(self.paths["downloads_dir"])
        self.backups_dir = self.resolve_path(self.paths["backups_dir"])
        self.runtime_dir = self.resolve_path(self.paths["runtime_dir"])

        self.bootstrap_file = self.resolve_path(self.paths["bootstrap_file"])
        self.payload_json_file = self.resolve_path(self.paths["payload_json_file"])
        self.payload_text_file = self.resolve_path(self.paths["payload_text_file"])
        self.launch_file = self.resolve_path(self.paths["launch_file"])
        self.manifest_file = self.resolve_path(self.paths["manifest_file"])
        self.state_file = self.resolve_path(self.paths["state_file"])
        self.exe_path = self.resolve_path(self.paths["exe_path"])

    def resolve_path(self, raw_path: str) -> Path:
        path = Path(raw_path)
        if path.is_absolute():
            return path
        return SCRIPT_DIR / path

    def ensure_dirs(self) -> None:
        for directory in [self.cache_dir, self.downloads_dir, self.backups_dir, self.runtime_dir]:
            directory.mkdir(parents=True, exist_ok=True)

    def load_state(self) -> Dict[str, Any]:
        return load_json_file(self.state_file, default={}) or {}

    def save_state(self, state: Dict[str, Any]) -> None:
        save_json_file(self.state_file, state)

    def build_url(self, path_or_url: str, params: Optional[Dict[str, Any]] = None) -> str:
        if path_or_url.startswith(("http://", "https://")):
            base = path_or_url
        else:
            base = f"{self.base_url}/{path_or_url.lstrip('/')}"

        query: Dict[str, Any] = {}
        if params:
            query.update({k: v for k, v in params.items() if v not in (None, "")})
        if self.use_query_token and self.auth_token:
            query.setdefault("auth_token", self.auth_token)

        if not query:
            return base

        parsed = urllib.parse.urlsplit(base)
        current_query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
        for key, value in query.items():
            current_query[key] = [str(value)]
        encoded_query = urllib.parse.urlencode(current_query, doseq=True)
        return urllib.parse.urlunsplit(
            (parsed.scheme, parsed.netloc, parsed.path, encoded_query, parsed.fragment)
        )

    def request_json(self, url: str) -> Dict[str, Any]:
        request = urllib.request.Request(url, method="GET")
        if self.auth_token and not self.use_query_token:
            request.add_header("X-Auth-Token", self.auth_token)

        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                body = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "ignore")
            raise RuntimeError(f"请求失败: HTTP {exc.code} {detail}") from exc
        except Exception as exc:
            raise RuntimeError(f"请求失败: {exc}") from exc

        try:
            data = json.loads(body)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"返回内容不是合法 JSON: {body[:200]}") from exc

        if not isinstance(data, dict):
            raise RuntimeError("返回 JSON 不是对象")
        return data

    def fetch_bootstrap(self) -> Dict[str, Any]:
        if not self.agent_id:
            raise RuntimeError("配置里的 server.agent_id 不能为空")
        url = self.build_url("/api/bootstrap", {"agent_id": self.agent_id})
        return self.request_json(url)

    def fetch_manifest(self, bootstrap: Dict[str, Any]) -> Dict[str, Any]:
        manifest_info = bootstrap.get("downloads", {}).get("resources_manifest", {})
        manifest_url = str(manifest_info.get("url", "")).strip()
        if not manifest_url:
            manifest_url = self.build_url("/api/resources/manifest", {"agent_id": self.agent_id})
        return self.request_json(manifest_url)

    def write_bootstrap_outputs(self, bootstrap: Dict[str, Any], manifest: Optional[Dict[str, Any]]) -> None:
        save_json_file(self.bootstrap_file, bootstrap)
        save_json_file(self.launch_file, bootstrap.get("launch", {}))
        if manifest is not None:
            save_json_file(self.manifest_file, manifest)

        config_block = bootstrap.get("config", {})
        payload_json = config_block.get("payload_json")
        payload_text = config_block.get("payload_text") or ""

        if payload_json is not None:
            save_json_file(self.payload_json_file, payload_json)
        elif self.payload_json_file.exists():
            self.payload_json_file.unlink()

        if payload_text:
            self.payload_text_file.parent.mkdir(parents=True, exist_ok=True)
            self.payload_text_file.write_text(str(payload_text), encoding="utf-8")
        elif self.payload_text_file.exists():
            self.payload_text_file.unlink()

    def download_to_temp(self, url: str) -> Path:
        file_name = Path(urllib.parse.urlsplit(url).path).name or "download.bin"
        temp_path = self.downloads_dir / f".{int(time.time())}_{file_name}.tmp"
        temp_path.parent.mkdir(parents=True, exist_ok=True)

        request = urllib.request.Request(url, method="GET")
        if self.auth_token and not self.use_query_token:
            request.add_header("X-Auth-Token", self.auth_token)

        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                with temp_path.open("wb") as handle:
                    shutil.copyfileobj(response, handle)
        except Exception as exc:
            if temp_path.exists():
                temp_path.unlink()
            raise RuntimeError(f"下载失败: {url} -> {exc}") from exc

        return temp_path

    def verify_sha256(self, file_path: Path, expected_sha256: str) -> None:
        expected_sha256 = expected_sha256.strip().lower()
        if not expected_sha256:
            return
        actual_sha256 = sha256_of_file(file_path).lower()
        if actual_sha256 != expected_sha256:
            raise RuntimeError(
                f"SHA256 校验失败: {file_path.name} expected={expected_sha256} actual={actual_sha256}"
            )

    def sync_resource_items(self, manifest: Dict[str, Any], state: Dict[str, Any]) -> None:
        if not self.behavior.get("download_resources", True):
            print_line("[SYNC] 已跳过资源下载（behavior.download_resources=false）")
            return

        resource_state = state.setdefault("resources", {})
        items = manifest.get("items", [])
        if not isinstance(items, list):
            raise RuntimeError("manifest.items 必须是数组")

        for item in items:
            if not isinstance(item, dict):
                continue

            name = str(item.get("name", "")).strip() or "unnamed"
            url = str(item.get("url", "")).strip()
            target_path_text = str(item.get("target_path", "")).strip()
            version = str(item.get("version", "")).strip()
            sha256 = str(item.get("sha256", "")).strip()

            if not url or not target_path_text:
                print_line(f"[SYNC] 跳过资源 {name}，因为 url 或 target_path 为空")
                continue

            target_path = self.resolve_path(target_path_text)
            old_version = str(resource_state.get(name, {}).get("version", ""))
            if target_path.exists() and version and old_version == version:
                print_line(f"[SYNC] 资源未变化，跳过: {name} ({version})")
                continue

            print_line(f"[SYNC] 下载资源: {name}")
            temp_file = self.download_to_temp(url)
            try:
                self.verify_sha256(temp_file, sha256)
                target_path.parent.mkdir(parents=True, exist_ok=True)
                backup = backup_file(target_path, self.backups_dir)
                shutil.move(str(temp_file), str(target_path))
                resource_state[name] = {
                    "version": version,
                    "target_path": str(target_path),
                    "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                }
                if backup:
                    print_line(f"[SYNC] 已备份旧文件 -> {backup}")
                print_line(f"[SYNC] 资源已更新 -> {target_path}")
            finally:
                if temp_file.exists():
                    temp_file.unlink()

    def sync_exe(self, bootstrap: Dict[str, Any], state: Dict[str, Any]) -> None:
        if not self.behavior.get("download_exe", False):
            return

        exe_info = bootstrap.get("downloads", {}).get("exe", {})
        exe_url = str(exe_info.get("url", "")).strip()
        exe_version = str(exe_info.get("version", "")).strip()
        exe_sha256 = str(exe_info.get("sha256", "")).strip()

        if not exe_url:
            print_line("[SYNC] 未配置 exe 下载地址，跳过 exe 更新")
            return

        old_version = str(state.get("exe_version", ""))
        if self.exe_path.exists() and exe_version and old_version == exe_version:
            print_line(f"[SYNC] EXE 未变化，跳过: {exe_version}")
            return

        print_line(f"[SYNC] 下载 EXE 更新: {exe_url}")
        temp_file = self.download_to_temp(exe_url)
        try:
            self.verify_sha256(temp_file, exe_sha256)
            if temp_file.suffix.lower() != ".exe":
                staged_path = self.downloads_dir / temp_file.name.replace(".tmp", "")
                shutil.move(str(temp_file), str(staged_path))
                print_line(f"[SYNC] 已下载 EXE 更新包，但第一版只自动替换 .exe 文件: {staged_path}")
                return

            self.exe_path.parent.mkdir(parents=True, exist_ok=True)
            backup = backup_file(self.exe_path, self.backups_dir)
            shutil.move(str(temp_file), str(self.exe_path))
            state["exe_version"] = exe_version
            if backup:
                print_line(f"[SYNC] 已备份旧 EXE -> {backup}")
            print_line(f"[SYNC] EXE 已更新 -> {self.exe_path}")
        finally:
            if temp_file.exists():
                temp_file.unlink()

    def print_bootstrap_summary(self, bootstrap: Dict[str, Any]) -> None:
        task = bootstrap.get("task", {})
        config_info = bootstrap.get("config", {})
        launch = bootstrap.get("launch", {})
        downloads = bootstrap.get("downloads", {})

        print_line("=" * 56)
        print_line(f"Agent ID: {bootstrap.get('agent_id', '-')}")
        print_line(f"任务启用: {task.get('enabled', False)}")
        print_line(f"区服: {task.get('region', '-')}")
        print_line(f"Group: {task.get('group_start', 0)} -> {task.get('group_end', 0)}")
        print_line(f"任务模式: {task.get('task_mode', '-')}")
        print_line(f"配置版本: {bootstrap.get('profile_version', '-')}")
        print_line(f"脚本配置版本: {config_info.get('version', '-')}")
        print_line(f"资源清单版本: {downloads.get('resources_manifest', {}).get('version', '-')}")
        print_line(f"启动 EXE: {launch.get('startup_exe', '-')}")
        print_line("=" * 56)

    def do_bootstrap(self) -> Dict[str, Any]:
        self.ensure_dirs()
        bootstrap = self.fetch_bootstrap()
        self.print_bootstrap_summary(bootstrap)
        return bootstrap

    def do_sync(self) -> Dict[str, Any]:
        self.ensure_dirs()
        state = self.load_state()

        bootstrap = self.fetch_bootstrap()
        manifest: Optional[Dict[str, Any]] = None
        try:
            manifest = self.fetch_manifest(bootstrap)
        except Exception as exc:
            if self.behavior.get("fail_on_missing_manifest", False):
                raise
            print_line(f"[SYNC] 资源清单获取失败，已跳过: {exc}")

        self.write_bootstrap_outputs(bootstrap, manifest)
        if manifest is not None:
            self.sync_resource_items(manifest, state)

        self.sync_exe(bootstrap, state)

        state["agent_id"] = bootstrap.get("agent_id", self.agent_id)
        state["profile_version"] = bootstrap.get("profile_version", "")
        state["config_version"] = bootstrap.get("config", {}).get("version", "")
        state["manifest_version"] = (
            bootstrap.get("downloads", {}).get("resources_manifest", {}).get("version", "")
        )
        state["last_sync_time"] = time.strftime("%Y-%m-%d %H:%M:%S")
        self.save_state(state)

        print_line(f"[SYNC] bootstrap 已保存 -> {self.bootstrap_file}")
        if self.payload_json_file.exists():
            print_line(f"[SYNC] JSON 配置已写入 -> {self.payload_json_file}")
        if self.payload_text_file.exists():
            print_line(f"[SYNC] 文本配置已写入 -> {self.payload_text_file}")
        if self.launch_file.exists():
            print_line(f"[SYNC] 启动参数已写入 -> {self.launch_file}")
        if self.manifest_file.exists():
            print_line(f"[SYNC] 资源清单已写入 -> {self.manifest_file}")

        self.print_bootstrap_summary(bootstrap)
        return bootstrap

    def do_launch(self) -> None:
        launch_info = load_json_file(self.launch_file, default={}) or {}
        startup_exe = str(launch_info.get("startup_exe", "")).strip()
        startup_args = str(launch_info.get("startup_args", "")).strip()

        exe_path = self.resolve_path(startup_exe) if startup_exe else self.exe_path
        if not exe_path.exists():
            raise RuntimeError(f"找不到要启动的 EXE: {exe_path}")

        command = [str(exe_path)]
        if startup_args:
            command.extend(shlex.split(startup_args, posix=False))

        print_line(f"[LAUNCH] 启动程序: {exe_path}")
        subprocess.Popen(command, cwd=str(exe_path.parent))

    def do_run(self) -> None:
        bootstrap = self.do_sync()
        if not bootstrap.get("task", {}).get("enabled", True):
            print_line("[RUN] 当前 VM 任务处于停用状态，已停止，不启动 EXE")
            return
        self.do_launch()


def create_example_config() -> None:
    if not EXAMPLE_CONFIG_FILE.exists():
        save_json_file(EXAMPLE_CONFIG_FILE, DEFAULT_CONFIG)


def create_runtime_config_if_missing() -> bool:
    if CONFIG_FILE.exists():
        return False
    save_json_file(CONFIG_FILE, DEFAULT_CONFIG)
    return True


def load_config() -> Dict[str, Any]:
    if not CONFIG_FILE.exists():
        raise RuntimeError(
            f"找不到配置文件: {CONFIG_FILE}。请先执行: python game_tool.py init"
        )
    user_config = load_json_file(CONFIG_FILE, default={}) or {}
    if not isinstance(user_config, dict):
        raise RuntimeError("game_tool_config.json 必须是 JSON 对象")
    return merge_dict(DEFAULT_CONFIG, user_config)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="game_tool 第一版")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("init", help="生成示例配置并创建目录")
    subparsers.add_parser("bootstrap", help="拉取 bootstrap 并打印摘要")
    subparsers.add_parser("sync", help="拉取 bootstrap、manifest，并写入本地文件")
    subparsers.add_parser("launch", help="直接启动 EXE")
    subparsers.add_parser("run", help="先 sync，再启动 EXE")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        create_example_config()

        if args.command == "init":
            created = create_runtime_config_if_missing()
            config = load_config()
            tool = GameTool(config)
            tool.ensure_dirs()
            if created:
                print_line(f"已生成配置文件: {CONFIG_FILE}")
            else:
                print_line(f"配置文件已存在: {CONFIG_FILE}")
            print_line(f"示例配置文件: {EXAMPLE_CONFIG_FILE}")
            print_line("请先修改 game_tool_config.json，再执行 bootstrap 或 sync")
            return 0

        config = load_config()
        tool = GameTool(config)

        if args.command == "bootstrap":
            tool.do_bootstrap()
            return 0
        if args.command == "sync":
            tool.do_sync()
            return 0
        if args.command == "launch":
            tool.do_launch()
            return 0
        if args.command == "run":
            tool.do_run()
            return 0

        parser.print_help()
        return 1
    except Exception as exc:
        print_line(f"[ERROR] {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
