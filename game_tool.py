import argparse
import csv
import ctypes
import hashlib
import io
import json
import os
import shlex
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from ctypes import wintypes
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "game_tool_config.json"
EXAMPLE_CONFIG_FILE = SCRIPT_DIR / "game_tool_config.example.json"
IS_WINDOWS = os.name == "nt"

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
        "control_poll_seconds": 15,
        "control_error_retry_seconds": 30,
        "process_stop_timeout_seconds": 20,
        "window_find_timeout_seconds": 60,
        "launch_ready_seconds": 20,
        "post_load_delay_seconds": 2,
        "launch_settle_seconds": 8,
        "startup_grace_seconds_fallback": 300,
    },
    "qiannian": {
        "ui_enabled": True,
        "launch_button": "gongzi",
        "control_ids": {
            "region_combo": 1005,
            "load_button": 1007,
            "current_group_edit": 1008,
            "max_group_edit": 1021,
            "role_index_edit": 1016,
            "start_button": 1002,
            "runtask_button": 1013,
            "gongzi_button": 1020,
        },
    },
}

VALID_ONE_SHOT_ACTIONS = {"", "start_once", "restart_once", "sync_once", "stop_once"}
BUTTON_ID_MAP = {
    "start": "start_button",
    "runtask": "runtask_button",
    "gongzi": "gongzi_button",
}


def print_line(message: str) -> None:
    print(message, flush=True)


def now_str() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def today_str() -> str:
    return time.strftime("%Y-%m-%d")


def parse_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def normalize_action(value: Any) -> str:
    text = str(value or "").strip().lower()
    if text in VALID_ONE_SHOT_ACTIONS:
        return text
    return ""


def parse_hhmm(value: Any) -> Optional[Tuple[int, int]]:
    text = str(value or "").strip()
    if not text:
        return None
    parts = text.split(":", 1)
    if len(parts) != 2:
        return None
    try:
        hour = int(parts[0])
        minute = int(parts[1])
    except ValueError:
        return None
    if hour < 0 or hour > 23 or minute < 0 or minute > 59:
        return None
    return hour, minute


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

class Win32DialogController:
    BM_CLICK = 0x00F5
    WM_COMMAND = 0x0111
    WM_SETTEXT = 0x000C
    WM_GETTEXT = 0x000D
    WM_GETTEXTLENGTH = 0x000E
    EN_CHANGE = 0x0300
    CBN_SELCHANGE = 1
    CB_FINDSTRING = 0x014C
    CB_FINDSTRINGEXACT = 0x0158
    CB_SETCURSEL = 0x014E
    GW_OWNER = 4
    SW_RESTORE = 9

    def __init__(self) -> None:
        if not IS_WINDOWS:
            raise RuntimeError("Win32 UI 控制仅支持 Windows")
        self.user32 = ctypes.windll.user32

    @staticmethod
    def make_wparam(low_word: int, high_word: int) -> int:
        return (high_word << 16) | (low_word & 0xFFFF)

    def find_main_window(self, pid: int, timeout_seconds: int) -> int:
        deadline = time.time() + max(1, timeout_seconds)
        while time.time() < deadline:
            hwnd = self._find_main_window_once(pid)
            if hwnd:
                self.user32.ShowWindow(hwnd, self.SW_RESTORE)
                return hwnd
            time.sleep(0.5)
        raise RuntimeError(f"在 {timeout_seconds} 秒内未找到进程 {pid} 的主窗口")

    def _find_main_window_once(self, pid: int) -> int:
        result: List[int] = []
        enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        @enum_proc_type
        def enum_proc(hwnd: int, _lparam: int) -> bool:
            proc_id = wintypes.DWORD(0)
            self.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
            if proc_id.value != pid:
                return True
            if not self.user32.IsWindowVisible(hwnd):
                return True
            if self.user32.GetWindow(hwnd, self.GW_OWNER):
                return True
            result.append(hwnd)
            return False

        self.user32.EnumWindows(enum_proc, 0)
        return result[0] if result else 0

    def get_control(self, parent_hwnd: int, control_id: int) -> int:
        child_hwnd = self.user32.GetDlgItem(parent_hwnd, control_id)
        if not child_hwnd:
            raise RuntimeError(f"找不到控件 ID={control_id}")
        return child_hwnd

    def set_edit_text(self, parent_hwnd: int, control_id: int, text: str) -> None:
        child_hwnd = self.get_control(parent_hwnd, control_id)
        if not self.user32.SendMessageW(
            child_hwnd,
            self.WM_SETTEXT,
            0,
            ctypes.c_wchar_p(str(text)),
        ):
            raise RuntimeError(f"设置文本失败，控件 ID={control_id}")
        self.user32.SendMessageW(
            parent_hwnd,
            self.WM_COMMAND,
            self.make_wparam(control_id, self.EN_CHANGE),
            child_hwnd,
        )
        current_text = self.get_edit_text(parent_hwnd, control_id)
        if current_text != str(text):
            raise RuntimeError(
                f"控件文本回读不一致，控件 ID={control_id} expected={text} actual={current_text}"
            )

    def get_edit_text(self, parent_hwnd: int, control_id: int) -> str:
        child_hwnd = self.get_control(parent_hwnd, control_id)
        length = int(self.user32.SendMessageW(child_hwnd, self.WM_GETTEXTLENGTH, 0, 0))
        buffer = ctypes.create_unicode_buffer(length + 1)
        self.user32.SendMessageW(child_hwnd, self.WM_GETTEXT, length + 1, buffer)
        return buffer.value

    def select_combo_text(self, parent_hwnd: int, control_id: int, text: str) -> None:
        child_hwnd = self.get_control(parent_hwnd, control_id)
        keyword = str(text or "").strip()
        if not keyword:
            return

        index = self.user32.SendMessageW(
            child_hwnd,
            self.CB_FINDSTRINGEXACT,
            -1,
            ctypes.c_wchar_p(keyword),
        )
        if index == -1:
            index = self.user32.SendMessageW(
                child_hwnd,
                self.CB_FINDSTRING,
                -1,
                ctypes.c_wchar_p(keyword),
            )
        if index == -1:
            raise RuntimeError(f"下拉框中找不到选项: {keyword}")

        self.user32.SendMessageW(child_hwnd, self.CB_SETCURSEL, index, 0)
        self.user32.SendMessageW(
            parent_hwnd,
            self.WM_COMMAND,
            self.make_wparam(control_id, self.CBN_SELCHANGE),
            child_hwnd,
        )

    def click_button(self, parent_hwnd: int, control_id: int) -> None:
        child_hwnd = self.get_control(parent_hwnd, control_id)
        self.user32.SendMessageW(child_hwnd, self.BM_CLICK, 0, 0)

class GameTool:
    def __init__(self, config: Dict[str, Any]) -> None:
        self.config = config
        self.server = config["server"]
        self.paths = config["paths"]
        self.behavior = config["behavior"]
        self.qiannian = config.get("qiannian", {})
        self.control_ids = self.qiannian.get("control_ids", {})

        self.base_url = str(self.server["base_url"]).rstrip("/")
        self.agent_id = str(self.server["agent_id"]).strip()
        self.auth_token = str(self.server.get("auth_token", "")).strip()
        self.use_query_token = bool(self.server.get("use_query_token", False))
        self.timeout_seconds = max(5, parse_int(self.server.get("timeout_seconds"), 15))

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

        self.control_poll_seconds = max(5, parse_int(self.behavior.get("control_poll_seconds"), 15))
        self.control_error_retry_seconds = max(5, parse_int(self.behavior.get("control_error_retry_seconds"), 30))
        self.process_stop_timeout_seconds = max(5, parse_int(self.behavior.get("process_stop_timeout_seconds"), 20))
        self.window_find_timeout_seconds = max(5, parse_int(self.behavior.get("window_find_timeout_seconds"), 60))
        self.launch_ready_seconds = max(0, parse_int(self.behavior.get("launch_ready_seconds"), 20))
        self.post_load_delay_seconds = max(0, parse_int(self.behavior.get("post_load_delay_seconds"), 2))
        self.launch_settle_seconds = max(1, parse_int(self.behavior.get("launch_settle_seconds"), 8))
        self.startup_grace_seconds_fallback = max(
            30,
            parse_int(self.behavior.get("startup_grace_seconds_fallback"), 300),
        )

        self.ui_enabled = bool(self.qiannian.get("ui_enabled", True))
        self.default_launch_button = str(self.qiannian.get("launch_button", "gongzi")).strip().lower() or "gongzi"
        self.dialog_controller = Win32DialogController() if self.ui_enabled else None

    def resolve_path(self, raw_path: str) -> Path:
        path = Path(raw_path)
        if path.is_absolute():
            return path
        return SCRIPT_DIR / path

    def ensure_dirs(self) -> None:
        for directory in [self.cache_dir, self.downloads_dir, self.backups_dir, self.runtime_dir]:
            directory.mkdir(parents=True, exist_ok=True)

    def load_state(self) -> Dict[str, Any]:
        state = load_json_file(self.state_file, default={}) or {}
        if not isinstance(state, dict):
            return {}
        return state

    def save_state(self, state: Dict[str, Any]) -> None:
        save_json_file(self.state_file, state)

    def get_runtime_state(self, state: Dict[str, Any]) -> Dict[str, Any]:
        runtime = state.setdefault("agent_runtime", {})
        if not isinstance(runtime, dict):
            runtime = {}
            state["agent_runtime"] = runtime
        return runtime

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
        return self.request_json(self.build_url("/api/bootstrap", {"agent_id": self.agent_id}))

    def fetch_control(self) -> Dict[str, Any]:
        if not self.agent_id:
            raise RuntimeError("配置里的 server.agent_id 不能为空")
        return self.request_json(self.build_url("/api/agent/control", {"agent_id": self.agent_id}))

    def fetch_manifest(self, bootstrap: Dict[str, Any]) -> Dict[str, Any]:
        manifest_info = bootstrap.get("downloads", {}).get("resources_manifest", {})
        manifest_url = str(manifest_info.get("url", "")).strip()
        if not manifest_url:
            manifest_url = self.build_url("/api/resources/manifest", {"agent_id": self.agent_id})
        return self.request_json(manifest_url)

    def do_bootstrap(self) -> Dict[str, Any]:
        self.ensure_dirs()
        bootstrap = self.fetch_bootstrap()
        self.print_bootstrap_summary(bootstrap)
        return bootstrap

    def do_control(self) -> Dict[str, Any]:
        control = self.fetch_control()
        print_line(json.dumps(control, ensure_ascii=False, indent=2))
        return control

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
            print_line("[SYNC] 已跳过资源下载 (behavior.download_resources=false)")
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
                    "updated_at": now_str(),
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
                print_line(f"[SYNC] 已下载 EXE 更新包，但当前仅自动替换 .exe 文件: {staged_path}")
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
        control = bootstrap.get("control", {})
        config_info = bootstrap.get("config", {})
        launch = bootstrap.get("launch", {})
        downloads = bootstrap.get("downloads", {})

        print_line("=" * 56)
        print_line(f"Agent ID: {bootstrap.get('agent_id', '-')}")
        print_line(f"任务启用: {task.get('enabled', False)}")
        print_line(f"区服: {task.get('region', '-')}")
        print_line(f"Group: {task.get('group_start', 0)} -> {task.get('group_end', 0)}")
        print_line(f"任务模式: {task.get('task_mode', '-')}")
        print_line(f"运行状态: {control.get('desired_run_state', '-')}")
        print_line(f"每日启动: {control.get('schedule_daily_start', '-') or '-'}")
        print_line(f"配置版本: {bootstrap.get('profile_version', '-')}")
        print_line(f"脚本配置版本: {config_info.get('version', '-')}")
        print_line(f"资源清单版本: {downloads.get('resources_manifest', {}).get('version', '-')}")
        print_line(f"启动 EXE: {launch.get('startup_exe', '-')}")
        print_line("=" * 56)

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
        state["manifest_version"] = bootstrap.get("downloads", {}).get("resources_manifest", {}).get("version", "")
        state["last_sync_time"] = now_str()
        self.save_state(state)

        print_line(f"[SYNC] bootstrap 已写入 -> {self.bootstrap_file}")
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

    def load_cached_bootstrap(self) -> Dict[str, Any]:
        bootstrap = load_json_file(self.bootstrap_file, default={}) or {}
        if not isinstance(bootstrap, dict) or not bootstrap:
            raise RuntimeError("本地 bootstrap.json 不存在或无效，无法执行 warm restart")
        return bootstrap

    def build_launch_command(self) -> Tuple[Path, str, List[str]]:
        launch_info = load_json_file(self.launch_file, default={}) or {}
        startup_exe = str(launch_info.get("startup_exe", "")).strip()
        startup_args = str(launch_info.get("startup_args", "")).strip()

        exe_path = self.exe_path
        if startup_exe:
            requested_path = self.resolve_path(startup_exe)
            if requested_path.exists():
                exe_path = requested_path
            elif self.exe_path.exists():
                print_line(
                    f"[LAUNCH] startup_exe 不存在，回退到 paths.exe_path: {requested_path} -> {self.exe_path}"
                )
            else:
                exe_path = requested_path
        if not exe_path.exists():
            raise RuntimeError(f"找不到要启动的 EXE: {exe_path}")

        command = [str(exe_path)]
        if startup_args:
            command.extend(shlex.split(startup_args, posix=False))
        return exe_path, exe_path.name, command

    def do_launch(self) -> int:
        return self.launch_process()

    def launch_process(self) -> int:
        exe_path, _image_name, command = self.build_launch_command()
        print_line(f"[LAUNCH] 启动程序: {exe_path}")
        try:
            process = subprocess.Popen(command, cwd=str(exe_path.parent))
        except OSError as exc:
            if getattr(exc, "winerror", None) == 740:
                raise RuntimeError(
                    "QianNian.exe 需要管理员权限。请用“管理员身份”启动当前终端，再运行 python game_tool.py run。"
                ) from exc
            raise
        return process.pid

    def resolve_image_name(self) -> str:
        try:
            _exe_path, image_name, _command = self.build_launch_command()
            return image_name
        except Exception:
            return self.exe_path.name

    def list_process_pids(self, image_name: Optional[str] = None) -> List[int]:
        target = image_name or self.resolve_image_name()
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {target}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode not in (0, 1):
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "tasklist 失败")

        pids: List[int] = []
        reader = csv.reader(io.StringIO(result.stdout))
        for row in reader:
            if not row or row[0].startswith("INFO:"):
                continue
            if str(row[0]).strip().lower() != target.lower():
                continue
            if len(row) >= 2:
                pids.append(parse_int(row[1], 0))
        return [pid for pid in pids if pid > 0]

    def is_process_running(self, image_name: Optional[str] = None) -> bool:
        return bool(self.list_process_pids(image_name))

    def terminate_processes(self, image_name: Optional[str] = None) -> bool:
        target = image_name or self.resolve_image_name()
        pids = self.list_process_pids(target)
        if not pids:
            return False

        print_line(f"[STOP] 关闭进程: {target} -> {pids}")
        subprocess.run(
            ["taskkill", "/IM", target, "/T", "/F"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )

        deadline = time.time() + self.process_stop_timeout_seconds
        while time.time() < deadline:
            if not self.list_process_pids(target):
                return True
            time.sleep(1)
        raise RuntimeError(f"关闭进程超时: {target}")

    def stop_qiannian(self, state: Optional[Dict[str, Any]], reason: str, clear_session: bool) -> None:
        stopped = self.terminate_processes()
        if state is not None:
            runtime = self.get_runtime_state(state)
            runtime["last_stop_reason"] = reason
            runtime["last_stop_time"] = now_str()
            runtime["last_pid"] = 0
            if clear_session:
                runtime["session_active"] = False
            self.save_state(state)
        if stopped:
            print_line(f"[STOP] 已关闭 qiannian, reason={reason}")
        else:
            print_line(f"[STOP] 未发现正在运行的 qiannian, reason={reason}")

    def extract_ui_overrides(self, bootstrap: Dict[str, Any]) -> Dict[str, Any]:
        config_json = bootstrap.get("config", {}).get("payload_json")
        if not isinstance(config_json, dict):
            return {}
        merged = dict(config_json)
        nested = config_json.get("qiannian_ui")
        if isinstance(nested, dict):
            merged.update(nested)
        return merged

    def pick_first_value(self, values: List[Any], default: Any = "") -> Any:
        for value in values:
            if value not in (None, ""):
                return value
        return default

    def build_ui_plan(self, bootstrap: Dict[str, Any]) -> Dict[str, Any]:
        task = bootstrap.get("task", {}) if isinstance(bootstrap.get("task", {}), dict) else {}
        control = bootstrap.get("control", {}) if isinstance(bootstrap.get("control", {}), dict) else {}
        overrides = self.extract_ui_overrides(bootstrap)
        region = self.pick_first_value([overrides.get("region"), task.get("region")], "")
        group_start = self.pick_first_value(
            [overrides.get("group_start"), overrides.get("start_group"), overrides.get("current_group"), overrides.get("group_id"), task.get("group_start")],
            0,
        )
        group_end = self.pick_first_value(
            [overrides.get("group_end"), overrides.get("max_group"), overrides.get("max_group_id"), task.get("group_end")],
            0,
        )
        role_index = self.pick_first_value(
            [overrides.get("role_index"), overrides.get("selorder"), overrides.get("start_role_index")],
            0,
        )
        launch_button = str(self.pick_first_value([overrides.get("launch_button"), self.default_launch_button], self.default_launch_button)).strip().lower() or self.default_launch_button
        return {
            "region": str(region or "").strip(),
            "group_start": parse_int(group_start, 0),
            "group_end": parse_int(group_end, 0),
            "role_index": parse_int(role_index, 0),
            "launch_button": launch_button,
            "desired_run_state": str(control.get("desired_run_state", "run")).strip().lower() or "run",
        }

    def apply_bootstrap_to_window(self, pid: int, bootstrap: Dict[str, Any]) -> None:
        if not self.ui_enabled:
            return
        if self.dialog_controller is None:
            raise RuntimeError("UI 控制器未初始化")

        plan = self.build_ui_plan(bootstrap)
        hwnd = self.dialog_controller.find_main_window(pid, self.window_find_timeout_seconds)
        region_combo_id = parse_int(self.control_ids.get("region_combo"), 1005)
        load_button_id = parse_int(self.control_ids.get("load_button"), 1007)
        current_group_edit_id = parse_int(self.control_ids.get("current_group_edit"), 1008)
        max_group_edit_id = parse_int(self.control_ids.get("max_group_edit"), 1021)
        role_index_edit_id = parse_int(self.control_ids.get("role_index_edit"), 1016)

        if plan["region"]:
            self.dialog_controller.select_combo_text(hwnd, region_combo_id, plan["region"])
        if plan["group_start"] > 0:
            self.dialog_controller.set_edit_text(hwnd, current_group_edit_id, str(plan["group_start"]))
            print_line(f"[UI] 起始组已写入控件 {current_group_edit_id}: {plan['group_start']}")
            self.dialog_controller.set_edit_text(hwnd, role_index_edit_id, str(plan["role_index"]))
            print_line(f"[UI] 角色索引已写入控件 {role_index_edit_id}: {plan['role_index']}")
            self.dialog_controller.click_button(hwnd, load_button_id)
            print_line(f"[UI] 已点击加载账号按钮 {load_button_id}")
            if self.post_load_delay_seconds > 0:
                print_line(f"[UI] 加载账号后等待 {self.post_load_delay_seconds} 秒")
                time.sleep(self.post_load_delay_seconds)
        if plan["group_end"] > 0:
            self.dialog_controller.set_edit_text(hwnd, max_group_edit_id, str(plan["group_end"]))
            print_line(f"[UI] 结束组已写入控件 {max_group_edit_id}: {plan['group_end']}")

        button_key = BUTTON_ID_MAP.get(plan["launch_button"], BUTTON_ID_MAP[self.default_launch_button])
        button_id = parse_int(self.control_ids.get(button_key), 0)
        if button_id <= 0:
            raise RuntimeError(f"未配置按钮 ID: {button_key}")
        time.sleep(0.5)
        self.dialog_controller.click_button(hwnd, button_id)
        print_line(
            f"[UI] 已设置 region={plan['region'] or '-'} group_start={plan['group_start']} role_index={plan['role_index']} group_end={plan['group_end']} button={plan['launch_button']}"
        )

    def prepare_bootstrap(self, use_sync: bool) -> Dict[str, Any]:
        return self.do_sync() if use_sync else self.load_cached_bootstrap()

    def start_session(self, state: Dict[str, Any], reason: str, use_sync: bool, count_as_restart: bool) -> bool:
        bootstrap = self.prepare_bootstrap(use_sync=use_sync)
        task = bootstrap.get("task", {}) if isinstance(bootstrap.get("task", {}), dict) else {}
        control = bootstrap.get("control", {}) if isinstance(bootstrap.get("control", {}), dict) else {}
        runtime = self.get_runtime_state(state)

        if str(control.get("desired_run_state", "run")).strip().lower() == "stop":
            print_line(f"[AGENT] 当前控制状态为 stop，跳过启动, reason={reason}")
            runtime["session_active"] = False
            self.save_state(state)
            return False
        if not task.get("enabled", True):
            print_line(f"[AGENT] 当前任务未启用，跳过启动, reason={reason}")
            runtime["session_active"] = False
            self.save_state(state)
            return False

        self.stop_qiannian(state, reason=f"pre-start:{reason}", clear_session=False)
        pid = self.launch_process()
        if self.launch_ready_seconds > 0:
            print_line(f"[LAUNCH] 等待 qiannian 完成启动初始化: {self.launch_ready_seconds} 秒")
            time.sleep(self.launch_ready_seconds)
        self.apply_bootstrap_to_window(pid, bootstrap)
        time.sleep(self.launch_settle_seconds)

        runtime["status"] = "running"
        runtime["session_active"] = True
        runtime["session_date"] = today_str()
        runtime["last_launch_time"] = now_str()
        runtime["last_launch_epoch"] = int(time.time())
        runtime["last_launch_reason"] = reason
        runtime["last_pid"] = pid
        runtime["last_seen_report_time"] = ""
        runtime["last_seen_stale"] = False
        if count_as_restart:
            if runtime.get("restart_count_date", "") != today_str():
                runtime["restart_count_date"] = today_str()
                runtime["restart_count_today"] = 0
            runtime["restart_count_today"] = parse_int(runtime.get("restart_count_today"), 0) + 1
            runtime["last_restart_time"] = now_str()
            runtime["last_restart_epoch"] = int(time.time())
            runtime["last_restart_reason"] = reason
        self.save_state(state)
        print_line(f"[AGENT] 已启动 qiannian, pid={pid}, reason={reason}")
        return True

    def ensure_restart_counter(self, runtime: Dict[str, Any]) -> None:
        if runtime.get("restart_count_date", "") != today_str():
            runtime["restart_count_date"] = today_str()
            runtime["restart_count_today"] = 0

    def can_auto_restart(self, runtime: Dict[str, Any], control: Dict[str, Any]) -> bool:
        self.ensure_restart_counter(runtime)
        max_restart_per_day = max(0, parse_int(control.get("max_restart_per_day"), 0))
        restart_cooldown_seconds = max(0, parse_int(control.get("restart_cooldown_seconds"), 0))
        if max_restart_per_day <= 0:
            return False
        if parse_int(runtime.get("restart_count_today"), 0) >= max_restart_per_day:
            return False
        last_restart_epoch = parse_int(runtime.get("last_restart_epoch"), 0)
        if last_restart_epoch > 0 and (time.time() - last_restart_epoch) < restart_cooldown_seconds:
            return False
        return True

    def should_run_daily_schedule(self, runtime: Dict[str, Any], control: Dict[str, Any]) -> bool:
        schedule = parse_hhmm(control.get("schedule_daily_start", ""))
        if schedule is None or runtime.get("last_schedule_date", "") == today_str():
            return False
        now_local = time.localtime()
        return (now_local.tm_hour * 60 + now_local.tm_min) >= (schedule[0] * 60 + schedule[1])

    def should_restart_for_missing_first_report(self, runtime: Dict[str, Any], remote_runtime: Dict[str, Any], control: Dict[str, Any]) -> bool:
        if not runtime.get("session_active", False):
            return False
        if remote_runtime.get("has_report", False):
            return False
        last_launch_epoch = parse_int(runtime.get("last_launch_epoch"), 0)
        if last_launch_epoch <= 0:
            return False
        startup_grace_seconds = max(30, parse_int(control.get("startup_grace_seconds"), self.startup_grace_seconds_fallback))
        return (time.time() - last_launch_epoch) >= startup_grace_seconds

    def handle_one_shot_action(self, state: Dict[str, Any], control_doc: Dict[str, Any]) -> bool:
        runtime = self.get_runtime_state(state)
        control = control_doc.get("control", {}) if isinstance(control_doc.get("control", {}), dict) else {}
        action = normalize_action(control.get("desired_action", ""))
        action_seq = max(0, parse_int(control.get("action_seq"), 0))
        last_action_seq = max(0, parse_int(runtime.get("last_action_seq"), 0))
        if not action or action_seq <= 0 or action_seq <= last_action_seq:
            return False

        print_line(f"[AGENT] 收到动作: action={action} seq={action_seq}")
        if action == "sync_once":
            self.do_sync()
        elif action == "start_once":
            self.start_session(state, reason="action:start_once", use_sync=True, count_as_restart=False)
        elif action == "restart_once":
            self.start_session(state, reason="action:restart_once", use_sync=False, count_as_restart=True)
        elif action == "stop_once":
            self.stop_qiannian(state, reason="action:stop_once", clear_session=True)

        runtime["last_action_seq"] = action_seq
        runtime["last_action"] = action
        runtime["last_action_time"] = now_str()
        self.save_state(state)
        return True

    def update_runtime_from_control(self, state: Dict[str, Any], control_doc: Dict[str, Any]) -> None:
        runtime = self.get_runtime_state(state)
        remote_runtime = control_doc.get("runtime", {}) if isinstance(control_doc.get("runtime", {}), dict) else {}
        runtime["last_control_time"] = now_str()
        runtime["last_control_ok"] = True
        runtime["last_control_error"] = ""
        runtime["last_server_time"] = control_doc.get("server_time", "")
        runtime["last_remote_stale"] = bool(remote_runtime.get("stale", False))
        runtime["last_remote_has_report"] = bool(remote_runtime.get("has_report", False))
        if remote_runtime.get("has_report", False):
            runtime["last_seen_report_time"] = remote_runtime.get("server_time", "")
            runtime["last_seen_report_elapsed"] = remote_runtime.get("elapsed")
            runtime["last_seen_group"] = remote_runtime.get("current_group", 0)
            runtime["last_seen_role_index"] = remote_runtime.get("role_index", 0)
        self.save_state(state)

    def mark_control_error(self, state: Dict[str, Any], exc: Exception) -> None:
        runtime = self.get_runtime_state(state)
        runtime["last_control_ok"] = False
        runtime["last_control_time"] = now_str()
        runtime["last_control_error"] = str(exc)
        self.save_state(state)

    def run_agent_loop(self) -> None:
        self.ensure_dirs()
        print_line(f"[AGENT] 启动常驻模式, agent_id={self.agent_id}")
        while True:
            state = self.load_state()
            runtime = self.get_runtime_state(state)
            self.ensure_restart_counter(runtime)
            self.save_state(state)

            try:
                control_doc = self.fetch_control()
                self.update_runtime_from_control(state, control_doc)
            except Exception as exc:
                self.mark_control_error(state, exc)
                print_line(f"[AGENT] 拉取控制信息失败: {exc}")
                time.sleep(self.control_error_retry_seconds)
                continue

            state = self.load_state()
            runtime = self.get_runtime_state(state)
            control = control_doc.get("control", {}) if isinstance(control_doc.get("control", {}), dict) else {}
            remote_runtime = control_doc.get("runtime", {}) if isinstance(control_doc.get("runtime", {}), dict) else {}
            desired_run_state = str(control.get("desired_run_state", "run")).strip().lower() or "run"

            if desired_run_state == "stop":
                self.stop_qiannian(state, reason="desired_run_state=stop", clear_session=True)
                runtime["status"] = "stopped"
                self.save_state(state)
                time.sleep(self.control_poll_seconds)
                continue

            if self.handle_one_shot_action(state, control_doc):
                time.sleep(self.control_poll_seconds)
                continue
            state = self.load_state()
            runtime = self.get_runtime_state(state)
            self.ensure_restart_counter(runtime)

            if self.should_run_daily_schedule(runtime, control):
                if self.start_session(state, reason="daily_schedule", use_sync=True, count_as_restart=False):
                    runtime = self.get_runtime_state(state)
                    runtime["last_schedule_date"] = today_str()
                    self.save_state(state)

            if bool(control.get("auto_restart_on_stale", False)):
                if self.should_restart_for_missing_first_report(runtime, remote_runtime, control):
                    if self.can_auto_restart(runtime, control):
                        self.start_session(state, reason="startup_no_report", use_sync=False, count_as_restart=True)
                    else:
                        print_line("[AGENT] 启动后未上报，但已达到自动重启阈值或冷却限制，暂不处理")
                elif remote_runtime.get("has_report", False) and remote_runtime.get("stale", False):
                    if self.can_auto_restart(runtime, control):
                        self.start_session(state, reason="report_stale", use_sync=False, count_as_restart=True)
                    else:
                        print_line("[AGENT] 检测到 stale，但已达到自动重启阈值或冷却限制，暂不处理")

            time.sleep(self.control_poll_seconds)

    def do_run(self) -> None:
        self.start_session(self.load_state(), reason="manual_run", use_sync=True, count_as_restart=False)

    def do_restart(self) -> None:
        self.start_session(self.load_state(), reason="manual_restart", use_sync=False, count_as_restart=True)


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
        raise RuntimeError(f"找不到配置文件: {CONFIG_FILE}。请先执行: python game_tool.py init")
    user_config = load_json_file(CONFIG_FILE, default={}) or {}
    if not isinstance(user_config, dict):
        raise RuntimeError("game_tool_config.json 必须是 JSON 对象")
    return merge_dict(DEFAULT_CONFIG, user_config)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="game_tool")
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("init", help="生成示例配置并创建目录")
    subparsers.add_parser("bootstrap", help="拉取 bootstrap 并打印摘要")
    subparsers.add_parser("control", help="拉取 agent control 信息")
    subparsers.add_parser("sync", help="拉取 bootstrap、manifest，并写入本地文件")
    subparsers.add_parser("launch", help="直接启动 EXE")
    subparsers.add_parser("run", help="先 sync，再启动 qiannian 并触发按钮")
    subparsers.add_parser("restart", help="使用本地缓存执行 warm restart")
    subparsers.add_parser("agent", help="常驻轮询 local_report，负责定时启动和自动恢复")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        create_example_config()
        if args.command == "init":
            created = create_runtime_config_if_missing()
            tool = GameTool(load_config())
            tool.ensure_dirs()
            print_line(f"已生成配置文件: {CONFIG_FILE}" if created else f"配置文件已存在: {CONFIG_FILE}")
            print_line(f"示例配置文件: {EXAMPLE_CONFIG_FILE}")
            print_line("请先修改 game_tool_config.json，再执行 bootstrap、control、run 或 agent")
            return 0

        tool = GameTool(load_config())
        if args.command == "bootstrap":
            tool.do_bootstrap()
            return 0
        if args.command == "control":
            tool.do_control()
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
        if args.command == "restart":
            tool.do_restart()
            return 0
        if args.command == "agent":
            tool.run_agent_loop()
            return 0
        return 1
    except KeyboardInterrupt:
        print_line("[INFO] 已收到中断信号，退出")
        return 130
    except Exception as exc:
        print_line(f"[ERROR] {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
