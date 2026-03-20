import argparse
import configparser
import datetime as dt
import csv
import ctypes
import hashlib
import io
import json
import os
import re
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


if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
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
        "heartbeat_interval_seconds": 30,
        "process_stop_timeout_seconds": 20,
        "window_find_timeout_seconds": 60,
        "launch_ready_seconds": 20,
        "post_load_delay_seconds": 2,
        "load_confirm_timeout_seconds": 3,
        "launch_settle_seconds": 8,
        "post_clear_game_delay_seconds": 5,
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
            "trade_setting_checkbox": 1018,
            "gumu_exit_checkbox": 1022,
            "log_file_checkbox": 1015,
            "log_detail_checkbox": 1017,
        },
    },
    "window_cleanup": {
        "enabled": False,
        "timeout_seconds": 20,
        "targets": []
    },
}

VALID_ONE_SHOT_ACTIONS = {"", "start_once", "restart_once", "sync_once", "stop_once"}
BUTTON_ID_MAP = {
    "start": "start_button",
    "runtask": "runtask_button",
    "gongzi": "gongzi_button",
}
CHECKBOX_ID_MAP = {
    "trade_setting": "trade_setting_checkbox",
    "gumu_exit": "gumu_exit_checkbox",
    "log_file": "log_file_checkbox",
    "log_detail": "log_detail_checkbox",
}
VALID_LAUNCH_BUTTONS = set(BUTTON_ID_MAP.keys()) | {"none"}
LOCAL_REPORT_SECTION = "LocalReport"
LOCAL_REPORT_DEFAULTS = {
    "Enable": "1",
    "Host": "127.0.0.1",
    "Port": "18080",
    "Path": "/api/report",
    "AgentId": "",
    "TimeoutMs": "2500",
}


def get_hidden_subprocess_kwargs() -> Dict[str, Any]:
    if not IS_WINDOWS:
        return {}
    kwargs: Dict[str, Any] = {}
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if creationflags:
        kwargs["creationflags"] = creationflags
    startupinfo_cls = getattr(subprocess, "STARTUPINFO", None)
    if startupinfo_cls is not None:
        startupinfo = startupinfo_cls()
        startupinfo.dwFlags |= getattr(subprocess, "STARTF_USESHOWWINDOW", 0)
        startupinfo.wShowWindow = 0
        kwargs["startupinfo"] = startupinfo
    return kwargs


def run_hidden_subprocess(command: List[str], **kwargs):
    merged = get_hidden_subprocess_kwargs()
    merged.update(kwargs)
    return subprocess.run(command, **merged)


def read_text_with_fallback(
    path: Path,
    encodings: Optional[List[str]] = None,
) -> Tuple[str, str]:
    tried = encodings or ["utf-8-sig", "gbk", "utf-16", "latin-1"]
    last_error: Optional[Exception] = None
    for encoding in tried:
        try:
            return path.read_text(encoding=encoding), encoding
        except Exception as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"无法读取文本文件: {path}")


def replace_ini_section(raw_text: str, section_name: str, section_text: str) -> str:
    normalized = raw_text.replace("\r\n", "\n")
    pattern = re.compile(
        rf"(?ms)^\[{re.escape(section_name)}\]\s*$.*?(?=^\[|\Z)"
    )
    replacement = section_text.rstrip() + "\n"
    if pattern.search(normalized):
        updated = pattern.sub(replacement, normalized, count=1)
    else:
        updated = normalized.rstrip("\n")
        if updated:
            updated += "\n\n"
        updated += replacement
    return updated.replace("\n", "\r\n")




class RemoteRequestError(RuntimeError):
    def __init__(self, message: str, kind: str = "request_failed") -> None:
        super().__init__(message)
        self.kind = kind


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


def parse_bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off"}:
        return False
    return default


def normalize_action(value: Any) -> str:
    text = str(value or "").strip().lower()
    if text in VALID_ONE_SHOT_ACTIONS:
        return text
    return ""


def normalize_launch_button(value: Any, fallback: str = "gongzi") -> str:
    text = str(value or "").strip().lower()
    normalized_fallback = str(fallback or "gongzi").strip().lower() or "gongzi"
    if text == "":
        return normalized_fallback
    if text in VALID_LAUNCH_BUTTONS:
        return text
    return normalized_fallback


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
    return json.loads(path.read_text(encoding="utf-8-sig"))


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
    BM_GETCHECK = 0x00F0
    BM_CLICK = 0x00F5
    BST_CHECKED = 1
    IDOK = 1
    IDYES = 6
    WM_COMMAND = 0x0111
    WM_SETTEXT = 0x000C
    WM_GETTEXT = 0x000D
    WM_GETTEXTLENGTH = 0x000E
    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    VK_RETURN = 0x0D
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

    def find_main_window_ready(self, pid: int, timeout_seconds: int) -> Tuple[int, bool]:
        deadline = time.time() + max(1, timeout_seconds)
        confirmed_any = False
        while time.time() < deadline:
            hwnd = self._find_main_window_once(pid)
            dialog_hwnd = self._find_top_window_once(
                pid,
                class_name="#32770",
                exclude_hwnd=hwnd if hwnd else 0,
            )
            if dialog_hwnd:
                self._confirm_dialog_once(dialog_hwnd)
                confirmed_any = True
                time.sleep(0.3)
                continue
            if hwnd:
                self.user32.ShowWindow(hwnd, self.SW_RESTORE)
                return hwnd, confirmed_any
            time.sleep(0.3)
        raise RuntimeError(f"在 {timeout_seconds} 秒内未找到进程 {pid} 的主窗口")

    def _find_main_window_once(self, pid: int) -> int:
        result: List[int] = []
        enum_proc_type = ctypes.WINFUNCTYPE(
            wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
        )

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
            raise RuntimeError(f"找不到控件，ID={control_id}")
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

    def post_click_button(self, parent_hwnd: int, control_id: int) -> None:
        child_hwnd = self.get_control(parent_hwnd, control_id)
        if not self.user32.PostMessageW(child_hwnd, self.BM_CLICK, 0, 0):
            raise RuntimeError(f"failed to post click button, control_id={control_id}")

    def get_checkbox_state(self, parent_hwnd: int, control_id: int) -> bool:
        child_hwnd = self.get_control(parent_hwnd, control_id)
        state = int(self.user32.SendMessageW(child_hwnd, self.BM_GETCHECK, 0, 0))
        return state == self.BST_CHECKED

    def set_checkbox_state(
        self, parent_hwnd: int, control_id: int, checked: bool
    ) -> None:
        desired = bool(checked)
        current = self.get_checkbox_state(parent_hwnd, control_id)
        if current != desired:
            self.click_button(parent_hwnd, control_id)
            time.sleep(0.1)
        actual = self.get_checkbox_state(parent_hwnd, control_id)
        if actual != desired:
            raise RuntimeError(
                f"checkbox state mismatch, control_id={control_id} expected={desired} actual={actual}"
            )

    def get_window_class_name(self, hwnd: int) -> str:
        buffer = ctypes.create_unicode_buffer(256)
        self.user32.GetClassNameW(hwnd, buffer, 256)
        return buffer.value

    def get_window_text(self, hwnd: int) -> str:
        length = int(self.user32.GetWindowTextLengthW(hwnd))
        buffer = ctypes.create_unicode_buffer(length + 1)
        self.user32.GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value

    def list_visible_windows(self) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        enum_proc_type = ctypes.WINFUNCTYPE(
            wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
        )

        @enum_proc_type
        def enum_proc(hwnd: int, _lparam: int) -> bool:
            if not self.user32.IsWindowVisible(hwnd):
                return True
            if self.user32.GetWindow(hwnd, self.GW_OWNER):
                return True
            proc_id = wintypes.DWORD(0)
            self.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
            result.append(
                {
                    "hwnd": hwnd,
                    "pid": int(proc_id.value),
                    "class_name": self.get_window_class_name(hwnd),
                    "title": self.get_window_text(hwnd),
                }
            )
            return True

        self.user32.EnumWindows(enum_proc, 0)
        return result

    def find_top_window(
        self,
        pid: int,
        timeout_seconds: int,
        class_name: str = "",
        exclude_hwnd: int = 0,
    ) -> int:
        deadline = time.time() + max(1, timeout_seconds)
        while time.time() < deadline:
            hwnd = self._find_top_window_once(pid, class_name, exclude_hwnd)
            if hwnd:
                return hwnd
            time.sleep(0.2)
        return 0

    def _find_top_window_once(
        self,
        pid: int,
        class_name: str = "",
        exclude_hwnd: int = 0,
    ) -> int:
        result: List[int] = []
        expected_class = str(class_name or "").strip()
        excluded = int(exclude_hwnd or 0)
        enum_proc_type = ctypes.WINFUNCTYPE(
            wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
        )

        @enum_proc_type
        def enum_proc(hwnd: int, _lparam: int) -> bool:
            proc_id = wintypes.DWORD(0)
            self.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
            if proc_id.value != pid:
                return True
            if hwnd == excluded:
                return True
            if not self.user32.IsWindowVisible(hwnd):
                return True
            if expected_class and self.get_window_class_name(hwnd) != expected_class:
                return True
            result.append(hwnd)
            return False

        self.user32.EnumWindows(enum_proc, 0)
        return result[0] if result else 0

    def _confirm_dialog_once(self, dialog_hwnd: int) -> bool:
        self.user32.ShowWindow(dialog_hwnd, self.SW_RESTORE)
        self.user32.SetForegroundWindow(dialog_hwnd)

        for button_id in (self.IDOK, self.IDYES):
            button_hwnd = self.user32.GetDlgItem(dialog_hwnd, button_id)
            if button_hwnd:
                self.user32.SendMessageW(button_hwnd, self.BM_CLICK, 0, 0)
                return True

        self.user32.PostMessageW(dialog_hwnd, self.WM_KEYDOWN, self.VK_RETURN, 0)
        self.user32.PostMessageW(dialog_hwnd, self.WM_KEYUP, self.VK_RETURN, 0)
        return True

    def confirm_message_box(
        self, pid: int, owner_hwnd: int = 0, timeout_seconds: int = 5
    ) -> str:
        detect_deadline = time.time() + max(0, timeout_seconds)
        dialog_hwnd = 0
        while time.time() < detect_deadline:
            dialog_hwnd = self._find_top_window_once(
                pid,
                class_name="#32770",
                exclude_hwnd=owner_hwnd,
            )
            if dialog_hwnd:
                break
            time.sleep(0.2)

        if not dialog_hwnd:
            return "not_found"

        close_deadline = time.time() + max(1, timeout_seconds)
        while time.time() < close_deadline:
            self._confirm_dialog_once(dialog_hwnd)
            time.sleep(0.2)
            if not self.user32.IsWindow(dialog_hwnd) or not self.user32.IsWindowVisible(dialog_hwnd):
                return "confirmed"
            next_dialog_hwnd = self._find_top_window_once(
                pid,
                class_name="#32770",
                exclude_hwnd=owner_hwnd,
            )
            if next_dialog_hwnd:
                dialog_hwnd = next_dialog_hwnd

        return "timeout"



class GameTool:
    def __init__(self, config: Dict[str, Any]) -> None:
        self.config = config
        self.server = config["server"]
        self.paths = config["paths"]
        self.behavior = config["behavior"]
        self.qiannian = config.get("qiannian", {})
        self.window_cleanup = config.get("window_cleanup", {})
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

        self.control_poll_seconds = max(
            5, parse_int(self.behavior.get("control_poll_seconds"), 15)
        )
        self.control_error_retry_seconds = max(
            5, parse_int(self.behavior.get("control_error_retry_seconds"), 30)
        )
        self.heartbeat_interval_seconds = max(
            5, parse_int(self.behavior.get("heartbeat_interval_seconds"), 30)
        )
        self.process_stop_timeout_seconds = max(
            5, parse_int(self.behavior.get("process_stop_timeout_seconds"), 20)
        )
        self.window_find_timeout_seconds = max(
            5, parse_int(self.behavior.get("window_find_timeout_seconds"), 60)
        )
        self.launch_ready_seconds = max(
            0, parse_int(self.behavior.get("launch_ready_seconds"), 20)
        )
        self.post_load_delay_seconds = max(
            0, parse_int(self.behavior.get("post_load_delay_seconds"), 2)
        )
        self.load_confirm_timeout_seconds = max(
            0, parse_int(self.behavior.get("load_confirm_timeout_seconds"), 3)
        )
        self.launch_settle_seconds = max(
            1, parse_int(self.behavior.get("launch_settle_seconds"), 8)
        )
        self.post_clear_game_delay_seconds = max(
            0,
            parse_int(self.behavior.get("post_clear_game_delay_seconds"), 5),
        )
        self.startup_grace_seconds_fallback = max(
            30,
            parse_int(self.behavior.get("startup_grace_seconds_fallback"), 300),
        )

        self.ui_enabled = bool(self.qiannian.get("ui_enabled", True))
        self.default_launch_button = normalize_launch_button(
            self.qiannian.get("launch_button", "gongzi"), "gongzi"
        )
        self.dialog_controller = Win32DialogController() if IS_WINDOWS else None
        self.window_cleanup_enabled = parse_bool(
            self.window_cleanup.get("enabled", False), False
        )
        self.window_cleanup_timeout_seconds = max(
            5,
            parse_int(
                self.window_cleanup.get("timeout_seconds"),
                self.process_stop_timeout_seconds,
            ),
        )
        self.window_cleanup_targets = self.normalize_window_cleanup_targets(
            self.window_cleanup.get("targets", [])
        )
        self.agent_started_epoch = int(time.time())

    def resolve_path(self, raw_path: str) -> Path:
        path = Path(raw_path)
        if path.is_absolute():
            return path
        return SCRIPT_DIR / path

    def ensure_dirs(self) -> None:
        for directory in [
            self.cache_dir,
            self.downloads_dir,
            self.backups_dir,
            self.runtime_dir,
        ]:
            directory.mkdir(parents=True, exist_ok=True)

    def load_state(self) -> Dict[str, Any]:
        state = load_json_file(self.state_file, default={}) or {}
        if not isinstance(state, dict):
            return {}
        return state

    def normalize_state_for_agent(self, state: Dict[str, Any]) -> Dict[str, Any]:
        saved_agent_id = str(state.get("agent_id", "")).strip()
        if saved_agent_id and saved_agent_id != self.agent_id:
            print_line(
                f"[STATE] 检测到 state.json 属于其他 agent，重置运行态: {saved_agent_id} -> {self.agent_id}"
            )
            state["agent_runtime"] = {}
            state["agent_id"] = self.agent_id
        elif not saved_agent_id:
            state["agent_id"] = self.agent_id
        return state

    def save_state(self, state: Dict[str, Any]) -> None:
        state["agent_id"] = self.agent_id
        save_json_file(self.state_file, state)

    def get_runtime_state(self, state: Dict[str, Any]) -> Dict[str, Any]:
        runtime = state.setdefault("agent_runtime", {})
        if not isinstance(runtime, dict):
            runtime = {}
            state["agent_runtime"] = runtime
        return runtime

    def normalize_runtime_for_today(self, state: Dict[str, Any]) -> bool:
        runtime = self.get_runtime_state(state)
        session_date = str(runtime.get("session_date", "")).strip()
        if not session_date or session_date == today_str():
            return False
        if self.is_process_running():
            return False

        changed = False
        if runtime.get("session_active", False):
            runtime["session_active"] = False
            changed = True
        if session_date:
            runtime["session_date"] = ""
            changed = True
        if parse_int(runtime.get("last_pid"), 0) != 0:
            runtime["last_pid"] = 0
            changed = True
        if str(runtime.get("status", "")).strip().lower() in {"running", "completed"}:
            runtime["status"] = "idle"
            changed = True

        reset_fields = {
            "last_seen_report_time": "",
            "last_seen_report_elapsed": None,
            "last_seen_group": 0,
            "last_seen_role_index": 0,
            "last_remote_stale": False,
            "last_remote_has_report": False,
            "last_remote_completed": False,
            "last_seen_stale": False,
            "last_local_completed": False,
            "last_local_group": 0,
            "last_local_role_index": 0,
            "last_local_status_date": "",
            "last_local_target_group_end": 0,
            "last_local_complete_role_index": 0,
            "last_status_exists": False,
            "last_status_mtime_epoch": 0,
            "last_status_signature": "",
            "last_progress_change_at": "",
            "last_progress_change_epoch": 0,
            "last_heartbeat_at": "",
            "last_heartbeat_epoch": 0,
            "intent": "none",
            "intent_reason": "",
            "intent_at": "",
            "intent_epoch": 0,
            "resume_task_fingerprint": "",
            "resume_task_kind": "",
            "resume_task_label": "",
        }
        for key, expected in reset_fields.items():
            if runtime.get(key) != expected:
                runtime[key] = expected
                changed = True

        if changed:
            print_line(
                f"[STATE] cleared previous-day runtime without active process: session_date={session_date} -> {today_str()}"
            )
        return changed

    def build_url(
        self, path_or_url: str, params: Optional[Dict[str, Any]] = None
    ) -> str:
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

    @staticmethod
    def unwrap_request_exception(exc: Exception) -> Exception:
        current = exc
        seen = set()
        while isinstance(current, Exception) and id(current) not in seen:
            seen.add(id(current))
            if (
                isinstance(current, urllib.error.URLError)
                and current.reason not in (None, current)
                and isinstance(current.reason, Exception)
            ):
                current = current.reason
                continue
            cause = getattr(current, "__cause__", None)
            if isinstance(cause, Exception):
                current = cause
                continue
            break
        return current

    @classmethod
    def classify_request_exception(cls, exc: Exception) -> Tuple[str, str]:
        root = cls.unwrap_request_exception(exc)
        detail = str(root or exc).strip() or exc.__class__.__name__
        winerror = getattr(root, "winerror", None)
        if isinstance(root, TimeoutError) or winerror == 10060:
            return (
                "timeout",
                f"请求 local_report 超时，请检查网络或服务状态: {detail}",
            )
        if isinstance(root, ConnectionRefusedError) or winerror == 10061:
            return (
                "connection_refused",
                f"无法连接 local_report，请检查 base_url / 服务地址: {detail}",
            )
        if isinstance(root, ConnectionResetError) or winerror == 10054:
            return (
                "connection_reset",
                f"连接 local_report 被重置，请检查服务稳定性: {detail}",
            )
        if isinstance(exc, urllib.error.URLError):
            return ("url_error", f"请求失败: {detail}")
        return ("request_failed", detail)

    def build_request_error(self, prefix: str, exc: Exception) -> RemoteRequestError:
        kind, detail = self.classify_request_exception(exc)
        return RemoteRequestError(f"{prefix}: {detail}", kind=kind)

    def request_json(self, url: str) -> Dict[str, Any]:
        request = urllib.request.Request(url, method="GET")
        if self.auth_token and not self.use_query_token:
            request.add_header("X-Auth-Token", self.auth_token)

        try:
            with urllib.request.urlopen(
                request, timeout=self.timeout_seconds
            ) as response:
                body = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "ignore")
            raise RemoteRequestError(
                f"请求失败: HTTP {exc.code} {detail}", kind=f"http_{exc.code}"
            ) from exc
        except Exception as exc:
            raise self.build_request_error("请求失败", exc) from exc

        try:
            data = json.loads(body)
        except json.JSONDecodeError as exc:
            raise RemoteRequestError(
                f"响应不是合法 JSON: {body[:200]}", kind="invalid_json"
            ) from exc

        if not isinstance(data, dict):
            raise RemoteRequestError("响应 JSON 不是对象", kind="invalid_json")
        return data

    def post_json(self, path: str, payload: Dict[str, Any]) -> Dict[str, Any]:
        url = self.build_url(path)
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(url, data=body, method="POST")
        request.add_header("Content-Type", "application/json; charset=utf-8")
        if self.auth_token and not self.use_query_token:
            request.add_header("X-Auth-Token", self.auth_token)

        try:
            with urllib.request.urlopen(
                request, timeout=self.timeout_seconds
            ) as response:
                body_text = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "ignore")
            raise RemoteRequestError(
                f"请求失败: HTTP {exc.code} {detail}", kind=f"http_{exc.code}"
            ) from exc
        except Exception as exc:
            raise self.build_request_error("请求失败", exc) from exc

        try:
            data = json.loads(body_text)
        except json.JSONDecodeError as exc:
            raise RemoteRequestError(
                f"响应不是合法 JSON: {body_text[:200]}", kind="invalid_json"
            ) from exc
        if not isinstance(data, dict):
            raise RemoteRequestError("响应 JSON 不是对象", kind="invalid_json")
        return data

    def fetch_bootstrap(self) -> Dict[str, Any]:
        if not self.agent_id:
            raise RuntimeError("配置里的 server.agent_id 不能为空")
        return self.request_json(
            self.build_url("/api/bootstrap", {"agent_id": self.agent_id})
        )

    def fetch_control(self) -> Dict[str, Any]:
        if not self.agent_id:
            raise RuntimeError("配置里的 server.agent_id 不能为空")
        return self.request_json(
            self.build_url("/api/agent/control", {"agent_id": self.agent_id})
        )

    def notify_recovering(
        self,
        reason: str,
        plan: Dict[str, Any],
        control: Dict[str, Any],
    ) -> None:
        hold_seconds = max(
            parse_int(
                control.get("startup_grace_seconds"),
                self.startup_grace_seconds_fallback,
            ),
            self.launch_ready_seconds
            + self.launch_settle_seconds
            + self.post_clear_game_delay_seconds,
        )
        payload = {
            "agent_id": self.agent_id,
            "reason": reason,
            "region": str(plan.get("region", "")),
            "current_group": parse_int(plan.get("group_start"), 0),
            "role_index": parse_int(plan.get("role_index"), 0),
            "hold_seconds": hold_seconds,
            "ts": now_str(),
        }
        try:
            self.post_json("/api/agent/recovering", payload)
            print_line(
                f"[REPORT] marked recovering: reason={reason} group={payload['current_group']} role_index={payload['role_index']} hold={hold_seconds}s"
            )
        except Exception as exc:
            print_line(f"[REPORT] failed to mark recovering: {exc}")

    def fetch_manifest(self, bootstrap: Dict[str, Any]) -> Dict[str, Any]:
        manifest_info = bootstrap.get("downloads", {}).get("resources_manifest", {})
        manifest_url = str(manifest_info.get("url", "")).strip()
        if not manifest_url:
            manifest_url = self.build_url(
                "/api/resources/manifest", {"agent_id": self.agent_id}
            )
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

    def do_status(self) -> None:
        self.ensure_dirs()
        state = self.normalize_state_for_agent(self.load_state())
        if self.normalize_runtime_for_today(state):
            self.save_state(state)
        runtime = self.get_runtime_state(state)
        self.ensure_restart_counter(runtime)
        self.save_state(state)

        print_line("=" * 56)
        print_line(f"Agent ID: {self.agent_id}")
        print_line(f"State File: {self.state_file}")
        print_line(f"status.ini current_group: {self.read_status_ini_group() or '-'}")
        print_line(f"Local Status: {runtime.get('status', '-')}")
        print_line(f"Session Active: {bool(runtime.get('session_active', False))}")
        print_line(f"Last Launch: {runtime.get('last_launch_time', '-')}")
        print_line(f"Launch Reason: {runtime.get('last_launch_reason', '-')}")
        print_line(f"Last Stop: {runtime.get('last_stop_time', '-')}")
        print_line(f"Stop Reason: {runtime.get('last_stop_reason', '-')}")
        print_line(
            f"Next Schedule Date: {runtime.get('next_schedule_date', '-') or '-'}"
        )
        print_line(
            f"Restart Count: {parse_int(runtime.get('restart_count_today'), 0)} / {runtime.get('restart_count_date', '-') or '-'}"
        )
        print_line(
            f"Intent: {runtime.get('intent', 'none')} reason={runtime.get('intent_reason', '-') or '-'} at={runtime.get('intent_at', '-') or '-'}"
        )
        print_line(f"Last Heartbeat: {runtime.get('last_heartbeat_at', '-') or '-'}")
        print_line(f"Last Local Evidence: {runtime.get('last_progress_change_at', '-') or '-'}")
        print_line(f"Last Report: {runtime.get('last_seen_report_time', '-')}")
        print_line(
            f"Last Progress: group={runtime.get('last_seen_group', '-')} role_index={runtime.get('last_seen_role_index', '-')}"
        )
        print_line(
            f"Local Completion: completed={bool(runtime.get('last_local_completed', False))} group={runtime.get('last_local_group', '-')} role_index={runtime.get('last_local_role_index', '-')} target_group_end={runtime.get('last_local_target_group_end', '-')} complete_role_index={runtime.get('last_local_complete_role_index', '-')} date={runtime.get('last_local_status_date', '-') or '-'}"
        )
        print_line(
            f"Remote Snapshot: has_report={bool(runtime.get('last_remote_has_report', False))} stale={bool(runtime.get('last_remote_stale', False))} completed={bool(runtime.get('last_remote_completed', False))}"
        )
        print_line(
            f"Control Fetch: ok={bool(runtime.get('last_control_ok', False))} time={runtime.get('last_control_time', '-') or '-'}"
        )
        if runtime.get("last_control_error"):
            print_line(f"Control Error: {runtime.get('last_control_error')}")

        try:
            control_doc = self.fetch_control()
            control = (
                control_doc.get("control", {})
                if isinstance(control_doc.get("control", {}), dict)
                else {}
            )
            remote_runtime = (
                control_doc.get("runtime", {})
                if isinstance(control_doc.get("runtime", {}), dict)
                else {}
            )
            supervision = (
                control_doc.get("supervision", {})
                if isinstance(control_doc.get("supervision", {}), dict)
                else {}
            )
            result_snapshot = (
                control_doc.get("result", {})
                if isinstance(control_doc.get("result", {}), dict)
                else {}
            )
            task = (
                control_doc.get("task", {})
                if isinstance(control_doc.get("task", {}), dict)
                else {}
            )
            local_completion = self.get_local_completion_state(task, control)
            print_line(
                f"Remote Control: run_state={control.get('desired_run_state', '-')} auto_restart={bool(control.get('auto_restart_on_stale', False))} schedule={control.get('schedule_daily_start', '-') or '-'}"
            )
            print_line(
                f"Remote Report: has_report={bool(remote_runtime.get('has_report', False))} stale={bool(remote_runtime.get('stale', False))} completed={bool(remote_runtime.get('completed', False))} elapsed={remote_runtime.get('elapsed', '-') if remote_runtime.get('elapsed') is not None else '-'}"
            )
            print_line(
                f"Local Status.ini: completed={local_completion.get('completed', False)} group={local_completion.get('current_group', 0)} role_index={local_completion.get('role_index', 0)} target_group_end={local_completion.get('target_group_end', 0)} complete_role_index={local_completion.get('complete_role_index', 0)} date={local_completion.get('status_date', '-') or '-'}"
            )
            print_line(
                f"Local Detail: {self.describe_local_status(runtime, control, task, remote_runtime=remote_runtime, local_completion=local_completion)}"
            )
            if supervision:
                print_line(
                    f"Remote Supervision: {supervision.get('state_label', '-')} detail={supervision.get('detail', '-')}"
                )
            if result_snapshot:
                print_line(
                    f"Remote Result Status: {result_snapshot.get('state_label', '-')} detail={result_snapshot.get('detail', '-')}"
                )
            if self.should_resume_pending_session(runtime, control, task):
                print_line(
                    "[HINT] 当前若直接运行 agent，会按今天未完成进度立即续跑；如果今天任务已在其他设备完成，请先执行 `game_tool.exe skip-today`。"
                )

            pending_reason = self.get_pending_restart_reason(
                runtime, remote_runtime, control, task
            )
            if pending_reason:
                if self.should_defer_auto_restart_until_schedule(runtime, control):
                    print_line(
                        f"Auto Restart: triggered={pending_reason} deferred_until_schedule next_date={runtime.get('next_schedule_date', '-') or '-'} schedule={control.get('schedule_daily_start', '-') or '-'}"
                    )
                else:
                    block_reason = self.get_auto_restart_block_reason(runtime, control)
                    if block_reason:
                        print_line(
                            f"Auto Restart: triggered={pending_reason} blocked={block_reason}"
                        )
                    else:
                        print_line(
                            f"Auto Restart: triggered={pending_reason} allowed=true"
                        )
            elif bool(control.get("auto_restart_on_stale", False)) and runtime.get(
                "session_active", False
            ):
                last_launch_epoch = parse_int(runtime.get("last_launch_epoch"), 0)
                if last_launch_epoch > 0 and not remote_runtime.get(
                    "has_report", False
                ):
                    startup_grace_seconds = max(
                        30,
                        parse_int(
                            control.get("startup_grace_seconds"),
                            self.startup_grace_seconds_fallback,
                        ),
                    )
                    elapsed = max(0, int(time.time() - last_launch_epoch))
                    remaining = max(0, startup_grace_seconds - elapsed)
                    print_line(
                        f"Startup Grace: elapsed={elapsed}s grace={startup_grace_seconds}s remaining={remaining}s"
                    )
                else:
                    print_line("Auto Restart: not triggered")
            else:
                print_line("Auto Restart: not triggered")
        except Exception as exc:
            print_line(f"Remote Control Error: {exc}")

        print_line("=" * 56)

    def do_reset_runtime(self) -> None:
        self.ensure_dirs()
        state = self.normalize_state_for_agent(self.load_state())
        state["agent_runtime"] = {}
        self.save_state(state)
        print_line(f"[STATE] 已清空 agent_runtime: {self.state_file}")

    def do_stop(self) -> None:
        self.ensure_dirs()
        state = self.normalize_state_for_agent(self.load_state())
        self.stop_qiannian(state, reason="manual_stop", clear_session=True)
        runtime = self.get_runtime_state(state)
        runtime["status"] = "stopped"
        self.set_intent(runtime, "stop_requested", "manual_stop")
        self.save_state(state)
        self.maybe_post_heartbeat(state, force=True)

    def do_skip_today(self) -> None:
        self.ensure_dirs()
        state = self.normalize_state_for_agent(self.load_state())
        if self.normalize_runtime_for_today(state):
            self.save_state(state)
        self.stop_qiannian(state, reason="manual_skip_today", clear_session=True)
        runtime = self.get_runtime_state(state)
        runtime["status"] = "stopped"
        runtime["session_active"] = False
        self.set_intent(runtime, "skip_today", "manual_skip_today")

        schedule_text = str(runtime.get("schedule_daily_start", "")).strip()
        if not schedule_text:
            try:
                control_doc = self.fetch_control()
                control = (
                    control_doc.get("control", {})
                    if isinstance(control_doc.get("control", {}), dict)
                    else {}
                )
                schedule_text = str(control.get("schedule_daily_start", "")).strip()
            except Exception as exc:
                print_line(f"[SKIP] 拉取远端 schedule 失败，继续使用本地状态: {exc}")

        schedule = parse_hhmm(schedule_text)
        if schedule is not None:
            next_date = (dt.date.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")
            runtime["schedule_daily_start"] = schedule_text
            runtime["last_schedule_date"] = today_str()
            runtime["next_schedule_date"] = next_date
            print_line(
                f"[SKIP] 已跳过今天，agent 将等待下一次定时启动: date={next_date} schedule={schedule_text}"
            )
        else:
            print_line(
                "[SKIP] 未找到有效 schedule_daily_start；已阻止今天续跑，但请确认远端定时配置是否正确"
            )

        self.save_state(state)
        self.maybe_post_heartbeat(state, force=True)

    def write_bootstrap_outputs(
        self, bootstrap: Dict[str, Any], manifest: Optional[Dict[str, Any]]
    ) -> None:
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
            with urllib.request.urlopen(
                request, timeout=self.timeout_seconds
            ) as response:
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

    def sync_resource_items(
        self, manifest: Dict[str, Any], state: Dict[str, Any]
    ) -> None:
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
                print_line(
                    f"[SYNC] 已下载 EXE 到临时文件，等待后续替换 .exe: {staged_path}"
                )
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
        print_line(
            f"资源清单版本: {downloads.get('resources_manifest', {}).get('version', '-')}"
        )
        print_line(f"启动 EXE: {launch.get('startup_exe', '-')}")
        print_line("=" * 56)

    def do_sync(self) -> Dict[str, Any]:
        self.ensure_dirs()
        state = self.normalize_state_for_agent(self.load_state())

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
            bootstrap.get("downloads", {})
            .get("resources_manifest", {})
            .get("version", "")
        )
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
            raise RuntimeError(
                "未找到 bootstrap.json 中的 launch 配置，无法 warm restart"
            )
        return bootstrap

    def build_launch_command(
        self, emit_fallback_log: bool = True
    ) -> Tuple[Path, str, List[str]]:
        launch_info = load_json_file(self.launch_file, default={}) or {}
        startup_exe = str(launch_info.get("startup_exe", "")).strip()
        startup_args = str(launch_info.get("startup_args", "")).strip()

        exe_path = self.exe_path
        if startup_exe:
            requested_path = self.resolve_path(startup_exe)
            if requested_path.exists():
                exe_path = requested_path
            elif self.exe_path.exists():
                if emit_fallback_log:
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
                    "QianNian.exe 需要管理员权限。请用管理员身份启动当前终端，然后再运行 game_tool。"
                ) from exc
            raise
        return process.pid

    def resolve_image_name(self) -> str:
        try:
            _exe_path, image_name, _command = self.build_launch_command(
                emit_fallback_log=False
            )
            return image_name
        except Exception:
            return self.exe_path.name

    def list_process_pids(self, image_name: Optional[str] = None) -> List[int]:
        target = image_name or self.resolve_image_name()
        result = run_hidden_subprocess(
            ["tasklist", "/FI", f"IMAGENAME eq {target}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode not in (0, 1):
            raise RuntimeError(
                result.stderr.strip() or result.stdout.strip() or "tasklist 失败"
            )

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

    def normalize_window_cleanup_targets(self, raw_targets: Any) -> List[Dict[str, str]]:
        if isinstance(raw_targets, dict):
            raw_targets = [raw_targets]
        if not isinstance(raw_targets, list):
            return []

        targets: List[Dict[str, str]] = []
        for index, item in enumerate(raw_targets, start=1):
            if not isinstance(item, dict):
                continue
            target = {
                "name": str(item.get("name") or f"target_{index}").strip()
                or f"target_{index}",
                "process_name": str(item.get("process_name", "")).strip(),
                "window_class": str(item.get("window_class", "")).strip(),
                "window_title": str(
                    item.get(
                        "window_title",
                        item.get("window_title_contains", item.get("title", "")),
                    )
                ).strip(),
            }
            if (
                target["process_name"]
                or target["window_class"]
                or target["window_title"]
            ):
                targets.append(target)
        return targets

    def list_process_pids_by_window_match(
        self,
        window_class: str = "",
        window_title: str = "",
    ) -> List[int]:
        if self.dialog_controller is None:
            return []

        class_keyword = str(window_class or "").strip().lower()
        title_keyword = str(window_title or "").strip().lower()
        if not class_keyword and not title_keyword:
            return []

        matched_pids = set()
        for window in self.dialog_controller.list_visible_windows():
            pid = parse_int(window.get("pid"), 0)
            if pid <= 0:
                continue
            class_name = str(window.get("class_name", "")).lower()
            title = str(window.get("title", "")).lower()
            matched = False
            if class_keyword and class_keyword in class_name:
                matched = True
            if title_keyword and title_keyword in title:
                matched = True
            if matched:
                matched_pids.add(pid)
        return sorted(matched_pids)

    def is_pid_running(self, pid: int) -> bool:
        if pid <= 0:
            return False
        result = run_hidden_subprocess(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode not in (0, 1):
            raise RuntimeError(
                result.stderr.strip() or result.stdout.strip() or "tasklist failed"
            )
        reader = csv.reader(io.StringIO(result.stdout))
        for row in reader:
            if not row or row[0].startswith("INFO:"):
                continue
            if len(row) >= 2 and parse_int(row[1], 0) == pid:
                return True
        return False

    def terminate_pid_list(self, pids: List[int], reason: str) -> bool:
        unique_pids = sorted({pid for pid in pids if pid > 0})
        if not unique_pids:
            return False

        print_line(f"[CLEANUP] close external processes ({reason}): {unique_pids}")
        for pid in unique_pids:
            run_hidden_subprocess(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
            )

        deadline = time.time() + self.window_cleanup_timeout_seconds
        remaining = list(unique_pids)
        while time.time() < deadline:
            remaining = [pid for pid in unique_pids if self.is_pid_running(pid)]
            if not remaining:
                return True
            time.sleep(1)
        raise RuntimeError(f"external cleanup timeout, remaining pids: {remaining}")

    def cleanup_external_windows(self, reason: str) -> bool:
        if not self.window_cleanup_enabled:
            print_line(f"[CLEANUP] external cleanup disabled, skip: reason={reason}")
            return False
        if not self.window_cleanup_targets:
            print_line(f"[CLEANUP] no cleanup targets configured, skip: reason={reason}")
            return False

        matched_pids = set()
        for target in self.window_cleanup_targets:
            target_pids = set()
            process_name = target.get("process_name", "")
            if process_name:
                target_pids.update(self.list_process_pids(process_name))
            window_class = target.get("window_class", "")
            window_title = target.get("window_title", "")
            if window_class or window_title:
                target_pids.update(
                    self.list_process_pids_by_window_match(window_class, window_title)
                )
            if target_pids:
                print_line(
                    f"[CLEANUP] matched target {target.get('name', '-')}: {sorted(target_pids)}"
                )
                matched_pids.update(target_pids)

        if not matched_pids:
            print_line(f"[CLEANUP] no external windows matched: reason={reason}")
            return False
        return self.terminate_pid_list(sorted(matched_pids), reason)

    def is_process_running(self, image_name: Optional[str] = None) -> bool:
        return bool(self.list_process_pids(image_name))

    def terminate_processes(self, image_name: Optional[str] = None) -> bool:
        target = image_name or self.resolve_image_name()
        pids = self.list_process_pids(target)
        if not pids:
            return False

        print_line(f"[STOP] 关闭进程: {target} -> {pids}")
        run_hidden_subprocess(
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

    def stop_qiannian(
        self, state: Optional[Dict[str, Any]], reason: str, clear_session: bool
    ) -> None:
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

    def pick_positive_int(self, values: List[Any], default: int = 0) -> int:
        for value in values:
            parsed = parse_int(value, 0)
            if parsed > 0:
                return parsed
        return default

    def get_launch_base_dir(self) -> Path:
        try:
            exe_path, _image_name, _command = self.build_launch_command(
                emit_fallback_log=False
            )
            return exe_path.parent
        except Exception:
            return self.exe_path.parent

    def get_status_ini_path(self) -> Path:
        return self.get_launch_base_dir() / "account" / "status.ini"

    def get_global_config_ini_path(self) -> Path:
        config_dir = self.get_launch_base_dir() / "config"
        preferred = config_dir / "GlobalConfig.ini"
        legacy = config_dir / "GlobalCofig.ini"
        if preferred.exists():
            return preferred
        if legacy.exists():
            return legacy
        return preferred

    def read_global_local_report_config(self) -> Dict[str, Any]:
        ini_path = self.get_global_config_ini_path()
        result = {
            "path": str(ini_path),
            "exists": ini_path.exists(),
            "encoding": "",
            "section_exists": False,
            "enable": parse_bool(LOCAL_REPORT_DEFAULTS["Enable"], True),
            "host": LOCAL_REPORT_DEFAULTS["Host"],
            "port": parse_int(LOCAL_REPORT_DEFAULTS["Port"], 18080),
            "path_value": LOCAL_REPORT_DEFAULTS["Path"],
            "agent_id": LOCAL_REPORT_DEFAULTS["AgentId"],
            "timeout_ms": parse_int(LOCAL_REPORT_DEFAULTS["TimeoutMs"], 2500),
        }
        if not ini_path.exists():
            return result

        raw_text, encoding = read_text_with_fallback(ini_path)
        result["encoding"] = encoding

        parser = configparser.ConfigParser()
        parser.optionxform = str
        parser.read_string(raw_text)
        if not parser.has_section(LOCAL_REPORT_SECTION):
            return result

        result["section_exists"] = True
        result["enable"] = parse_bool(
            parser.get(LOCAL_REPORT_SECTION, "Enable", fallback=LOCAL_REPORT_DEFAULTS["Enable"]),
            True,
        )
        result["host"] = str(
            parser.get(LOCAL_REPORT_SECTION, "Host", fallback=LOCAL_REPORT_DEFAULTS["Host"])
        ).strip()
        result["port"] = parse_int(
            parser.get(LOCAL_REPORT_SECTION, "Port", fallback=LOCAL_REPORT_DEFAULTS["Port"]),
            18080,
        )
        result["path_value"] = str(
            parser.get(LOCAL_REPORT_SECTION, "Path", fallback=LOCAL_REPORT_DEFAULTS["Path"])
        ).strip()
        result["agent_id"] = str(
            parser.get(LOCAL_REPORT_SECTION, "AgentId", fallback=LOCAL_REPORT_DEFAULTS["AgentId"])
        ).strip()
        result["timeout_ms"] = parse_int(
            parser.get(
                LOCAL_REPORT_SECTION,
                "TimeoutMs",
                fallback=LOCAL_REPORT_DEFAULTS["TimeoutMs"],
            ),
            2500,
        )
        return result

    def write_global_local_report_config(self, payload: Dict[str, Any]) -> Path:
        ini_path = self.get_global_config_ini_path()
        ini_path.parent.mkdir(parents=True, exist_ok=True)

        raw_text = ""
        encoding = "utf-8-sig"
        if ini_path.exists():
            raw_text, encoding = read_text_with_fallback(ini_path)
            backup_file(ini_path, self.backups_dir)

        values = {
            "Enable": "1" if parse_bool(payload.get("enable"), True) else "0",
            "Host": str(payload.get("host", LOCAL_REPORT_DEFAULTS["Host"])).strip(),
            "Port": str(max(1, parse_int(payload.get("port"), 18080))),
            "Path": str(payload.get("path_value", LOCAL_REPORT_DEFAULTS["Path"])).strip()
            or LOCAL_REPORT_DEFAULTS["Path"],
            "AgentId": str(payload.get("agent_id", "")).strip(),
            "TimeoutMs": str(max(100, parse_int(payload.get("timeout_ms"), 2500))),
        }
        section_lines = [f"[{LOCAL_REPORT_SECTION}]"] + [
            f"{key}={value}" for key, value in values.items()
        ]
        updated_text = replace_ini_section(raw_text, LOCAL_REPORT_SECTION, "\n".join(section_lines))
        ini_path.write_text(updated_text, encoding=encoding)
        return ini_path

    def read_status_ini_progress(self) -> Dict[str, Any]:
        ini_path = self.get_status_ini_path()
        result = {
            "path": str(ini_path),
            "exists": ini_path.exists(),
            "mtime_epoch": 0,
            "group": 0,
            "role_index": 0,
            "last_reset_date": "",
            "is_today": False,
        }
        if not ini_path.exists():
            return result

        try:
            result["mtime_epoch"] = int(ini_path.stat().st_mtime)
        except OSError:
            result["mtime_epoch"] = 0

        parser = configparser.ConfigParser()
        last_error = None
        for encoding in ("utf-8-sig", "gbk", "utf-16", "latin-1"):
            try:
                parser.read(ini_path, encoding=encoding)
                if parser.has_section("WorkStatus"):
                    break
            except Exception as exc:
                last_error = exc

        if not parser.has_section("WorkStatus"):
            if last_error is not None:
                print_line(
                    f"[RESUME] failed to read status.ini: {ini_path} -> {last_error}"
                )
            return result

        result["group"] = parse_int(
            parser.get("WorkStatus", "CurrentProcessGroupID", fallback="0"), 0
        )
        result["role_index"] = parse_int(
            parser.get("WorkStatus", "CurrentProcessRoleIndex", fallback="0"), 0
        )
        result["last_reset_date"] = str(
            parser.get("WorkStatus", "LastReSetDate", fallback="")
        ).strip()
        result["is_today"] = result["last_reset_date"] == today_str()
        return result

    def read_status_ini_group(self) -> int:
        return parse_int(self.read_status_ini_progress().get("group"), 0)

    def get_local_completion_state(
        self, task: Dict[str, Any], control: Dict[str, Any]
    ) -> Dict[str, Any]:
        progress = self.read_status_ini_progress()
        current_group = parse_int(progress.get("group"), 0)
        role_index = parse_int(progress.get("role_index"), 0)
        target_group_end = max(0, parse_int(task.get("group_end"), 0))
        complete_role_index = max(
            1,
            parse_int(control.get("complete_role_index"), 5),
        )
        completed = bool(progress.get("is_today", False)) and target_group_end > 0 and (
            current_group > target_group_end
            or (
                current_group == target_group_end
                and role_index >= complete_role_index
            )
        )
        return {
            "completed": completed,
            "current_group": current_group,
            "role_index": role_index,
            "is_today": bool(progress.get("is_today", False)),
            "status_date": str(progress.get("last_reset_date", "")).strip(),
            "target_group_end": target_group_end,
            "complete_role_index": complete_role_index,
        }

    def set_intent(self, runtime: Dict[str, Any], intent: str, reason: str) -> None:
        runtime["intent"] = str(intent or "none").strip() or "none"
        runtime["intent_reason"] = str(reason or "").strip()
        runtime["intent_at"] = now_str()
        runtime["intent_epoch"] = int(time.time())
        runtime["last_heartbeat_epoch"] = 0

    def update_progress_evidence(
        self, state: Dict[str, Any]
    ) -> Tuple[Dict[str, Any], bool]:
        runtime = self.get_runtime_state(state)
        progress = self.read_status_ini_progress()
        signature = "|".join(
            [
                "1" if progress.get("exists", False) else "0",
                str(parse_int(progress.get("group"), 0)),
                str(parse_int(progress.get("role_index"), 0)),
                str(progress.get("last_reset_date", "") or ""),
                str(parse_int(progress.get("mtime_epoch"), 0)),
            ]
        )
        changed = False
        if runtime.get("last_status_signature", "") != signature:
            runtime["last_status_signature"] = signature
            runtime["last_progress_change_at"] = now_str()
            runtime["last_progress_change_epoch"] = int(time.time())
            changed = True
        current_exists = bool(progress.get("exists", False))
        current_mtime = parse_int(progress.get("mtime_epoch"), 0)
        if runtime.get("last_status_exists") != current_exists:
            runtime["last_status_exists"] = current_exists
            changed = True
        if parse_int(runtime.get("last_status_mtime_epoch"), 0) != current_mtime:
            runtime["last_status_mtime_epoch"] = current_mtime
            changed = True
        current_group = parse_int(progress.get("group"), 0)
        current_role_index = parse_int(progress.get("role_index"), 0)
        current_status_date = str(progress.get("last_reset_date", "")).strip()
        if parse_int(runtime.get("last_local_group"), 0) != current_group:
            runtime["last_local_group"] = current_group
            changed = True
        if parse_int(runtime.get("last_local_role_index"), 0) != current_role_index:
            runtime["last_local_role_index"] = current_role_index
            changed = True
        if str(runtime.get("last_local_status_date", "")).strip() != current_status_date:
            runtime["last_local_status_date"] = current_status_date
            changed = True
        return progress, changed

    def build_heartbeat_payload(
        self, state: Dict[str, Any], progress: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        runtime = self.get_runtime_state(state)
        progress = progress or self.read_status_ini_progress()
        process_pids = self.list_process_pids()
        process_pid = process_pids[0] if process_pids else 0
        return {
            "agent_id": self.agent_id,
            "heartbeat_at": now_str(),
            "intent": str(runtime.get("intent", "none") or "none"),
            "intent_reason": str(runtime.get("intent_reason", "") or ""),
            "intent_at": str(runtime.get("intent_at", "") or ""),
            "process_exists": bool(process_pid),
            "process_pid": process_pid,
            "status_exists": bool(progress.get("exists", False)),
            "status_group": parse_int(progress.get("group"), 0),
            "status_role_index": parse_int(progress.get("role_index"), 0),
            "status_date": str(progress.get("last_reset_date", "") or ""),
            "status_mtime_epoch": parse_int(progress.get("mtime_epoch"), 0),
            "last_progress_change_at": str(
                runtime.get("last_progress_change_at", "") or ""
            ),
            "last_restart_at": str(runtime.get("last_restart_time", "") or ""),
            "restart_count_today": parse_int(runtime.get("restart_count_today"), 0),
        }

    def maybe_post_heartbeat(
        self,
        state: Dict[str, Any],
        progress: Optional[Dict[str, Any]] = None,
        force: bool = False,
    ) -> bool:
        runtime = self.get_runtime_state(state)
        now_epoch_value = int(time.time())
        last_heartbeat_epoch = parse_int(runtime.get("last_heartbeat_epoch"), 0)
        if (
            not force
            and last_heartbeat_epoch > 0
            and (now_epoch_value - last_heartbeat_epoch) < self.heartbeat_interval_seconds
        ):
            return False
        payload = self.build_heartbeat_payload(state, progress=progress)
        try:
            self.post_json("/api/agent/heartbeat", payload)
        except Exception as exc:
            print_line(f"[HEARTBEAT] failed to post heartbeat: {exc}")
            return False
        runtime["last_heartbeat_at"] = now_str()
        runtime["last_heartbeat_epoch"] = now_epoch_value
        self.save_state(state)
        return True

    def build_resume_task_identity(self, task: Dict[str, Any]) -> Dict[str, Any]:
        task = task if isinstance(task, dict) else {}
        if not task:
            return {}
        assist = task.get("assist", {}) if isinstance(task.get("assist", {}), dict) else {}
        assist_active = bool(assist.get("active", False))
        assist_role = str(assist.get("role", "") or "").strip().lower()
        region = str(task.get("region", "") or "").strip()
        task_mode = str(task.get("task_mode", "normal") or "normal").strip().lower() or "normal"
        if assist_active and assist_role == "helper":
            helper_ids_raw = assist.get("helper_agent_ids", [])
            helper_ids: List[str] = []
            if isinstance(helper_ids_raw, list):
                helper_ids = [
                    str(item or "").strip()
                    for item in helper_ids_raw
                    if str(item or "").strip()
                ]
            identity: Dict[str, Any] = {
                "kind": "assist_helper",
                "region": region,
                "task_mode": task_mode,
                "target_agent_id": str(assist.get("target_agent_id", "") or "").strip(),
                "helper_agent_id": str(assist.get("helper_agent_id", "") or "").strip(),
                "delegate_start": parse_int(assist.get("delegate_start"), 0),
                "delegate_end": parse_int(assist.get("delegate_end"), 0),
                "original_target_group_end": parse_int(
                    assist.get("original_target_group_end"), 0
                ),
                "effective_target_group_end": parse_int(
                    assist.get("effective_target_group_end"), 0
                ),
                "assist_id": parse_int(assist.get("id"), 0),
                "work_date": str(assist.get("work_date", "") or "").strip(),
                "created_at": str(assist.get("created_at", "") or "").strip(),
            }
            if helper_ids:
                identity["helper_agent_ids"] = helper_ids
            return identity

        profile_group_start = parse_int(
            task.get("profile_group_start"),
            parse_int(task.get("group_start"), 0),
        )
        profile_group_end = parse_int(
            task.get("profile_group_end"),
            parse_int(task.get("group_end"), 0),
        )
        return {
            "kind": "main_task",
            "region": region,
            "task_mode": task_mode,
            "profile_group_start": max(0, profile_group_start),
            "profile_group_end": max(0, profile_group_end),
        }

    def compute_resume_task_fingerprint(self, task: Dict[str, Any]) -> str:
        identity = self.build_resume_task_identity(task)
        if not identity:
            return ""
        payload = json.dumps(identity, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha1(payload.encode("utf-8")).hexdigest()

    def describe_resume_task(self, task: Dict[str, Any]) -> str:
        identity = self.build_resume_task_identity(task)
        if not identity:
            return "-"
        kind = str(identity.get("kind", "") or "")
        if kind == "assist_helper":
            target_agent_id = str(identity.get("target_agent_id", "") or "-")
            delegate_start = parse_int(identity.get("delegate_start"), 0)
            delegate_end = parse_int(identity.get("delegate_end"), 0)
            region = str(identity.get("region", "") or "")
            assist_id = parse_int(identity.get("assist_id"), 0)
            pieces = [f"assist:{target_agent_id} {delegate_start}->{delegate_end}"]
            if region:
                pieces.append(f"region={region}")
            if assist_id > 0:
                pieces.append(f"id={assist_id}")
            return " ".join(pieces)
        return (
            f"main:{identity.get('region', '') or '-'} "
            f"{parse_int(identity.get('profile_group_start'), 0)}->{parse_int(identity.get('profile_group_end'), 0)} "
            f"mode={identity.get('task_mode', 'normal') or 'normal'}"
        )

    def resolve_resume_from_status_ini(
        self,
        bootstrap: Dict[str, Any],
        runtime: Dict[str, Any],
        count_as_restart: bool,
    ) -> Dict[str, Any]:
        plan = self.build_ui_plan(bootstrap)
        task = (
            bootstrap.get("task", {})
            if isinstance(bootstrap.get("task", {}), dict)
            else {}
        )
        configured_group = parse_int(plan.get("group_start"), 0)
        configured_role_index = parse_int(plan.get("role_index"), 0)
        progress = self.read_status_ini_progress()
        group = parse_int(progress.get("group"), 0)
        role_index = parse_int(progress.get("role_index"), 0)
        ahead_of_config = group > configured_group or (
            group == configured_group and role_index > configured_role_index
        )
        current_task_fingerprint = self.compute_resume_task_fingerprint(task)
        current_task_label = self.describe_resume_task(task)
        current_task_kind = str(
            self.build_resume_task_identity(task).get("kind", "") or ""
        )
        stored_task_fingerprint = str(
            runtime.get("resume_task_fingerprint", "") or ""
        ).strip()
        stored_task_label = str(runtime.get("resume_task_label", "") or "").strip() or "-"
        same_task = bool(
            current_task_fingerprint
            and stored_task_fingerprint
            and current_task_fingerprint == stored_task_fingerprint
        )
        same_day_session = str(runtime.get("session_date", "")).strip() == today_str()
        should_resume = False
        same_day_resume = False
        resume_reason = ""
        resume_block_reason = ""
        resume_block_detail = ""

        if count_as_restart and group > 0 and same_task and same_day_session:
            should_resume = True
            resume_reason = "restart"
        elif progress.get("is_today", False) and group > 0 and ahead_of_config and same_task:
            should_resume = True
            same_day_resume = True
            resume_reason = "same_day_progress"
        elif group <= 0:
            resume_block_reason = "status_group_empty"
        elif not current_task_fingerprint:
            resume_block_reason = "current_task_missing"
            resume_block_detail = f"current_task={current_task_label}"
        elif not stored_task_fingerprint:
            resume_block_reason = "previous_task_missing"
            resume_block_detail = f"current_task={current_task_label}"
        elif not same_task:
            resume_block_reason = "task_changed"
            resume_block_detail = (
                f"previous_task={stored_task_label}({stored_task_fingerprint[:8] or '-'}) "
                f"current_task={current_task_label}({current_task_fingerprint[:8] or '-'})"
            )
        elif count_as_restart and not same_day_session:
            resume_block_reason = "previous_session_not_today"
        elif not progress.get("is_today", False):
            resume_block_reason = "status_not_today"
        elif not ahead_of_config:
            resume_block_reason = "status_not_ahead_of_config"

        return {
            "configured_group": configured_group,
            "configured_role_index": configured_role_index,
            "resume_group": group if should_resume else 0,
            "status_group": group,
            "status_role_index": role_index,
            "skip_load_button": should_resume,
            "same_day_resume": same_day_resume,
            "resume_reason": resume_reason,
            "resume_block_reason": resume_block_reason,
            "resume_block_detail": resume_block_detail,
            "status_date": str(progress.get("last_reset_date", "")).strip(),
            "is_today": bool(progress.get("is_today", False)),
            "current_task_fingerprint": current_task_fingerprint,
            "stored_task_fingerprint": stored_task_fingerprint,
            "current_task_label": current_task_label,
            "stored_task_label": stored_task_label,
            "current_task_kind": current_task_kind,
            "same_task": same_task,
            "same_day_session": same_day_session,
        }

    def build_ui_plan(
        self,
        bootstrap: Dict[str, Any],
        resume_group: int = 0,
        skip_load_button: bool = False,
        force_load_button: bool = False,
    ) -> Dict[str, Any]:
        task = (
            bootstrap.get("task", {})
            if isinstance(bootstrap.get("task", {}), dict)
            else {}
        )
        control = (
            bootstrap.get("control", {})
            if isinstance(bootstrap.get("control", {}), dict)
            else {}
        )
        overrides = self.extract_ui_overrides(bootstrap)
        region = self.pick_first_value(
            [task.get("region"), overrides.get("region")], ""
        )
        group_start = self.pick_positive_int(
            [
                task.get("group_start"),
                overrides.get("group_start"),
                overrides.get("start_group"),
                overrides.get("current_group"),
                overrides.get("group_id"),
            ],
            0,
        )
        group_end = self.pick_positive_int(
            [
                task.get("group_end"),
                overrides.get("group_end"),
                overrides.get("max_group"),
                overrides.get("max_group_id"),
            ],
            0,
        )
        role_index = self.pick_first_value(
            [
                overrides.get("role_index"),
                overrides.get("selorder"),
                overrides.get("start_role_index"),
            ],
            0,
        )
        launch_button = normalize_launch_button(
            self.pick_first_value(
                [overrides.get("launch_button"), self.default_launch_button],
                self.default_launch_button,
            ),
            self.default_launch_button,
        )
        checkbox_source = overrides.get("checkboxes")
        checkbox_plan: Optional[Dict[str, bool]] = None
        if isinstance(checkbox_source, dict):
            checkbox_plan = {
                key: parse_bool(checkbox_source.get(key), False)
                for key in CHECKBOX_ID_MAP.keys()
            }
        if resume_group > 0:
            group_start = resume_group
        return {
            "region": str(region or "").strip(),
            "group_start": parse_int(group_start, 0),
            "group_end": parse_int(group_end, 0),
            "role_index": parse_int(role_index, 0),
            "launch_button": launch_button,
            "checkboxes": checkbox_plan,
            "skip_load_button": bool(skip_load_button),
            "force_load_button": bool(force_load_button),
            "desired_run_state": str(control.get("desired_run_state", "run"))
            .strip()
            .lower()
            or "run",
        }

    def apply_bootstrap_to_window(
        self,
        pid: int,
        bootstrap: Dict[str, Any],
        resume_group: int = 0,
        skip_load_button: bool = False,
    ) -> Dict[str, Any]:
        if not self.ui_enabled:
            return
        if self.dialog_controller is None:
            raise RuntimeError("UI controller is not initialized")

        hwnd, startup_dialog_confirmed = self.dialog_controller.find_main_window_ready(
            pid, self.window_find_timeout_seconds
        )
        effective_resume_group = resume_group
        if startup_dialog_confirmed:
            if resume_group > 0:
                print_line(
                    f"[UI] completion dialog detected; ignore resume_group={resume_group} and restart from configured group"
                )
            effective_resume_group = 0
        plan = self.build_ui_plan(
            bootstrap,
            resume_group=effective_resume_group,
            skip_load_button=skip_load_button,
            force_load_button=startup_dialog_confirmed,
        )
        load_button_id = parse_int(self.control_ids.get("load_button"), 1007)
        region_combo_id = parse_int(self.control_ids.get("region_combo"), 1005)
        current_group_edit_id = parse_int(
            self.control_ids.get("current_group_edit"), 1008
        )
        max_group_edit_id = parse_int(self.control_ids.get("max_group_edit"), 1021)
        role_index_edit_id = parse_int(self.control_ids.get("role_index_edit"), 1016)

        if plan["region"]:
            self.dialog_controller.select_combo_text(
                hwnd, region_combo_id, plan["region"]
            )
        if plan["group_start"] > 0:
            self.dialog_controller.set_edit_text(
                hwnd, current_group_edit_id, str(plan["group_start"])
            )
            print_line(
                f"[UI] group_start edit <- {plan['group_start']} (control_id={current_group_edit_id})"
            )
            should_click_load_button = (not plan["skip_load_button"]) or bool(
                plan.get("force_load_button", False)
            )
            if should_click_load_button:
                role_index_to_write = 0 if plan.get("force_load_button", False) else plan["role_index"]
                self.dialog_controller.set_edit_text(
                    hwnd, role_index_edit_id, str(role_index_to_write)
                )
                print_line(
                    f"[UI] role_index edit <- {role_index_to_write} (control_id={role_index_edit_id})"
                )
                if plan["skip_load_button"] and plan.get("force_load_button", False):
                    print_line(
                        "[UI] startup completion dialog accepted; reset role_index to 0 and click load_button once to clear qiannian completion prompt"
                    )
                self.dialog_controller.click_button(hwnd, load_button_id)
                print_line(f"[UI] clicked load_button (control_id={load_button_id})")
                if self.load_confirm_timeout_seconds > 0:
                    confirm_status = self.dialog_controller.confirm_message_box(
                        pid,
                        owner_hwnd=hwnd,
                        timeout_seconds=self.load_confirm_timeout_seconds,
                    )
                    if confirm_status == "confirmed":
                        print_line("[UI] load_button confirm dialog accepted")
                    elif confirm_status == "timeout":
                        raise RuntimeError(
                            f"加载账号确认弹窗未能在 {self.load_confirm_timeout_seconds} 秒内自动关闭"
                        )
                if self.post_load_delay_seconds > 0:
                    print_line(f"[UI] wait after load: {self.post_load_delay_seconds}s")
                    time.sleep(self.post_load_delay_seconds)
            else:
                print_line(
                    "[UI] recovery mode: skip load_button and keep the group restored by qiannian startup"
                )
        if plan["group_end"] > 0:
            self.dialog_controller.set_edit_text(
                hwnd, max_group_edit_id, str(plan["group_end"])
            )
            print_line(
                f"[UI] group_end edit <- {plan['group_end']} (control_id={max_group_edit_id})"
            )

        checkbox_plan = plan.get("checkboxes")
        if isinstance(checkbox_plan, dict):
            for checkbox_name, control_key in CHECKBOX_ID_MAP.items():
                control_id = parse_int(self.control_ids.get(control_key), 0)
                if control_id <= 0:
                    raise RuntimeError(f"checkbox id is not configured: {control_key}")
                self.dialog_controller.set_checkbox_state(
                    hwnd,
                    control_id,
                    parse_bool(checkbox_plan.get(checkbox_name), False),
                )
                print_line(
                    f"[UI] checkbox {checkbox_name} <- {parse_bool(checkbox_plan.get(checkbox_name), False)} (control_id={control_id})"
                )

        if plan["launch_button"] == "none":
            print_line(
                f"[UI] plan region={plan['region'] or '-'} group_start={plan['group_start']} role_index={plan['role_index']} group_end={plan['group_end']} button=none(skip task button click)"
            )
            return plan

        button_key = BUTTON_ID_MAP.get(
            plan["launch_button"], BUTTON_ID_MAP[self.default_launch_button]
        )
        button_id = parse_int(self.control_ids.get(button_key), 0)
        if button_id <= 0:
            raise RuntimeError(f"button id is not configured: {button_key}")
        time.sleep(0.5)
        self.dialog_controller.click_button(hwnd, button_id)
        print_line(
            f"[UI] plan region={plan['region'] or '-'} group_start={plan['group_start']} role_index={plan['role_index']} group_end={plan['group_end']} button={plan['launch_button']}"
        )
        return plan

    def prepare_bootstrap(self, use_sync: bool) -> Dict[str, Any]:
        return self.do_sync() if use_sync else self.load_cached_bootstrap()

    def start_session(
        self, state: Dict[str, Any], reason: str, use_sync: bool, count_as_restart: bool
    ) -> bool:
        bootstrap = self.prepare_bootstrap(use_sync=use_sync)
        task = (
            bootstrap.get("task", {})
            if isinstance(bootstrap.get("task", {}), dict)
            else {}
        )
        control = (
            bootstrap.get("control", {})
            if isinstance(bootstrap.get("control", {}), dict)
            else {}
        )
        runtime = self.get_runtime_state(state)

        if str(control.get("desired_run_state", "run")).strip().lower() == "stop":
            print_line(f"[AGENT] desired_run_state=stop, skip start, reason={reason}")
            runtime["session_active"] = False
            self.save_state(state)
            return False
        if not task.get("enabled", True):
            print_line(f"[AGENT] task is disabled, skip start, reason={reason}")
            runtime["session_active"] = False
            self.save_state(state)
            return False

        self.set_intent(
            runtime,
            "restart_requested" if count_as_restart else "start_requested",
            reason,
        )
        self.save_state(state)

        resume_state = self.resolve_resume_from_status_ini(
            bootstrap,
            runtime,
            count_as_restart=count_as_restart,
        )
        resume_group = parse_int(resume_state.get("resume_group"), 0)
        skip_load_button = bool(resume_state.get("skip_load_button", False))
        status_group = parse_int(resume_state.get("status_group"), 0)
        status_role_index = parse_int(resume_state.get("status_role_index"), 0)
        status_date = str(resume_state.get("status_date", "")).strip() or "-"

        if resume_group > 0:
            print_line(
                f"[RESUME] status.ini group={resume_group} role_index={status_role_index} date={status_date}; reason={resume_state.get('resume_reason', '-') or '-'}; skip configured start group"
            )
        else:
            block_reason = str(resume_state.get("resume_block_reason", "") or "")
            block_detail = str(resume_state.get("resume_block_detail", "") or "")
            detail_suffix = f"; {block_detail}" if block_detail else ""
            if block_reason:
                print_line(
                    f"[RESUME] status.ini not used: block={block_reason}; configured_group={resume_state.get('configured_group', 0)} status_group={status_group} role_index={status_role_index} date={status_date}{detail_suffix}"
                )
            else:
                print_line(
                    f"[RESUME] status.ini not used: configured_group={resume_state.get('configured_group', 0)} status_group={status_group} role_index={status_role_index} date={status_date}"
                )

        self.stop_qiannian(state, reason=f"pre-start:{reason}", clear_session=False)

        need_external_cleanup = bool(
            count_as_restart
            or reason == "daily_rollover"
            or resume_state.get("same_day_resume", False)
        )
        if need_external_cleanup:
            self.cleanup_external_windows(reason)

        pid = self.launch_process()
        if self.launch_ready_seconds > 0:
            print_line(f"[LAUNCH] wait for qiannian init: {self.launch_ready_seconds}s")
            time.sleep(self.launch_ready_seconds)

        applied_plan = self.apply_bootstrap_to_window(
            pid,
            bootstrap,
            resume_group=resume_group,
            skip_load_button=skip_load_button,
        )
        time.sleep(self.launch_settle_seconds)
        self.notify_recovering(reason, applied_plan, control)

        runtime["status"] = "running"
        runtime["session_active"] = True
        runtime["session_date"] = today_str()
        runtime["last_launch_time"] = now_str()
        runtime["last_launch_epoch"] = int(time.time())
        runtime["last_launch_reason"] = reason
        runtime["last_pid"] = pid
        runtime["last_seen_report_time"] = ""
        runtime["last_seen_stale"] = False
        runtime["resume_task_fingerprint"] = str(
            resume_state.get("current_task_fingerprint", "") or ""
        )
        runtime["resume_task_kind"] = str(
            resume_state.get("current_task_kind", "") or ""
        )
        runtime["resume_task_label"] = str(
            resume_state.get("current_task_label", "") or ""
        )
        if count_as_restart:
            if runtime.get("restart_count_date", "") != today_str():
                runtime["restart_count_date"] = today_str()
                runtime["restart_count_today"] = 0
            runtime["restart_count_today"] = (
                parse_int(runtime.get("restart_count_today"), 0) + 1
            )
            runtime["last_restart_time"] = now_str()
            runtime["last_restart_epoch"] = int(time.time())
            runtime["last_restart_reason"] = reason
        self.save_state(state)
        self.maybe_post_heartbeat(state, force=True)
        print_line(f"[AGENT] started qiannian, pid={pid}, reason={reason}")
        return True

    def ensure_restart_counter(self, runtime: Dict[str, Any]) -> None:
        if runtime.get("restart_count_date", "") != today_str():
            runtime["restart_count_date"] = today_str()
            runtime["restart_count_today"] = 0

    def get_auto_restart_block_reason(
        self, runtime: Dict[str, Any], control: Dict[str, Any]
    ) -> str:
        self.ensure_restart_counter(runtime)
        max_restart_per_day = max(0, parse_int(control.get("max_restart_per_day"), 0))
        restart_cooldown_seconds = max(
            0, parse_int(control.get("restart_cooldown_seconds"), 0)
        )
        if max_restart_per_day <= 0:
            return "max_restart_per_day<=0，自动重启未启用"

        restart_count_today = parse_int(runtime.get("restart_count_today"), 0)
        if restart_count_today >= max_restart_per_day:
            return f"已达到当日重启上限: {restart_count_today}/{max_restart_per_day}"

        last_restart_epoch = parse_int(runtime.get("last_restart_epoch"), 0)
        if last_restart_epoch > 0:
            remaining = restart_cooldown_seconds - int(time.time() - last_restart_epoch)
            if remaining > 0:
                return f"冷却中，还需等待 {remaining} 秒 (cooldown={restart_cooldown_seconds}s)"
        return ""

    def can_auto_restart(
        self, runtime: Dict[str, Any], control: Dict[str, Any]
    ) -> bool:
        return self.get_auto_restart_block_reason(runtime, control) == ""

    def has_schedule_run_today(self, runtime: Dict[str, Any]) -> bool:
        return str(runtime.get("last_schedule_date", "")).strip() == today_str()

    def compute_next_schedule_date(
        self, runtime: Dict[str, Any], schedule: Tuple[int, int]
    ) -> str:
        if self.has_schedule_run_today(runtime):
            return (dt.date.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")
        return today_str()

    def advance_schedule_date(self, runtime: Dict[str, Any]) -> None:
        next_schedule_date = str(runtime.get("next_schedule_date", "")).strip()
        if next_schedule_date:
            try:
                base_date = dt.datetime.strptime(next_schedule_date, "%Y-%m-%d").date()
            except ValueError:
                base_date = dt.date.today()
        else:
            base_date = dt.date.today()
        runtime["last_schedule_date"] = base_date.strftime("%Y-%m-%d")
        runtime["next_schedule_date"] = (base_date + dt.timedelta(days=1)).strftime(
            "%Y-%m-%d"
        )

    def sync_schedule_runtime(
        self, runtime: Dict[str, Any], control: Dict[str, Any]
    ) -> bool:
        schedule_text = str(control.get("schedule_daily_start", "")).strip()
        schedule = parse_hhmm(schedule_text)
        saved_schedule_text = str(runtime.get("schedule_daily_start", "")).strip()
        changed = False

        if schedule is None:
            if (
                saved_schedule_text
                or str(runtime.get("next_schedule_date", "")).strip()
            ):
                runtime["schedule_daily_start"] = ""
                runtime["next_schedule_date"] = ""
                changed = True
            return changed

        desired_next_schedule_date = self.compute_next_schedule_date(runtime, schedule)
        current_next_schedule_date = str(runtime.get("next_schedule_date", "")).strip()
        if (
            saved_schedule_text != schedule_text
            or current_next_schedule_date != desired_next_schedule_date
        ):
            runtime["schedule_daily_start"] = schedule_text
            runtime["next_schedule_date"] = desired_next_schedule_date
            changed = True
        return changed

    def should_run_daily_schedule(
        self, runtime: Dict[str, Any], control: Dict[str, Any]
    ) -> bool:
        schedule = parse_hhmm(control.get("schedule_daily_start", ""))
        if schedule is None:
            return False
        if self.is_process_running():
            return False
        next_schedule_date = str(runtime.get("next_schedule_date", "")).strip()
        if next_schedule_date != today_str():
            return False
        now_local = time.localtime()
        return (now_local.tm_hour * 60 + now_local.tm_min) >= (
            schedule[0] * 60 + schedule[1]
        )

    def should_defer_auto_restart_until_schedule(
        self, runtime: Dict[str, Any], control: Dict[str, Any]
    ) -> bool:
        schedule = parse_hhmm(control.get("schedule_daily_start", ""))
        if schedule is None:
            return False
        if str(runtime.get("session_date", "")).strip() == today_str():
            return False
        next_schedule_date = str(runtime.get("next_schedule_date", "")).strip()
        if not next_schedule_date:
            return False
        if next_schedule_date != today_str():
            return True
        now_local = time.localtime()
        return (now_local.tm_hour * 60 + now_local.tm_min) < (
            schedule[0] * 60 + schedule[1]
        )

    def should_resume_pending_session(
        self, runtime: Dict[str, Any], control: Dict[str, Any], task: Dict[str, Any]
    ) -> bool:
        if str(control.get("desired_run_state", "run")).strip().lower() == "stop":
            return False
        assist = task.get("assist", {}) if isinstance(task.get("assist", {}), dict) else {}
        assist_active = bool(assist.get("active", False))
        if not task.get("enabled", True) and not assist_active:
            return False
        if self.is_process_running():
            return False
        if str(runtime.get("session_date", "")).strip() != today_str():
            return False
        if str(runtime.get("status", "")).strip().lower() in {"completed", "stopped"}:
            return False
        if bool(runtime.get("last_remote_completed", False)):
            return False
        current_task_fingerprint = self.compute_resume_task_fingerprint(task)
        stored_task_fingerprint = str(runtime.get("resume_task_fingerprint", "") or "").strip()
        if not current_task_fingerprint or not stored_task_fingerprint:
            return False
        if current_task_fingerprint != stored_task_fingerprint:
            return False
        local_completion = self.get_local_completion_state(task, control)
        if local_completion.get("completed", False):
            return False
        progress = self.read_status_ini_progress()
        return bool(progress.get("is_today", False)) and parse_int(
            progress.get("group"), 0
        ) > 0

    def describe_schedule_wait(
        self, runtime: Dict[str, Any], schedule_text: str
    ) -> str:
        schedule_text = str(schedule_text or "").strip()
        next_schedule_date = str(runtime.get("next_schedule_date", "")).strip()
        schedule = parse_hhmm(schedule_text)
        if next_schedule_date:
            if next_schedule_date == today_str():
                if schedule is None:
                    return "等待今天启动"
                now_local = time.localtime()
                current_minutes = now_local.tm_hour * 60 + now_local.tm_min
                schedule_minutes = schedule[0] * 60 + schedule[1]
                if current_minutes < schedule_minutes:
                    return f"等待今天 {schedule_text} 启动"
                return f"已到计划时间 {schedule_text}，等待 Agent 启动"
            if schedule_text:
                return f"等待 {next_schedule_date} {schedule_text} 启动"
            return f"等待 {next_schedule_date} 启动"
        if schedule_text:
            return f"等待设定时间 {schedule_text} 启动"
        return ""

    def describe_local_status(
        self,
        runtime: Dict[str, Any],
        control: Dict[str, Any],
        task: Dict[str, Any],
        remote_runtime: Optional[Dict[str, Any]] = None,
        local_completion: Optional[Dict[str, Any]] = None,
        progress: Optional[Dict[str, Any]] = None,
        qiannian_running: Optional[bool] = None,
    ) -> str:
        remote_runtime = remote_runtime or {}
        local_completion = local_completion or self.get_local_completion_state(task, control)
        progress = progress or self.read_status_ini_progress()
        if qiannian_running is None:
            qiannian_running = self.is_process_running()

        schedule_text = str(
            control.get("schedule_daily_start") or runtime.get("schedule_daily_start") or ""
        ).strip()
        wait_detail = self.describe_schedule_wait(runtime, schedule_text)
        status = str(runtime.get("status", "")).strip().lower()
        desired_run_state = (
            str(control.get("desired_run_state", "run")).strip().lower() or "run"
        )
        intent = str(runtime.get("intent", "")).strip().lower()
        intent_reason = str(runtime.get("intent_reason", "")).strip().lower()
        assist = task.get("assist", {}) if isinstance(task.get("assist", {}), dict) else {}
        assist_active = bool(assist.get("active", False))
        assist_summary = str(assist.get("summary", "") or "").strip()
        pending_restart_reason = self.get_pending_restart_reason(
            runtime, remote_runtime, control, task
        )
        restart_reason_label = {
            "startup_no_report": "启动后未收到结果上报",
            "report_stale": "结果上报超时",
        }.get(pending_restart_reason, pending_restart_reason or "")

        if desired_run_state == "stop":
            return "远端已下发停止，等待重新启用"
        if (
            remote_runtime.get("completed", False)
            or local_completion.get("completed", False)
            or status == "completed"
        ):
            if wait_detail:
                return f"本轮任务已完成，{wait_detail}"
            return "本轮任务已完成，等待下一次计划"
        if qiannian_running:
            if assist_active:
                return assist_summary or "协助任务运行中"
            return "主任务运行中"
        if self.should_resume_pending_session(runtime, control, task):
            current_group = max(0, parse_int(progress.get("group"), 0))
            if current_group > 0:
                return f"检测到今天未完成进度，等待从组 {current_group} 续跑"
            return "检测到今天未完成进度，等待续跑"
        if pending_restart_reason and self.should_defer_auto_restart_until_schedule(
            runtime, control
        ):
            if wait_detail:
                return f"{restart_reason_label}，已延后到计划时间；{wait_detail}"
            return f"{restart_reason_label}，已延后到下次计划时间"
        if intent == "skip_today" or intent_reason == "manual_skip_today":
            if wait_detail:
                return f"今天已跳过，{wait_detail}"
            return "今天已跳过，等待下一次计划"
        if pending_restart_reason:
            block_reason = self.get_auto_restart_block_reason(runtime, control)
            if block_reason:
                return f"{restart_reason_label}，但当前阻塞：{block_reason}"
            return f"{restart_reason_label}，等待自动恢复"
        if not task.get("enabled", True) and not assist_active:
            if wait_detail:
                return f"当前待命，{wait_detail}"
            return "当前待命，等待远端启用或协助任务"
        if wait_detail and not bool(runtime.get("session_active", False)):
            if assist_active:
                return f"待命机已接到协助任务，{wait_detail}"
            return wait_detail
        if assist_active:
            return assist_summary or "已接到临时协助任务"
        if status == "stopped":
            return "已停止"
        if status == "idle":
            return "空闲，等待下一次同步"
        if status == "running":
            return "等待 Agent 拉起千年"
        return "等待下一次同步"

    def should_force_daily_relaunch(
        self, runtime: Dict[str, Any], control: Dict[str, Any]
    ) -> bool:
        if str(control.get("desired_run_state", "run")).strip().lower() == "stop":
            return False
        if not self.is_process_running():
            return False
        session_date = str(runtime.get("session_date", "")).strip()
        if session_date:
            return session_date != today_str()
        last_launch_epoch = parse_int(runtime.get("last_launch_epoch"), 0)
        if last_launch_epoch <= 0:
            return False
        last_launch_date = time.strftime("%Y-%m-%d", time.localtime(last_launch_epoch))
        return last_launch_date != today_str()

    def should_restart_for_missing_first_report(
        self,
        runtime: Dict[str, Any],
        remote_runtime: Dict[str, Any],
        control: Dict[str, Any],
    ) -> bool:
        if not runtime.get("session_active", False):
            return False
        if remote_runtime.get("has_report", False):
            return False
        last_launch_epoch = parse_int(runtime.get("last_launch_epoch"), 0)
        if last_launch_epoch <= 0:
            return False
        if last_launch_epoch < self.agent_started_epoch:
            return False
        startup_grace_seconds = max(
            30,
            parse_int(
                control.get("startup_grace_seconds"),
                self.startup_grace_seconds_fallback,
            ),
        )
        return (time.time() - last_launch_epoch) >= startup_grace_seconds

    def is_previous_day_stale(self, remote_runtime: Dict[str, Any]) -> bool:
        if not (
            remote_runtime.get("has_report", False)
            and remote_runtime.get("stale", False)
        ):
            return False
        server_time = str(remote_runtime.get("server_time", "")).strip()
        if len(server_time) < 10:
            return False
        report_date = server_time[:10]
        if report_date == today_str():
            return False
        return True

    def get_pending_restart_reason(
        self,
        runtime: Dict[str, Any],
        remote_runtime: Dict[str, Any],
        control: Dict[str, Any],
        task: Dict[str, Any],
    ) -> str:
        if remote_runtime.get("completed", False):
            return ""
        local_completion = self.get_local_completion_state(task, control)
        if local_completion.get("completed", False):
            return ""
        if self.should_restart_for_missing_first_report(
            runtime, remote_runtime, control
        ):
            return "startup_no_report"
        if self.is_previous_day_stale(remote_runtime):
            return ""
        if remote_runtime.get("has_report", False) and remote_runtime.get(
            "stale", False
        ):
            return "report_stale"
        return ""

    def maybe_stop_for_completed_task(
        self,
        state: Dict[str, Any],
        runtime: Dict[str, Any],
        remote_runtime: Dict[str, Any],
        control: Dict[str, Any],
        task: Dict[str, Any],
    ) -> bool:
        if not self.is_process_running():
            return False

        local_completion = self.get_local_completion_state(task, control)
        remote_completed = bool(remote_runtime.get("completed", False))
        local_completed = bool(local_completion.get("completed", False))
        if not remote_completed and not local_completed:
            return False

        assist = task.get("assist", {}) if isinstance(task.get("assist", {}), dict) else {}
        assist_active = bool(assist.get("active", False))
        assist_role = str(assist.get("role", "") or "").strip().lower()
        current_group = parse_int(local_completion.get("current_group"), 0)
        role_index = parse_int(local_completion.get("role_index"), 0)
        target_group_end = parse_int(local_completion.get("target_group_end"), 0)
        complete_role_index = parse_int(local_completion.get("complete_role_index"), 0)

        reason_parts = ["task_completed"]
        if assist_active:
            reason_parts.append(f"assist_{assist_role or 'active'}")
        reason_parts.append("local" if local_completed else "remote")
        stop_reason = ":".join(reason_parts)

        print_line(
            f"[AGENT] task completed, stop qiannian: reason={stop_reason} "
            f"current_group={current_group} role_index={role_index} "
            f"target_group_end={target_group_end} complete_role_index={complete_role_index}"
        )
        self.stop_qiannian(state, reason=stop_reason, clear_session=True)
        runtime = self.get_runtime_state(state)
        runtime["status"] = "completed"
        runtime["session_active"] = False
        self.save_state(state)
        return True

    def handle_one_shot_action(
        self, state: Dict[str, Any], control_doc: Dict[str, Any]
    ) -> bool:
        runtime = self.get_runtime_state(state)
        control = (
            control_doc.get("control", {})
            if isinstance(control_doc.get("control", {}), dict)
            else {}
        )
        action = normalize_action(control.get("desired_action", ""))
        action_seq = max(0, parse_int(control.get("action_seq"), 0))
        last_action_seq = max(0, parse_int(runtime.get("last_action_seq"), 0))
        if not action or action_seq <= 0 or action_seq <= last_action_seq:
            return False

        print_line(f"[AGENT] 收到动作: action={action} seq={action_seq}")
        if action == "sync_once":
            self.do_sync()
        elif action == "start_once":
            self.start_session(
                state, reason="action:start_once", use_sync=True, count_as_restart=False
            )
        elif action == "restart_once":
            self.start_session(
                state,
                reason="action:restart_once",
                use_sync=False,
                count_as_restart=True,
            )
        elif action == "stop_once":
            self.stop_qiannian(state, reason="action:stop_once", clear_session=True)

        runtime["last_action_seq"] = action_seq
        runtime["last_action"] = action
        runtime["last_action_time"] = now_str()
        self.save_state(state)
        return True

    def update_runtime_from_control(
        self, state: Dict[str, Any], control_doc: Dict[str, Any]
    ) -> bool:
        runtime = self.get_runtime_state(state)
        recovered = (not bool(runtime.get("last_control_ok", True))) or (
            parse_int(runtime.get("control_error_streak"), 0) > 0
        )
        remote_runtime = (
            control_doc.get("runtime", {})
            if isinstance(control_doc.get("runtime", {}), dict)
            else {}
        )
        control = (
            control_doc.get("control", {})
            if isinstance(control_doc.get("control", {}), dict)
            else {}
        )
        task = (
            control_doc.get("task", {})
            if isinstance(control_doc.get("task", {}), dict)
            else {}
        )
        local_completion = self.get_local_completion_state(task, control)
        runtime["last_control_time"] = now_str()
        runtime["last_control_ok"] = True
        runtime["last_control_error"] = ""
        runtime["last_control_error_kind"] = ""
        runtime["control_error_streak"] = 0
        runtime["last_server_time"] = control_doc.get("server_time", "")
        runtime["last_remote_stale"] = bool(remote_runtime.get("stale", False))
        runtime["last_remote_has_report"] = bool(
            remote_runtime.get("has_report", False)
        )
        runtime["last_remote_completed"] = bool(remote_runtime.get("completed", False))
        runtime["last_local_completed"] = bool(local_completion.get("completed", False))
        runtime["last_local_group"] = parse_int(
            local_completion.get("current_group"), 0
        )
        runtime["last_local_role_index"] = parse_int(
            local_completion.get("role_index"), 0
        )
        runtime["last_local_status_date"] = str(
            local_completion.get("status_date", "")
        ).strip()
        runtime["last_local_target_group_end"] = parse_int(
            local_completion.get("target_group_end"), 0
        )
        runtime["last_local_complete_role_index"] = parse_int(
            local_completion.get("complete_role_index"), 0
        )
        if remote_runtime.get("has_report", False):
            runtime["last_seen_report_time"] = remote_runtime.get("server_time", "")
            runtime["last_seen_report_elapsed"] = remote_runtime.get("elapsed")
            runtime["last_seen_group"] = remote_runtime.get("current_group", 0)
            runtime["last_seen_role_index"] = remote_runtime.get("role_index", 0)
        if remote_runtime.get("completed", False) or local_completion.get("completed", False):
            runtime["status"] = "completed"
            runtime["session_active"] = False
        self.save_state(state)
        return recovered

    def mark_control_error(
        self, state: Dict[str, Any], exc: Exception
    ) -> Dict[str, Any]:
        runtime = self.get_runtime_state(state)
        if isinstance(exc, RemoteRequestError):
            error_text = str(exc)
            error_kind = str(getattr(exc, "kind", "request_failed")).strip() or "request_failed"
        else:
            error_kind, detail = self.classify_request_exception(exc)
            error_text = f"控制拉取失败: {detail}" if error_kind != "request_failed" else str(exc)
        same_error = (
            runtime.get("last_control_error") == error_text
            and runtime.get("last_control_error_kind") == error_kind
        )
        runtime["last_control_ok"] = False
        runtime["last_control_time"] = now_str()
        runtime["last_control_error"] = error_text
        runtime["last_control_error_kind"] = error_kind
        runtime["control_error_streak"] = (
            parse_int(runtime.get("control_error_streak"), 0) + 1 if same_error else 1
        )
        self.save_state(state)
        return runtime

    def should_emit_control_error_log(self, runtime: Dict[str, Any]) -> bool:
        streak = max(1, parse_int(runtime.get("control_error_streak"), 1))
        return streak <= 3 or streak % 10 == 0

    def build_control_error_log(self, runtime: Dict[str, Any]) -> str:
        message = str(runtime.get("last_control_error", "")).strip() or "未知错误"
        streak = max(1, parse_int(runtime.get("control_error_streak"), 1))
        if self.is_process_running():
            local_hint = "本地 qiannian 仍在运行，请优先检查 local_report 服务。"
        else:
            local_hint = "本地未检测到 qiannian 进程，请确认服务和启动链路。"
        return (
            f"[AGENT] 拉取控制信息失败: {message} {local_hint} "
            f"{self.control_error_retry_seconds} 秒后重试 (连续 {streak} 次)"
        )

    def run_agent_loop(self) -> None:
        self.ensure_dirs()
        print_line(f"[AGENT] 启动常驻模式, agent_id={self.agent_id}")
        while True:
            state = self.normalize_state_for_agent(self.load_state())
            if self.normalize_runtime_for_today(state):
                self.save_state(state)
            runtime = self.get_runtime_state(state)
            self.ensure_restart_counter(runtime)
            progress, progress_changed = self.update_progress_evidence(state)
            if progress_changed:
                self.save_state(state)
            else:
                self.save_state(state)

            try:
                control_doc = self.fetch_control()
                control_recovered = self.update_runtime_from_control(state, control_doc)
                if control_recovered:
                    print_line("[AGENT] 已恢复从 local_report 拉取控制信息")
            except Exception as exc:
                runtime = self.mark_control_error(state, exc)
                if self.should_emit_control_error_log(runtime):
                    print_line(self.build_control_error_log(runtime))
                time.sleep(self.control_error_retry_seconds)
                continue

            state = self.normalize_state_for_agent(self.load_state())
            runtime = self.get_runtime_state(state)
            progress, progress_changed = self.update_progress_evidence(state)
            if progress_changed:
                self.save_state(state)
            self.maybe_post_heartbeat(state, progress=progress)
            control = (
                control_doc.get("control", {})
                if isinstance(control_doc.get("control", {}), dict)
                else {}
            )
            remote_runtime = (
                control_doc.get("runtime", {})
                if isinstance(control_doc.get("runtime", {}), dict)
                else {}
            )
            if self.sync_schedule_runtime(runtime, control):
                self.save_state(state)
            desired_run_state = (
                str(control.get("desired_run_state", "run")).strip().lower() or "run"
            )

            if desired_run_state == "stop":
                self.stop_qiannian(
                    state, reason="desired_run_state=stop", clear_session=True
                )
                runtime["status"] = "stopped"
                self.save_state(state)
                time.sleep(self.control_poll_seconds)
                continue

            if self.handle_one_shot_action(state, control_doc):
                time.sleep(self.control_poll_seconds)
                continue
            state = self.normalize_state_for_agent(self.load_state())
            runtime = self.get_runtime_state(state)
            self.ensure_restart_counter(runtime)

            if self.should_force_daily_relaunch(runtime, control):
                if self.start_session(
                    state,
                    reason="daily_rollover",
                    use_sync=True,
                    count_as_restart=False,
                ):
                    runtime = self.get_runtime_state(state)
                    runtime["last_schedule_date"] = today_str()
                    self.save_state(state)
                time.sleep(self.control_poll_seconds)
                continue

            task = (
                control_doc.get("task", {})
                if isinstance(control_doc.get("task", {}), dict)
                else {}
            )

            if self.maybe_stop_for_completed_task(
                state, runtime, remote_runtime, control, task
            ):
                time.sleep(self.control_poll_seconds)
                continue

            if self.should_resume_pending_session(runtime, control, task):
                print_line(
                    "[AGENT] 检测到今天存在未完成进度，准备自动续跑；如果这不是预期行为，请先停止 agent 并执行 `game_tool.exe skip-today`。"
                )
                self.start_session(
                    state,
                    reason="resume_pending_session",
                    use_sync=True,
                    count_as_restart=True,
                )
                time.sleep(self.control_poll_seconds)
                continue

            if self.should_run_daily_schedule(runtime, control):
                started = self.start_session(
                    state,
                    reason="daily_schedule",
                    use_sync=True,
                    count_as_restart=False,
                )
                runtime = self.get_runtime_state(state)
                runtime["last_schedule_result"] = (
                    "started" if started else "skipped_disabled"
                )
                self.advance_schedule_date(runtime)
                self.save_state(state)

            if bool(control.get("auto_restart_on_stale", False)):
                pending_reason = self.get_pending_restart_reason(
                    runtime, remote_runtime, control, task
                )
                if pending_reason:
                    if self.should_defer_auto_restart_until_schedule(runtime, control):
                        print_line(
                            f"[AGENT] auto restart deferred until schedule: pending={pending_reason} next_date={runtime.get('next_schedule_date', '-') or '-'} schedule={control.get('schedule_daily_start', '-') or '-'} session_date={runtime.get('session_date', '-') or '-'}"
                        )
                    else:
                        block_reason = self.get_auto_restart_block_reason(
                            runtime, control
                        )
                        if not block_reason:
                            self.start_session(
                                state,
                                reason=pending_reason,
                                use_sync=False,
                                count_as_restart=True,
                            )
                        else:
                            print_line(
                                f"[AGENT] 自动重启已触发({pending_reason})，但当前阻塞: {block_reason}"
                            )
            time.sleep(self.control_poll_seconds)

    def do_run(self) -> None:
        self.start_session(
            self.normalize_state_for_agent(self.load_state()),
            reason="manual_run",
            use_sync=True,
            count_as_restart=False,
        )

    def do_restart(self) -> None:
        self.start_session(
            self.normalize_state_for_agent(self.load_state()),
            reason="manual_restart",
            use_sync=False,
            count_as_restart=True,
        )


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
            f"找不到配置文件: {CONFIG_FILE}。请先执行 python game_tool.py init"
        )
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
    subparsers.add_parser(
        "status", help="show local runtime and remote control summary"
    )
    subparsers.add_parser("sync", help="拉取 bootstrap 和 manifest，并写入本地文件")
    subparsers.add_parser("launch", help="直接启动 EXE")
    subparsers.add_parser("stop", help="stop local qiannian process and update runtime")
    subparsers.add_parser("skip-today", help="跳过今天，不续跑，等待下一次定时启动")
    subparsers.add_parser("run", help="先 sync，再启动 qiannian 并触发按钮")
    subparsers.add_parser("restart", help="使用本地缓存执行 warm restart")
    subparsers.add_parser("reset-runtime", help="clear local agent_runtime for testing")
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
            print_line(
                f"已生成配置文件: {CONFIG_FILE}"
                if created
                else f"配置文件已存在: {CONFIG_FILE}"
            )
            print_line(f"示例配置文件: {EXAMPLE_CONFIG_FILE}")
            print_line(
                "请先修改 game_tool_config.json，再执行 bootstrap、control、run 或 agent"
            )
            return 0

        tool = GameTool(load_config())
        if args.command == "bootstrap":
            tool.do_bootstrap()
            return 0
        if args.command == "control":
            tool.do_control()
            return 0
        if args.command == "status":
            tool.do_status()
            return 0
        if args.command == "sync":
            tool.do_sync()
            return 0
        if args.command == "launch":
            tool.do_launch()
            return 0
        if args.command == "stop":
            tool.do_stop()
            return 0
        if args.command == "skip-today":
            tool.do_skip_today()
            return 0
        if args.command == "run":
            tool.do_run()
            return 0
        if args.command == "restart":
            tool.do_restart()
            return 0
        if args.command == "reset-runtime":
            tool.do_reset_runtime()
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
