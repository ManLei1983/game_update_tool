import json
import json
import locale
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk
from typing import Any, Dict, List, Optional

import game_tool as core
from game_tool_gui_config import ConfigEditorWindow


CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


MOJIBAKE_MARKERS: tuple[str, ...] = ()
COMMON_CJK_HINTS = set(
    "的一是了在人有我他这中大来上个们到说国和地也子时道出而要于就下得可你年生自会那后能对着事其里所去行过家学用同控制面板刷新状态启动停止运行配置今天协助任务区服完成失败目标当前结束开始组日期计划本地远端错误重启进度时间跳过后台"
)

STATUS_COLORS = {
    "ok": "#167c2b",
    "warn": "#b26a00",
    "error": "#c62828",
    "muted": "#666666",
}


def _text_score(value: str) -> int:
    cjk_count = sum(1 for ch in value if "一" <= ch <= "鿿")
    common_count = sum(1 for ch in value if ch in COMMON_CJK_HINTS)
    replacement_count = value.count("�")
    question_count = value.count("?")
    return common_count * 4 + cjk_count - replacement_count * 8 - question_count * 4


def repair_text(value: Any) -> str:
    text = str(value or "")
    if not text:
        return ""
    candidates = [text]
    for src, dst in (("gbk", "utf-8"), ("latin-1", "utf-8"), ("cp1252", "utf-8")):
        try:
            candidates.append(text.encode(src).decode(dst))
        except Exception:
            continue
    best = max(candidates, key=_text_score)
    return best if _text_score(best) > _text_score(text) else text


def repair_payload(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: repair_payload(item) for key, item in value.items()}
    if isinstance(value, list):
        return [repair_payload(item) for item in value]
    if isinstance(value, tuple):
        return [repair_payload(item) for item in value]
    if isinstance(value, str):
        return repair_text(value)
    return value


def decode_cli_output(raw: bytes) -> str:
    if not raw:
        return ""
    encodings: List[str] = ["utf-8-sig"]
    preferred = locale.getpreferredencoding(False) or ""
    if preferred and preferred.lower() not in {item.lower() for item in encodings}:
        encodings.append(preferred)
    for extra in ("gbk", "utf-16", "latin-1"):
        if extra.lower() not in {item.lower() for item in encodings}:
            encodings.append(extra)
    decoded = ""
    for encoding in encodings:
        try:
            decoded = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if not decoded:
        decoded = raw.decode(encodings[0], errors="replace")
    return repair_text(decoded)


class GameToolGui:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("game_tool 控制面板")
        self.root.geometry("1380x960")
        self.root.minsize(1180, 768)

        core.create_example_config()
        if not core.CONFIG_FILE.exists():
            core.create_runtime_config_if_missing()

        self.refresh_running = False
        self.refresh_after_id: Optional[str] = None
        self.agent_launching = False
        self.last_state_warning = ""
        self.auto_refresh_var = tk.BooleanVar(value=True)
        self.refresh_interval_var = tk.IntVar(value=5)
        self.log_line_limit = 400
        self._overview_canvas: Optional[tk.Canvas] = None
        self.header_agent_id_var = tk.StringVar(value="-")

        self.summary_vars: Dict[str, tk.StringVar] = {}
        self.startup_check_vars: Dict[str, tk.StringVar] = {}
        self.startup_check_labels: Dict[str, tk.Label] = {}
        self.startup_check_titles: Dict[str, str] = {
            "config": "配置",
            "state": "STATE",
            "bootstrap": "BOOT",
            "exe": "EXE",
            "agent": "AGENT",
        }
        core.print_line = self._forward_core_log
        self.config_window: Optional[ConfigEditorWindow] = None
        self._build_ui()
        self._schedule_log("GUI 已就绪")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(200, self.refresh_snapshot)
        self._bind_mousewheel_for_overview()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)
        self.root.rowconfigure(3, weight=0)

        top_bar = ttk.Frame(self.root, padding=(12, 8))
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.columnconfigure(0, weight=1)
        top_bar.columnconfigure(1, weight=0)

        header_info = ttk.Frame(top_bar)
        header_info.grid(row=0, column=0, sticky="w")
        ttk.Label(header_info, text="当前 AgentID:").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header_info,
            textvariable=self.header_agent_id_var,
            font=("Segoe UI", 11, "bold"),
        ).grid(row=0, column=1, sticky="w", padx=(6, 0))

        toolbar = ttk.Frame(top_bar)
        toolbar.grid(row=0, column=1, sticky="e")
        for idx in range(11):
            toolbar.columnconfigure(idx, weight=0)

        ttk.Button(
            toolbar,
            text="刷新状态",
            command=lambda: self.refresh_snapshot(log_success=True),
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(toolbar, text="启动 Agent", command=self.start_agent).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(
            toolbar,
            text="跳过今天并启动",
            command=self.skip_today_and_start_agent,
        ).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(toolbar, text="停止 Agent", command=self.stop_agent).grid(
            row=0, column=3, padx=(0, 8)
        )
        ttk.Button(toolbar, text="停止千年", command=self.stop_qiannian).grid(
            row=0, column=4, padx=(0, 8)
        )
        ttk.Button(toolbar, text="立即同步配置", command=self.sync_once).grid(
            row=0, column=5, padx=(0, 8)
        )
        ttk.Button(toolbar, text="同步并运行一次", command=self.run_once).grid(
            row=0, column=6, padx=(0, 8)
        )
        ttk.Button(toolbar, text="配置面板", command=self.open_config_window).grid(
            row=0, column=7, padx=(0, 8)
        )
        ttk.Checkbutton(
            toolbar,
            text="自动刷新",
            variable=self.auto_refresh_var,
            command=self._on_toggle_auto_refresh,
        ).grid(row=0, column=8, padx=(12, 6))
        ttk.Label(toolbar, text="刷新间隔(秒)").grid(row=0, column=9, sticky="e")
        ttk.Spinbox(
            toolbar,
            from_=3,
            to=60,
            textvariable=self.refresh_interval_var,
            width=6,
            command=self._on_toggle_auto_refresh,
        ).grid(row=0, column=10, sticky="w")

        quick_bar = ttk.LabelFrame(
            self.root,
            text="当前协助 / 本机快捷入口",
            padding=(10, 4),
        )
        quick_bar.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 6))
        quick_bar.columnconfigure(1, weight=1)

        self.assist_status_var = tk.StringVar(value="当前未接到临时协助任务")

        ttk.Label(quick_bar, text="当前协助").grid(
            row=0, column=0, sticky="nw", padx=(0, 8)
        )
        ttk.Label(
            quick_bar,
            textvariable=self.assist_status_var,
            justify=tk.LEFT,
            wraplength=1160,
        ).grid(row=0, column=1, sticky="ew")

        ttk.Label(quick_bar, text="快捷操作").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(4, 0)
        )
        quick_actions = ttk.Frame(quick_bar)
        quick_actions.grid(row=1, column=1, sticky="w", pady=(4, 0))
        for idx in range(7):
            quick_actions.columnconfigure(idx, weight=0)
        ttk.Button(
            quick_actions,
            text="game_tool 配置",
            command=self.open_game_tool_config,
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")
        ttk.Button(
            quick_actions,
            text="GlobalConfig.ini",
            command=self.open_global_config_ini,
        ).grid(row=0, column=1, padx=(0, 8), sticky="w")
        ttk.Button(
            quick_actions,
            text="state.json",
            command=self.open_state_file,
        ).grid(row=0, column=2, padx=(0, 8), sticky="w")
        ttk.Button(
            quick_actions,
            text="bootstrap.json",
            command=self.open_bootstrap_file,
        ).grid(row=0, column=3, padx=(0, 8), sticky="w")
        ttk.Button(
            quick_actions,
            text="QianNian 目录",
            command=self.open_qiannian_dir,
        ).grid(row=0, column=4, padx=(0, 8), sticky="w")
        ttk.Button(
            quick_actions,
            text="account 目录",
            command=self.open_account_dir,
        ).grid(row=0, column=5, sticky="w")
        ttk.Button(
            quick_actions,
            text="agent.log",
            command=self.open_agent_log,
        ).grid(row=0, column=6, padx=(8, 0), sticky="w")

        ttk.Label(quick_bar, text="启动自检").grid(
            row=2, column=0, sticky="nw", padx=(0, 8), pady=(6, 0)
        )
        startup_checks = ttk.Frame(quick_bar)
        startup_checks.grid(row=2, column=1, sticky="ew", pady=(6, 0))
        for idx, (key, title) in enumerate(self.startup_check_titles.items()):
            startup_checks.columnconfigure(idx, weight=1)
            var = tk.StringVar(value=f"{title}: -")
            self.startup_check_vars[key] = var
            value_label = tk.Label(
                startup_checks,
                textvariable=var,
                anchor="w",
                justify=tk.LEFT,
                wraplength=220,
                padx=8,
                pady=4,
                bd=1,
                relief=tk.GROOVE,
                fg=STATUS_COLORS["muted"],
            )
            value_label.grid(
                row=0,
                column=idx,
                sticky="ew",
                padx=(0, 8) if idx < len(self.startup_check_titles) - 1 else (0, 0),
            )
            self.startup_check_labels[key] = value_label

        overview_host = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        overview_host.grid(row=2, column=0, sticky="nsew")
        overview_host.columnconfigure(0, weight=1)
        overview_host.rowconfigure(0, weight=1)

        overview_canvas = tk.Canvas(overview_host, highlightthickness=0, borderwidth=0)
        self._overview_canvas = overview_canvas
        overview_canvas.grid(row=0, column=0, sticky="nsew")
        overview_scrollbar = ttk.Scrollbar(
            overview_host,
            orient=tk.VERTICAL,
            command=overview_canvas.yview,
        )
        overview_scrollbar.grid(row=0, column=1, sticky="ns")
        overview_canvas.configure(yscrollcommand=overview_scrollbar.set)

        overview = ttk.Frame(overview_canvas)
        overview_window = overview_canvas.create_window(
            (0, 0), window=overview, anchor="nw"
        )
        overview.bind(
            "<Configure>",
            lambda event: overview_canvas.configure(
                scrollregion=overview_canvas.bbox("all")
            ),
        )
        overview_canvas.bind(
            "<Configure>",
            lambda event: overview_canvas.itemconfigure(
                overview_window, width=event.width
            ),
        )

        overview.columnconfigure(0, weight=1)
        overview.columnconfigure(1, weight=1)
        overview.rowconfigure(0, weight=1)
        overview.rowconfigure(1, weight=1)

        left_top = ttk.LabelFrame(overview, text="任务配置", padding=6)
        left_top.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        right_top = ttk.LabelFrame(overview, text="本地状态", padding=6)
        right_top.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))
        left_bottom = ttk.LabelFrame(overview, text="远端状态", padding=6)
        left_bottom.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(6, 0))
        right_bottom = ttk.LabelFrame(overview, text="执行判断", padding=6)
        right_bottom.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(6, 0))

        self._build_kv_grid(
            left_top,
            [
                ("Agent ID", "agent_id"),
                ("区服", "region"),
                ("生效范围", "group_range"),
                ("原配置范围", "profile_group_range"),
                ("边界变化", "task_boundary_note"),
                ("计划时间", "schedule_daily_start"),
                ("期望状态", "desired_run_state"),
                ("自动恢复", "auto_restart"),
                ("Server", "base_url"),
                ("临时协助", "assist_summary"),
                ("启动 EXE", "exe_path"),
            ],
            wraplength=300,
            wrap_max_by_key={"task_boundary_note": 300, "exe_path": 280},
        )
        self._build_kv_grid(
            right_top,
            [
                ("本地状态", "local_status"),
                ("状态说明", "local_status_detail"),
                ("STATE异常", "state_warning"),
                ("Agent状态", "agent_health"),
                ("Agent说明", "agent_health_detail"),
                ("Session Active", "session_active"),
                ("千年进程", "qiannian_running"),
                ("本地心跳", "last_heartbeat_time"),
                ("本地证据", "local_progress_evidence"),
                ("status.ini", "status_ini"),
                ("下次计划", "next_schedule_date"),
                ("最近启动", "last_launch_time"),
                ("启动原因", "last_launch_reason"),
                ("最近停止", "last_stop_time"),
                ("停止原因", "last_stop_reason"),
                ("Agent进程", "agent_process"),
            ],
            wraplength=380,
            wrap_max_by_key={
                "state_warning": 320,
                "agent_health_detail": 320,
                "local_status_detail": 320,
                "local_progress_evidence": 320,
            },
        )
        self._build_kv_grid(
            left_bottom,
            [
                ("监督状态", "remote_supervision_state"),
                ("监督说明", "remote_supervision_detail"),
                ("结果状态", "remote_result_state"),
                ("结果说明", "remote_result_detail"),
                ("远端心跳", "remote_heartbeat"),
                ("心跳证据", "remote_heartbeat_progress"),
                ("结果进度", "remote_result_progress"),
                ("远端事件", "remote_event"),
                ("拉取错误", "control_error"),
            ],
            wraplength=380,
            wrap_max_by_key={
                "remote_supervision_detail": 320,
                "remote_result_detail": 320,
                "remote_heartbeat_progress": 320,
                "remote_result_progress": 320,
                "control_error": 300,
            },
        )
        self._build_kv_grid(
            right_bottom,
            [
                ("本地完成", "local_completed"),
                ("目标结束组", "local_target_group_end"),
                ("完成阈值", "local_complete_role_index"),
                ("今天续跑", "should_resume_today"),
                ("重启计数", "restart_counter"),
                ("本地日期", "status_date"),
                ("最近进度", "last_progress"),
                ("待重启原因", "pending_restart_reason"),
                ("链路对比", "evidence_comparison"),
                ("处理建议", "evidence_hint"),
            ],
            wraplength=380,
            wrap_max_by_key={
                "evidence_comparison": 320,
                "evidence_hint": 320,
            },
        )

        bottom = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        bottom.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 12))

        raw_frame = ttk.LabelFrame(bottom, text="原始快照", padding=8)
        log_frame = ttk.LabelFrame(bottom, text="GUI 日志", padding=8)
        bottom.add(raw_frame, weight=3)
        bottom.add(log_frame, weight=2)

        raw_frame.rowconfigure(0, weight=1)
        raw_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.raw_text = scrolledtext.ScrolledText(
            raw_frame, wrap=tk.WORD, font=("Consolas", 10), height=8
        )
        self.raw_text.grid(row=0, column=0, sticky="nsew")
        self.raw_text.configure(state=tk.DISABLED)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=("Consolas", 10), height=8
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state=tk.DISABLED)

    def _refresh_header_agent_id(self, value: Any = None) -> None:
        if value in (None, ""):
            try:
                config = core.load_config()
                value = (
                    config.get("server", {}).get("agent_id", "")
                    if isinstance(config, dict)
                    else ""
                )
            except Exception:
                value = ""
        display = repair_text(str(value or "-").strip() or "-")
        self.header_agent_id_var.set(display)

    def _on_local_config_changed(self) -> None:
        self._refresh_header_agent_id()
        self.refresh_snapshot()

    def _build_kv_grid(
        self,
        parent: ttk.LabelFrame,
        rows: List[tuple[str, str]],
        columns: int = 1,
        wraplength: Optional[int] = None,
        full_width_keys: Optional[set[str]] = None,
        full_width_wraplength: Optional[int] = None,
        wrap_max_by_key: Optional[Dict[str, int]] = None,
    ) -> None:
        columns = max(1, int(columns or 1))
        full_width_keys = set(full_width_keys or set())
        wrap_max_by_key = dict(wrap_max_by_key or {})
        total_columns = columns * 2
        for column_group in range(columns):
            label_column = column_group * 2
            value_column = label_column + 1
            parent.columnconfigure(label_column, weight=0)
            parent.columnconfigure(value_column, weight=1)

        effective_wraplength = (
            wraplength if wraplength is not None else (460 if columns == 1 else 220)
        )
        effective_full_width_wraplength = (
            full_width_wraplength
            if full_width_wraplength is not None
            else max(520, effective_wraplength * max(2, columns))
        )

        current_row = 0
        current_slot = 0
        for label_text, key in rows:
            is_full_width = columns > 1 and key in full_width_keys
            if is_full_width and current_slot:
                current_row += 1
                current_slot = 0

            if is_full_width:
                ttk.Label(parent, text=label_text).grid(
                    row=current_row,
                    column=0,
                    sticky="nw",
                    padx=(0, 12),
                    pady=2,
                )
                var = tk.StringVar(value="-")
                self.summary_vars[key] = var
                value_label = ttk.Label(
                    parent,
                    textvariable=var,
                    anchor="w",
                    justify=tk.LEFT,
                    wraplength=effective_full_width_wraplength,
                )
                value_label.grid(
                    row=current_row,
                    column=1,
                    columnspan=total_columns - 1,
                    sticky="ew",
                    pady=2,
                )
                self._bind_dynamic_wrap(
                    value_label,
                    fallback=effective_full_width_wraplength,
                    max_width=wrap_max_by_key.get(key),
                )
                current_row += 1
                current_slot = 0
                continue

            label_column = current_slot * 2
            value_column = label_column + 1
            label_padx = (0, 12) if current_slot == 0 else (18, 12)
            value_padx = (0, 6) if current_slot == 0 else (0, 0)
            ttk.Label(parent, text=label_text).grid(
                row=current_row,
                column=label_column,
                sticky="nw",
                padx=label_padx,
                pady=2,
            )
            var = tk.StringVar(value="-")
            self.summary_vars[key] = var
            value_label = ttk.Label(
                parent,
                textvariable=var,
                anchor="w",
                justify=tk.LEFT,
                wraplength=effective_wraplength,
            )
            value_label.grid(
                row=current_row,
                column=value_column,
                sticky="ew",
                padx=value_padx,
                pady=2,
            )
            self._bind_dynamic_wrap(
                value_label,
                fallback=effective_wraplength,
                max_width=wrap_max_by_key.get(key),
            )
            current_slot += 1
            if current_slot >= columns:
                current_row += 1
                current_slot = 0

    def _bind_dynamic_wrap(
        self,
        widget: ttk.Label,
        fallback: int,
        max_width: Optional[int] = None,
    ) -> None:
        widget.configure(wraplength=fallback)
        widget.bind(
            "<Configure>",
            lambda event, target=widget, default_wrap=fallback, max_wrap=max_width: self._update_dynamic_wrap(
                target,
                default_wrap,
                max_wrap,
            ),
            add="+",
        )

    def _update_dynamic_wrap(
        self,
        widget: ttk.Label,
        fallback: int,
        max_width: Optional[int] = None,
    ) -> None:
        try:
            current_width = int(widget.winfo_width())
        except Exception:
            current_width = 0
        target_wrap = max(120, current_width - 8) if current_width > 0 else fallback
        if max_width is not None and max_width > 0:
            target_wrap = min(target_wrap, max_width)
        try:
            existing = int(float(widget.cget("wraplength")))
        except Exception:
            existing = 0
        if abs(existing - target_wrap) > 4:
            widget.configure(wraplength=target_wrap)

    def _bind_mousewheel_for_overview(self) -> None:
        self.root.bind_all("<MouseWheel>", self._on_overview_mousewheel, add="+")
        self.root.bind_all(
            "<Button-4>",
            lambda event: self._on_overview_mousewheel(event, forced_delta=-1),
            add="+",
        )
        self.root.bind_all(
            "<Button-5>",
            lambda event: self._on_overview_mousewheel(event, forced_delta=1),
            add="+",
        )

    def _is_overview_widget(self, widget: Any) -> bool:
        canvas = self._overview_canvas
        while widget is not None:
            if widget is canvas:
                return True
            widget = getattr(widget, "master", None)
        return False

    def _on_overview_mousewheel(
        self,
        event: Any,
        forced_delta: Optional[int] = None,
    ) -> None:
        canvas = self._overview_canvas
        if not canvas or not canvas.winfo_exists():
            return
        pointer_widget = self.root.winfo_containing(
            self.root.winfo_pointerx(),
            self.root.winfo_pointery(),
        )
        if not self._is_overview_widget(pointer_widget):
            return
        delta = forced_delta
        if delta is None:
            raw_delta = getattr(event, "delta", 0)
            if raw_delta == 0:
                return
            delta = -1 if raw_delta > 0 else 1
        canvas.yview_scroll(int(delta), "units")

    def _append_log(self, message: str) -> None:
        timestamp = time.strftime("%H:%M:%S")
        clean_message = repair_text(message)
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {clean_message}\n")
        try:
            line_count = int(float(self.log_text.index("end-1c").split(".")[0]))
        except Exception:
            line_count = 0
        if line_count > self.log_line_limit:
            overflow = line_count - self.log_line_limit
            self.log_text.delete("1.0", f"{overflow + 1}.0")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _schedule_log(self, message: str) -> None:
        self.root.after(0, self._append_log, repair_text(message))

    def _forward_core_log(self, message: str) -> None:
        text = str(message)
        core.append_agent_log(text)
        self._schedule_log(text)

    def _set_var(self, key: str, value: Any) -> None:
        display = "-" if value in (None, "") else repair_text(value)
        self.summary_vars[key].set(display)

    def _set_startup_check(self, key: str, text: str, level: str) -> None:
        title = self.startup_check_titles.get(key, key.upper())
        display = f"{title}: {repair_text(text) if text else '-'}"
        if key in self.startup_check_vars:
            self.startup_check_vars[key].set(display)
        if key in self.startup_check_labels:
            self.startup_check_labels[key].configure(
                fg=STATUS_COLORS.get(level, STATUS_COLORS["muted"])
            )

    def _inspect_exe_path(self, tool: core.GameTool) -> Dict[str, str]:
        raw_path = str(tool.paths.get("exe_path", "") or "").strip()
        if not raw_path:
            return {"level": "error", "text": "未配置 exe_path"}
        exe_path = tool.exe_path
        if exe_path.exists() and exe_path.is_file():
            return {"level": "ok", "text": f"存在: {exe_path.name}"}
        if exe_path.exists():
            return {"level": "error", "text": f"不是文件: {exe_path.name}"}
        return {"level": "error", "text": f"不存在: {exe_path.name or exe_path}"}

    def _inspect_bootstrap_file(self, tool: core.GameTool) -> Dict[str, str]:
        if not tool.bootstrap_file.exists():
            return {
                "level": "warn",
                "text": "缺失；定时启动前需先同步一次",
            }
        try:
            bootstrap = core.load_json_file(tool.bootstrap_file, default={}) or {}
        except Exception as exc:
            return {"level": "error", "text": f"读取失败: {repair_text(str(exc))}"}
        if not isinstance(bootstrap, dict) or not bootstrap:
            return {"level": "error", "text": "文件为空或格式错误"}
        bootstrap_agent_id = str(bootstrap.get("agent_id", "") or "").strip()
        if bootstrap_agent_id and bootstrap_agent_id != tool.agent_id:
            return {
                "level": "error",
                "text": f"agent 不一致: {bootstrap_agent_id}->{tool.agent_id}",
            }
        timestamp = str(
            bootstrap.get("updated_at", "") or bootstrap.get("server_time", "") or ""
        ).strip()
        if timestamp:
            return {"level": "ok", "text": f"已就绪: {timestamp}"}
        return {"level": "ok", "text": "已就绪"}

    def _render_startup_checks(
        self, snapshot: Dict[str, Any], agent_health: Dict[str, str]
    ) -> None:
        agent_id = str(snapshot.get("agent_id", "") or "").strip()
        base_url = str(snapshot.get("base_url", "") or "").strip()
        state_warning = str(snapshot.get("state_warning", "") or "").strip()
        state_file = str(snapshot.get("state_file", "") or "").strip()
        bootstrap_check = snapshot.get("bootstrap_check", {})
        exe_check = snapshot.get("exe_check", {})
        agent_status = str(agent_health.get("status", "") or "").strip()

        if not core.CONFIG_FILE.exists():
            self._set_startup_check("config", "配置文件不存在", "error")
        elif not agent_id:
            self._set_startup_check("config", "agent_id 未填写", "error")
        elif not base_url:
            self._set_startup_check("config", "base_url 未填写", "error")
        else:
            self._set_startup_check("config", f"agent_id={agent_id}", "ok")

        if state_warning:
            self._set_startup_check("state", state_warning, "error")
        elif state_file and Path(state_file).exists():
            self._set_startup_check("state", "已初始化", "ok")
        else:
            self._set_startup_check("state", "待初始化", "warn")

        self._set_startup_check(
            "bootstrap",
            str(bootstrap_check.get("text", "-") or "-"),
            str(bootstrap_check.get("level", "muted") or "muted"),
        )
        self._set_startup_check(
            "exe",
            str(exe_check.get("text", "-") or "-"),
            str(exe_check.get("level", "muted") or "muted"),
        )

        if agent_status == "在线":
            self._set_startup_check("agent", "在线", "ok")
        elif agent_status in {"未启动", "已停止", "可能卡住"}:
            self._set_startup_check("agent", agent_status, "warn")
        elif agent_status:
            self._set_startup_check("agent", agent_status, "error")
        else:
            self._set_startup_check("agent", "未知", "muted")

    def _format_remote_state(self, payload: Dict[str, Any]) -> str:
        if not isinstance(payload, dict) or not payload:
            return "-"
        label = str(payload.get("state_label", "") or "").strip()
        code = str(payload.get("state_code", "") or "").strip()
        if label and code:
            return f"{label} ({code})"
        return label or code or "-"

    def _build_task_boundary_note(
        self, task: Dict[str, Any], assist: Dict[str, Any]
    ) -> str:
        task = task if isinstance(task, dict) else {}
        assist = assist if isinstance(assist, dict) else {}

        current_start = core.parse_int(task.get("group_start"), 0)
        current_end = core.parse_int(task.get("group_end"), 0)
        profile_start = core.parse_int(
            task.get("profile_group_start", task.get("group_start")),
            current_start,
        )
        profile_end = core.parse_int(
            task.get("profile_group_end", task.get("group_end")),
            current_end,
        )

        assist_active = bool(assist.get("active", False))
        assist_role = str(assist.get("role", "") or "").strip().lower()
        delegate_start = core.parse_int(assist.get("delegate_start"), 0)
        delegate_end = core.parse_int(assist.get("delegate_end"), 0)
        original_target_group_end = core.parse_int(
            assist.get("original_target_group_end"),
            profile_end,
        )
        effective_target_group_end = core.parse_int(
            assist.get("effective_target_group_end"),
            current_end,
        )

        if assist_active and assist_role == "target":
            helper_ids: List[str] = []
            raw_helper_ids = assist.get("helper_agent_ids", [])
            if isinstance(raw_helper_ids, list):
                helper_ids = [
                    str(item).strip() for item in raw_helper_ids if str(item).strip()
                ]
            helper_agent_id = str(assist.get("helper_agent_id", "") or "").strip()
            if helper_agent_id and helper_agent_id not in helper_ids:
                helper_ids.append(helper_agent_id)
            helper_label = "、".join(helper_ids) if helper_ids else "helper"
            if profile_end > current_end > 0:
                note = f"尾段已交给 {helper_label}：原结束组 {profile_end} -> 当前有效结束组 {current_end}"
                if delegate_start > 0 and delegate_end >= delegate_start:
                    note += f"；接手区间 {delegate_start}->{delegate_end}"
                return note
            return (
                f"当前作为被协助目标机运行；有效结束组 {current_end or '-'}，"
                f"原配置结束组 {profile_end or '-'}"
            )

        if assist_active and assist_role == "helper":
            target_agent_id = str(assist.get("target_agent_id", "") or "").strip()
            note = f"当前处于协助模式：执行 {current_start}->{current_end}"
            if target_agent_id:
                note += f"，目标 {target_agent_id}"
            if original_target_group_end > 0:
                note += f"；目标原结束组 {original_target_group_end}"
                if effective_target_group_end > 0:
                    note += f"，当前有效结束组 {effective_target_group_end}"
            return note

        if profile_start != current_start or profile_end != current_end:
            return (
                f"当前任务范围已调整：{profile_start}->{profile_end} -> "
                f"{current_start}->{current_end}"
            )
        return "当前按原配置范围执行"

    def _build_evidence_insight(self, snapshot: Dict[str, Any]) -> Dict[str, Any]:
        remote_runtime = snapshot.get("remote_runtime", {})
        remote_supervision = snapshot.get("remote_supervision", {})
        remote_result = snapshot.get("remote_result", {})
        remote_heartbeat = snapshot.get("remote_heartbeat", {})
        progress = snapshot.get("status_progress", {})
        control_error = str(snapshot.get("control_error", "") or "").strip()
        qiannian_running = bool(snapshot.get("qiannian_running", False))
        task = snapshot.get("task", {})
        assist = snapshot.get("assist", {})

        supervision_label = (
            str(remote_supervision.get("state_label", "") or "-").strip() or "-"
        )
        supervision_code = str(remote_supervision.get("state_code", "") or "").strip()
        result_label = str(remote_result.get("state_label", "") or "-").strip() or "-"
        result_code = str(remote_result.get("state_code", "") or "").strip()

        local_group = core.parse_int(progress.get("group"), 0)
        remote_group = core.parse_int(
            remote_result.get("current_group", remote_runtime.get("current_group", 0)),
            0,
        )
        heartbeat_group = core.parse_int(remote_heartbeat.get("status_group"), 0)
        current_group_end = core.parse_int(task.get("group_end"), 0)
        profile_group_end = core.parse_int(
            task.get("profile_group_end", task.get("group_end")),
            current_group_end,
        )
        assist_active = bool(assist.get("active", False))
        assist_role = str(assist.get("role", "") or "").strip().lower()
        delegate_start = core.parse_int(assist.get("delegate_start"), 0)
        delegate_end = core.parse_int(assist.get("delegate_end"), 0)
        target_agent_id = str(assist.get("target_agent_id", "") or "").strip()

        comparison_parts: List[str] = []
        if supervision_label != "-" or result_label != "-":
            comparison_parts.append(f"监督={supervision_label} / 结果={result_label}")
        if heartbeat_group > 0 or remote_group > 0 or local_group > 0:
            comparison_parts.append(
                f"heartbeat={heartbeat_group or '-'} / report={remote_group or '-'} / 本地status={local_group or '-'}"
            )
        if (
            profile_group_end > 0
            and current_group_end > 0
            and profile_group_end != current_group_end
        ):
            comparison_parts.append(
                f"有效结束组={current_group_end} / 原结束组={profile_group_end}"
            )

        hints: List[str] = []
        if control_error:
            if qiannian_running:
                hints.append(
                    "local_report 当前不可达，但本地千年仍在运行；优先检查 report 服务，不要急着重启。"
                )
            else:
                hints.append(
                    "local_report 当前不可达，且本地未检测到千年进程；先恢复 report 服务，再看启动链路。"
                )
        if (
            assist_active
            and assist_role == "target"
            and profile_group_end > current_group_end > 0
        ):
            hints.append(
                f"当前尾段已外包：本机只应跑到 group={current_group_end}；原结束组 {profile_group_end} 之后的任务已交给 helper。"
            )
            max_seen_group = max(local_group, remote_group, heartbeat_group)
            if max_seen_group > current_group_end:
                hints.append(
                    f"检测到进度已越过当前有效结束组 {current_group_end}；新版 agent 会在下一次轮询中自动停止 qiannian。"
                )
        elif assist_active and assist_role == "helper":
            hint = f"当前处于协助模式，执行区间 {task.get('group_start', '-')}->{task.get('group_end', '-')}"
            if target_agent_id:
                hint += f"，目标 {target_agent_id}"
            if delegate_start > 0 and delegate_end >= delegate_start:
                hint += f"；协助段 {delegate_start}->{delegate_end}"
            hint += "；不要按本机原配置范围判断完成。"
            hints.append(hint)
        if (
            supervision_code in {"suspected_stuck", "startup_failed"}
            and result_code == "fresh"
        ):
            hints.append(
                "结果链路仍在刷新，但监督链路已落后；优先检查 game_tool agent 是否常驻、heartbeat 是否持续上报。"
            )
        if remote_group > 0 and local_group > 0:
            if remote_group > local_group:
                hints.append(
                    f"report 结果已到 group={remote_group}，本地 status.ini 仍是 {local_group}；更像是本地证据刷新落后。"
                )
            elif local_group > remote_group:
                hints.append(
                    f"本地 status.ini 已到 group={local_group}，report 结果仍是 {remote_group}；通常表示当前组尚未完成，属于正常中间态。"
                )
        seen_groups = [
            value for value in [local_group, remote_group, heartbeat_group] if value > 0
        ]
        if heartbeat_group > 0 and remote_group > 0 and heartbeat_group != remote_group:
            if remote_group > heartbeat_group and result_code == "fresh":
                hints.append(
                    f"report 侧 heartbeat 仍停在 group={heartbeat_group}，但结果上报已到 group={remote_group}；不要仅凭监督状态判死。"
                )
            elif heartbeat_group > remote_group:
                hints.append(
                    f"heartbeat 证据已到 group={heartbeat_group}，结果上报仍是 group={remote_group}；说明当前组可能仍在执行中。"
                )
        if len(set(seen_groups)) >= 2:
            comparison_parts.append("restart续跑=本地status.ini")
            hints.append(
                f"【重启判定】发生链路分叉时，restart 续跑以本地 status.ini 为准；当前 本地={local_group or '-'} / 结果={remote_group or '-'} / heartbeat={heartbeat_group or '-'}，不要按远端 heartbeat 直接判断续跑组号。"
            )
        if not hints and result_code == "completed":
            hints.append("本轮任务已完成，等待下一次计划或新的协助任务。")
        if (
            not hints
            and not qiannian_running
            and local_group <= 0
            and remote_group <= 0
        ):
            hints.append("当前未看到运行证据，等待计划时间或手动启动。")
        return {
            "comparison": " ； ".join(comparison_parts) if comparison_parts else "-",
            "hint": " ".join(hints) if hints else "-",
        }

    def _build_agent_health(self, snapshot: Dict[str, Any]) -> Dict[str, str]:
        runtime = snapshot.get("runtime", {})
        control = snapshot.get("control", {})
        agent_processes = snapshot.get("agent_processes", [])
        qiannian_running = bool(snapshot.get("qiannian_running", False))
        threshold = max(
            60,
            core.parse_int(snapshot.get("agent_alive_threshold_seconds"), 180),
        )
        last_loop_epoch = core.parse_int(runtime.get("last_agent_loop_epoch"), 0)
        last_loop_time = str(runtime.get("last_agent_loop_time", "") or "").strip()
        agent_phase = str(runtime.get("agent_phase", "") or "").strip()
        schedule_text = str(
            control.get("schedule_daily_start")
            or runtime.get("schedule_daily_start")
            or ""
        ).strip()
        next_schedule_date = str(runtime.get("next_schedule_date", "") or "").strip()
        desired_run_state = (
            str(control.get("desired_run_state", "run")).strip().lower() or "run"
        )
        session_active = bool(runtime.get("session_active", False))
        has_agent_history = bool(last_loop_time or next_schedule_date or schedule_text)
        if agent_processes:
            stale_seconds = (
                int(time.time()) - last_loop_epoch if last_loop_epoch > 0 else 0
            )
            if last_loop_epoch > 0 and stale_seconds > threshold:
                detail = (
                    f"检测到 agent 进程仍在，但后台循环已超过 {stale_seconds} 秒未刷新"
                )
                if last_loop_time:
                    detail += f"；最近轮询 {last_loop_time}"
                if agent_phase:
                    detail += f"；阶段 {agent_phase}"
                return {"status": "可能卡住", "detail": detail}
            detail = "agent 正在后台运行"
            if last_loop_time:
                detail += f"；最近轮询 {last_loop_time}"
            if agent_phase:
                detail += f"；阶段 {agent_phase}"
            return {"status": "在线", "detail": detail}
        if desired_run_state == "stop" and not session_active and not qiannian_running:
            return {"status": "已停止", "detail": "远端当前为停止状态，未运行 agent"}
        if has_agent_history or session_active or qiannian_running:
            detail_parts = ["未检测到 agent 进程"]
            if last_loop_time:
                detail_parts.append(f"最近轮询 {last_loop_time}")
            if agent_phase:
                detail_parts.append(f"最近阶段 {agent_phase}")
            if next_schedule_date:
                detail_parts.append(f"下次计划 {next_schedule_date}")
            elif schedule_text:
                detail_parts.append(f"计划时间 {schedule_text}")
            return {"status": "掉线", "detail": "；".join(detail_parts)}
        return {"status": "未启动", "detail": "当前未检测到 agent 常驻进程"}

    def _render_snapshot(self, snapshot: Dict[str, Any]) -> None:
        task = snapshot.get("task", {})
        control = snapshot.get("control", {})
        runtime = snapshot.get("runtime", {})
        remote = snapshot.get("remote_runtime", {})
        remote_supervision = snapshot.get("remote_supervision", {})
        remote_result = snapshot.get("remote_result", {})
        remote_heartbeat = snapshot.get("remote_heartbeat", {})
        local = snapshot.get("local_completion", {})
        progress = snapshot.get("status_progress", {})
        agent_processes = snapshot.get("agent_processes", [])
        control_error = snapshot.get("control_error", "")
        state_warning = str(snapshot.get("state_warning", "") or "").strip()
        assist = snapshot.get("assist", {})
        profile_group_start = task.get(
            "profile_group_start", task.get("group_start", "-")
        )
        profile_group_end = task.get("profile_group_end", task.get("group_end", "-"))
        insight = self._build_evidence_insight(snapshot)
        agent_health = self._build_agent_health(snapshot)
        self._render_startup_checks(snapshot, agent_health)

        assist_summary = assist.get("summary", "") if isinstance(assist, dict) else ""
        self.assist_status_var.set(
            repair_text(str(assist_summary or "当前未接到临时协助任务"))
        )
        self._set_var("agent_id", snapshot.get("agent_id"))
        self._refresh_header_agent_id(snapshot.get("agent_id"))
        self._set_var("base_url", snapshot.get("base_url"))
        self._set_var("region", task.get("region", "-"))
        self._set_var(
            "group_range",
            f"{task.get('group_start', '-')} -> {task.get('group_end', '-')}",
        )
        self._set_var(
            "profile_group_range",
            f"{profile_group_start} -> {profile_group_end}",
        )
        self._set_var(
            "task_boundary_note",
            self._build_task_boundary_note(task, assist),
        )
        self._set_var(
            "assist_summary",
            assist_summary or "当前未接到临时协助任务",
        )
        self._set_var("schedule_daily_start", control.get("schedule_daily_start", "-"))
        self._set_var("desired_run_state", control.get("desired_run_state", "-"))
        self._set_var("auto_restart", bool(control.get("auto_restart_on_stale", False)))
        self._set_var("exe_path", snapshot.get("exe_path"))

        self._set_var("local_status", runtime.get("status", "-"))
        self._set_var("local_status_detail", snapshot.get("local_status_detail", "-"))
        self._set_var("state_warning", state_warning or "-")
        self._set_var("agent_health", agent_health.get("status", "-"))
        self._set_var("agent_health_detail", agent_health.get("detail", "-"))
        self._set_var("session_active", bool(runtime.get("session_active", False)))
        self._set_var("next_schedule_date", runtime.get("next_schedule_date", "-"))
        self._set_var("last_launch_time", runtime.get("last_launch_time", "-"))
        self._set_var("last_launch_reason", runtime.get("last_launch_reason", "-"))
        self._set_var("last_stop_time", runtime.get("last_stop_time", "-"))
        self._set_var("last_stop_reason", runtime.get("last_stop_reason", "-"))
        self._set_var(
            "agent_process",
            (
                f"running ({', '.join(str(item.get('pid')) for item in agent_processes)})"
                if agent_processes
                else "not running"
            ),
        )
        self._set_var("qiannian_running", snapshot.get("qiannian_running", False))
        self._set_var(
            "last_heartbeat_time",
            runtime.get("last_heartbeat_at", "-") or "-",
        )
        self._set_var(
            "local_progress_evidence",
            f"group={progress.get('group', 0)} role={progress.get('role_index', 0)} / last_change={runtime.get('last_progress_change_at', '-') or '-'}",
        )
        self._set_var(
            "status_ini",
            f"group={progress.get('group', 0)} role={progress.get('role_index', 0)} is_today={progress.get('is_today', False)} mtime={progress.get('mtime_epoch', 0)}",
        )

        remote_result_elapsed = (
            remote_result.get("elapsed", "-")
            if remote_result.get("elapsed") is not None
            else "-"
        )
        remote_heartbeat_elapsed = (
            remote_heartbeat.get("heartbeat_elapsed", "-")
            if remote_heartbeat.get("heartbeat_elapsed") is not None
            else "-"
        )
        self._set_var(
            "remote_supervision_state",
            self._format_remote_state(remote_supervision),
        )
        self._set_var(
            "remote_supervision_detail",
            remote_supervision.get("detail", "-") or "-",
        )
        self._set_var(
            "remote_result_state",
            self._format_remote_state(remote_result),
        )
        self._set_var(
            "remote_result_detail",
            remote_result.get("detail", "-") or "-",
        )
        self._set_var(
            "remote_heartbeat",
            f"has={bool(remote_heartbeat.get('has_heartbeat', False))} / at={remote_heartbeat.get('server_time', '-') or '-'} / elapsed={remote_heartbeat_elapsed}s",
        )
        self._set_var(
            "remote_heartbeat_progress",
            f"group={remote_heartbeat.get('status_group', '-')} role={remote_heartbeat.get('status_role_index', '-')} / last_progress={remote_heartbeat.get('last_progress_change_at', '-') or '-'}",
        )
        self._set_var(
            "remote_result_progress",
            f"group={remote.get('current_group', '-')} role={remote.get('role_index', '-')} / elapsed={remote_result_elapsed}s / time={remote.get('server_time', '-') or '-'}",
        )
        self._set_var("remote_event", remote.get("event", "-"))
        self._set_var("control_error", control_error or "-")

        self._set_var("local_completed", bool(local.get("completed", False)))
        self._set_var("local_target_group_end", local.get("target_group_end", "-"))
        self._set_var(
            "local_complete_role_index", local.get("complete_role_index", "-")
        )
        self._set_var("should_resume_today", snapshot.get("should_resume_today", False))
        self._set_var(
            "pending_restart_reason", snapshot.get("pending_restart_reason", "-")
        )
        self._set_var(
            "restart_counter",
            f"{runtime.get('restart_count_today', 0)} / {runtime.get('restart_count_date', '-') or '-'}",
        )
        self._set_var(
            "last_progress",
            f"group={runtime.get('last_seen_group', '-')} role={runtime.get('last_seen_role_index', '-')} / report_time={remote.get('server_time', '-') or '-'}",
        )
        self._set_var("status_date", progress.get("last_reset_date", "-"))
        self._set_var("evidence_comparison", insight.get("comparison", "-"))
        self._set_var("evidence_hint", insight.get("hint", "-"))

        self.raw_text.configure(state=tk.NORMAL)
        self.raw_text.delete("1.0", tk.END)
        self.raw_text.insert(
            tk.END,
            json.dumps(
                repair_payload(snapshot), ensure_ascii=False, indent=2, default=str
            ),
        )
        self.raw_text.configure(state=tk.DISABLED)

        if state_warning and state_warning != self.last_state_warning:
            gui_message = f"[STATE] {state_warning}"
            self._append_log(gui_message)
            messagebox.showwarning("STATE 异常", repair_text(state_warning))
        self.last_state_warning = state_warning

    def _build_snapshot(self) -> Dict[str, Any]:
        config = core.load_config()
        tool = core.GameTool(config)
        tool.ensure_dirs()

        loaded_state = tool.load_state()
        state_warning = tool.get_state_agent_mismatch_message(loaded_state)
        state = tool.normalize_state_for_agent(loaded_state)
        state_changed = bool(state_warning)
        if tool.normalize_runtime_for_today(state):
            state_changed = True
        runtime = tool.get_runtime_state(state)
        if tool.ensure_restart_counter(runtime):
            state_changed = True
        if state_changed:
            tool.save_state(state)

        control_doc: Dict[str, Any] = {}
        control: Dict[str, Any] = {}
        task: Dict[str, Any] = {}
        remote_runtime: Dict[str, Any] = {}
        remote_supervision: Dict[str, Any] = {}
        remote_result: Dict[str, Any] = {}
        remote_heartbeat: Dict[str, Any] = {}
        assist: Dict[str, Any] = {}
        control_error = ""

        try:
            control_doc = tool.fetch_control()
            tool.update_runtime_from_control(state, control_doc)
            state = tool.normalize_state_for_agent(tool.load_state())
            runtime = tool.get_runtime_state(state)
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
            remote_runtime = (
                control_doc.get("runtime", {})
                if isinstance(control_doc.get("runtime", {}), dict)
                else {}
            )
            remote_supervision = (
                control_doc.get("supervision", {})
                if isinstance(control_doc.get("supervision", {}), dict)
                else {}
            )
            remote_result = (
                control_doc.get("result", {})
                if isinstance(control_doc.get("result", {}), dict)
                else {}
            )
            remote_heartbeat = (
                control_doc.get("heartbeat", {})
                if isinstance(control_doc.get("heartbeat", {}), dict)
                else {}
            )
            assist = (
                control_doc.get("assist", {})
                if isinstance(control_doc.get("assist", {}), dict)
                else {}
            )
            if not assist and isinstance(task.get("assist", {}), dict):
                assist = dict(task.get("assist", {}))
        except Exception as exc:
            control_error = repair_text(str(exc))

        progress = tool.read_status_ini_progress()
        local_completion = tool.get_local_completion_state(task, control)
        qiannian_running = tool.is_process_running()
        local_status_detail = tool.describe_local_status(
            runtime,
            control,
            task,
            remote_runtime=remote_runtime,
            local_completion=local_completion,
            progress=progress,
            qiannian_running=qiannian_running,
        )
        should_resume_today = False
        pending_restart_reason = ""
        if task or control:
            should_resume_today = tool.should_resume_pending_session(
                runtime, control, task
            )
            pending_restart_reason = tool.get_pending_restart_reason(
                runtime, remote_runtime, control, task
            )

        return {
            "agent_id": tool.agent_id,
            "base_url": tool.base_url,
            "exe_path": str(tool.exe_path),
            "state_file": str(tool.state_file),
            "bootstrap_file": str(tool.bootstrap_file),
            "state_warning": state_warning,
            "bootstrap_check": self._inspect_bootstrap_file(tool),
            "exe_check": self._inspect_exe_path(tool),
            "task": task,
            "assist": assist,
            "control": control,
            "runtime": dict(runtime),
            "remote_runtime": remote_runtime,
            "remote_supervision": remote_supervision,
            "remote_result": remote_result,
            "remote_heartbeat": remote_heartbeat,
            "status_progress": progress,
            "local_completion": local_completion,
            "local_status_detail": local_status_detail,
            "should_resume_today": should_resume_today,
            "pending_restart_reason": pending_restart_reason,
            "control_error": control_error,
            "agent_processes": self._find_agent_processes(),
            "agent_alive_threshold_seconds": max(
                180,
                tool.control_poll_seconds * 4,
                tool.process_stop_timeout_seconds
                + tool.window_find_timeout_seconds
                + tool.launch_ready_seconds
                + tool.post_load_delay_seconds
                + tool.load_confirm_timeout_seconds
                + tool.launch_settle_seconds
                + tool.window_cleanup_timeout_seconds,
            ),
            "qiannian_running": qiannian_running,
        }

    def refresh_snapshot(self, log_success: bool = False) -> None:
        if self.refresh_running:
            return
        self.refresh_running = True

        def worker() -> None:
            try:
                snapshot = self._build_snapshot()
                self.root.after(0, self._render_snapshot, snapshot)
                if log_success:
                    self._schedule_log("状态已刷新")
            except Exception as exc:
                self._schedule_log(f"刷新失败: {exc}")
            finally:
                self.refresh_running = False
                if self.auto_refresh_var.get():
                    self._schedule_next_refresh()

        threading.Thread(target=worker, daemon=True).start()

    def _schedule_next_refresh(self) -> None:
        if self.refresh_after_id is not None:
            self.root.after_cancel(self.refresh_after_id)
            self.refresh_after_id = None
        interval = max(3, int(self.refresh_interval_var.get() or 5))
        self.refresh_after_id = self.root.after(interval * 1000, self.refresh_snapshot)

    def _on_toggle_auto_refresh(self) -> None:
        if self.auto_refresh_var.get():
            self.refresh_snapshot()
        elif self.refresh_after_id is not None:
            self.root.after_cancel(self.refresh_after_id)
            self.refresh_after_id = None

    def _run_cli_once(self, subcommand: str) -> str:
        cmd = self._build_cli_command(subcommand)
        completed = subprocess.run(
            cmd,
            cwd=str(core.SCRIPT_DIR),
            capture_output=True,
            text=False,
            creationflags=CREATE_NO_WINDOW,
            check=False,
        )
        output = decode_cli_output(
            (completed.stdout or b"") + (completed.stderr or b"")
        )
        output = output.strip()
        if output:
            for line in output.splitlines():
                self._schedule_log(line)
        if completed.returncode != 0:
            raise RuntimeError(f"{subcommand} failed: exit_code={completed.returncode}")
        return output

    def _run_background_action(self, title: str, func) -> None:
        def worker() -> None:
            try:
                self._schedule_log(f"{title}: 开始")
                func()
                self._schedule_log(f"{title}: 完成")
            except Exception as exc:
                self._schedule_log(f"{title}: 失败 -> {exc}")
                self.root.after(
                    0, lambda: messagebox.showerror(title, repair_text(str(exc)))
                )
            finally:
                self.root.after(0, self.refresh_snapshot)

        threading.Thread(target=worker, daemon=True).start()

    def _build_cli_command(self, subcommand: str) -> List[str]:
        cli_exe = core.SCRIPT_DIR / "game_tool.exe"
        cli_py = core.SCRIPT_DIR / "game_tool.py"
        if getattr(sys, "frozen", False) and cli_exe.exists():
            return [str(cli_exe), subcommand]
        if (
            cli_exe.exists()
            and core.SCRIPT_DIR != Path(sys.executable).resolve().parent
        ):
            return [str(cli_exe), subcommand]
        return [sys.executable, str(cli_py), subcommand]

    def _find_agent_processes(self) -> List[Dict[str, Any]]:
        if os.name != "nt":
            return []
        command = [
            "powershell.exe",
            "-Command",
            "$ErrorActionPreference='SilentlyContinue'; "
            "$items = Get-CimInstance Win32_Process | Where-Object { "
            "($_.Name -ieq 'game_tool.exe') -or "
            "((($_.Name -ieq 'python.exe') -or ($_.Name -ieq 'pythonw.exe')) -and ([string]$_.CommandLine -match 'game_tool\\.py\"?\\s+agent')) "
            "} | Select-Object @{Name='pid';Expression={$_.ProcessId}}, @{Name='name';Expression={$_.Name}}, @{Name='command_line';Expression={$_.CommandLine}}; "
            "if ($items) { $items | ConvertTo-Json -Compress }",
        ]
        try:
            completed = subprocess.run(
                command,
                cwd=str(core.SCRIPT_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=CREATE_NO_WINDOW,
                check=False,
            )
            raw = str(completed.stdout or "").strip()
            if raw:
                parsed = json.loads(raw)
                if isinstance(parsed, dict):
                    parsed = [parsed]
                if isinstance(parsed, list):
                    return [
                        {
                            "pid": core.parse_int(item.get("pid"), 0),
                            "name": str(item.get("name", "") or "").strip(),
                            "command_line": str(
                                item.get("command_line", "") or ""
                            ).strip(),
                        }
                        for item in parsed
                        if core.parse_int(item.get("pid"), 0) > 0
                    ]
        except Exception:
            pass
        tool = core.GameTool(core.load_config())
        results: List[Dict[str, Any]] = []
        for pid in tool.list_process_pids("game_tool.exe"):
            if pid > 0:
                results.append({"pid": pid, "name": "game_tool.exe"})
        return results

    def _kill_agent_processes(self) -> int:
        processes = self._find_agent_processes()
        killed = 0
        for item in processes:
            pid = core.parse_int(item.get("pid"), 0)
            if pid <= 0:
                continue
            completed = core.run_hidden_subprocess(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=False,
                check=False,
            )
            if completed.returncode == 0:
                killed += 1
                self._schedule_log(f"已终止 agent 进程 pid={pid}")
        return killed

    def _spawn_agent_process(self) -> None:
        cmd = self._build_cli_command("agent")
        subprocess.Popen(
            cmd,
            cwd=str(core.SCRIPT_DIR),
            creationflags=CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

    def _wait_for_agent_processes(
        self, timeout_seconds: int = 15
    ) -> List[Dict[str, Any]]:
        deadline = time.time() + max(1, timeout_seconds)
        while time.time() < deadline:
            processes = self._find_agent_processes()
            if processes:
                return processes
            time.sleep(0.5)
        return []

    def start_agent(self) -> None:
        if self.agent_launching:
            return
        self.agent_launching = True

        def action() -> None:
            processes = self._find_agent_processes()
            if processes:
                raise RuntimeError("检测到 agent 已在运行，请先停止旧 agent 再重新启动")
            self._spawn_agent_process()
            self._schedule_log("已在后台启动 agent")

        def wrapped() -> None:
            try:
                action()
            finally:
                self.agent_launching = False

        self._run_background_action("启动 Agent", wrapped)

    def skip_today_and_start_agent(self) -> None:
        def action() -> None:
            self._run_cli_once("skip-today")
            time.sleep(0.5)
            processes = self._find_agent_processes()
            if processes:
                raise RuntimeError(
                    "skip-today 执行后检测到 agent 已在运行，请先停止旧 agent"
                )
            self._spawn_agent_process()
            self._schedule_log("已执行 skip-today 并在后台启动 agent")

        self._run_background_action("跳过今天并启动", action)

    def stop_agent(self) -> None:
        def action() -> None:
            killed = self._kill_agent_processes()
            self._run_cli_once("stop")
            if killed == 0:
                self._schedule_log("未检测到运行中的 agent 进程，已执行 stop")

        self._run_background_action("停止 Agent", action)

    def stop_qiannian(self) -> None:
        self._run_background_action("停止千年", lambda: self._run_cli_once("stop"))

    def open_config_window(self) -> None:
        if self.config_window is not None and self.config_window.winfo_exists():
            self.config_window.focus_force()
            self.config_window.lift()
            return
        self.config_window = ConfigEditorWindow(
            self.root,
            self._schedule_log,
            on_local_config_changed=self._on_local_config_changed,
        )

    def _create_tool(self) -> core.GameTool:
        tool = core.GameTool(core.load_config())
        tool.ensure_dirs()
        return tool

    def _open_path(self, target: Path, title: str) -> None:
        def action() -> None:
            path = Path(target)
            open_target = path
            if not open_target.exists():
                if path.parent.exists():
                    open_target = path.parent
                    self._schedule_log(f"{title} 不存在，已打开上级目录: {open_target}")
                else:
                    raise RuntimeError(f"{title} 不存在: {path}")
            if os.name == "nt":
                os.startfile(str(open_target))
            else:
                subprocess.Popen(["xdg-open", str(open_target)])
            self._schedule_log(f"已打开 {title}: {open_target}")

        self._run_background_action(f"打开 {title}", action)

    def open_game_tool_config(self) -> None:
        self._open_path(core.CONFIG_FILE, "game_tool_config.json")

    def open_global_config_ini(self) -> None:
        tool = self._create_tool()
        self._open_path(tool.get_global_config_ini_path(), "GlobalConfig.ini")

    def open_state_file(self) -> None:
        tool = self._create_tool()
        self._open_path(tool.state_file, "state.json")

    def open_bootstrap_file(self) -> None:
        tool = self._create_tool()
        self._open_path(tool.bootstrap_file, "bootstrap.json")

    def open_agent_log(self) -> None:
        self._open_path(core.AGENT_LOG_FILE, "agent.log")

    def open_qiannian_dir(self) -> None:
        tool = self._create_tool()
        self._open_path(tool.get_launch_base_dir(), "QianNian 目录")

    def open_account_dir(self) -> None:
        tool = self._create_tool()
        self._open_path(tool.get_launch_base_dir() / "account", "account 目录")

    def sync_once(self) -> None:
        self._run_background_action("立即同步配置", lambda: self._run_cli_once("sync"))

    def run_once(self) -> None:
        def action() -> None:
            processes = self._find_agent_processes()
            if processes:
                self._schedule_log(
                    "检测到 agent 正在运行，改为委托 agent 执行“同步并运行一次”"
                )
            else:
                self._schedule_log("未检测到 agent，先在后台拉起 agent，再委托执行")
                self._spawn_agent_process()
                processes = self._wait_for_agent_processes(timeout_seconds=15)
                if not processes:
                    raise RuntimeError(
                        "已尝试拉起 agent，但在 15 秒内未检测到 agent 进程"
                    )
                self._schedule_log("agent 已启动，继续委托本次同步并运行一次")
            tool = self._create_tool()
            request = tool.enqueue_local_agent_action(
                "sync_run_once", source="gui.run_once"
            )
            self._schedule_log(
                f"已提交本地请求给 agent: action={request.get('action')} seq={request.get('seq')}"
            )

        self._run_background_action("同步并运行一次", action)

    def on_close(self) -> None:
        if messagebox.askyesno(
            "关闭控制面板",
            "关闭后不会自动停止已经在后台运行的 agent。\n是否继续关闭？",
        ):
            self.root.destroy()


def main() -> int:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    gui = GameToolGui(root)
    gui._schedule_log(f"Config: {core.CONFIG_FILE}")
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
