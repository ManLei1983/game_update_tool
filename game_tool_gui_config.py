import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable, Dict
from urllib.parse import urlsplit

import game_tool as core


STATUS_COLORS = {
    "ok": "#167c2b",
    "warn": "#b26a00",
    "error": "#c62828",
    "muted": "#666666",
}


def parse_base_url(base_url: str) -> Dict[str, Any]:
    text = str(base_url or "").strip()
    if not text:
        return {"host": "", "port": 0}
    parsed = urlsplit(text)
    host = (parsed.hostname or "").strip()
    port = parsed.port
    if port is None:
        if parsed.scheme.lower() == "https":
            port = 443
        elif parsed.scheme.lower() == "http":
            port = 80
        else:
            port = 0
    return {"host": host, "port": port}


class ConfigEditorWindow(tk.Toplevel):
    def __init__(self, parent: tk.Misc, log_func: Callable[[str], None]) -> None:
        super().__init__(parent)
        self.log = log_func
        self.title("game_tool 配置面板")
        self.geometry("980x760")
        self.minsize(900, 700)

        self.base_url_var = tk.StringVar()
        self.agent_id_var = tk.StringVar()
        self.exe_path_var = tk.StringVar()
        self.timeout_var = tk.StringVar()
        self.poll_var = tk.StringVar()
        self.retry_var = tk.StringVar()
        self.ready_var = tk.StringVar()
        self.load_delay_var = tk.StringVar()

        self.ini_enable_var = tk.BooleanVar(value=True)
        self.ini_host_var = tk.StringVar()
        self.ini_port_var = tk.StringVar()
        self.ini_path_var = tk.StringVar()
        self.ini_agent_id_var = tk.StringVar()
        self.ini_timeout_ms_var = tk.StringVar()
        self.ini_file_var = tk.StringVar(value="-")
        self.ini_encoding_var = tk.StringVar(value="-")
        self.config_file_var = tk.StringVar(value=str(core.CONFIG_FILE))

        self.compare_vars: Dict[str, tk.StringVar] = {}
        self.compare_labels: Dict[str, tk.Label] = {}

        self._build_ui()
        self.reload_local_config(show_message=False)
        self.reload_ini(show_message=False)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        info = ttk.LabelFrame(self, text="文件路径", padding=8)
        info.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        info.columnconfigure(1, weight=1)
        ttk.Label(info, text="game_tool JSON").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(info, textvariable=self.config_file_var, state="readonly").grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Label(info, text="GlobalConfig.ini").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(info, textvariable=self.ini_file_var, state="readonly").grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Label(info, text="INI 编码").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(info, textvariable=self.ini_encoding_var, state="readonly").grid(row=2, column=1, sticky="ew", pady=3)

        forms = ttk.Frame(self)
        forms.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
        forms.columnconfigure(0, weight=1)
        forms.columnconfigure(1, weight=1)

        left = ttk.LabelFrame(forms, text="game_tool 本地配置", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.columnconfigure(1, weight=1)
        right = ttk.LabelFrame(forms, text="QianNian LocalReport", padding=10)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right.columnconfigure(1, weight=1)

        self._add_entry(left, 0, "Server Base URL", self.base_url_var)
        self._add_entry(left, 1, "Agent ID", self.agent_id_var)
        self._add_entry(left, 2, "QianNian EXE", self.exe_path_var)
        self._add_entry(left, 3, "HTTP 超时秒", self.timeout_var, 12)
        self._add_entry(left, 4, "control_poll_seconds", self.poll_var, 12)
        self._add_entry(left, 5, "control_error_retry_seconds", self.retry_var, 12)
        self._add_entry(left, 6, "launch_ready_seconds", self.ready_var, 12)
        self._add_entry(left, 7, "post_load_delay_seconds", self.load_delay_var, 12)

        ttk.Button(left, text="重新读取 JSON", command=self.reload_local_config).grid(row=8, column=0, padx=(0, 8), pady=(10, 0), sticky="w")
        ttk.Button(left, text="保存 JSON", command=self.save_local_config).grid(row=8, column=1, pady=(10, 0), sticky="w")

        ttk.Label(right, text="Enable").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Checkbutton(right, variable=self.ini_enable_var).grid(row=0, column=1, sticky="w", pady=4)
        self._add_entry(right, 1, "Host", self.ini_host_var, 22)
        self._add_entry(right, 2, "Port", self.ini_port_var, 12)
        self._add_entry(right, 3, "Path", self.ini_path_var, 22)
        self._add_entry(right, 4, "AgentId", self.ini_agent_id_var, 22)
        self._add_entry(right, 5, "TimeoutMs", self.ini_timeout_ms_var, 12)

        actions = ttk.Frame(right)
        actions.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(actions, text="读取 INI", command=self.reload_ini).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="按 game_tool 填充", command=self.fill_ini_from_tool).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="保存 INI", command=self.save_ini).grid(row=0, column=2)

        compare = ttk.LabelFrame(self, text="一致性对比", padding=10)
        compare.grid(row=2, column=0, sticky="nsew", padx=10, pady=(6, 10))
        compare.columnconfigure(1, weight=1)
        for row, key, label in [
            (0, "summary", "综合结论"),
            (1, "host", "Host 对比"),
            (2, "port", "Port 对比"),
            (3, "agent", "AgentId 对比"),
            (4, "path", "Path 检查"),
            (5, "enable", "Enable 检查"),
            (6, "url", "实际上报地址"),
        ]:
            ttk.Label(compare, text=label).grid(row=row, column=0, sticky="nw", padx=(0, 10), pady=4)
            var = tk.StringVar(value="-")
            self.compare_vars[key] = var
            value = tk.Label(compare, textvariable=var, anchor="w", justify=tk.LEFT, wraplength=640, fg=STATUS_COLORS["muted"])
            value.grid(row=row, column=1, sticky="ew", pady=4)
            self.compare_labels[key] = value

    def _add_entry(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, width: int = 40) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=1, sticky="ew", pady=4)

    def reload_local_config(self, show_message: bool = True) -> None:
        config = core.load_config()
        self.base_url_var.set(str(config["server"].get("base_url", "")))
        self.agent_id_var.set(str(config["server"].get("agent_id", "")))
        self.exe_path_var.set(str(config["paths"].get("exe_path", "")))
        self.timeout_var.set(str(config["server"].get("timeout_seconds", 15)))
        self.poll_var.set(str(config["behavior"].get("control_poll_seconds", 15)))
        self.retry_var.set(str(config["behavior"].get("control_error_retry_seconds", 30)))
        self.ready_var.set(str(config["behavior"].get("launch_ready_seconds", 20)))
        self.load_delay_var.set(str(config["behavior"].get("post_load_delay_seconds", 2)))
        self._refresh_compare()
        if show_message:
            self.log("已重新读取 game_tool_config.json")

    def _build_tool_from_form(self) -> core.GameTool:
        config = core.merge_dict(core.DEFAULT_CONFIG, core.load_config())
        config["server"]["base_url"] = self.base_url_var.get().strip()
        config["server"]["agent_id"] = self.agent_id_var.get().strip()
        config["paths"]["exe_path"] = self.exe_path_var.get().strip()
        config["server"]["timeout_seconds"] = max(5, core.parse_int(self.timeout_var.get(), 15))
        config["behavior"]["control_poll_seconds"] = max(5, core.parse_int(self.poll_var.get(), 15))
        config["behavior"]["control_error_retry_seconds"] = max(5, core.parse_int(self.retry_var.get(), 30))
        config["behavior"]["launch_ready_seconds"] = max(0, core.parse_int(self.ready_var.get(), 20))
        config["behavior"]["post_load_delay_seconds"] = max(0, core.parse_int(self.load_delay_var.get(), 2))
        tool = core.GameTool(config)
        tool.ensure_dirs()
        return tool

    def save_local_config(self) -> None:
        try:
            tool = self._build_tool_from_form()
            core.save_json_file(core.CONFIG_FILE, tool.config)
            self.log(f"已保存 game_tool 配置 -> {core.CONFIG_FILE}")
            self.reload_ini(show_message=False)
            messagebox.showinfo("保存成功", f"已写入:\n{core.CONFIG_FILE}")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def reload_ini(self, show_message: bool = True) -> None:
        tool = self._build_tool_from_form()
        info = tool.read_global_local_report_config()
        self.ini_file_var.set(str(info.get("path", "-")))
        self.ini_encoding_var.set(str(info.get("encoding", "-") or "-"))
        self.ini_enable_var.set(bool(info.get("enable", True)))
        self.ini_host_var.set(str(info.get("host", "")))
        self.ini_port_var.set(str(info.get("port", 18080)))
        self.ini_path_var.set(str(info.get("path_value", "/api/report")))
        self.ini_agent_id_var.set(str(info.get("agent_id", "")))
        self.ini_timeout_ms_var.set(str(info.get("timeout_ms", 2500)))
        self._refresh_compare()
        if show_message:
            self.log("已读取 GlobalConfig.ini 的 [LocalReport]")

    def fill_ini_from_tool(self) -> None:
        parsed = parse_base_url(self.base_url_var.get())
        self.ini_enable_var.set(True)
        self.ini_host_var.set(str(parsed.get("host", "")))
        self.ini_port_var.set(str(parsed.get("port", 18080) or 18080))
        self.ini_path_var.set(self.ini_path_var.get().strip() or "/api/report")
        self.ini_agent_id_var.set(self.agent_id_var.get().strip())
        self.ini_timeout_ms_var.set(str(max(100, core.parse_int(self.ini_timeout_ms_var.get(), 2500))))
        self._refresh_compare()
        self.log("已按 game_tool 配置填充 INI 表单")

    def save_ini(self) -> None:
        try:
            tool = self._build_tool_from_form()
            payload = {
                "enable": self.ini_enable_var.get(),
                "host": self.ini_host_var.get().strip(),
                "port": max(1, core.parse_int(self.ini_port_var.get(), 18080)),
                "path_value": self.ini_path_var.get().strip() or "/api/report",
                "agent_id": self.ini_agent_id_var.get().strip(),
                "timeout_ms": max(100, core.parse_int(self.ini_timeout_ms_var.get(), 2500)),
            }
            if not str(payload["path_value"]).startswith("/"):
                payload["path_value"] = "/" + str(payload["path_value"])
            path = tool.write_global_local_report_config(payload)
            self.log(f"已写入 GlobalConfig.ini -> {path}")
            self.reload_ini(show_message=False)
            messagebox.showinfo("保存成功", f"已写入:\n{path}")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def _set_compare(self, key: str, text: str, level: str) -> None:
        self.compare_vars[key].set(text)
        self.compare_labels[key].configure(fg=STATUS_COLORS.get(level, STATUS_COLORS["muted"]))

    def _refresh_compare(self) -> None:
        parsed = parse_base_url(self.base_url_var.get())
        host = self.ini_host_var.get().strip()
        port = max(0, core.parse_int(self.ini_port_var.get(), 0))
        agent_id = self.ini_agent_id_var.get().strip()
        path_value = self.ini_path_var.get().strip()
        enable = bool(self.ini_enable_var.get())

        host_match = bool(parsed.get("host")) and parsed.get("host", "").lower() == host.lower()
        port_match = int(parsed.get("port", 0) or 0) == port and port > 0
        agent_match = bool(self.agent_id_var.get().strip()) and self.agent_id_var.get().strip() == agent_id
        path_ok = bool(path_value) and path_value.startswith("/")

        self._set_compare("host", f"一致: {host}" if host_match else f"不一致: game_tool={parsed.get('host') or '-'} / ini={host or '-'}", "ok" if host_match else "error")
        self._set_compare("port", f"一致: {port}" if port_match else f"不一致: game_tool={parsed.get('port') or '-'} / ini={port or '-'}", "ok" if port_match else "error")
        self._set_compare("agent", f"一致: {agent_id}" if agent_match else f"不一致: game_tool={self.agent_id_var.get().strip() or '-'} / ini={agent_id or '-'}", "ok" if agent_match else "error")
        if path_ok and path_value == "/api/report":
            self._set_compare("path", f"正常: {path_value}", "ok")
        elif path_ok:
            self._set_compare("path", f"可用但非常规: {path_value}", "warn")
        else:
            self._set_compare("path", "缺失或格式错误，建议使用 /api/report", "error")
        self._set_compare("enable", "已开启" if enable else "未开启，QianNian 不会上报", "ok" if enable else "error")

        report_url = f"http://{host}:{port}{path_value}" if host and port and path_ok else "-"
        self._set_compare("url", report_url, "muted" if report_url == "-" else "ok")

        if all([host_match, port_match, agent_match, path_ok, enable]):
            self._set_compare("summary", "一致性通过。game_tool 与 QianNian LocalReport 配置匹配。", "ok")
        else:
            missing = []
            if not host_match:
                missing.append("Host")
            if not port_match:
                missing.append("Port")
            if not agent_match:
                missing.append("AgentId")
            if not path_ok:
                missing.append("Path")
            if not enable:
                missing.append("Enable")
            self._set_compare("summary", "存在不一致项: " + ", ".join(missing), "error")
