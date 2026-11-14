"""Tkinter desktop GUI.

Layout: toolbar / file table / options / progress + live log.
Capabilities: task queue with configurable concurrency, streamed log, cancel (kills the
child process tree), failed rows stay in the table for a retry, settings dialog.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pdfmerge
from pdfmerge import __version__
from pdfmerge import config as pdfconfig
from pdfmerge.wordcheck import word_available

LEVELS = pdfconfig.VALID_LEVELS
STAGE_LABEL = {"convert": "转换 docx", "merge": "合并", "compress": "压缩"}
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} GB"


def worker_command() -> list[str]:
    """Command that runs the CLI worker; frozen builds reuse their own executable."""
    if getattr(sys, "frozen", False):
        return [sys.executable, "--cli"]
    return [sys.executable, "-m", "pdfmerge.cli"]


def worker_env() -> dict:
    env = os.environ.copy()
    if not getattr(sys, "frozen", False):
        source_root = str(Path(pdfmerge.__file__).resolve().parent.parent)
        existing = env.get("PYTHONPATH", "")
        env["PYTHONPATH"] = source_root + (os.pathsep + existing if existing else "")
    return env


class SettingsDialog(tk.Toplevel):
    def __init__(self, app: App):
        super().__init__(app.root)
        self.app = app
        self.title("设置")
        self.transient(app.root)
        self.resizable(False, False)
        cfg = app.cfg
        self.vars = {
            "out_dir": tk.StringVar(value=cfg["general"]["out_dir"]),
            "name_template": tk.StringVar(value=cfg["general"]["name_template"]),
            "max_parallel": tk.IntVar(value=int(cfg["general"]["max_parallel"])),
            "open_output_folder": tk.BooleanVar(value=bool(cfg["general"]["open_output_folder"])),
            "keep_merged": tk.BooleanVar(value=bool(cfg["general"]["keep_merged"])),
            "level": tk.StringVar(value=cfg["compress"]["level"]),
            "public_key": tk.StringVar(value=cfg["ilovepdf"]["public_key"]),
            "secret_key": tk.StringVar(value=cfg["ilovepdf"]["secret_key"]),
            "proxy": tk.StringVar(value=cfg["ilovepdf"]["proxy"]),
            "timeout_seconds": tk.IntVar(value=int(cfg["ilovepdf"]["timeout_seconds"])),
            "sdk_path": tk.StringVar(value=cfg["ilovepdf"]["sdk_path"]),
        }
        pad = {"padx": 8, "pady": 4}
        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        row = 0

        def label(text):
            nonlocal row
            ttk.Label(body, text=text).grid(row=row, column=0, sticky="w", **pad)

        label("默认输出目录")
        line = ttk.Frame(body)
        line.grid(row=row, column=1, sticky="ew", **pad)
        ttk.Entry(line, textvariable=self.vars["out_dir"], width=40).pack(side="left", fill="x", expand=True)
        ttk.Button(line, text="…", width=3, command=self._pick_out).pack(side="left", padx=(4, 0))
        row += 1
        ttk.Label(body, text="留空 = 主体文件所在目录", foreground="#777").grid(
            row=row, column=1, sticky="w", padx=8
        )

        row += 1
        label("输出名模板")
        ttk.Entry(body, textvariable=self.vars["name_template"], width=30).grid(
            row=row, column=1, sticky="w", **pad
        )
        row += 1
        ttk.Label(body, text="可用 {main} 代表主体文件名", foreground="#777").grid(
            row=row, column=1, sticky="w", padx=8
        )

        row += 1
        label("压缩级别")
        ttk.Combobox(body, textvariable=self.vars["level"], values=LEVELS, width=16, state="readonly").grid(
            row=row, column=1, sticky="w", **pad
        )

        row += 1
        label("最大并发任务")
        ttk.Spinbox(body, from_=1, to=8, textvariable=self.vars["max_parallel"], width=6).grid(
            row=row, column=1, sticky="w", **pad
        )

        row += 1
        ttk.Checkbutton(body, text="完成后打开输出目录", variable=self.vars["open_output_folder"]).grid(
            row=row, column=1, sticky="w", **pad
        )
        row += 1
        ttk.Checkbutton(body, text="保留合并中间文件（_merged.pdf）", variable=self.vars["keep_merged"]).grid(
            row=row, column=1, sticky="w", **pad
        )

        row += 1
        ttk.Separator(body, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)

        row += 1
        label("iLovePDF 公钥")
        ttk.Entry(body, textvariable=self.vars["public_key"], width=44).grid(
            row=row, column=1, sticky="w", **pad
        )
        row += 1
        label("iLovePDF 私钥")
        self.secret_entry = ttk.Entry(body, textvariable=self.vars["secret_key"], width=44, show="*")
        self.secret_entry.grid(row=row, column=1, sticky="w", **pad)
        row += 1
        self.show_secret = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            body,
            text="显示私钥",
            variable=self.show_secret,
            command=lambda: self.secret_entry.configure(show="" if self.show_secret.get() else "*"),
        ).grid(row=row, column=1, sticky="w", padx=8)

        row += 1
        ttk.Button(
            body, text="获取免费密钥（developer.ilovepdf.com）", command=self._open_keys_page
        ).grid(row=row, column=1, sticky="w", **pad)

        row += 1
        label("代理")
        ttk.Entry(body, textvariable=self.vars["proxy"], width=30).grid(row=row, column=1, sticky="w", **pad)
        row += 1
        ttk.Label(body, text="留空 = 直连；例：http://127.0.0.1:7890", foreground="#777").grid(
            row=row, column=1, sticky="w", padx=8
        )

        row += 1
        label("超时（秒）")
        ttk.Spinbox(
            body, from_=60, to=1800, increment=60, textvariable=self.vars["timeout_seconds"], width=8
        ).grid(row=row, column=1, sticky="w", **pad)

        row += 1
        label("SDK 目录")
        sdk_line = ttk.Frame(body)
        sdk_line.grid(row=row, column=1, sticky="ew", **pad)
        ttk.Entry(sdk_line, textvariable=self.vars["sdk_path"], width=34).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(sdk_line, text="…", width=3, command=self._pick_sdk).pack(side="left", padx=(4, 0))
        row += 1
        ttk.Label(body, text="留空 = 使用已安装的 ilovepdf 包", foreground="#777").grid(
            row=row, column=1, sticky="w", padx=8
        )

        row += 1
        buttons = ttk.Frame(body)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="保存", command=self._save).pack(side="right")

        self.grab_set()

    def _pick_out(self):
        chosen = filedialog.askdirectory(parent=self, title="选择默认输出目录")
        if chosen:
            self.vars["out_dir"].set(chosen)

    def _pick_sdk(self):
        chosen = filedialog.askdirectory(parent=self, title="选择 ilovepdf SDK 目录")
        if chosen:
            self.vars["sdk_path"].set(chosen)

    def _open_keys_page(self):
        webbrowser.open(pdfconfig.ILOVEPDF_KEYS_URL)

    def _save(self):
        cfg = self.app.cfg
        cfg["general"]["out_dir"] = self.vars["out_dir"].get().strip()
        cfg["general"]["name_template"] = self.vars["name_template"].get().strip() or "{main}"
        cfg["general"]["max_parallel"] = max(1, min(8, int(self.vars["max_parallel"].get())))
        cfg["general"]["open_output_folder"] = bool(self.vars["open_output_folder"].get())
        cfg["general"]["keep_merged"] = bool(self.vars["keep_merged"].get())
        cfg["compress"]["level"] = self.vars["level"].get()
        cfg["ilovepdf"]["public_key"] = self.vars["public_key"].get().strip()
        cfg["ilovepdf"]["secret_key"] = self.vars["secret_key"].get().strip()
        cfg["ilovepdf"]["proxy"] = self.vars["proxy"].get().strip()
        cfg["ilovepdf"]["timeout_seconds"] = max(60, int(self.vars["timeout_seconds"].get()))
        cfg["ilovepdf"]["sdk_path"] = self.vars["sdk_path"].get().strip()
        pdfconfig.save(cfg)
        self.app.apply_config()
        self.app.log("设置已保存: " + str(pdfconfig.config_path()))
        self.destroy()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.cfg = pdfconfig.load()
        self.max_parallel = max(1, int(self.cfg["general"]["max_parallel"]))
        self.rows: dict[str, dict] = {}
        self.main_path: str | None = None
        self.next_id = 0
        self.pending: list[dict] = []
        self.running: dict[int, subprocess.Popen | None] = {}
        self.ui_queue: queue.Queue[tuple] = queue.Queue()
        self.scheduler_lock = threading.Lock()
        self.word_ok = word_available()

        root.title(f"PDF 合并压缩工具 {__version__}")
        root.geometry("920x640")
        root.minsize(760, 520)
        self._build_ui()
        self.apply_config()
        self.root.after(80, self._pump)

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=(10, 8, 10, 4))
        top.pack(fill="x")
        ttk.Button(top, text="添加文件…", command=self.add_files).pack(side="left")
        ttk.Button(top, text="添加文件夹…", command=self.add_folder).pack(side="left", padx=4)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(top, text="移除选中", command=self.remove_selected).pack(side="left")
        ttk.Button(top, text="上移", command=lambda: self.move(-1)).pack(side="left", padx=2)
        ttk.Button(top, text="下移", command=lambda: self.move(1)).pack(side="left")
        ttk.Button(top, text="清空", command=self.clear_rows).pack(side="left", padx=4)
        ttk.Button(top, text="重试失败", command=self.retry_failed).pack(side="left", padx=(12, 0))
        ttk.Button(top, text="设置…", command=lambda: SettingsDialog(self)).pack(side="right")

        table_area = ttk.Frame(self.root, padding=(10, 0))
        table_area.pack(fill="both", expand=True)
        columns = ("main", "name", "kind", "size", "state")
        self.tree = ttk.Treeview(table_area, columns=columns, show="headings", selectmode="extended")
        headings = {
            "main": ("", 40),
            "name": ("文件名", 460),
            "kind": ("类型", 70),
            "size": ("大小", 90),
            "state": ("状态", 80),
        }
        for key, (text, width) in headings.items():
            self.tree.heading(key, text=text)
            self.tree.column(
                key, width=width, anchor="w" if key == "name" else "center", stretch=key == "name"
            )
        scroll = ttk.Scrollbar(table_area, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.tag_configure("ok", foreground="#1a7f37")
        self.tree.tag_configure("fail", foreground="#c0392b")
        self.tree.tag_configure("warn", foreground="#b26a00")
        self.tree.tag_configure("main", background="#eef5ff")

        opts = ttk.Frame(self.root, padding=(10, 6))
        opts.pack(fill="x")
        self.compress_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="iLovePDF 压缩", variable=self.compress_var).pack(side="left")
        ttk.Label(opts, text="  级别").pack(side="left")
        self.level_var = tk.StringVar(value=self.cfg["compress"]["level"])
        ttk.Combobox(opts, textvariable=self.level_var, values=LEVELS, width=12, state="readonly").pack(
            side="left", padx=4
        )
        ttk.Label(opts, text="  输出目录").pack(side="left", padx=(12, 0))
        self.outdir_var = tk.StringVar(value=self.cfg["general"]["out_dir"])
        ttk.Entry(opts, textvariable=self.outdir_var, width=34).pack(side="left", padx=4)
        ttk.Button(opts, text="…", width=3, command=self._pick_outdir).pack(side="left")

        action = ttk.Frame(self.root, padding=(10, 4))
        action.pack(fill="x")
        self.start_btn = ttk.Button(action, text="开始合并压缩", command=self.start)
        self.start_btn.pack(side="left")
        self.cancel_btn = ttk.Button(action, text="取消", command=self.cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(action, mode="indeterminate", length=240)
        self.progress.pack(side="left", padx=10)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(action, textvariable=self.status_var).pack(side="left", padx=6)

        log_area = ttk.Frame(self.root, padding=(10, 4, 10, 10))
        log_area.pack(fill="both", expand=True)
        self.log_text = tk.Text(
            log_area,
            height=10,
            wrap="none",
            state="disabled",
            font=("Consolas", 9),
            background="#1e1e1e",
            foreground="#d4d4d4",
            insertbackground="#d4d4d4",
        )
        log_scroll = ttk.Scrollbar(log_area, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.log_text.tag_configure("err", foreground="#f48771")
        self.log_text.tag_configure("ok", foreground="#89d185")

    def apply_config(self):
        self.max_parallel = max(1, int(self.cfg["general"]["max_parallel"]))
        self.level_var.set(self.cfg["compress"]["level"])
        self.outdir_var.set(self.cfg["general"]["out_dir"])
        self.compress_var.set(bool(self.cfg["compress"]["enabled"]))

    def _kind_of(self, path: str) -> str:
        return "DOCX" if os.path.splitext(path)[1].lower() in (".docx", ".doc") else "PDF"

    def add_files(self):
        chosen = filedialog.askopenfilenames(
            title="选择 PDF / docx",
            filetypes=[("PDF 或 Word", "*.pdf *.docx *.doc"), ("所有文件", "*.*")],
        )
        self._add_paths(list(chosen))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 PDF / docx 的文件夹")
        if not folder:
            return
        found = []
        with os.scandir(folder) as entries:
            for entry in entries:
                if entry.is_file() and entry.name.lower().endswith((".pdf", ".docx", ".doc")):
                    found.append(entry.path)
        self._add_paths(sorted(found, key=lambda p: os.path.basename(p).lower()))

    def _add_paths(self, paths):
        added = 0
        existing = {row["path"] for row in self.rows.values()}
        for path in paths:
            path = os.path.abspath(path)
            if path in existing:
                continue
            try:
                size = os.path.getsize(path)
            except OSError:
                size = 0
            iid = self.tree.insert(
                "",
                "end",
                values=("", os.path.basename(path), self._kind_of(path), human_size(size), "待处理"),
            )
            state = "待处理"
            if not self.word_ok and path.lower().endswith((".docx", ".doc")):
                state = "缺 Word"
            self.rows[iid] = {"path": path, "state": state}
            self.tree.set(iid, "state", state)
            existing.add(path)
            added += 1
        if self.main_path is None and self.rows:
            self._set_main(self.tree.get_children()[0])
        self._refresh_marks()
        if added:
            self.log(f"已添加 {added} 个文件")
        if added and not self.word_ok and any(p.lower().endswith((".docx", ".doc")) for p in paths):
            self.log("未检测到 Microsoft Word：docx 无法转换，请先用 Word 另存为 PDF。", "err")

    def remove_selected(self):
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.rows.pop(iid, None)
        if self.main_path and not any(row["path"] == self.main_path for row in self.rows.values()):
            self.main_path = None
            remaining = self.tree.get_children()
            if remaining:
                self._set_main(remaining[0])
        self._refresh_marks()

    def clear_rows(self):
        if self.running and not messagebox.askyesno("确认", "仍有任务在运行，确定清空列表？"):
            return
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.main_path = None

    def move(self, delta: int):
        selected = self.tree.selection()
        if len(selected) != 1:
            messagebox.showinfo("提示", "上移/下移请只选中一行。")
            return
        iid = selected[0]
        kids = list(self.tree.get_children())
        index = kids.index(iid)
        target = index + delta
        if target < 0 or target >= len(kids):
            return
        self.tree.move(iid, "", target)
        self._refresh_marks()

    def _on_double_click(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self._set_main(iid)

    def _set_main(self, iid: str):
        self.main_path = self.rows[iid]["path"]
        self.tree.selection_set(iid)
        self._refresh_marks()

    def _refresh_marks(self):
        for iid, row in self.rows.items():
            self.tree.set(iid, "main", "●" if row["path"] == self.main_path else "")
            tags = []
            if row["path"] == self.main_path:
                tags.append("main")
            if row["state"] == "完成":
                tags.append("ok")
            elif row["state"] == "失败":
                tags.append("fail")
            elif row["state"] == "缺 Word":
                tags.append("warn")
            self.tree.item(iid, tags=tuple(tags))

    def _row_order(self) -> list[str]:
        return [iid for iid in self.tree.get_children() if iid in self.rows]

    def _ordered_paths(self) -> list[str]:
        paths = [self.rows[iid]["path"] for iid in self._row_order()]
        if self.main_path in paths:
            paths.remove(self.main_path)
            paths.insert(0, self.main_path)
        return paths

    def _pick_outdir(self):
        chosen = filedialog.askdirectory(title="选择输出目录")
        if chosen:
            self.outdir_var.set(chosen)

    def log(self, text: str, tag: str = ""):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text.rstrip() + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def start(self, only_failed: bool = False):
        if only_failed:
            paths = [self.rows[iid]["path"] for iid in self._row_order() if self.rows[iid]["state"] == "失败"]
        else:
            paths = self._ordered_paths()
        if not paths:
            messagebox.showinfo("提示", "列表为空，请先添加文件。")
            return
        if not self.word_ok and any(p.lower().endswith((".docx", ".doc")) for p in paths):
            messagebox.showerror(
                "未检测到 Microsoft Word",
                "列表中有 docx 文件，但本机没有可用的 Microsoft Word，无法转换。\n"
                "请先用 Word 另存为 PDF 再添加，或改用 PDF 输入。",
            )
            return
        task = {
            "id": self.next_id,
            "paths": paths,
            "out_dir": self.outdir_var.get().strip() or self.cfg["general"]["out_dir"].strip(),
            "compress": bool(self.compress_var.get()),
            "level": self.level_var.get() or "low",
        }
        self.next_id += 1
        self.pending.append(task)
        for iid in self._row_order():
            if self.rows[iid]["path"] in paths:
                self.rows[iid]["state"] = "排队"
        self._refresh_marks()
        self._dispatch()

    def retry_failed(self):
        self.start(only_failed=True)

    def _dispatch(self):
        with self.scheduler_lock:
            while self.pending and len(self.running) < self.max_parallel:
                task = self.pending.pop(0)
                self.running[task["id"]] = None
                threading.Thread(target=self._run_task, args=(task,), daemon=True).start()
            active = len(self.running) + len(self.pending)
        if active:
            self.progress.start(12)
            self.cancel_btn.configure(state="normal")
            self.status_var.set(f"运行中 {len(self.running)} / 排队 {len(self.pending)}")
        else:
            self.progress.stop()
            self.cancel_btn.configure(state="disabled")
            self.status_var.set("就绪")

    def _run_task(self, task: dict):
        head = os.path.basename(task["paths"][0]) if task["paths"] else "-"
        self.ui_queue.put(
            ("log", task["id"], f"── 任务 #{task['id']} 开始（{len(task['paths'])} 个文件，主体 {head}）")
        )
        cmd = worker_command() + task["paths"]
        if task["out_dir"]:
            cmd += ["--out", task["out_dir"]]
        if not task["compress"]:
            cmd += ["--no-compress"]
        else:
            cmd += ["--level", task["level"]]
        out_path = None
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=worker_env(),
                creationflags=NO_WINDOW,
            )
            with self.scheduler_lock:
                self.running[task["id"]] = proc
            assert proc.stdout is not None
            for line in proc.stdout:
                line = line.rstrip()
                if not line:
                    continue
                if line.startswith("STAGE:"):
                    self.ui_queue.put(("stage", task["id"], line))
                    continue
                if line.startswith("DONE:"):
                    out_path = line.split("DONE:", 1)[1].strip()
                self.ui_queue.put(("log", task["id"], line))
            code = proc.wait()
        except Exception as exc:
            code = -1
            self.ui_queue.put(("log", task["id"], f"启动失败: {exc}", "err"))
        self.ui_queue.put(("done", task["id"], code, out_path, task["paths"]))

    def cancel(self):
        with self.scheduler_lock:
            pending_ids = [task["id"] for task in self.pending]
            procs = [proc for proc in self.running.values() if proc]
            self.pending.clear()
        for proc in procs:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    capture_output=True,
                    creationflags=NO_WINDOW,
                )
            except Exception:
                proc.kill()
        for task_id in pending_ids:
            self.running.setdefault(task_id, None)
        self.log("已请求取消" + (f"（终止 {len(procs)} 个进程）" if procs else ""), "err")

    def _pump(self):
        try:
            while True:
                message = self.ui_queue.get_nowait()
                kind = message[0]
                if kind == "log":
                    self.log(message[2], message[3] if len(message) > 3 else "")
                elif kind == "stage":
                    parts = message[2].split(":")
                    name = STAGE_LABEL.get(parts[1], parts[1])
                    detail = parts[2] if len(parts) > 2 else ""
                    self.status_var.set(f"任务 #{message[1]} · {name} {detail}".strip())
                elif kind == "done":
                    self._on_task_done(*message[1:])
        except queue.Empty:
            pass
        self.root.after(80, self._pump)

    def _on_task_done(self, task_id: int, code: int, out_path, paths):
        with self.scheduler_lock:
            self.running.pop(task_id, None)
        state = "完成" if code == 0 else "失败"
        for iid in self._row_order():
            if self.rows[iid]["path"] in paths:
                self.rows[iid]["state"] = state
                self.tree.set(iid, "state", state)
        if code == 0:
            self.log(f"✔ 任务 #{task_id} 完成" + (f" → {out_path}" if out_path else ""), "ok")
            if out_path and self.cfg["general"]["open_output_folder"]:
                try:
                    os.startfile(os.path.dirname(out_path))
                except OSError:
                    pass
        else:
            self.log(f"✘ 任务 #{task_id} 失败（退出码 {code}）；失败文件已保留，可点「重试失败」", "err")
        self._refresh_marks()
        self._dispatch()


def main() -> int:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
