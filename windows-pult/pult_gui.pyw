# -*- coding: utf-8 -*-
"""Окно пульта. Запускается на Windows без браузера."""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import urllib.request
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

ROOT = Path(__file__).resolve().parent
RESULTS = ROOT / "results"
ICON = ROOT / "pult.ico"
VERSION_PATH = ROOT / "version.txt"
UPDATE_BASE = "https://raw.githubusercontent.com/YarBogos-svg/pult/main"
UPDATE_FILES = ("pult_gui.pyw", "parsers/gng.py", "parsers/partsagro.py", "requirements.txt")
JOBS = {
    "gng": {
        "label": "G&G",
        "site": "John Deere · gngzap.com",
        "script": ROOT / "parsers" / "gng.py",
        "dir": RESULTS / "gng",
        "file": "gng.xlsx",
    },
    "parts": {
        "label": "PartsAgro",
        "site": "Все бренды · data.partsagro.ru",
        "script": ROOT / "parsers" / "partsagro.py",
        "dir": RESULTS / "partsagro",
        "file": "partsagro.xlsx",
    },
}

BG = "#14160f"
SURFACE = "#1d2116"
RAISED = "#272d1e"
FG = "#eceee4"
MUTED = "#9aa18e"
LINE = "#3a4230"
ACCENT = "#c6e05a"
INK = "#1a2208"

LABELS = {
    "idle": "Ждёт",
    "running": "Идёт",
    "done": "Готово",
    "error": "Ошибка",
    "stopped": "Стоп",
}


class Desk(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Пульт")
        self.configure(bg=BG)
        self.geometry("1120x760")
        self.minsize(900, 640)
        if ICON.is_file():
            try:
                self.iconbitmap(str(ICON))
            except tk.TclError:
                pass

        self.status = "idle"
        self.rows = 0
        self.proc = None
        self.out_file = None
        self.events = queue.Queue()
        self.limited = tk.BooleanVar(value=False)
        self.pages = tk.StringVar(value="10")
        self.logins = {job_id: tk.StringVar(value=self._saved_login(job_id)) for job_id in JOBS}
        self.passwords = {job_id: tk.StringVar() for job_id in JOBS}
        self.jobs = {}
        for job_id, spec in JOBS.items():
            self.jobs[job_id] = {
                "status": "idle",
                "proc": None,
                "rows": 0,
                "out": spec["dir"] / spec["file"],
            }

        self._build()
        self.after(200, self._pump)
        self.after(1200, self._check_update)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _latest_sha(self):
        request = urllib.request.Request(
            "https://github.com/YarBogos-svg/pult.git/info/refs?service=git-upload-pack",
            headers={"User-Agent": "pult"},
        )
        with urllib.request.urlopen(request, timeout=15) as response:
            text = response.read().decode("utf-8", "replace")
        for line in text.splitlines():
            if "refs/heads/main" in line:
                token = line.split()[0]
                return token[-40:]
        raise RuntimeError("не нашёл свежий коммит")

    def _fetch(self, path, sha):
        url = f"https://raw.githubusercontent.com/YarBogos-svg/pult/{sha}/{path}"
        request = urllib.request.Request(url, headers={"User-Agent": "pult"})
        with urllib.request.urlopen(request, timeout=20) as response:
            return response.read()

    def _local_version(self):
        if VERSION_PATH.is_file():
            return VERSION_PATH.read_text(encoding="utf-8").strip()
        return "0"

    def _check_update(self):
        if self._running_count():
            self._append("Сначала останови сбор, потом обновление.")
            return
        self.update_btn.configure(text="Проверяю…")
        threading.Thread(target=self._update_worker, daemon=True).start()

    def _update_worker(self):
        try:
            sha = self._latest_sha()
            remote = self._fetch("version.txt", sha).decode("utf-8").strip()
        except Exception as error:
            text = str(error)
            if "rate limit" in text or "403" in text:
                text = "GitHub временно ограничил запросы. Подожди минуту и нажми «Обновить» ещё раз."
            self.events.put(("log", f"Обновление недоступно: {text}"))
            self.events.put(("update_idle", ""))
            return
        if not remote or remote == self._local_version():
            self.events.put(("log", "Обновлений нет"))
            self.events.put(("update_idle", ""))
            return
        try:
            for name in UPDATE_FILES:
                data = self._fetch(f"windows-pult/{name}", sha)
                target = ROOT / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
            VERSION_PATH.write_text(remote + "\n", encoding="utf-8")
        except Exception as error:
            self.events.put(("log", f"Не удалось обновиться: {error}"))
            self.events.put(("update_idle", ""))
            return
        self.events.put(("restart", remote))

    def _restart(self):
        subprocess.Popen([sys.executable, str(ROOT / "pult_gui.pyw")], cwd=str(ROOT))
        self.destroy()

    def _saved_login(self, job_id):
        path = ROOT / f"login_{job_id}.txt"
        if path.is_file():
            return path.read_text(encoding="utf-8").strip()
        legacy = ROOT / "login.txt"
        if job_id == "gng" and legacy.is_file():
            return legacy.read_text(encoding="utf-8").strip()
        return ""

    def _build(self):
        shell = tk.Frame(self, bg=BG)
        shell.pack(fill="both", expand=True, padx=18, pady=16)

        header = tk.Frame(shell, bg=BG)
        header.pack(fill="x")
        tk.Label(header, text="ЛИЧНЫЙ ПУЛЬТ", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
        title_row = tk.Frame(header, bg=BG)
        title_row.pack(fill="x")
        tk.Label(title_row, text="Парсеры", bg=BG, fg=FG, font=("Segoe UI", 22, "bold")).pack(side="left")
        self.count = tk.Label(title_row, text="в работе 0", bg=SURFACE, fg=MUTED, font=("Segoe UI", 9), padx=10, pady=4)
        self.count.pack(side="right")
        self.update_btn = tk.Button(
            title_row,
            text="Обновить",
            command=self._check_update,
            bg=SURFACE,
            fg=FG,
            relief="flat",
            font=("Segoe UI", 9),
        )
        self.update_btn.pack(side="right", padx=(0, 8))
        tk.Label(
            header,
            text="Окно программы на этом компьютере. Браузер не нужен.",
            bg=BG,
            fg=MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        body = tk.Frame(shell, bg=BG)
        body.pack(fill="both", expand=True, pady=(16, 0))
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        side = tk.Frame(body, bg=BG, width=320)
        side.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        side.grid_propagate(False)
        side.rowconfigure(0, weight=1)
        side.columnconfigure(0, weight=1)

        canvas = tk.Canvas(side, bg=SURFACE, highlightthickness=1, highlightbackground=LINE, bd=0)
        scroll = tk.Scrollbar(side, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        card = tk.Frame(canvas, bg=SURFACE, padx=14, pady=14)
        card_window = canvas.create_window((0, 0), window=card, anchor="nw")

        def _fit(event):
            canvas.itemconfigure(card_window, width=event.width)
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _grow(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", _fit)
        card.bind("<Configure>", _grow)
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-event.delta / 120), "units"))

        self.badges = {}
        self.actions = {}
        self.row_labels = {}
        self.file_btns = {}
        tk.Label(card, text="Парсеры", bg=SURFACE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 8))
        for job_id, spec in JOBS.items():
            block = tk.Frame(card, bg=RAISED, padx=10, pady=10)
            block.pack(fill="x", pady=(0, 8))
            top = tk.Frame(block, bg=RAISED)
            top.pack(fill="x")
            tk.Label(top, text=spec["label"], bg=RAISED, fg=FG, font=("Segoe UI", 14, "bold")).pack(side="left")
            badge = tk.Label(top, text="Ждёт", bg=SURFACE, fg=FG, font=("Segoe UI", 9), padx=8, pady=2)
            badge.pack(side="right")
            self.badges[job_id] = badge
            tk.Label(block, text=spec["site"], bg=RAISED, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 8))
            action = tk.Button(block, text=spec["label"], command=lambda key=job_id: self._toggle(key), font=("Segoe UI", 11, "bold"))
            action.pack(fill="x", ipady=6)
            self.actions[job_id] = action
            meta = tk.Frame(block, bg=RAISED)
            meta.pack(fill="x", pady=(6, 0))
            rows = tk.Label(meta, text="0 строк", bg=RAISED, fg=MUTED, font=("Consolas", 10))
            rows.pack(side="left")
            self.row_labels[job_id] = rows
            file_btn = tk.Button(
                meta, text="файла нет", command=lambda key=job_id: self._open_job(key), bg=RAISED, fg=MUTED, relief="flat", font=("Segoe UI", 10)
            )
            file_btn.pack(side="right")
            self.file_btns[job_id] = file_btn
            self._field(block, self.logins[job_id], "Логин", bg=RAISED)
            self._field(block, self.passwords[job_id], "Пароль", show="•", bg=RAISED)

        tk.Frame(card, bg=LINE, height=1).pack(fill="x", pady=(6, 10))
        tk.Label(card, text="Объём", bg=SURFACE, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w")
        modes = tk.Frame(card, bg=SURFACE)
        modes.pack(fill="x", pady=(4, 8))
        self.all_btn = tk.Button(modes, text="Весь каталог", command=lambda: self._mode(False))
        self.limit_btn = tk.Button(modes, text="Лимит", command=lambda: self._mode(True))
        self.all_btn.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self.limit_btn.pack(side="left", expand=True, fill="x")

        self.limit_box = tk.Frame(card, bg=SURFACE)
        self.limit_entry = tk.Entry(
            self.limit_box, textvariable=self.pages, bg=BG, fg=FG, insertbackground=FG, relief="flat", font=("Consolas", 12)
        )
        self.limit_entry.pack(fill="x", ipady=6, pady=(0, 8))
        self._clipboard(self.limit_entry)
        self.hint = tk.Label(
            card,
            text="G&G считает страницы, PartsAgro — товары на бренд.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            wraplength=280,
            justify="left",
        )
        self.hint.pack(anchor="w", pady=(0, 8))

        self.action = self.actions["gng"]
        self.pause_btn = tk.Button(
            card,
            text="Пауза",
            command=self._pause,
            bg=RAISED,
            fg=FG,
            relief="flat",
            font=("Segoe UI", 12, "bold"),
        )
        self.resume = tk.Button(
            card,
            text="Продолжить",
            command=self._resume,
            bg=ACCENT,
            fg=INK,
            relief="flat",
            font=("Segoe UI", 12, "bold"),
        )

        meta = tk.Frame(card, bg=SURFACE)
        self.rows_label = self.row_labels["gng"]
        self.file_btn = self.file_btns["gng"]

        self.error = tk.Label(card, text="", bg=SURFACE, fg=FG, font=("Segoe UI", 9), wraplength=280, justify="left")
        self.error.pack(anchor="w", pady=(8, 0))

        log_wrap = tk.Frame(body, bg=BG, highlightbackground=LINE, highlightthickness=1)
        log_wrap.grid(row=0, column=1, sticky="nsew")
        bar = tk.Frame(log_wrap, bg=BG)
        bar.pack(fill="x")
        self.log_tabs = {}
        self.logs = {}
        tabs = tk.Frame(bar, bg=BG)
        tabs.pack(side="left", padx=8, pady=8)
        for job_id, spec in JOBS.items():
            tab = tk.Button(
                tabs,
                text=spec["label"],
                command=lambda key=job_id: self._show_log(key),
                relief="flat",
                font=("Segoe UI", 11, "bold"),
                padx=14,
                pady=6,
            )
            tab.pack(side="left", padx=(0, 6))
            self.log_tabs[job_id] = tab
        tk.Button(
            bar,
            text="Лог ошибок",
            command=self._open_log,
            bg=BG,
            fg=FG,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(side="right", padx=12, pady=8)
        self.log_host = tk.Frame(log_wrap, bg=BG)
        self.log_host.pack(fill="both", expand=True)
        for job_id, spec in JOBS.items():
            log = tk.Text(
                self.log_host, bg=BG, fg=MUTED, insertbackground=FG, relief="flat", font=("Consolas", 10), wrap="word", padx=12, pady=8
            )
            log.tag_configure("err", foreground="#ffb4a2")
            log.tag_configure("pause", foreground="#c6e05a")
            log.insert("1.0", f"Журнал {spec['label']}. Запусти его — здесь будут только его строки.")
            log.configure(state="disabled")
            self.logs[job_id] = log
        self.active_log = "gng"
        self._show_log("gng")
        self._paint_mode()

    def _field(self, parent, variable, caption, show=None, bg=SURFACE):
        tk.Label(parent, text=caption, bg=bg, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))
        entry = tk.Entry(
            parent, textvariable=variable, bg=BG, fg=FG, insertbackground=FG, relief="flat", font=("Segoe UI", 11), show=show
        )
        entry.pack(fill="x", ipady=7)
        self._clipboard(entry)
        return entry

    def _clipboard(self, entry):
        def paste(_event=None):
            try:
                text = entry.clipboard_get()
            except tk.TclError:
                return "break"
            try:
                entry.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            entry.insert("insert", text)
            return "break"

        def copy(_event=None):
            try:
                text = entry.selection_get()
            except tk.TclError:
                return "break"
            entry.clipboard_clear()
            entry.clipboard_append(text)
            return "break"

        def cut(_event=None):
            copy()
            try:
                entry.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            return "break"

        def select_all(_event=None):
            entry.select_range(0, "end")
            entry.icursor("end")
            return "break"

        def on_ctrl(event):
            action = {86: paste, 67: copy, 88: cut, 65: select_all}.get(event.keycode)
            if action:
                return action()
            return None

        entry.bind("<Control-Key>", on_ctrl)
        entry.bind("<Shift-Insert>", paste)
        menu = tk.Menu(entry, tearoff=0, bg=SURFACE, fg=FG, activebackground=ACCENT, activeforeground=INK)
        menu.add_command(label="Вставить", command=paste)
        menu.add_command(label="Копировать", command=copy)
        menu.add_command(label="Вырезать", command=cut)
        menu.add_separator()
        menu.add_command(label="Выделить всё", command=select_all)

        def popup(event):
            entry.focus_set()
            menu.tk_popup(event.x_root, event.y_root)
            return "break"

        entry.bind("<Button-3>", popup)

    def _style_button(self, button, primary):
        if primary:
            button.configure(bg=ACCENT, fg=INK, activebackground=ACCENT, activeforeground=INK, relief="flat", bd=0)
        else:
            button.configure(bg=BG, fg=MUTED, activebackground=RAISED, activeforeground=FG, relief="flat", bd=0, highlightbackground=LINE, highlightthickness=1)

    def _mode(self, limited):
        self.limited.set(limited)
        self._paint_mode()

    def _paint_mode(self):
        self._style_button(self.all_btn, not self.limited.get())
        self._style_button(self.limit_btn, self.limited.get())
        if self.limited.get():
            self.hint.pack_forget()
            if not self.limit_box.winfo_ismapped():
                self.limit_box.pack(fill="x", before=self.error)
        else:
            self.limit_box.pack_forget()
            if not self.hint.winfo_ismapped():
                self.hint.pack(fill="x", before=self.error)
        for job_id, spec in JOBS.items():
            running = self.jobs[job_id]["status"] == "running"
            button = self.actions[job_id]
            if running:
                button.configure(text="Остановить", bg=RAISED, fg=FG, activebackground=RAISED, activeforeground=FG, relief="flat")
            else:
                button.configure(text=spec["label"], bg=ACCENT, fg=INK, activebackground=ACCENT, activeforeground=INK, relief="flat")

    def _running_count(self):
        return sum(1 for job in self.jobs.values() if job["status"] == "running")

    def _set_status(self, job_id, status):
        job = self.jobs[job_id]
        job["status"] = status
        self.badges[job_id].configure(text=LABELS.get(status, status))
        self.count.configure(text=f"в работе {self._running_count()}")
        if self._running_count() == 0:
            self.resume.pack_forget()
            self.pause_btn.pack_forget()
        else:
            self._show_pause()
        self._paint_mode()

    def _show_log(self, job_id):
        self.active_log = job_id
        for key, log in self.logs.items():
            log.pack_forget()
        self.logs[job_id].pack(fill="both", expand=True)
        for key, tab in self.log_tabs.items():
            if key == job_id:
                tab.configure(bg=ACCENT, fg=INK, activebackground=ACCENT, activeforeground=INK)
            else:
                tab.configure(bg=RAISED, fg=FG, activebackground=RAISED, activeforeground=FG)

    def _append(self, line, job_id=None):
        target = job_id or self.active_log
        log = self.logs[target]
        log.configure(state="normal")
        current = log.get("1.0", "end").strip()
        if current.startswith("Журнал "):
            log.delete("1.0", "end")
        tag = ""
        if line.startswith("ОШИБКА") or line.startswith("Код выхода"):
            tag = "err"
        elif line.startswith("ПАУЗА"):
            tag = "pause"
        log.insert("end", line + "\n", tag)
        log.see("end")
        log.configure(state="disabled")
        if job_id and job_id != self.active_log and self.jobs[job_id]["status"] == "running":
            self._show_log(job_id)
        if line.startswith("ПАУЗА"):
            self.pause_btn.pack_forget()
            if not self.resume.winfo_ismapped():
                self.resume.pack(fill="x", ipady=8, pady=(0, 8), before=self.error)
            self.error.configure(text="Пауза. Когда будешь готов, нажми «Продолжить».", fg=ACCENT)
        elif line.startswith("Продолжаю") or line.startswith("Остановка") or line.startswith("DONE") or line.startswith("Готово"):
            if not any(job["status"] == "running" and self.resume.winfo_ismapped() for job in self.jobs.values()):
                self.resume.pack_forget()
            if self._running_count():
                self._show_pause()
        if line.startswith("ОШИБКА") or line.startswith("Код выхода"):
            self.error.configure(text=line, fg="#ffb4a2")
        if job_id and line.startswith("ROWS "):
            tail = line.split()[-1]
            if tail.isdigit():
                self.jobs[job_id]["rows"] = int(tail)
                self.row_labels[job_id].configure(text=f"{tail} строк")
        if job_id and self.jobs[job_id]["out"].is_file():
            self.file_btns[job_id].configure(text="Открыть Excel", fg=FG)

    def _show_pause(self):
        if self.resume.winfo_ismapped():
            return
        if not self.pause_btn.winfo_ismapped():
            self.pause_btn.pack(fill="x", ipady=8, pady=(0, 8), before=self.error)

    def _pause(self):
        for job_id, job in self.jobs.items():
            if job["status"] != "running":
                continue
            folder = JOBS[job_id]["dir"]
            folder.mkdir(parents=True, exist_ok=True)
            (folder / "pause.flag").write_text("1", encoding="utf-8")
        self.pause_btn.pack_forget()
        self._append("Пауза запрошена. Дождусь конца текущего шага.")

    def _resume(self):
        for spec in JOBS.values():
            spec["dir"].mkdir(parents=True, exist_ok=True)
            (spec["dir"] / "continue.flag").write_text("1", encoding="utf-8")
        self.resume.pack_forget()
        self.error.configure(text="")
        self._append("Сигнал «Продолжить» отправлен")
        if self._running_count():
            self._show_pause()

    def _toggle(self, job_id):
        self.error.configure(text="")
        if self.jobs[job_id]["status"] == "running":
            self._stop(job_id)
        else:
            self._start(job_id)

    def _start(self, job_id):
        spec = JOBS[job_id]
        if not spec["script"].is_file():
            self.error.configure(text=f"Не найден парсер {spec['label']}.")
            return
        pages = 0
        if self.limited.get():
            try:
                pages = int(self.pages.get())
            except ValueError:
                self.error.configure(text="Лимит — целое число.")
                return
            if pages < 1 or pages > 5000:
                self.error.configure(text="Лимит — от 1 до 5000.")
                return
        login = self.logins[job_id].get().strip()
        (ROOT / f"login_{job_id}.txt").write_text(login, encoding="utf-8")
        spec["dir"].mkdir(parents=True, exist_ok=True)
        out_file = spec["dir"] / spec["file"]
        self.jobs[job_id]["out"] = out_file
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["GNG_LOGIN"] = self.logins["gng"].get().strip()
        env["GNG_PASSWORD"] = self.passwords["gng"].get()
        env["PARTS_LOGIN"] = self.logins["parts"].get().strip()
        env["PARTS_PASSWORD"] = self.passwords["parts"].get()
        env.pop("PLAYWRIGHT_BROWSERS_PATH", None)
        flags = subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
        try:
            proc = subprocess.Popen(
                [sys.executable, "-u", str(spec["script"]), "--pages", str(pages), "--out", str(out_file), "--show"],
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                creationflags=flags,
            )
        except OSError as error:
            self.error.configure(text=str(error))
            return
        self.jobs[job_id]["proc"] = proc
        self.jobs[job_id]["rows"] = 0
        self.row_labels[job_id].configure(text="0 строк")
        self.file_btns[job_id].configure(text="файла нет", fg=MUTED)
        self._set_status(job_id, "running")
        limit_text = f"лимит {pages}" if pages else "весь каталог"
        self._append(f"Старт, {limit_text}", job_id)
        threading.Thread(target=self._read, args=(job_id, proc), daemon=True).start()

    def _read(self, job_id, proc):
        stream = proc.stdout
        if stream is not None:
            for raw in iter(stream.readline, b""):
                self.events.put(("log", job_id, raw.decode("utf-8", "replace").rstrip()))
            stream.close()
        self.events.put(("exit", job_id, proc.wait()))

    def _stop(self, job_id):
        proc = self.jobs[job_id]["proc"]
        if proc is None:
            return
        self._append("Останавливаю", job_id)
        if os.name == "nt":
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
        else:
            proc.terminate()
        self._set_status(job_id, "stopped")

    def _pump(self):
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            kind = event[0]
            if kind == "log":
                if len(event) == 3:
                    _, job_id, payload = event
                else:
                    job_id, payload = None, event[1]
                if payload:
                    self._append(payload, job_id)
            elif kind == "update_idle":
                self.update_btn.configure(text="Обновить")
            elif kind == "restart":
                self._append(f"Обновление {event[1]} установлено. Перезапускаю.")
                self.after(700, self._restart)
            elif kind == "exit":
                _, job_id, code = event
                if self.jobs[job_id]["status"] == "running":
                    self.jobs[job_id]["proc"] = None
                    self._set_status(job_id, "done" if code == 0 else "error")
                    self._append("Готово" if code == 0 else f"Код выхода {code}", job_id)
                if self.jobs[job_id]["out"].is_file():
                    self.file_btns[job_id].configure(text="Открыть Excel", fg=FG)
        self.after(200, self._pump)

    def _open_log(self):
        spec = JOBS[self.active_log]
        path = spec["dir"] / spec["file"].replace(".xlsx", ".log")
        if path.is_file():
            self._open_path(path)
            return
        self.error.configure(text=f"Лога {spec['label']} ещё нет — он появится после запуска.")

    def _open_job(self, job_id):
        path = self.jobs[job_id]["out"]
        if path.is_file():
            self._open_path(path)

    def _open_path(self, path):
        if os.name == "nt":
            os.startfile(path)  # noqa: S606 — открыть свой файл
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def _close(self):
        if self._running_count():
            if not messagebox.askokcancel("Пульт", "Сбор ещё идёт. Закрыть и остановить?"):
                return
            for job_id, job in list(self.jobs.items()):
                if job["status"] == "running":
                    self._stop(job_id)
        self.destroy()


if __name__ == "__main__":
    Desk().mainloop()
