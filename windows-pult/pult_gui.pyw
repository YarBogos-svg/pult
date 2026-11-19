# -*- coding: utf-8 -*-
"""Окно пульта. Запускается на Windows без браузера."""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import urllib.request
from pathlib import Path
from tkinter import messagebox

ROOT = Path(__file__).resolve().parent
SCRIPT = ROOT / "parsers" / "gng.py"
RESULTS = ROOT / "results" / "gng"
ICON = ROOT / "pult.ico"
VERSION_PATH = ROOT / "version.txt"
UPDATE_BASE = "https://raw.githubusercontent.com/YarBogos-svg/pult/main"
UPDATE_FILES = ("pult_gui.pyw", "parsers/gng.py", "requirements.txt")

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
        self.geometry("980x640")
        self.minsize(760, 520)
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
        self.login = tk.StringVar(value=self._saved_login())
        self.password = tk.StringVar()

        self._build()
        self.after(200, self._pump)
        self.after(1200, self._check_update)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _fetch(self, path):
        request = urllib.request.Request(
            f"{UPDATE_BASE}/{path}",
            headers={"User-Agent": "pult", "Cache-Control": "no-cache"},
        )
        with urllib.request.urlopen(request, timeout=15) as response:
            return response.read()

    def _local_version(self):
        if VERSION_PATH.is_file():
            return VERSION_PATH.read_text(encoding="utf-8").strip()
        return "0"

    def _check_update(self):
        if self.status == "running":
            self._append("Сначала останови сбор, потом обновление.")
            return
        self.update_btn.configure(text="Проверяю…")
        threading.Thread(target=self._update_worker, daemon=True).start()

    def _update_worker(self):
        try:
            remote = self._fetch("version.txt").decode("utf-8").strip()
        except Exception as error:
            self.events.put(("log", f"Обновление недоступно: {error}"))
            self.events.put(("update_idle", ""))
            return
        if not remote or remote == self._local_version():
            self.events.put(("log", "Обновлений нет"))
            self.events.put(("update_idle", ""))
            return
        try:
            for name in UPDATE_FILES:
                data = self._fetch(f"windows-pult/{name}")
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

    def _saved_login(self):
        path = ROOT / "login.txt"
        if path.is_file():
            return path.read_text(encoding="utf-8").strip()
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

        card = tk.Frame(body, bg=SURFACE, highlightbackground=LINE, highlightthickness=1, padx=14, pady=14)
        card.grid(row=0, column=0, sticky="ns", padx=(0, 14))

        top = tk.Frame(card, bg=SURFACE)
        top.pack(fill="x")
        tk.Label(top, text="G&G", bg=SURFACE, fg=FG, font=("Segoe UI", 16, "bold")).pack(side="left")
        self.badge = tk.Label(top, text="Ждёт", bg=RAISED, fg=FG, font=("Segoe UI", 9), padx=8, pady=3)
        self.badge.pack(side="right")
        tk.Label(card, text="John Deere · gngzap.com", bg=SURFACE, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))

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
            text="Число страниц берётся с сайта, до последней.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            wraplength=260,
            justify="left",
        )
        self.hint.pack(anchor="w", pady=(0, 8))

        self.action = tk.Button(card, text="G&G", command=self._toggle, font=("Segoe UI", 12, "bold"))
        self.action.pack(fill="x", ipady=8, pady=(4, 8))
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
        meta.pack(fill="x")
        self.rows_label = tk.Label(meta, text="0 строк", bg=SURFACE, fg=MUTED, font=("Consolas", 10))
        self.rows_label.pack(side="left")
        self.file_btn = tk.Button(
            meta, text="файла нет", command=self._open_file, bg=SURFACE, fg=MUTED, relief="flat", font=("Segoe UI", 10)
        )
        self.file_btn.pack(side="right")

        tk.Label(card, text="Вход на сайт", bg=SURFACE, fg=FG, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(14, 2))
        tk.Label(card, text="Необязательно.", bg=SURFACE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
        self._field(card, self.login, "Логин")
        self._field(card, self.password, "Пароль", show="•")
        self.error = tk.Label(card, text="", bg=SURFACE, fg=FG, font=("Segoe UI", 9), wraplength=260, justify="left")
        self.error.pack(anchor="w", pady=(8, 0))

        log_wrap = tk.Frame(body, bg=BG, highlightbackground=LINE, highlightthickness=1)
        log_wrap.grid(row=0, column=1, sticky="nsew")
        bar = tk.Frame(log_wrap, bg=BG)
        bar.pack(fill="x")
        tk.Label(bar, text="Журнал G&G", bg=BG, fg=FG, font=("Segoe UI", 11, "bold")).pack(side="left", padx=12, pady=10)
        tk.Button(
            bar,
            text="Лог ошибок",
            command=self._open_log,
            bg=BG,
            fg=FG,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(side="right", padx=12, pady=8)
        self.log = tk.Text(
            log_wrap, bg=BG, fg=MUTED, insertbackground=FG, relief="flat", font=("Consolas", 10), wrap="word", padx=12, pady=8
        )
        self.log.tag_configure("err", foreground="#ffb4a2")
        self.log.tag_configure("pause", foreground="#c6e05a")
        self.log.pack(fill="both", expand=True)
        self.log.insert("1.0", "Нажми кнопку G&G — здесь пойдут страницы, строки и ошибки.")
        self.log.configure(state="disabled")
        self._paint_mode()

    def _field(self, parent, variable, caption, show=None):
        tk.Label(parent, text=caption, bg=SURFACE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))
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
        if self.status == "running":
            return
        self.limited.set(limited)
        self._paint_mode()

    def _paint_mode(self):
        self._style_button(self.all_btn, not self.limited.get())
        self._style_button(self.limit_btn, self.limited.get())
        if self.limited.get():
            self.hint.pack_forget()
            self.limit_box.pack(fill="x", before=self.action)
        else:
            self.limit_box.pack_forget()
            self.hint.pack(fill="x", before=self.action)
        if self.status == "running":
            self.action.configure(text="Остановить", bg=RAISED, fg=FG, activebackground=RAISED, activeforeground=FG, relief="flat")
        else:
            self.action.configure(text="G&G", bg=ACCENT, fg=INK, activebackground=ACCENT, activeforeground=INK, relief="flat")

    def _set_status(self, status):
        self.status = status
        self.badge.configure(text=LABELS.get(status, status))
        self.count.configure(text="в работе 1" if status == "running" else "в работе 0")
        if status != "running":
            self.resume.pack_forget()
            self.pause_btn.pack_forget()
        else:
            self._show_pause()
        self._paint_mode()

    def _append(self, line):
        self.log.configure(state="normal")
        current = self.log.get("1.0", "end").strip()
        if current.startswith("Нажми кнопку"):
            self.log.delete("1.0", "end")
        tag = ""
        if line.startswith("ОШИБКА") or line.startswith("Код выхода"):
            tag = "err"
        elif line.startswith("ПАУЗА"):
            tag = "pause"
        self.log.insert("end", line + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        if line.startswith("ПАУЗА"):
            self.pause_btn.pack_forget()
            self.resume.pack(fill="x", ipady=8, pady=(0, 8), before=self.rows_label.master)
            self.error.configure(text="Пауза. Когда будешь готов, нажми «Продолжить».", fg=ACCENT)
        elif line.startswith("Продолжаю") or line.startswith("Остановка") or line.startswith("DONE") or line.startswith("Готово"):
            self.resume.pack_forget()
            if self.status == "running":
                self._show_pause()
        if line.startswith("ОШИБКА") or line.startswith("Код выхода"):
            self.error.configure(text=line, fg="#ffb4a2")
        if line.startswith("ROWS "):
            tail = line.split()[-1]
            if tail.isdigit():
                self.rows = int(tail)
                self.rows_label.configure(text=f"{self.rows} строк")
        if self.out_file and self.out_file.is_file():
            self.file_btn.configure(text="Открыть Excel", fg=FG)

    def _show_pause(self):
        if self.resume.winfo_ismapped():
            return
        if not self.pause_btn.winfo_ismapped():
            self.pause_btn.pack(fill="x", ipady=8, pady=(0, 8), before=self.rows_label.master)

    def _pause(self):
        RESULTS.mkdir(parents=True, exist_ok=True)
        (RESULTS / "pause.flag").write_text("1", encoding="utf-8")
        self.pause_btn.pack_forget()
        self._append("Пауза запрошена. Дождусь конца текущего шага.")

    def _resume(self):
        RESULTS.mkdir(parents=True, exist_ok=True)
        (RESULTS / "continue.flag").write_text("1", encoding="utf-8")
        self.resume.pack_forget()
        self.error.configure(text="")
        self._append("Сигнал «Продолжить» отправлен")

    def _toggle(self):
        self.error.configure(text="")
        if self.status == "running":
            self._stop()
        else:
            self._start()

    def _start(self):
        if not SCRIPT.is_file():
            self.error.configure(text="Не найден парсер G&G.")
            return
        pages = 0
        if self.limited.get():
            try:
                pages = int(self.pages.get())
            except ValueError:
                self.error.configure(text="Лимит — целое число страниц.")
                return
            if pages < 1 or pages > 5000:
                self.error.configure(text="Лимит страниц — от 1 до 5000.")
                return
        login = self.login.get().strip()
        (ROOT / "login.txt").write_text(login, encoding="utf-8")
        RESULTS.mkdir(parents=True, exist_ok=True)
        self.out_file = RESULTS / "gng.xlsx"
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["GNG_LOGIN"] = login
        env["GNG_PASSWORD"] = self.password.get()
        env.pop("PLAYWRIGHT_BROWSERS_PATH", None)
        flags = 0
        if os.name == "nt":
            flags = subprocess.CREATE_NEW_PROCESS_GROUP
        try:
            self.proc = subprocess.Popen(
                [sys.executable, "-u", str(SCRIPT), "--pages", str(pages), "--out", str(self.out_file), "--show"],
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                creationflags=flags,
            )
        except OSError as error:
            self.error.configure(text=str(error))
            return
        self.rows = 0
        self.rows_label.configure(text="0 строк")
        self.file_btn.configure(text="файла нет", fg=MUTED)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._set_status("running")
        self._append("Старт G&G, " + (f"не больше {pages} стр." if pages else "весь каталог"))
        threading.Thread(target=self._read, args=(self.proc,), daemon=True).start()

    def _read(self, proc):
        stream = proc.stdout
        if stream is not None:
            for raw in iter(stream.readline, b""):
                self.events.put(("log", raw.decode("utf-8", "replace").rstrip()))
            stream.close()
        code = proc.wait()
        self.events.put(("exit", code))

    def _stop(self):
        proc = self.proc
        if proc is None:
            return
        self._append("Останавливаю")
        if os.name == "nt":
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
        else:
            proc.terminate()
        self._set_status("stopped")

    def _pump(self):
        while True:
            try:
                kind, payload = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == "log" and payload:
                self._append(payload)
            elif kind == "update_idle":
                self.update_btn.configure(text="Обновить")
            elif kind == "restart":
                self._append(f"Обновление {payload} установлено. Перезапускаю.")
                self.after(700, self._restart)
            elif kind == "exit" and self.status == "running":
                self.proc = None
                if payload == 0:
                    self._set_status("done")
                    self._append("Готово")
                else:
                    self._set_status("error")
                    self._append(f"Код выхода {payload}")
                if self.out_file and self.out_file.is_file():
                    self.file_btn.configure(text="Открыть Excel", fg=FG)
        self.after(200, self._pump)

    def _open_log(self):
        path = RESULTS / "gng.log"
        if not path.is_file():
            self.error.configure(text="Файла ошибок ещё нет — он появится после запуска.")
            return
        self._open_path(path)

    def _open_file(self):
        if not self.out_file or not self.out_file.is_file():
            return
        self._open_path(self.out_file)

    def _open_path(self, path):
        if os.name == "nt":
            os.startfile(path)  # noqa: S606 — открыть свой файл
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def _close(self):
        if self.status == "running":
            if not messagebox.askokcancel("Пульт", "Сбор ещё идёт. Закрыть и остановить?"):
                return
            self._stop()
        self.destroy()


if __name__ == "__main__":
    Desk().mainloop()
