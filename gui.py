import os
import sys
import threading
from typing import Any, cast
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import winreg
import pystray
from PIL import Image, ImageDraw, ImageTk

from core import NetDaemon, load_config, save_config, setup_logging

APP_NAME = "苏州大学网关自动登录工具"
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
ICON_PATH = os.path.join(os.path.dirname(__file__), "resources", "suda-logo.png")


def _get_executable_command():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    script = os.path.abspath(sys.argv[0])
    return f'"{sys.executable}" "{script}"'


def is_autostart_enabled():
    if not winreg:
        return False
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_READ
        ) as key:
            value, _ = winreg.QueryValueEx(key, APP_NAME)
        return bool(value)
    except FileNotFoundError:
        return False
    except Exception:
        return False


def set_autostart_enabled(enabled: bool):
    if not winreg:
        return
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE
    ) as key:
        if enabled:
            winreg.SetValueEx(
                key, APP_NAME, 0, winreg.REG_SZ, _get_executable_command()
            )
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass


def create_tray_icon(on_open, on_exit):
    if os.path.exists(ICON_PATH):
        image = Image.open(ICON_PATH).convert("RGBA")
        image = image.resize((64, 64))
    else:
        image = Image.new("RGB", (64, 64), color=(30, 136, 229))
        draw = ImageDraw.Draw(image)
        draw.ellipse((12, 12, 52, 52), fill=(255, 255, 255))
        draw.text((22, 20), "S", fill=(30, 136, 229))

    menu = pystray.Menu(
        pystray.MenuItem("打开", on_open, default=True),
        pystray.MenuItem("退出", on_exit),
    )
    icon = pystray.Icon(APP_NAME, image, APP_NAME, menu)
    try:
        cast(Any, icon).on_clicked = on_open
    except Exception:
        pass
    return icon


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("480x300")
        self.root.minsize(480, 300)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self._set_window_icon()
        self._init_style()

        self.daemon = None
        self.tray_icon = None

        self.status_var = tk.StringVar(value="未启动")
        self.autostart_var = tk.BooleanVar(value=is_autostart_enabled())
        self._log_lines = 0
        self._log_limit = 500

        self._build_ui()
        self._load_config()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(main)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        basic_tab = ttk.Frame(notebook, padding=12)
        advanced_tab = ttk.Frame(notebook, padding=12)
        logs_tab = ttk.Frame(notebook, padding=12)
        about_tab = ttk.Frame(notebook, padding=12)
        notebook.add(basic_tab, text="基础设置")
        notebook.add(advanced_tab, text="高级设置")
        notebook.add(logs_tab, text="日志")
        notebook.add(about_tab, text="关于")

        row = 0
        ttk.Label(basic_tab, text="网关地址").grid(row=row, column=0, sticky=tk.W)
        self.host_var = tk.StringVar()
        ttk.Entry(basic_tab, textvariable=self.host_var, width=50).grid(
            row=row, column=1, columnspan=2, sticky=tk.EW, padx=(8, 0)
        )

        row += 1
        ttk.Label(basic_tab, text="账号").grid(row=row, column=0, sticky=tk.W)
        self.account_var = tk.StringVar()
        ttk.Entry(basic_tab, textvariable=self.account_var, width=30).grid(
            row=row, column=1, sticky=tk.W, padx=(8, 0)
        )

        row += 1
        ttk.Label(basic_tab, text="密码").grid(row=row, column=0, sticky=tk.W)
        self.password_var = tk.StringVar()
        ttk.Entry(basic_tab, textvariable=self.password_var, width=30, show="*").grid(
            row=row, column=1, sticky=tk.W, padx=(8, 0)
        )

        row += 1
        ttk.Label(basic_tab, text="运营商").grid(row=row, column=0, sticky=tk.W)
        self.operator_var = tk.StringVar()
        self.operator_combo = ttk.Combobox(
            basic_tab,
            textvariable=self.operator_var,
            values=["校园网", "中国电信", "中国移动", "中国联通"],
            state="readonly",
            width=28,
        )
        self.operator_combo.grid(row=row, column=1, sticky=tk.W, padx=(8, 0))

        row += 1
        ttk.Label(basic_tab, text="检测间隔(秒)").grid(row=row, column=0, sticky=tk.W)
        self.freq_var = tk.StringVar()
        ttk.Entry(basic_tab, textvariable=self.freq_var, width=10).grid(
            row=row, column=1, sticky=tk.W, padx=(8, 0)
        )

        row += 1
        ttk.Label(advanced_tab, text="运营商 XPath(可选)").grid(
            row=row, column=0, sticky=tk.W
        )
        self.operator_xpath_var = tk.StringVar()
        ttk.Entry(advanced_tab, textvariable=self.operator_xpath_var, width=60).grid(
            row=row, column=1, columnspan=2, sticky=tk.EW, padx=(8, 0)
        )

        row += 1
        ttk.Label(advanced_tab, text="账号 XPath(可选)").grid(
            row=row, column=0, sticky=tk.W
        )
        self.account_xpath_var = tk.StringVar()
        ttk.Entry(advanced_tab, textvariable=self.account_xpath_var, width=60).grid(
            row=row, column=1, columnspan=2, sticky=tk.EW, padx=(8, 0)
        )

        row += 1
        ttk.Label(advanced_tab, text="密码 XPath(可选)").grid(
            row=row, column=0, sticky=tk.W
        )
        self.password_xpath_var = tk.StringVar()
        ttk.Entry(advanced_tab, textvariable=self.password_xpath_var, width=60).grid(
            row=row, column=1, columnspan=2, sticky=tk.EW, padx=(8, 0)
        )

        row += 1
        ttk.Label(advanced_tab, text="登录按钮 XPath(可选)").grid(
            row=row, column=0, sticky=tk.W
        )
        self.submit_xpath_var = tk.StringVar()
        ttk.Entry(advanced_tab, textvariable=self.submit_xpath_var, width=60).grid(
            row=row, column=1, columnspan=2, sticky=tk.EW, padx=(8, 0)
        )

        row += 1
        ttk.Checkbutton(
            basic_tab,
            text="开机启动",
            variable=self.autostart_var,
            command=self.toggle_autostart,
        ).grid(row=row, column=0, sticky=tk.W, pady=6)

        row += 1
        btn_frame = ttk.Frame(basic_tab)
        btn_frame.grid(row=row, column=0, columnspan=3, sticky=tk.W)

        ttk.Button(btn_frame, text="保存配置", command=self.save).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="启动", command=self.start).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(btn_frame, text="停止", command=self.stop).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="最小化", command=self.hide_to_tray).pack(
            side=tk.LEFT, padx=6
        )

        row += 1
        ttk.Label(basic_tab, text="状态").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(basic_tab, textvariable=self.status_var).grid(
            row=row, column=1, columnspan=2, sticky=tk.W
        )

        ttk.Label(logs_tab, text="日志").grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        self.log_text = ScrolledText(logs_tab, height=16, wrap=tk.WORD)
        self.log_text.grid(row=1, column=0, columnspan=3, sticky=tk.NSEW)
        self.log_text.configure(state=tk.DISABLED)

        ttk.Button(logs_tab, text="清空日志", command=self.clear_log).grid(
            row=2, column=0, sticky=tk.W, pady=8
        )

        basic_tab.columnconfigure(1, weight=1)
        advanced_tab.columnconfigure(1, weight=1)
        logs_tab.columnconfigure(0, weight=1)
        logs_tab.rowconfigure(1, weight=1)

        about_title = ttk.Label(about_tab, text="关于", style="Title.TLabel")
        about_title.pack(anchor=tk.W)
        about_text = (
            "本工具用于苏州大学网关自动登录与掉线重连。\n"
            "\n"
            "作者：Les1ie（原作者）\n"
            "维护：Allie\n"
            "\n"
        )
        ttk.Label(
            about_tab, text=about_text, style="Body.TLabel", justify=tk.LEFT
        ).pack(anchor=tk.W, pady=(8, 0))

    def _init_style(self):
        style = ttk.Style(self.root)
        for theme in ("vista", "xpnative", "clam", "default"):
            try:
                style.theme_use(theme)
                break
            except Exception:
                continue
        bg = "#FFFFFF"
        style.configure(".", background=bg)
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg)
        style.configure("TCheckbutton", background=bg)
        style.configure("TNotebook", background=bg)
        style.configure("TNotebook.Tab", padding=(10, 4))
        style.map("TNotebook.Tab", background=[("selected", bg)])
        style.configure("TEntry", fieldbackground=bg)
        style.configure("TCombobox", fieldbackground=bg)
        style.configure("Title.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Body.TLabel", font=("Segoe UI", 10))
        try:
            self.root.configure(bg=bg)
        except Exception:
            pass

    def _load_config(self):
        cfg = load_config("config.json")
        self.host_var.set(cfg["daemon"]["host"])
        self.freq_var.set(str(cfg["daemon"]["frequencies"]))

        login = cfg.get("login", {})
        self.account_var.set(login.get("account", ""))
        self.password_var.set(login.get("password", ""))
        self.operator_var.set(login.get("operator", ""))
        self.operator_xpath_var.set(login.get("operator_xpath", ""))
        self.account_xpath_var.set(login.get("account_xpath", ""))
        self.password_xpath_var.set(login.get("password_xpath", ""))
        self.submit_xpath_var.set(login.get("submit_xpath", ""))

    def _build_config(self):
        return {
            "login": {
                "account": self.account_var.get().strip(),
                "password": self.password_var.get().strip(),
                "operator": self.operator_var.get().strip(),
                "operator_xpath": self.operator_xpath_var.get().strip(),
                "account_xpath": self.account_xpath_var.get().strip(),
                "password_xpath": self.password_xpath_var.get().strip(),
                "submit_xpath": self.submit_xpath_var.get().strip(),
            },
            "daemon": {
                "host": self.host_var.get().strip(),
                "frequencies": int(self.freq_var.get().strip() or 10),
            },
        }

    def save(self):
        cfg = self._build_config()
        save_config(cfg, "config.json")
        self._set_status("配置已保存")

    def start(self):
        if self.daemon and self.daemon.is_alive():
            self._set_status("已在运行")
            return
        cfg = self._build_config()
        save_config(cfg, "config.json")
        self.daemon = NetDaemon(cfg, on_status=self.handle_status)
        self.daemon.start()
        self._set_status("启动中...")

    def stop(self):
        if self.daemon:
            self.daemon.stop()
            self.daemon = None
            self._set_status("已停止")

    def handle_status(self, text):
        self.root.after(0, self._set_status, text)
        self.root.after(0, self.append_log, text)

    def _set_window_icon(self):
        if not os.path.exists(ICON_PATH):
            return
        try:
            img = Image.open(ICON_PATH)
            self._icon_img = ImageTk.PhotoImage(img)
            self.root.iconphoto(False, self._icon_img)
        except Exception:
            pass

    def _set_status(self, text):
        self.status_var.set(text)

    def append_log(self, text):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self._log_lines += 1
        if self._log_lines > self._log_limit:
            self.log_text.delete("1.0", "2.0")
            self._log_lines -= 1
        self.log_text.configure(state=tk.DISABLED)

    def clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self._log_lines = 0

    def toggle_autostart(self):
        try:
            set_autostart_enabled(self.autostart_var.get())
        except Exception as e:
            messagebox.showerror("错误", f"设置开机启动失败: {e}")
            self.autostart_var.set(is_autostart_enabled())

    def hide_to_tray(self):
        if self.tray_icon:
            return
        self.root.withdraw()

        def on_open(icon, _item=None):
            icon.stop()
            self.tray_icon = None
            self.root.after(0, self.root.deiconify)

        def on_exit(icon, _item=None):
            icon.stop()
            self.tray_icon = None
            self.root.after(0, self.root.destroy)

        self.tray_icon = create_tray_icon(on_open, on_exit)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()


if __name__ == "__main__":
    setup_logging()
    root = tk.Tk()
    app = App(root)
    root.mainloop()
