import os
import sys
import threading
from typing import Any, cast
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import pystray
from PIL import Image, ImageDraw, ImageTk

try:
    import win32event  # type: ignore
    import win32api  # type: ignore
    import winerror  # type: ignore

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

from core import NetDaemon, load_config, save_config, setup_logging, validate_config

APP_NAME = "苏州大学网关自动登录工具"
ICON_PATH = os.path.join(os.path.dirname(__file__), "resources", "suda-logo.png")
MUTEX_UI = "Local\\SUDA_Net_Daemon_UI_Mutex"


def ensure_single_instance(mutex_name: str, show_message: bool = True):
    """确保程序只运行一个实例"""
    if not HAS_WIN32:
        return None

    try:
        # CreateMutex的第一个参数应该省略，让其使用默认安全描述符
        mutex = win32event.CreateMutex(None, 1, mutex_name)  # type: ignore
        last_error = win32api.GetLastError()
        if last_error == winerror.ERROR_ALREADY_EXISTS:
            if show_message:
                messagebox.showerror(
                    "程序已运行", f"{APP_NAME}已经在运行中！\n\n请检查系统托盘图标。"
                )
            return None
        return mutex
    except Exception as e:
        print(f"创建互斥锁失败: {e}")
        return None


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
        # 注册退出清理
        import atexit

        atexit.register(self._safe_cleanup)
        self._set_window_icon()
        self._init_style()

        self.daemon = None
        self.tray_icon = None
        self._tray_thread = None
        self.config = {}  # 初始化配置字典

        self.status_var = tk.StringVar(value="未启动")
        self._log_lines = 0
        self._log_limit = 500

        self._build_ui()
        self._load_config()

        # 立即创建托盘图标（常驻）
        self.root.after(100, self._create_persistent_tray)

        # 自动启动网络连接（延迟2秒，确保界面加载完成）
        self.root.after(2000, self._auto_start_network)

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
            "原始作者：Les1ie\n"
            "维护作者：Allie\n"
            "\n"
            "项目地址：https://github.com/iAllie2002/SUDA-NetTool\n"
            "\n"
            "如果你觉得有所帮助，欢迎在项目页面点个Star支持一下！\n"
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
        self.config = load_config("config.json")
        self.host_var.set(self.config["daemon"]["host"])
        self.freq_var.set(str(self.config["daemon"]["frequencies"]))

        login = self.config.get("login", {})
        self.account_var.set(login.get("account", ""))
        self.password_var.set(login.get("password", ""))
        self.operator_var.set(login.get("operator", ""))
        self.operator_xpath_var.set(login.get("operator_xpath", ""))
        self.account_xpath_var.set(login.get("account_xpath", ""))
        self.password_xpath_var.set(login.get("password_xpath", ""))
        self.submit_xpath_var.set(login.get("submit_xpath", ""))

    def _create_persistent_tray(self):
        """创建常驻托盘图标"""
        if self.tray_icon:
            return

        def on_open(icon, _item=None):
            # 不停止托盘图标，只显示窗口
            self.root.after(0, self.show_window)

        def on_exit(icon, _item=None):
            # 退出时停止托盘图标
            icon.stop()
            self.root.after(0, self._cleanup_and_exit)

        self.tray_icon = create_tray_icon(on_open, on_exit)
        self._tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self._tray_thread.start()

    def _safe_cleanup(self):
        """安全清理，用于程序异常退出时的保护"""
        try:
            if self.daemon and self.daemon.is_alive():
                self.daemon.stop()
        except Exception:
            pass
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass

    def _auto_start_network(self):
        """自动启动网络连接"""
        try:
            # 检查是否已经在运行
            if self.daemon and self.daemon.is_alive():
                self.append_log("守护进程已在运行")
                return

            # 检查账号是否配置
            if not self.account_var.get().strip():
                self.append_log("提示：请先配置账号信息后再启动")
                return

            # 自动启动
            self.append_log("正在自动连接网络...")
            self.start()

            # 如果启动成功，最小化到托盘
            if self.daemon and self.daemon.is_alive():
                self.root.after(1000, self.hide_to_tray)
        except Exception as e:
            self.append_log(f"自动启动失败: {e}")

    def _build_config(self):
        # 验证并获取检测频率，确保在5-3600秒范围内
        freq = 10  # 默认值
        try:
            freq_str = self.freq_var.get().strip()
            if freq_str:
                freq = max(5, min(3600, int(freq_str)))  # 限制在5-3600范围内
        except (ValueError, TypeError):
            pass  # 使用默认值10

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
                "frequencies": freq,
            },
        }

    def save(self):
        try:
            cfg = self._build_config()
            # 先验证配置
            valid, error_msg = validate_config(cfg)
            if not valid:
                messagebox.showwarning(
                    "配置警告", f"配置可能有问题：\n{error_msg}\n\n仍将保存配置。"
                )

            save_config(cfg, "config.json")
            # 更新内存中的配置，保持一致性
            self.config = cfg
            self._set_status("配置已保存")
            messagebox.showinfo("成功", "配置已成功保存到 config.json")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")
            self._set_status("保存失败")

    def start(self):
        # 检查是否已经在运行
        if self.daemon and self.daemon.is_alive():
            self._set_status("已在运行")
            messagebox.showinfo("提示", "守护进程已经在运行中")
            return

        cfg = self._build_config()

        # 验证配置
        valid, error_msg = validate_config(cfg)
        if not valid:
            messagebox.showerror("配置错误", f"配置验证失败：\n{error_msg}")
            self._set_status("配置错误")
            return

        try:
            self.daemon = NetDaemon(cfg, on_status=self.handle_status)
            self.daemon.start()
            self._set_status("启动中...")
        except Exception as e:
            messagebox.showerror("启动失败", f"无法启动守护进程：{e}")
            self._set_status("启动失败")
            self.daemon = None

    def stop(self):
        if not self.daemon:
            self._set_status("未运行")
            messagebox.showinfo("提示", "守护进程未运行")
            return

        try:
            if self.daemon.is_alive():
                self._set_status("正在停止...")
                self.daemon.stop()
                self.daemon.join(timeout=2)
            self.daemon = None
            self._set_status("已停止")
        except Exception as e:
            messagebox.showerror("停止失败", f"停止守护进程时出错：{e}")
            self.daemon = None
            self._set_status("停止异常")

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

    def _cleanup_and_exit(self):
        """清理资源并退出程序"""
        try:
            if self.daemon and self.daemon.is_alive():
                self.daemon.stop()
                # 等待线程结束，最多等待3秒
                self.daemon.join(timeout=3)
        except Exception as e:
            print(f"清理资源时出错: {e}")
        finally:
            try:
                if self.tray_icon:
                    self.tray_icon.stop()
            except Exception:
                pass
            self.root.destroy()

    def hide_to_tray(self):
        """隐藏窗口到托盘（托盘图标已常驻）"""
        self.root.withdraw()

    def show_window(self):
        """显示窗口"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()


if __name__ == "__main__":
    setup_logging()

    # 确保只运行一个实例
    mutex = ensure_single_instance(MUTEX_UI)
    if mutex is None and HAS_WIN32:
        # 单实例检查失败，退出程序
        sys.exit(1)

    root = tk.Tk()
    app = App(root)
    root.mainloop()
