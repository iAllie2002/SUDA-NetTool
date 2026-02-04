import json
import logging
import os
import sys
import threading
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    SessionNotCreatedException,
    NoSuchElementException,
)
from selenium.webdriver.support.ui import Select


logger = logging.getLogger("SUDA-Net-Daemon")


def setup_logging():
    if logger.handlers:
        return
    logger.setLevel("INFO")
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        "[%(asctime)s %(levelname)s] %(message)s", "%d/%m/%Y %H:%M:%S"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    # 使用程序所在目录的绝对路径，避免权限问题
    try:
        # 如果是打包后的exe，使用exe所在目录
        if getattr(sys, "frozen", False):
            log_dir = os.path.dirname(sys.executable)
        else:
            log_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(log_dir, "daemon.log")
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        logger.warning(f"无法创建日志文件: {e}")


DEFAULT_CONFIG = {
    "login": {
        "account": "",
        "password": "",
        "operator": "",
        "operator_xpath": "",
        "account_xpath": "",
        "password_xpath": "",
        "submit_xpath": "",
    },
    "daemon": {
        "host": "http://10.9.1.3/",
        "frequencies": 10,
    },
}


def load_config(path="config.json"):
    # 使用程序所在目录的绝对路径
    if not os.path.isabs(path):
        # 如果是打包后的exe，使用exe所在目录
        if getattr(sys, "frozen", False):
            config_dir = os.path.dirname(sys.executable)
        else:
            config_dir = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(config_dir, path)
    if not os.path.exists(path):
        return DEFAULT_CONFIG.copy()
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    cfg = DEFAULT_CONFIG.copy()
    cfg["login"].update(data.get("login", {}))
    cfg["daemon"].update(data.get("daemon", {}))
    if "operator_index" in cfg["login"]:
        cfg["login"].pop("operator_index", None)
    return cfg


def save_config(cfg, path="config.json"):
    # 使用程序所在目录的绝对路径
    if not os.path.isabs(path):
        # 如果是打包后的exe，使用exe所在目录
        if getattr(sys, "frozen", False):
            config_dir = os.path.dirname(sys.executable)
        else:
            config_dir = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(config_dir, path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def validate_config(cfg):
    """验证配置完整性，返回(是否有效, 错误信息)"""
    errors = []

    login = cfg.get("login", {})
    if not login.get("account", "").strip():
        errors.append("账号不能为空")
    # 密码可以为空，某些网络不需要密码

    freq = cfg.get("daemon", {}).get(
        "frequencies", DEFAULT_CONFIG["daemon"]["frequencies"]
    )
    try:
        freq = int(freq)
        if freq < 5 or freq > 3600:
            errors.append("检测间隔必须在 5-3600 秒之间")
    except (ValueError, TypeError):
        errors.append("检测间隔必须是有效的数字")

    return len(errors) == 0, "；".join(errors)


def _find_first_by_xpath(chrome, xpaths):
    last_exc = None
    for xp in xpaths:
        if not xp:
            continue
        try:
            return chrome.find_element(By.XPATH, xp)
        except Exception as e:
            last_exc = e
            continue
    if last_exc:
        raise last_exc
    raise NoSuchElementException("No valid xpath provided")


def check(chrome, host):
    chrome.get(host)

    successed = False
    success_info_xpath = '//*[@id="edit_body"]/div/div[1]/form/div[1]'

    message_xpath = '//*[@id="message"]'
    succecc_msg = "您已经成功登录。"
    try:
        successed = (
            succecc_msg
            == chrome.find_element(By.XPATH, success_info_xpath).text.strip()
        )
    except NoSuchElementException:
        pass

    if successed:
        message = succecc_msg
    else:
        try:
            msg = chrome.find_element(By.XPATH, message_xpath).text
            message = msg
        except NoSuchElementException:
            message = "未登录，尝试登录。"
        except Exception:
            message = "页面状态解析失败。"
            logger.error(message, exc_info=True)
    return successed, message


def login(chrome, login_cfg=None):
    login_cfg = login_cfg or {}
    u = login_cfg.get("account", "")
    p = login_cfg.get("password", "")

    operator = login_cfg.get("operator")
    operator_xpath = login_cfg.get("operator_xpath")

    account_xpath = login_cfg.get("account_xpath")
    password_xpath = login_cfg.get("password_xpath")
    submit_xpath = login_cfg.get("submit_xpath")

    default_operator_xpaths = [
        '//*[@id="edit_body"]/div[2]/div[12]/select',
        '//*[@id="edit_body"]//select',
        "//select",
    ]

    default_account_xpaths = [
        '//*[@id="edit_body"]/div[2]/div[12]/form/input[3]',
        '//*[@id="edit_body"]//input[@type="text" or @name="username" or @id="username"]',
        '//input[@type="text" or @name="username" or @id="username"]',
    ]
    default_password_xpaths = [
        '//*[@id="edit_body"]/div[2]/div[12]/form/input[4]',
        '//*[@id="edit_body"]//input[@type="password" or @name="password" or @id="password"]',
        '//input[@type="password" or @name="password" or @id="password"]',
    ]
    default_submit_xpaths = [
        '//*[@id="edit_body"]/div[2]/div[12]/form/input[2]',
        '//*[@id="edit_body"]//input[@type="submit" or @value="登录" or @value="Login"]',
        '//input[@type="submit" or @value="登录" or @value="Login"]',
        '//button[contains(.,"登录") or contains(.,"Login")]',
    ]

    try:
        if operator or operator_xpath:
            op_xpaths = [operator_xpath] if operator_xpath else default_operator_xpaths
            dropdown = _find_first_by_xpath(chrome, op_xpaths)
            select = Select(dropdown)
            if operator:
                select.select_by_visible_text(operator)
            time.sleep(0.5)
    except Exception:
        logger.error(
            "选择运营商失败，请检查配置中的 operator 或 operator_xpath。", exc_info=True
        )
        return False

    try:
        account_input = _find_first_by_xpath(
            chrome, [account_xpath] if account_xpath else default_account_xpaths
        )
        password_input = _find_first_by_xpath(
            chrome, [password_xpath] if password_xpath else default_password_xpaths
        )
        login_bt = _find_first_by_xpath(
            chrome, [submit_xpath] if submit_xpath else default_submit_xpaths
        )

        account_input.click()
        time.sleep(0.5)
        account_input.clear()
        account_input.send_keys(u)

        password_input.click()
        time.sleep(0.5)
        password_input.clear()
        password_input.send_keys(p)

        chrome.execute_script("arguments[0].click()", login_bt)
        return True
    except Exception:
        logger.error("登录元素定位失败，请检查配置中的 XPath。", exc_info=True)
        return False


def init_chrome(host):
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("disable-cache")
    chrome_options.add_argument("log-level=3")
    chrome_options.add_argument("--headless=new")

    try:
        chrome = webdriver.Chrome(options=chrome_options)
        chrome.get(host)
        return chrome
    except SessionNotCreatedException:
        if os.path.exists("chromedriver.exe"):
            try:
                chrome = webdriver.Chrome(
                    service=Service("chromedriver.exe"), options=chrome_options
                )
                chrome.get(host)
                return chrome
            except SessionNotCreatedException:
                logger.error(
                    "ChromeDriver 与 Chrome 版本不一致，请更新 chromedriver.exe 或删除它以使用 Selenium Manager 自动匹配。",
                    exc_info=True,
                )
                return None
            except Exception:
                logger.error("ChromeDriver 初始化错误", exc_info=True)
                return None
        logger.error(
            "ChromeDriver 与 Chrome 版本不一致，请更新 chromedriver.exe 或删除它以使用 Selenium Manager 自动匹配。",
            exc_info=True,
        )
        return None
    except Exception:
        logger.error("ChromeDriver 初始化错误", exc_info=True)
        return None


class NetDaemon(threading.Thread):
    def __init__(self, config, on_status=None):
        super().__init__(daemon=True)
        self.config = config
        self.on_status = on_status
        self._stop_event = threading.Event()
        self.chrome = None

    def stop(self):
        """停止守护进程并清理资源"""
        logger.info("正在停止网络守护进程...")
        self._stop_event.set()
        try:
            if self.chrome:
                self.chrome.quit()
                self.chrome = None
        except Exception as e:
            logger.warning(f"关闭浏览器时出错: {e}")

    def _emit(self, text):
        if self.on_status:
            self.on_status(text)

    def run(self):
        host = self.config["daemon"]["host"]
        delay = int(self.config["daemon"]["frequencies"])
        login_cfg = self.config.get("login", {})

        self._emit("正在初始化浏览器...")
        self.chrome = init_chrome(host)
        if not self.chrome:
            self._emit("浏览器初始化失败")
            return
        self._emit("初始化完成，开始后台监控网络连接...")

        while not self._stop_event.is_set():
            try:
                s, m = check(self.chrome, host)
                dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                if not s:
                    self._emit(f"[{dt}] 状态：{m} 尝试登录...")
                    ok = login(self.chrome, login_cfg)
                    if not ok:
                        logger.error("登录流程失败，稍后重试。")
                    time.sleep(3)
                    s, m = check(self.chrome, host)
                    dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                if s:
                    msg = f"已成功登录。[{dt}]"
                else:
                    msg = f"尝试登录后仍未登录。[{dt}]"
                self._emit(msg)

                for _ in range(delay):
                    if self._stop_event.is_set():
                        break
                    time.sleep(1)
            except Exception as e:
                logger.error(f"主循环发生严重错误: {e}", exc_info=True)
                self._emit("主循环错误，30秒后重试...")
                for _ in range(30):
                    if self._stop_event.is_set():
                        break
                    time.sleep(1)
