# 苏州大学网关自动登录工具

本项目继承自 [Les1ie/SUDA-Net-Daemon](https://github.com/Les1ie/SUDA-Net-Daemon)，用于在 Windows 环境下自动检测并保持苏州大学网关登录状态。

妈妈再也不用担心工位电脑意外断网了！

## 功能
- 自动检测登录状态
- 自动登录与掉线重连
- GUI 管理界面（Tkinter）
- 托盘运行与开机启动
- 日志面板与配置保存

## 运行环境
- Windows 10/11
- Python 3.10+（建议使用 3.11/3.12）
- Google Chrome

## 安装依赖

```sh
pip install -r requirements.txt
```

## 配置
编辑 config.json：

```json
{
  "login": {
    "account": "学号",
    "password": "密码",
    "operator": "校园网",
    "operator_xpath": "",
    "account_xpath": "",
    "password_xpath": "",
    "submit_xpath": ""
  },
  "daemon": {
    "host": "http://10.9.1.3/",
    "frequencies": 10
  }
}
```

说明：
- `operator` 为运营商下拉框文本（校园网/中国电信/中国移动/中国联通）。
- XPath 字段为可选，只有在页面结构变化时才需要填写。

## 运行 GUI

```sh
python gui.py
```

## 打包为 EXE

```sh
.\shells\pack_gui.ps1
```

产物在 dist/ 下。

## 许可证
继承原项目许可证（MIT）。

