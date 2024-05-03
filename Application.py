import threading
import time
from tkinter import messagebox

import requests
from webdriver_manager.chrome import ChromeDriverManager

import config
from UI import UI


class Application:
    def __init__(self):
        self.ui = UI()
        self.ui.root.after(100, lambda: self.run())

    def run(self):
        self.log("初始化程序配置")
        self.log("正在检查网络状态")
        if not self.__check_network():
            self.show_error("无法访问网站，请检查网络设置")
            return
        self.log("网络正常")
        threading.Thread(target=self.__update_webdriver, daemon=True).start()

    def log(self, message):
        self.ui.root.after(5, lambda: self.ui.log(message))

    def __update_webdriver(self):
        self.log("开始下载安装更新webdriver")
        while True:
            try:
                ChromeDriverManager().install()
                break
            except Exception as e:
                print(e)
                self.log("下载安装webdriver失败，正在重试")
            time.sleep(10)
        self.log("webdriver更新成功")
        self.show_info("程序准备就绪")

    def start(self):
        self.ui.start()

    @staticmethod
    def __check_network() -> bool:
        try:
            response: requests.Response = requests.get(url=config.target_url, timeout=5)
            if response.status_code == 200:
                return True
            else:
                return False
        except Exception as e:
            return False

    @staticmethod
    def show_error(message: str):
        messagebox.showerror("错误", message)

    @staticmethod
    def show_info(message: str):
        messagebox.showinfo(title="信息", message=message)
