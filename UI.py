import tkinter
from datetime import datetime
from tkinter import Label
from tkinter import ttk
from importData import import_check_data


class UI:
    def __init__(self):
        self.__init_root_window()
        self.__init_log_window()

    def __init_root_window(self):
        self.root = tkinter.Tk()
        self.root.title("自动导入")
        self.root.geometry("240x210")
        # 创建用户名标签和输入框
        self.username_label = tkinter.Label(self.root, text="用户名", width=5)
        self.username_label.grid(row=0, column=0, padx=10, pady=10)
        self.username_entry = tkinter.Entry(self.root, width=20)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)

        # 创建密码标签和输入框
        self.password_label = tkinter.Label(self.root, text="密码", width=5)
        self.password_label.grid(row=1, column=0, padx=10, pady=10)
        self.password_entry = tkinter.Entry(self.root, show="*", width=20)  # show="*" 使密码以星号显示
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        # 创建下拉框
        self.mode_label = tkinter.Label(self.root, text="模式")
        self.mode_label.grid(row=2, column=0, padx=10, pady=10)
        self.mode_var = tkinter.StringVar()
        self.mode_combo = ttk.Combobox(self.root, textvariable=self.mode_var, state="readonly")
        self.mode_combo['values'] = ("新建模式", "覆盖模式")
        self.mode_combo.current(0)  # 设置默认选择第一个选项
        self.mode_combo.grid(row=2, column=1, padx=10, pady=10)

        # 创建登录按钮
        self.import_button = tkinter.Button(self.root, text="导入", command=self.__handle_import_button_click)
        self.import_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    def __handle_import_button_click(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        import_mode = self.mode_combo.current()
        import_check_data(username, password, import_mode)

    def __init_username_label(self):
        self.username_label = Label(self.root, text="用户名")
        self.username_label.pack()

    def __init_log_window(self):
        self.log_window = tkinter.Toplevel(self.root)
        self.log_window.title("日志")
        self.log_window.geometry("400x300")
        self.__init_log_text()

    def __init_log_text(self):
        self.log_text = tkinter.Text(self.log_window, width=200, height=100)
        self.log_text.config(state=tkinter.DISABLED)
        self.log_text.pack()

    def start(self):
        self.root.mainloop()

    def log(self, message: str):
        self.log_text.config(state=tkinter.NORMAL)
        self.log_text.insert(tkinter.END, self.log_time_stamp() + message + "\n")
        self.log_text.config(state=tkinter.DISABLED)

    @staticmethod
    def log_time_stamp() -> str:
        now = datetime.now()
        return "[" + now.strftime("%Y-%m-%d %H:%M:%S") + "]"
