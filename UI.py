import tkinter
from datetime import datetime
from tkinter import Label
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from importData import import_check_data
from config import current_row_index


class UI:
    def __init__(self):
        self.__init_properties()
        self.__init_root_window()
        self.__init_log_window()

    def __init_properties(self):
        self.select_file_path = None

    def __init_root_window(self):
        self.root = tkinter.Tk()
        self.root.title("自动导入")
        self.root.geometry("240x230")
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

        self.file_button = tkinter.Button(self.root, text="选择文件", command=self.__handle_file_button_click)
        self.file_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.import_button = tkinter.Button(self.root, text="导入", command=self.__handle_import_button_click)
        self.import_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    def __handle_file_button_click(self):
        self.select_file_path = filedialog.askopenfilename()
        self.log(f"所选文件：{self.select_file_path}")

    def __handle_import_button_click(self):
        if self.select_file_path is None or self.select_file_path == "":
            self.show_error("请先选择excel文件")
            return
        username = self.username_entry.get()
        password = self.password_entry.get()
        import_mode = self.mode_combo.current()
        while True:
            try:
                import_check_data(username, password, self.select_file_path, import_mode)
                break
            except Exception as e:
                self.log("导入过程异常，正在重新导入")
                self.log(f"从{current_row_index + 1}行继续导入")

    def __init_username_label(self):
        self.username_label = Label(self.root, text="用户名")
        self.username_label.pack()

    def __init_log_window(self):
        self.log_window = tkinter.Toplevel(self.root)
        self.log_window.title("日志")
        self.log_window.geometry("500x300")
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

    @staticmethod
    def show_error(message: str):
        messagebox.showerror("错误", message)

    @staticmethod
    def show_info(message: str):
        messagebox.showinfo(title="信息", message=message)
