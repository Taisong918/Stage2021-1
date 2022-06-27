# -*- coding: utf-8 -*-

from tkinter import *
import time
from tkinter import filedialog
import Separate

LOG_LINE_NUM = 0
ENTRY_WIDTH = 45

class MY_GUI:

    def __init__(self, init_window_name):
        self.output_path_entry = None
        self.output_path_button = None
        self.init_output_path_label = None
        self.folder_path_entry = None
        self.file_path_entry = None
        self.merger_button = None
        self.create_folder_button = None
        self.separate_button = None
        self.init_create_folder_label = None
        self.folder_path_button = None
        self.file_path_button = None
        self.log_data_Text = None
        self.log_label = None
        self.init_merger_label = None
        self.init_separate_label = None
        self.init_folder_path_label = None
        self.init_file_path_label = None

        self.file_path = StringVar()
        self.folder_path = StringVar()
        self.output_path = StringVar()
        self.init_window_name = init_window_name

    #设置窗口
    def set_init_window(self):

        self.init_window_name.title("LOGEFI SERVICES")           #窗口名
        self.init_window_name.geometry('400x600+10+10')
        #标签
        self.init_file_path_label = Label(self.init_window_name, text="选择Excel文件位置")
        self.init_file_path_label.grid(row=2, column=0)
        self.init_folder_path_label = Label(self.init_window_name, text="选择多个文件所在文件夹位置")
        self.init_folder_path_label.grid(row=4, column=0)
        self.init_output_path_label = Label(self.init_window_name, text="选择结果输出位置")
        self.init_output_path_label.grid(row=6, column=0)

        self.init_separate_label = Label(self.init_window_name, text="功能一: 分割总表")
        self.init_separate_label.grid(row=10, column=0)
        self.init_create_folder_label = Label(self.init_window_name, text="功能二: 分类M0PDF,自动生成Excel汇总文件")
        self.init_create_folder_label.grid(row=11, column=0)
        self.init_merger_label = Label(self.init_window_name, text="功能三: 合并JUJUBE扫描后的文件")
        self.init_merger_label.grid(row=12, column=0)

        self.log_label = Label(self.init_window_name, text="日志")
        self.log_label.grid(row=22, column=0)
        #文本框
        self.log_data_Text = Text(self.init_window_name, width=39, height=9)  # 日志框
        self.log_data_Text.grid(row=23, column=0, columnspan=10)
        #按钮
        self.file_path_button = Button(self.init_window_name, text="OpenFile", bg="lightblue", width=10, command=self.open_file)
        self.file_path_button.grid(row=2, column=11)
        self.folder_path_button = Button(self.init_window_name, text="OpenFolder", bg="lightblue", width=10, command=self.open_folder)
        self.folder_path_button.grid(row=4, column=11)
        self.output_path_button = Button(self.init_window_name, text="OutputFolder", bg="lightblue", width= 10, command=self.set_output_folder)
        self.output_path_button.grid(row=6, column=11)

        self.separate_button = Button(self.init_window_name, text="Separate", bg="lightblue", width=10, command=self.separate)
        self.separate_button.grid(row=10, column=11)
        self.create_folder_button = Button(self.init_window_name, text="CreateFloder", bg="lightblue", width=10)
        self.create_folder_button.grid(row=11, column=11)
        self.merger_button = Button(self.init_window_name, text="Merger", bg="lightblue", width=10)
        self.merger_button.grid(row=12, column=11)

        #Entry
        self.file_path_entry = Entry(self.init_window_name, width=ENTRY_WIDTH)
        self.file_path_entry.grid(row=3)
        self.folder_path_entry = Entry(self.init_window_name, width=ENTRY_WIDTH)
        self.folder_path_entry.grid(row=5)
        self.output_path_entry = Entry(self.init_window_name, width=ENTRY_WIDTH)
        self.output_path_entry.grid(row=7)

    #获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time

    #日志动态打印
    def write_log_to_Text(self, logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +" " + str(logmsg) + "\n"      #换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0,2.0)
            self.log_data_Text.insert(END, logmsg_in)

    # 选择文件
    def open_file(self):
        self.file_path = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xlsx'), ('All Files', '*')])
        self.file_path_entry.delete(0, "end")
        self.file_path_entry.insert(0, self.file_path)

    def open_folder(self):
        self.folder_path = filedialog.askdirectory(title='选择文件夹')
        self.folder_path_entry.delete(0, "end")
        self.folder_path_entry.insert(0, self.folder_path)

    def set_output_folder(self):
        self.output_path = filedialog.askdirectory(title='选择文件夹')
        self.output_path_entry.delete(0, "end")
        self.output_path_entry.insert(0, self.output_path)

    #实现功能函数
    def separate(self):
        Separate.Separate(self.output_path, self.file_path).classification()



def gui_start():
    init_window = Tk()              #实例化出一个父窗口
    GUI = MY_GUI(init_window)
    # 设置根窗口默认属性
    GUI.set_init_window()
    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示

gui_start()