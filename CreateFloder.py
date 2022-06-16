import os
import time
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet


class CreateFolder:
    dict_agence: dict
    date: str
    path: str
    wb: Workbook
    my_sheet: Worksheet
    cell: tuple

    def __init__(self, path):
        #初始化日期
        self.date = time.strftime("%Y.%m.%d", time.localtime())
        #读取文件
        self.path = path
        file_list = os.listdir(self.path)
        file_name: str = None
        for file in file_list:
            if ".xlsx" in file:
                file_name = file
                break
        self.wb = load_workbook(file_name, read_only=True)
        self.my_sheet = self.wb.active
        self.cell = tuple(self.my_sheet)
        #统计AGENCE数量
        self.dict_agence = {}
        for i in range(1, self.my_sheet.max_row):
            key = str(self.cell[i][1].value)
            value = self.dict_agence.get(key)
            if value is None:
                self.dict_agence[key] = 1
            else:
                self.dict_agence[key] = value + 1


    def mkdir(self, path):
        folder = os.path.exists(path)
        if not folder:
            os.makedirs(path)

    def create_folder(self):
        #1
        folder_name = self.date + " Paris 02 " + str(self.my_sheet.max_row - 1) + "家" #去除第一行
        self.mkdir(folder_name)
        self.path = self.path + os.sep + folder_name
        #2
        for k in self.dict_agence.keys():
            v = self.dict_agence.get(k)
            path_inner = self.path + os.sep + str(self.date) + " " + str(k) + " " + str(v) + "家"
            self.mkdir(path_inner)
            #3
            self.create_folder_sup(k, v, path_inner)

    def create_folder_sup(self, agence_name, agence_nbr, path_inner):
        cpt = 1
        while 0 < agence_nbr and cpt < self.my_sheet.max_row:
            if self.cell[cpt][1].value == agence_name:
                folder_name = str(self.cell[cpt][0].value) + " " + self.cell[cpt][3].value
                self.mkdir(path_inner + os.sep + folder_name)
                agence_nbr -= 1
            cpt += 1


file_name = os.getcwd()
cf = CreateFolder(file_name)
cf.create_folder()