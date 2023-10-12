import string
import os
import openpyxl
import configparser
from openpyxl.workbook import Workbook

config = configparser.ConfigParser()
config.read("Directory.ini")


# 遍历多个excel文件
def ReadDir(directory: string):
    filename = []
    filenames = []
    print(directory)
    for parents, dirnames, filenames in os.walk(directory):
        print(directory, ' & ', filenames, '&', len(filenames))
    for fileindex in range(len(filenames)):
        if ".xlsx" in filenames[fileindex]:
            filename.append(directory + filenames[fileindex])
    return filename


# 遍历读取表格的每一个 sheet 表
def ReadTable(files):
    ListTemp = []
    for fileindex in files:
        book = openpyxl.load_workbook(fileindex)
        sheetNamelist = book.sheetnames
        # 遍历~一张 sheet 表
        for name in sheetNamelist:
            sheet = book[name]
            print(sheet)
            # 遍历读取每一个 sheet 表中的 每一列
            for col in sheet.iter_cols():
                # print(col)
                # 若整列中含不为空的数据项，则保留数据
                if all(cell.value is None for cell in col):
                    continue
                else:
                    for cell in col:
                        if cell.value is not None:
                            ListTemp.append(cell.value)
    return ListTemp


# 填入临时表中
def WriteTable(listtemp: list):
    WorkTable = Workbook()
    sheet = WorkTable.active
    sheet.title = 'Sheet1'
    for raw in range(len(listtemp)):
        sheet.cell(raw + 1, 1, listtemp[raw])
    WorkTable.save(config.get("Directory","OutPath"))


if __name__ == "__main__":
    Directory = config.get("Directory", "ReadPath")
    Files = ReadDir(Directory)
    WriteTable(ReadTable(Files))