input("在运行此程序前, 请安装库: xlrd, 安装好后重启程序, 如果安装了, 单击回车继续")

import xlrd
import os
import shutil
import difflib
import threading

input("请将表格文件放在此程序下的sheet文件夹内, 放好后请单击回车")
input("请将所有需要检查的文件夹放到in文件夹内, 放好后单击回车")
input("使用方法：此程序会查找文件夹，并将in中的文件夹移动到out中, 单击回车以继续")

dir_path = "in"
dir_serial_number = []
table_serial_number = []
dir_to_move = []
thread_pool = []
dirs = None
data = None
e_table = None

for root, dir, files in os.walk(dir_path):
    dirs = dir
    break

for root, dir, files in os.walk("sheet"):
    file_name = None
    for i in files:
        if i.find("xls") >= 0:
            print(f"找到了文件: {i}")
            file_name = os.path.join("sheet", i)
            data = xlrd.open_workbook(file_name)
            e_table = data.sheets()
            break
    if file_name is None:
        print("没有找到表格文件")
        raise SystemExit


def traverseTable(table):
    buff = []
    for _ in table:
        for __ in range(_.nrows):
            for ___ in _.row_values(__):
                buff.append(___)
    return buff


def print_dir(dirs_):
    print("目录下的文件夹: ")
    count = 1
    for value in dirs_:
        print(f" -{count}-   {value}")
        count += 1
    print("")


def get_dir_snumber(dirs_):
    print("摘取文件夹序列号: ")
    for value3 in dirs_:
        if value3.find("”") >= 0:
            if len((value3.split("”")[1]).split(" ")[0]) != 0:
                dir_serial_number.append((value3.split("”")[1]).split(" ")[0])
        if value3.find("\"") >= 0:
            if len((value3.split('"')[1]).split(" ")[0]) != 0:
                dir_serial_number.append((value3.split('"')[1]).split(" ")[0])
    count = 1
    for value3 in dir_serial_number:
        print(f" -{count}-   {value3}")
        count += 1
    print("")


def _space_(_T):
    return " "


def calculateSimilarity(str1, str2, func=None):
    return difflib.SequenceMatcher(func, str1, str2).quick_ratio()


def calculate(i: str):
    global thread_pool

    def __calculate__(i: str):
        buff = ""
        for j in i.split(" "):
            j: str
            j = j.replace(" ", "")
            for value in dirs:
                if value.find("”") >= 0:
                    buff = (value.split("”")[1]).split(" ")[0]
                if value.find("\"") >= 0:
                    buff = (value.split('"')[1]).split(" ")[0]
                if len(buff) >= 1:
                    if j == buff:
                        print(f"  -   字符串1: {j} 字符串2: {buff} 相似度: {calculateSimilarity(j, buff, _space_)} 文件夹名称: {value}")
                        dir_to_move.append(value)

    thread = threading.Thread(target=__calculate__, args=[i], daemon=True)
    thread_pool.append(thread)
    thread.start()
    return thread


def calculateString(table_data: list):
    print("正在处理:")
    for i_ in table_data:
        calculate(i_)
    for thread in thread_pool:
        thread.join()
    print("")


def move_dirs():
    print("正在移动:")
    count = 1
    for value in dir_to_move:
        try:
            shutil.move(os.path.join(dir_path, value), "out")
        except:
            pass
        else:
            print(f"    -{count}-    文件夹: {value} 移动成功")
            count += 1
    print("")


print_dir(dirs)
get_dir_snumber(dirs)
calculateString(traverseTable(e_table))
move_dirs()