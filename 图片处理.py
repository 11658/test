import pandas as pd
from openpyxl import load_workbook
import pandas as pd

# 读取Excel文件
file_path = '迎宾信息.xls'
df = pd.read_excel(file_path)

# 按行输出文件内容
# for index, row in df.iterrows():
#     print(row)

# 或者按照特定的列输出
list_name = []
for index, row in df.iterrows():
    list_name.append([row[8], row[2], row[3], row[4]])

# 如果你有具体的列名，可以按需调整


import xlrd
from xlwt import Workbook

# 读取 Excel 文件
file_path = '迎宾信息.xls'
workbook = xlrd.open_workbook(file_path)

# 获取要修改的工作表
sheet = workbook.sheet_by_index(1)  # 第二个工作表，索引从0开始

# 创建一个新的 Excel 文件来写入修改后的内容
new_workbook = Workbook()
new_sheet = new_workbook.add_sheet('Sheet1')

#  list_name
i = 0
j = 0
for rpw, lists in enumerate(list_name):

    if j == 8:
        j = 0
        i += 4
    if str(lists[0]) == 'nan':
        aa = ' '
    else:
        aa = lists[0]

        aa = aa

    new_sheet.write(0 + i, j,aa)
    new_sheet.write(1 + i, j, lists[1])
    new_sheet.write(2 + i, j, lists[2])
    new_sheet.write(3 + i, j, '证件：' + str(lists[3]))
    j += 1
# 保存修改后的 Excel 文件
new_file_path = '修改后的迎宾信息.xls'
new_workbook.save(new_file_path)

print("修改已保存到:", new_file_path)
