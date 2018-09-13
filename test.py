import os
import sys
import re
import xlrd
from xlutils.copy import copy


#命令行操作

def cmd_command():

    command = input("Enter your input: ")

    temp = sys.stdout

    f = open('out.txt', 'w')

    sys.stdout = f

    result = os.popen(command)

    print(result.readlines())

    sys.stdout.close()  # 关闭管道

    sys.stdout = temp  # 重新连接至标准的输出



#正则匹配输出

def Reg():
    f = open('out.txt', 'r')

    st = f.read()

    result1 = re.compile('\d+\.\d+\s[A-Za-z]+/sec')

    results = result1.findall(st)

    rs = results[0]

    rd = results[1]

    f.close()

    return rs,rd


#将匹配字符输入excel表格

def write_excel():

    for i in range(1,get_rows()):

        cmd_command()

        Reg()

        workbook = xlrd.open_workbook(u'Net_test.xlsx', formatting_info=True)

        workbooknew = copy(workbook)

        ws = workbooknew.get_sheet(0)

        bandwidth = Reg()

        ws.write(i, 3, bandwidth[0])

        ws.write(i, 4, bandwidth[1])

        workbooknew.save(u'Net_test.xlsx')


#获取行数

def get_rows():
    wb = xlrd.open_workbook(u'Net_test.xlsx', formatting_info=True)

    sheet1 = wb.sheet_by_index(0)
    sheet1 = wb.sheet_by_name('Intranet')

    rows = sheet1.nrows
    return rows

write_excel()



