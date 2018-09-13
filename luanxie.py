
import os
import sys
import re
import xlrd
from xlutils.copy import copy

'''
temp=[u'Port',u'Date_size',u'Thread_Num',u'Up_bandwidth',u'Down_bandwidth',u'Packet_loss_rate']
row = 1
for pos,v in enumerate(temp):
    ws.write(row,pos,v)
    row = row+1
'''

    # 创建管道，将cmd命令输出至out.txt
#def cmd_command():
def xx():
    command = input("Enter your input: ")

    temp = sys.stdout

    f = open('out.txt', 'w')

    sys.stdout = f

    result = os.popen(command)

    print(result.read())

    sys.stdout.close()  # 关闭管道
    # while(1):
    # pass

    sys.stdout = temp  # 重新连接至标准的输出

    f.close()

#def Reg():

    f = open('out.txt', 'r')

    st = f.read()

#for line in st:

    result1 = re.compile('\d+\.\d+\s[A-Za-z]+/sec')

    results = result1.findall(st)
    rs = results[0]
    rd = results[1]
    return rs, rd

bandwidth = xx()
print(bandwidth[0])



#print(result1)

    #rs = result1.group(0)

    #rd = result.group(1)

#f.close()

    #print(rs, rd)
'''
    return rs, rd

cmd_command()

Reg()
#将匹配字符输入excel表格

workbook = xlrd.open_workbook(u'Net_test_copy.xlsx', formatting_info=True)

workbooknew = copy(workbook)

ws = workbooknew.get_sheet(1)

ws.write(1, 3, Reg())

workbooknew.save(u'Net_test_copy.xlsx')

'''










