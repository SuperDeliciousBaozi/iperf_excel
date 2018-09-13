import os
import sys
import re
import xlrd
from xlutils.copy import copy


#创建管道，将cmd命令输出至out.txt

temp = sys.stdout

f = open('out.txt', 'w+')

sys.stdout = f

result = os.popen("iperf -c check-test -p 80 -t 300 -P 1 -n 512b")

print(result.read())

sys.stdout.close()  #关闭管道
#while(1):
    #pass

sys.stdout = temp   #重新连接至标准的输出

f = open('out.txt', 'r')

st = f.read()

result1 = re.search('\d+\.\d+\s[A-Za-z]+/sec', st, re.M)

print(result1.group())

rs = result1.group()

f.close()

#将匹配字符输入excel表格

workbook = xlrd.open_workbook(u'Net_test.xlsx')

workbooknew = copy(workbook)

ws = workbooknew.get_sheet(0)

ws.write(3, 3, rs)

workbooknew.save(u'Net_test_new.xlsx')




"""
for i in range(1,31):

row=0
    sheet1.write(row,pos,v)
row+=1
excel = open('Net_test','w')
Net_test.Intranet.write(row,3,result1)
"""




