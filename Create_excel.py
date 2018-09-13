#import  os
import xlwt

#cmd命令操作
#result = os.popen("iperf -c check-test -p 80 -i 2 -t 10 -w 4k")
#print (result.read())

#excel表创建
f = xlwt.Workbook()
sheet1 = f.add_sheet(u'Intranet', cell_overwrite_ok=True)
sheet2 = f.add_sheet(u'Extranet', cell_overwrite_ok=True)

#####excel写入内容

#Intranet写入

row=0

temp=[u'Port',u'Window_size',u'Thread_Num',u'Up_bandwidth',u'Down_bandwidth',u'Packet_loss_rate']
for pos,v in enumerate(temp):
    sheet1.write(row,pos,v)
row+=1

sheet1.write(row,0,u'TCP80')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'TCP80')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet2.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'5')
row+=1
sheet1.write(row,0,u'TCP555')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'5')
row+=1
sheet1.write(row,0,u'UDP53')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'5')
row+=1
sheet1.write(row,0,u'UDP555')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'512b')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'4K')
sheet1.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'1')
row+=1
#sheet1.write(row,0,u'')
sheet1.write(row,1,u'1M')
sheet1.write(row,2,u'5')
row+=1
sheet1.write_merge(1, 8, 0, 0, u'TCP80')
sheet1.write_merge(9, 16, 0, 0, u'TCP555')
sheet1.write_merge(17, 24, 0, 0, u'UDP53')
sheet1.write_merge(25, 32, 0, 0, u'UDP555')
sheet1.write_merge(1, 2, 1, 1, u'512b')
sheet1.write_merge(3, 4, 1, 1, u'1K')
sheet1.write_merge(5, 6, 1, 1, u'4K')
sheet1.write_merge(7, 8, 1, 1, u'1M')
sheet1.write_merge(9, 10, 1, 1, u'512b')
sheet1.write_merge(11, 12, 1, 1, u'1K')
sheet1.write_merge(13, 14, 1, 1, u'4K')
sheet1.write_merge(15, 16, 1, 1, u'1M')
sheet1.write_merge(17, 18, 1, 1, u'512b')
sheet1.write_merge(19, 20, 1, 1, u'1K')
sheet1.write_merge(21, 22, 1, 1, u'4K')
sheet1.write_merge(23, 24, 1, 1, u'1M')
sheet1.write_merge(25, 26, 1, 1, u'512b')
sheet1.write_merge(27, 28, 1, 1, u'1K')
sheet1.write_merge(29, 30, 1, 1, u'4K')
sheet1.write_merge(31, 32, 1, 1, u'1M')

#Extranet写入
row=0
temp=[u'Port',u'Window_size',u'Thread_Num',u'Up_bandwidth',u'Down_bandwidth',u'Packet_loss_rate']
for pos,v in enumerate(temp):
    sheet2.write(row,pos,v)
row+=1

sheet2.write(row,0,u'TCP80')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'TCP80')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'5')
row+=1
sheet2.write(row,0,u'TCP555')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'5')
row+=1
sheet2.write(row,0,u'UDP53')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'5')
row+=1
sheet2.write(row,0,u'UDP555')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'512b')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1K')
sheet2.write(row,2,u'5')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'4K')
sheet2.write(row,2,u'5')
row+=1
#sheet1.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'1')
row+=1
#sheet2.write(row,0,u'')
sheet2.write(row,1,u'1M')
sheet2.write(row,2,u'5')
row+=1


sheet2.write_merge(1, 8, 0, 0, u'TCP80')
sheet2.write_merge(9, 16, 0, 0, u'TCP555')
sheet2.write_merge(17, 24, 0, 0, u'UDP53')
sheet2.write_merge(25, 32, 0, 0, u'UDP555')
sheet2.write_merge(1, 2, 1, 1, u'512b')
sheet2.write_merge(3, 4, 1, 1, u'1K')
sheet2.write_merge(5, 6, 1, 1, u'4K')
sheet2.write_merge(7, 8, 1, 1, u'1M')
sheet2.write_merge(9, 10, 1, 1, u'512b')
sheet2.write_merge(11, 12, 1, 1, u'1K')
sheet2.write_merge(13, 14, 1, 1, u'4K')
sheet2.write_merge(15, 16, 1, 1, u'1M')
sheet2.write_merge(17, 18, 1, 1, u'512b')
sheet2.write_merge(19, 20, 1, 1, u'1K')
sheet2.write_merge(21, 22, 1, 1, u'4K')
sheet2.write_merge(23, 24, 1, 1, u'1M')
sheet2.write_merge(25, 26, 1, 1, u'512b')
sheet2.write_merge(27, 28, 1, 1, u'1K')
sheet2.write_merge(29, 30, 1, 1, u'4K')
sheet2.write_merge(31, 32, 1, 1, u'1M')

f.save('Net_test.xlsx')





