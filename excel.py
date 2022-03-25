# -*- coding:utf-8 -*-
import xlrd, xlwt
import os
from xlutils.copy import copy
from openpyxl import load_workbook

if __name__ == '__main__':
    my_str=input("请输入您想定位的字符串(按回车确认):")
    path=r'C:\Users\pzl96\Desktop\python_excel'#你想获取数据的excel所在的文件夹绝对路径
    filenames=os.listdir(path)
    wt_cow = 0
    for filename in filenames:
        wt_cow=wt_cow+1
        print(wt_cow)
        # print(filename)
        try:
            data = xlrd.open_workbook(path+'\\'+filename)
        except Exception as e:
            print(e)
        s1 = data.sheet_by_index(0).col_values(-1)
        # print(len(s1))
        str_cow = 0
        # 从指定的列中获取我们想要的数据
        for i in range(len(s1)):
            if s1[i] == my_str:
                str_cow = i
        print(str_cow)
        data1 = s1[str_cow+1]
        data2 = s1[-1]
        #写数据方法1
        wr_table_path=r'C:\Users\pzl96\Desktop\python_wt1.xls' #你想保存数据的excel所在的文件夹绝对路径
        wr_table=xlrd.open_workbook(wr_table_path)
        wb=copy(wr_table)
        sheet1=wb.get_sheet(0)
        sheet1.write(str_cow,0,filename)
        sheet1.write(str_cow,1,data1)
        sheet1.write(str_cow,2,data2)
        os.remove(wr_table_path)
        wb.save(wr_table_path)
        print("第1个数据为：" + str(data1))
        print("第2个数据为：" + str(data2))
print("数据保存完成，表中有空白行，请在excel中使用Ctrl+G查找空值行，删除-删除所选行 ")




