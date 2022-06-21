#!/bin/python
#coding:utf-8
"""
Author: Michael Jin
Date: 2021-06-08
"""


import os
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

'''
def Search_table(tables,search_string):
    for table in tables:
#        print(table)
        for cells in table._cells:
            print(cells[0].text)
#            print(table._cells[0].text)
            for i in range(0,10):
                if cells[i].text == search_string:
                    print("Yes")
'''
####################################### main program #############################################
while True:
    try:
        print("This TDS is for IEC 60335-2-40:2018 in conjunction with IEC 60335-1:2010+A1:2013+A2:2016!")
        print("================================================== Program begin ==================================================")
#        job=input("Please input the project number:")
        document = Document('.\cr.docx') #Open the template document
        tables = document.tables #获取文档中表格对象列表
#        Search_table(tables,'24.1')
        table29 = tables[29] #获取第30个表格, table 24.1
        rows=table29.rows #获取表格所有行数
        columns=table29.columns #获取表格所有列数
#        print(rows[0].cells[1].paragraphs[0].runs[0].bold) #检查表格中字体是否加粗
        for i in range(2,len(rows)):
            print(rows[i].cells)
        if rows[0].cells[0].text=='24.1': #判断表格中文字的内容
            print("yes")
#        for i in range(0,10):
#            table29.add_row() #增加一行
        ###方法一：直接通过单元格来获取数据
#        cells = table5._cells #获取相应表格所有的单元格
#        j=1
#        for i in range(3,5863,4):
#            cells[i].text = str(j)
#            j = j+1

        ###方法二：先获取列对象，再获取列的单元格
#        cols = table5.columns #获取表格中的列对象
#        col3 = cols[3] #获取第四列
#        for i in range(0,1466):
#            col3.cells[i].text = "P"

        document.save('cdf.docx')

        print("==================================================Program END ==================================================") 
        flag=input("Press Enter to continue! Others to EXIT!")
        if flag!="":
            break
        else:
            os.system("cls")

    except:
        print("==================================================Program END ==================================================") 
        print("Error, please contact Michael.jc.jin@intertek.com")
        break

#document.save('D:\Downloads\Tools4Cert-master\demo2.docx')
