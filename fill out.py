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


####################################### main program #############################################
while True:
    try:
        print("This TDS is for IEC 60335-2-40:2018 in conjunction with IEC 60335-1:2010+A1:2013+A2:2016!")
        print("================================================== Program begin ==================================================")
#        job=input("Please input the project number:")
        document = Document('.\cr.docx') #Open the template document
        tables = document.tables #获取文档中表格对象列表
        table5 = tables[5] #获取第六个表格,即条款判断部分
        ###方法一：直接通过单元格来获取数据
        cells = table5._cells #获取相应表格所有的单元格
        j=1
        for i in range(3,5863,4):
            cells[i].text = str(j)
            j = j+1

        ###方法二：先获取列对象，再获取列的单元格
#        cols = table5.columns #获取表格中的列对象
#        col3 = cols[3] #获取第四列
#        for i in range(0,1466):
#            col3.cells[i].text = "P"

        document.save('cr_fill.docx')

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
