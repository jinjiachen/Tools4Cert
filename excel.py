#/bin/python 
#coding:utf-8

import xlrd
import xlwt
#import xlutils
from xlutils.copy import copy


def get_data(rpt_fn, data_fn):
    rpt_end=rpt_start+(data_end-data_start)
    xls_rpt=xlrd.open_workbook(rpt_fn,formatting_info=True)
    print(xls_rpt)
    xls_data=xlrd.open_workbook(data_fn)
    xls_new=copy(xls_rpt)
#    xls_new.save('output.xls')	

    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name='Arial'
    font.bold=True
    font.height=20*10  #10号字体,20为基数
    style.font=font
    borders=xlwt.Borders()
    borders.left=xlwt.Borders.THIN
    borders.right=xlwt.Borders.THIN
    borders.top=xlwt.Borders.THIN
    borders.bottom=xlwt.Borders.THIN
    style.borders=borders
    alignment=xlwt.Alignment()
    alignment.wrap=1
    alignment.horz=0x01 #0x01左对齐0x02居中0x03右对齐
    alignment.vert=0x01 #0x00上对齐0x01居中x02下对齐
    style.alignment=alignment

    print(xls_new)
    sheet_data=xls_data.sheet_by_name('4.0 Components')
    sheet_new=xls_new.get_sheet('4.0 Components')
    for i in range(rpt_start,rpt_end):
        sheet_new.write(i,2,sheet_data.cell_value(i-(rpt_start-data_start),data_col1).replace(',',', '),style)
        sheet_new.write(i,3,sheet_data.cell_value(i-(rpt_start-data_start),data_col2).replace(',',', ')+'\n('+sheet_data.cell_value(i-(rpt_start-data_start),9)+')',style)
        sheet_new.write(i,4,sheet_data.cell_value(i-(rpt_start-data_start),data_col3).replace(',',', '),style)
        sheet_new.write(i,5,sheet_data.cell_value(i-(rpt_start-data_start),data_col4).replace(',',', '),style)
    xls_new.save('output.xls')	

def get_name(filename):
    a=[]
    xls=xlrd.open_workbook(filename,formatting_info=True)
    sheet=xls.sheet_by_name('4.0 Components')
    for i in sheet.col_values(2):
        if i not in a:
            if i=='Name':
                pass
            else:
                a.append(i)
    return a

def get_manufacturer(filename):
    a=[]
    xls=xlrd.open_workbook(filename,formatting_info=True)
    sheet=xls.sheet_by_name('4.0 Components')
    for i in sheet.col_values(3):
        if i not in a:
            if i=='Manufacturer/ trademark2':
                pass
            else:
                a.append(i)
    return a

def sort_by_name(filename):
    xls=xlrd.open_workbook(filename,formatting_info=True)
    xls_new=copy(xls)
    sheet=xls.sheet_by_name('4.0 Components')
    sheet_new=xls_new.get_sheet('4.0 Components')

    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name='Arial'
    font.bold=True
    font.height=20*10  #10号字体,20为基数
    style.font=font
    borders=xlwt.Borders()
    borders.left=xlwt.Borders.THIN
    borders.right=xlwt.Borders.THIN
    borders.top=xlwt.Borders.THIN
    borders.bottom=xlwt.Borders.THIN
    style.borders=borders
    alignment=xlwt.Alignment()
    alignment.wrap=1
    alignment.horz=0x01 #0x01左对齐0x02居中0x03右对齐
    alignment.vert=0x01 #0x00上对齐0x01居中x02下对齐
    style.alignment=alignment

    k=1
    for i in get_name(filename):
        for j in range(1,360):
            if sheet.row_values(j)[2]==i:
                sheet_new.write(k,2,sheet.cell_value(j,2),style)
                sheet_new.write(k,3,sheet.cell_value(j,3),style)
                sheet_new.write(k,4,sheet.cell_value(j,4),style)
                sheet_new.write(k,5,sheet.cell_value(j,5),style)
                k=k+1
                
    xls_new.save('output.xls')
        
def sort_by_manufacturer(filename):
    xls=xlrd.open_workbook(filename,formatting_info=True)
    xls_new=copy(xls)
    sheet=xls.sheet_by_name('4.0 Components')
    sheet_new=xls_new.get_sheet('4.0 Components')

    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name='Arial'
    font.bold=True
    font.height=20*10  #10号字体,20为基数
    style.font=font
    borders=xlwt.Borders()
    borders.left=xlwt.Borders.THIN
    borders.right=xlwt.Borders.THIN
    borders.top=xlwt.Borders.THIN
    borders.bottom=xlwt.Borders.THIN
    style.borders=borders
    alignment=xlwt.Alignment()
    alignment.wrap=1
    alignment.horz=0x01 #0x01左对齐0x02居中0x03右对齐
    alignment.vert=0x01 #0x00上对齐0x01居中x02下对齐
    style.alignment=alignment

    k=1
    for i in get_manufacturer(filename):
        for j in range(1,360):
            if sheet.row_values(j)[2]==i:
                sheet_new.write(k,2,sheet.cell_value(j,2),style)
                sheet_new.write(k,3,sheet.cell_value(j,3),style)
                sheet_new.write(k,4,sheet.cell_value(j,4),style)
                sheet_new.write(k,5,sheet.cell_value(j,5),style)
                k=k+1
                
    xls_new.save('output.xls')
    

if __name__=='__main__':
    rpt=input("Please input the report path:")
#    rpt_start=int(input("Please input the start line of report:"))
#    data=input("Please input the data source path:")
#    data_start=int(input("Please input the start line of data:"))
#    data_end=int(input("Please input the end line of data:"))
#    data_col1=int(input("Please choose four columns of data (1/4):"))
#    data_col2=int(input("Please choose four columns of data (2/4):"))
#    data_col3=int(input("Please choose four columns of data(3/4):"))
#    data_col4=int(input("Please choose four columns of data (4/4):"))
#    get_data(rpt,data)
    sort_by_name(rpt)
