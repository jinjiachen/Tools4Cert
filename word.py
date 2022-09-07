#!/bin/python
#coding:utf-8
"""
Author: Michael Jin
Date: 2022-08-11
"""


import os
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from win32com.client import Dispatch

def doc2docx(path):#将doc文件转换成docx
    '''path具体到文件'''
    w = Dispatch('Word.Application') #打开word程序
    w.Visible = 0 #后台运行，不可见
    w.DisplayAlerts = 0
    doc = w.Documents.Open(path) #打开对应的doc文件
    newpath = os.path.splitext(path)[0] + '.docx' #生成docx文件路径
    doc.SaveAs(newpath, 12, False, "", True, "", False, False, False, False)#12为docx的类型
    doc.Close()
    w.Quit()
#    os.remove(path)#不删除源文件
    return newpath


def doc2pdf(path):#将doc文件转换成pdf
    '''path具体到文件'''
    w = Dispatch('Word.Application') #打开word程序
    w.Visible = 0 #后台运行，不可见
    w.DisplayAlerts = 0
    doc = w.Documents.Open(path) #打开对应的doc文件
    newpath = os.path.splitext(path)[0] + '.pdf' #生成docx文件路径
    doc.SaveAs(newpath, 17)#17为pdf的类型
    doc.Close()
    w.Quit()
#    os.remove(path)#不删除源文件
    return newpath

def docs2pdfs(path):#批量转换doc为pdf文件
    '''输入文件夹路径即可'''
    files=[f for f in os.listdir(path) if f.endswith('.doc')]
    file_path=[os.path.join(path, filename) for filename in files]
    for file in file_path:
        doc2pdf(file)


def find_table(tables,search_string):#关键词查找对应表格
    table_components=[]
    for table in tables:#遍历每一个表格
        if table.cell(0,0).text=='24.1' and table.cell(0,1).text=='TABLE: Critical components information':#标准TRF中零部件清单表格
            table_components.append(table)
        elif table.cell(0,0).text==search_string:#自定义关键字来搜索表格
            table_components.append(table)
    return table_components


def Search_table(tables,search_string):
    table_components=[]
    for table in tables:
#        print(type(table))
        for row in table.rows:
#            print(type(row.cells))
#            print(len(row.cells))
            for cell in row.cells:
#                print(cell.text)
                if cell.text.strip()==search_string:
#                if search_sring in cell.text:
#                    print(table)
                    table_components.append(table)
#                    break
    return table_components

def get_rows(table):#输入一个表格，返回一个以列表为元素的列表，每一个列表就是一行
    rows_values=[]
    for row in table.rows:#遍历表格的每一行
        if is_italic_row(row)==True:
            continue
        values=[]
#        if row.cells[0].text=='Object / part':
#            print('删除object所在行')
#            continue
#        elif row.cells[0].text=='hello':
#            print('删除hello所在行')
#            continue
#        elif row.cells[0].text=='Object / part No.':
#            print('删除Object / part No.所在行')
#            continue
        for i in range(1,7):#针对table24.1,会返回8个cell, 取其中的1-6个，头和尾是重复的，不知道原因
            values.append(row.cells[i].text)#把一行的数据放在value这个列表中
#            print(values)
#            break
        rows_values.append(values)#把value添加到rows_values这个总列表中
#            print(rows_values)
    return rows_values

def get_more(tables):#输入多个表格，返回每行的数据
    rows_values=[]
    for table in tables:
        rows_values=rows_values+get_rows(table)#把每个表格的行数据拼接起来
    return rows_values

def write_table(rows_value):#把获取的数据写入到新建的表格中
    for i in range(0,len(rows_value)-1):#遍历数据列表中的每一个元素
        for j in range(0,6):
            new_table.rows[i].cells[j].text=rows_value[i][j]
            new_table.rows[i].cells[j].paragraphs[0].runs[0].font.name='Arial'
            new_table.rows[i].cells[j].paragraphs[0].runs[0].font.size=Pt(10)

def is_italic_cell(cell):#判断表格中的cell是否为斜体,如果是空的，返回None，如果为斜体返回Ture，其他返回False
    if cell.text=='':
        return None
    if cell.paragraphs[0].runs[0].italic==True:
        return True
    else:
        return False
    
def is_bold_cell(cell):#判断表格中的cell是否为加粗
    if cell.paragraphs[0].runs[0].bold==True:
        return True
    else:
        return False

def is_italic_row(row):
    for cell in row.cells:
        if is_italic_cell(cell)==None:
            continue
        elif is_italic_cell(cell)==True:
            continue
        else:
            return False
            break 
    return True

#def obmit_row(list):
#    for l in list:
        

if __name__=='__main__':
    data=input('请输入要提取的数据文件的路径：')
#    docx=Document(r'B:\其他客户\220602760SHA_Schneider_IEC_report\Others\UL EN报告\SA12773-13CA18037 CB Report ACRC301S ACRC301H.docx')
    docx=Document(data)
    new_docx=Document()
    tables=docx.tables
    print('the numbers of tables:',len(tables))
#    table_components=Search_table(tables,'hello') 
    table_components=find_table(tables,'hello')
    print('找到部件清单的表格：',table_components)
#    print(len(table_components.rows))
#    rows_value=get_rows(table_components)
#    rows_value=[]
#    if False:
#        rows_value=get_rows(table_components)
#    else:
#        for table in table_components:
#            rows_value=rows_value+get_rows(table)
    rows_value=get_more(table_components)
    print('获取的数据行数：',len(rows_value))

    new_table=new_docx.add_table(rows=len(rows_value),cols=6,style="Table Grid")
    print('生成的新表格的行数：',len(new_table.rows))
    write_table(rows_value)
    new_docx.save(r'B:\其他客户\220602760SHA_Schneider_IEC_report\Others\UL EN报告\output.docx')