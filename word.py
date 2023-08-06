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
from excel import get_UC
from PyPDF2 import PdfFileMerger
import PyPDF2
import time
if os.name=='nt':
    from win32com.client import Dispatch
    import xlwings as xw


def Menu():
    choice=input('请输入你的选择：\n1.生成年检报告\n2.提取数据\n3.doc转docx\n4.批量doc转PDF\n5.合并pdf\n6.doc转pdf\n7.pdf加水印\ndft:生成草稿报告')
    if choice=='1':
        path_xls=input('请输入需要做年检的报告（excel)的文件夹路径')
        path_doc=input('请输入年检报告(word)的路径')
        path_doc=path_doc.replace('"','')#去除"号,做预处理
        app=xw.App(visible=False,add_book=False)#创建app对象，传入Annual_checks函数
        for component in ['compressor','motor','smps','transformer','pwb','power unit']:
            Annual_checks(app,path_xls,path_doc,component)
        app.kill()#关闭进程
    elif choice=='2':
        data=input('请输入要提取的数据文件的路径：')
        data=data.replace('"','')
        docx=Document(data)
        new_docx=Document()
        tables=docx.tables
        print('the numbers of tables:',len(tables))
        table_components=find_table(tables,'hello')
        print('找到部件清单的表格：',table_components)
        rows_value=get_more(table_components)
        print('获取的数据行数：',len(rows_value))
        new_table=new_docx.add_table(rows=len(rows_value),cols=6,style="Table Grid")
        print('生成的新表格的行数：',len(new_table.rows))
        write_table(new_table,rows_value)
        new_docx.save(data[:-5]+'output.docx')
    elif choice=='3':
        path=input('请输入需要转换的doc文件路径：')
        path=path.replace('"','')
        doc2docx(path)
    elif choice=='4':
        path=input('请输入需要转换的doc文件夹路径：')
        path=path.replace('"','')
        docs2pdfs(path)
    elif choice=='5':
        target_path = input('PDF的文件夹路径:')
        pdf_merge(target_path)
    elif choice=='6':
        path=input('请输入需要转换的doc文件夹路径：')
        path=path.replace('"','')
        doc2pdf(path)
    elif choice=='7':
        pdf_file=input('请输入pdf文件的路径:')
#        pdfWriter = PyPDF2.PdfFileWriter()      # 用于写pdf
#        pdfReader = PyPDF2.PdfFileReader(pdf_file)   # 读取pdf内容
        watermark='K:\Database\watermark.pdf'
        add_watermark(pdf_file,watermark)
#        # 遍历pdf的每一页,添加水印
#        for page in range(pdfReader.numPages):
#            page_pdf = add_watermark(watermark, pdfReader.getPage(page))
#            pdfWriter.addPage(page_pdf)
#        
#        with open(pdf_file, 'wb') as target_file:
#            pdfWriter.write(target_file)
    elif choice=='dft':
        path=input('请输入需要转换的doc文件路径：')
#        filename=os.path.basename(path)
#        dirname=os.path.dirname(path)
#        print(filename)
#        print(dirname)
        watermark='K:\Database\watermark.pdf'
        path=path.replace('"','')
        new_pdf=doc2pdf(path)
        print(new_pdf)
#        add_watermark(path[:-3]+'pdf',watermark)
        add_watermark(new_pdf,watermark)
        os.remove(new_pdf)



def add_watermark(pdf_path,watermark):
    """
    将水印pdf与pdf的一页进行合并
    :param water_file:
    :param page_pdf:
    :return:
    """
    pdf_water = PyPDF2.PdfFileReader(watermark)
    pdf_file = PyPDF2.PdfFileReader(pdf_path)   # 读取pdf内容
    pdfWriter = PyPDF2.PdfFileWriter()      # 用于写pdf
#    pdf_file.mergePage(pdf_file.getPage(0))
    # 遍历pdf的每一页,添加水印
    for page in range(pdf_file.numPages):
        pdf_page=pdf_file.getPage(page)
        pdf_page.mergePage(pdf_water.getPage(0))
        pdfWriter.addPage(pdf_page)

#    print(os.path.dirname(pdf_path))
    
    ###以新文件保存
    with open(pdf_path[:-4]+'_draft.pdf', 'wb') as target_file:
        pdfWriter.write(target_file)
    return target_file


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


def pdf_merge(path):
    pdf_lst = [f for f in os.listdir(path) if f.endswith('.pdf')]
    pdf_lst = [os.path.join(path, filename) for filename in pdf_lst]
    
    file_merger = PdfFileMerger()
    for pdf in pdf_lst:
        file_merger.append(pdf)     # 合并pdf文件
    
    file_merger.write(os.path.join(path,'Draft_report.pdf'))
    pass


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
#        for i in range(1,7):#针对table24.1,会返回8个cell, 取其中的1-6个，头和尾是重复的，不知道原因
        for i in range(0,6):#针对标准的TRF中table24.1
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

def write_table(new_table,rows_value):#把获取的数据写入到新建的表格中
    for i in range(0,len(rows_value)-1):#遍历数据列表中的每一个元素
#        print(f'写入数据第{i}行')
        for j in range(0,6):
            new_table.rows[i].cells[j].text=rows_value[i][j]
            new_table.rows[i].cells[j].paragraphs[0].runs[0].font.name='Arial'
            new_table.rows[i].cells[j].paragraphs[0].runs[0].font.size=Pt(10)
            print(f'写入数据第{i}行第{j}列数据')

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

        
def Annual_check(docx,data,component):#对具体的年检信息进行写入
#    table=docx.tables[0]
    table_content=find_table(docx.tables,'Unlisted Component')
    table_test=find_table(docx.tables,'Model No.')
    table_construction=find_table(docx.tables,'Model')
    print('找到目录表格:',len(table_content))
    print('找到耐压测试表格:',len(table_test))
    print('找到结构表格:',len(table_construction))
    flag=1#正常写入年检信息返回0，默认1
    for uc in data['uc_info']:
        if component in uc['name'].lower():#只处理指定的部件
            flag=0#检测到年检信息，更改返回值为0
            row_cells=table_content[0].add_row().cells#找到目录表格后增加一行，写入对应的数据
            row_cells[0].text=uc['name']
            row_cells[1].text=uc['manufacturer']
            row_cells[2].text=uc['model']
            row_cells[3].text=data['basic_info']['report']
            row_cells[4].text=str(uc['photo_no'])
            row_cells[5].text=str(uc['item_no'])


            #以下部分为多个绕组信息处理
            k=1
            while f'designation_{k}' in list(uc.keys()):
                row_cells=table_construction[0].add_row().cells#找到结构表格后增加一行，写入对应数据
                row_cells[0].text=uc['model']
                row_cells[1].text=uc[f'designation_{k}']
                row_cells[2].text=str(uc[f'wire_size_{k}'])
                row_cells[3].text=str(uc[f'resistance_{k}'])
#                row_cells[4].text=
                row_cells[5].text='Pass'
                k=k+1

            #当数据大于一行时，进行合并操作
            if k>2:
                total_rows=len(table_construction[0].rows)-1#获取结构表格总行数，-1是为了匹配索引
                col_0=table_construction[0].columns[0].cells#获取结构表格第一列的单元格

                #清除需要合并单元格中除第一格外的其他数据
                for col in range(total_rows-(k-3),total_rows+1):
                    print(f'清除结构表格第{col}行数据',col_0[col].text)
                    col_0[col].text=""

#                print('k:',k)
                #合并单元格
                print(f'合并结构表格第{total_rows-k+2}:{total_rows}行的单元格')
                table_construction[0].cell(total_rows-k+2,0).merge(table_construction[0].cell(total_rows,0))


            j=1
            while f'location_{j}' in list(uc.keys()):#找到测试表格后增加一行，写入对应数据
                row_cells=table_test[0].add_row().cells
                row_cells[0].text=uc['model']
                row_cells[1].text=uc['manufacturer']
                row_cells[2].text=uc[f'location_{j}']
                row_cells[3].text=uc[f'rating'].split(',')[0]
                row_cells[4].text=uc[f'voltage_{j}']
                row_cells[5].text='Pass'
                j=j+1

            #当数据大于一行时，进行合并操作
            if j>2:
                total_rows=len(table_test[0].rows)-1#获取测试表格总行数，-1是为了匹配索引
                col_0=table_test[0].columns[0].cells#获取测试表格第一列的单元格
                col_1=table_test[0].columns[1].cells#获取测试表格第二列的单元格

                #清除需要合并单元格中除第一格外的其他数据
                for col in range(total_rows-(j-3),total_rows+1):
                    print(f'清除测试表格第{col}行数据',col_0[col].text)
                    col_0[col].text=""
                    col_1[col].text=""

                print('j:',j)
                #合并单元格
                print(f'合并测试表格第{total_rows-j+2}:{total_rows}行的单元格')
                table_test[0].cell(total_rows-j+2,0).merge(table_test[0].cell(total_rows,0))
                table_test[0].cell(total_rows-j+2,1).merge(table_test[0].cell(total_rows,1))
    return flag


def exit_file(file_path):#判断一个文件是否存在
    dirname=os.path.dirname(file_path)
    filename=os.path.basename(file_path)
    for file in os.listdir(dirname):
        if filename==file:
            return True

            
def Annual_checks(app,path_xls,path_doc,component):#对多个报告生成对应部件的年检报告
    files=[f for f in os.listdir(path_xls) if f.endswith('.xls')] #列出所有的xls文件
    file_path=[os.path.join(path_xls, filename) for filename in files]#所有xls文件的绝对路径
    new_file=path_doc[:-4]+component+'.docx'#构造新的docx文件路径，即输出对应部件的年检报告
    print(file_path)
    for file in file_path:#遍历所有的xls文件，即所有需要做年检的报告
        print(f'正在处理{file}')
        wb=app.books.open(file)
        data=get_UC(wb)#提取相应的UC信息
        print(data)
        wb.close()
        if exit_file(new_file):
            docx=Document(new_file)#如果存在对应的年检报告，则在对应报告中添加
        else:
            docx=Document(path_doc)
        if not Annual_check(docx,data,component):#如果返回为0，则为正常写入数据，此时保存年检报告
            docx.save(new_file)
            

def update_components():#更新修改table24.1
    pass

def content_replace(documents,old_word,new_word):#替换对应文字，保持格式不变
    paragraphs=documents.paragraphs
    for paragraph in paragraphs:
        if old_word in paragraph.text:
            text=paragraph.text
            name=paragraph.runs[0].font.name
            size=paragraph.runs[0].font.size
            color=paragraph.runs[0].font.color.rgb
            print(text,name,size,color)
            update_text=text.replace(old_word,new_word)
            paragraph.text=update_text
            paragraph.runs[0].font.name=str(name)
            paragraph.runs[0].font.size=int(size)
#            paragraph.runs[0].font.color=str(color)

def Annual_init(path_doc):
    client_name='Yoau'
    report_No='202111002SHA-001'
    control_No='3061710'
    docx=Document(path_doc)
    content_replace(docx,'CUSTOMER NAME',client_name)
    content_replace(docx,'<report no.>',report_No)
    content_replace(docx,'<issue_date>',time.strftime('%d-%m-%Y'))
    content_replace(docx,'<Control Number>',control_No)
    docx.save(path_doc[:-4]+'init'+'.docx')



if __name__=='__main__':
    while True:
        os.system('cls')
        Menu()
