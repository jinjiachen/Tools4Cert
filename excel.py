#/bin/python 
#coding:utf-8

import xlrd
import xlwt
#import xlutils
from xlutils.copy import copy
import xlwings as xw

def Menu():
#    choice=input("1.��ȡ����\n2.�޸ı���")
    choice=input("1.Extract data\n2.Revise the report")
    if choice=='1':
        rpt=input("Please input the report path:")
        rpt_start=int(input("Please input the start line of report:"))
        data=input("Please input the data source path:")
        data_start=int(input("Please input the start line of data:"))
        data_end=int(input("Please input the end line of data:"))
        data_col1=int(input("Please choose four columns of data (1/4):"))
        data_col2=int(input("Please choose four columns of data (2/4):"))
        data_col3=int(input("Please choose four columns of data(3/4):"))
        data_col4=int(input("Please choose four columns of data (4/4):"))
        get_data(rpt,data)
    elif choice=='2':
#        app=xw.App(visible=True,add_book=False)
        rpt=input("Please input the report path:") #����Ҫ�޸ĵı����·��
        data=input("Please input the data source path:") #��������Դ��·��
        app=xw.App(visible=False,add_book=False)
        wb=app.books.open(rpt)
        sh=wb.sheets['4.0 Components']
        wb1=app.books.open(data)
        sh1=wb1.sheets['4.0 Components']
        update4(sh,sh1)
#        a=get_row_number(sh1,'h','A')
#        data=copy_line(sh1,a)
#        print(data)
#        print(sh['c44'].value)
#        print(len(data[0]))
#        print(len(sh['c44'].value))
#        print(data[0]==sh['c44'].value)
#        a=get_row_number(sh,'c',data[0])
#        a=lookdown(sh,'c',a)
#        print(a)
#        insert_line(sh,a,data)
#        print(row_range(sh,data))
        wb.save('output1.xls')
        wb.close()
        wb1.close()
        app.quit()
#        a=get_name('201100941SHA-001_R3.xls')
    

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
    font.height=20*10  #10������,20Ϊ����
    style.font=font
    borders=xlwt.Borders()
    borders.left=xlwt.Borders.THIN
    borders.right=xlwt.Borders.THIN
    borders.top=xlwt.Borders.THIN
    borders.bottom=xlwt.Borders.THIN
    style.borders=borders
    alignment=xlwt.Alignment()
    alignment.wrap=1
    alignment.horz=0x01 #0x01�����0x02����0x03�Ҷ���
    alignment.vert=0x01 #0x00�϶���0x01����x02�¶���
    style.alignment=alignment

    print(xls_new)
    sheet_data=xls_data.sheet_by_name('4.0 Components')
    sheet_new=xls_new.get_sheet('4.0 Components')
    for i in range(rpt_start,rpt_end):
        sheet_new.write(i,2,sheet_data.cell_value(i-(rpt_start-data_start),data_col1).replace(',',', '),style)
#        sheet_new.write(i,3,sheet_data.cell_value(i-(rpt_start-data_start),data_col2).replace(',',', ')+'\n('+sheet_data.cell_value(i-(rpt_start-data_start),9)+')',style)
        sheet_new.write(i,3,sheet_data.cell_value(i-(rpt_start-data_start),data_col2).replace(',',', '),style)
        sheet_new.write(i,4,sheet_data.cell_value(i-(rpt_start-data_start),data_col3).replace(',',', '),style)
        sheet_new.write(i,5,sheet_data.cell_value(i-(rpt_start-data_start),data_col4).replace(',',', '),style)
    xls_new.save('output.xls')	

def get_name(filename): #��ȡ�����б�
    a=[]
    xls=xlrd.open_workbook(filename,formatting_info=True)
    sheet=xls.sheet_by_name('4.0 Components')
    for i in sheet.col_values(2):
        if i not in a:
            if i=='Name':
                pass
            elif i=='':
                pass
            else:
                a.append(i)
    return a

def get_manufacturer(filename): #��ȡ�������б�
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

def sort_by_name(filename): #������������
    xls=xlrd.open_workbook(filename,formatting_info=True)
    xls_new=copy(xls)
    sheet=xls.sheet_by_name('4.0 Components')
    sheet_new=xls_new.get_sheet('4.0 Components')

    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name='Arial'
    font.bold=True
    font.height=20*10  #10������,20Ϊ����
    style.font=font
    borders=xlwt.Borders()
    borders.left=xlwt.Borders.THIN
    borders.right=xlwt.Borders.THIN
    borders.top=xlwt.Borders.THIN
    borders.bottom=xlwt.Borders.THIN
    style.borders=borders
    alignment=xlwt.Alignment()
    alignment.wrap=1
    alignment.horz=0x01 #0x01�����0x02����0x03�Ҷ���
    alignment.vert=0x01 #0x00�϶���0x01����x02�¶���
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
        
def sort_by_manufacturer(filename): #��������������
    xls=xlrd.open_workbook(filename,formatting_info=True)
    xls_new=copy(xls)
    sheet=xls.sheet_by_name('4.0 Components')
    sheet_new=xls_new.get_sheet('4.0 Components')

    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.name='Arial'
    font.bold=True
    font.height=20*10  #10������,20Ϊ����
    style.font=font
    borders=xlwt.Borders()
    borders.left=xlwt.Borders.THIN
    borders.right=xlwt.Borders.THIN
    borders.top=xlwt.Borders.THIN
    borders.bottom=xlwt.Borders.THIN
    style.borders=borders
    alignment=xlwt.Alignment()
    alignment.wrap=1
    alignment.horz=0x01 #0x01�����0x02����0x03�Ҷ���
    alignment.vert=0x01 #0x00�϶���0x01����x02�¶���
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
    
def copy_line(sheet,row): #xlwings:����ָ����
    index=f'C{row}:F{row}' #����C1:F1�ַ�������
    while sheet[f'C{row}'].value==None: #���ָ����C���Ƿ�Ϊ�գ����Ϊ�գ�����Ѱ��֪���ҵ��ǿ���ֵ
        row=row-1
    data=sheet[index].value #��C1:F1�����ݸ��Ƹ�data
    data[0]=sheet[f'C{row}'].value #��C�е���ֵ�����£��Է�Ϊ��
    return data

def paste_line(sheet,row,data): #xlwings:ָ����ճ��
#    row=str(row)
#    index='C'+str(row)+':'+'F'+str(row)
    index=f'C{row}:F{row}'
    sheet[index].value=data

def insert_line(sheet,row,data): #xlwings:��ָ���к������в�д������
    sheet.api.Rows(str(row+1)).Insert()
    paste_line(sheet,str(row+1),data)

def get_row_number(sheet,col,words): #xlwings:���ҹؼ��ʲ���������
    for i in range(1,100):
        cell=sheet[f'{col}{i}'].value
        if cell==words:
            return i
            break

def lookdown(sheet,col,row): #xlwings:��������Ѱ�ң��Ƿ��п�ֵ��ֱ���ҵ���һ���ǿյ�Ԫ��
    while(sheet[f'{col}{row+1}'].value==None):
        row=row+1
    return row

def fmt(sheet):#Ŀǰ��Ҫ�Ǻϲ�name�еĵ�Ԫ��
    for i in range(1,100):
        pass
    

def row_range(sheet,data): #xlwings:������ͬname�Ĳ�����������Χ
    rows=[]
    for i in range(1,100):#�ڱ���Ĵ�������Χ��ȥƥ��
        if sheet[f'c{i}'].value==data[0]:#c����Ѱ��data[0]����Name
            row_start=i #ͬname�Ĳ�������ʼ��
            rows.append(row_start)#�ҵ���Ӧ�Ĺؼ��ʣ���¼��ʼ��
            row_end=lookdown(sheet,'c',i)
            rows.append(row_end)#��¼�ݶ��Ľ����У�����·���ͬһ��������ᱻ����������������ǣ���������յ�����
            while(sheet[f'c{row_end+1}'].value==data[0]):
                row_end=row_end+1#ͬname�Ĳ����Ľ�����
                rows[1]=row_end #�ҵ�ͬ���Ĳ����������½�����
        if len(rows)==2:
            break
    return rows
#    return min(rows),max(rows)

def update4(sheet1,sheet2,):#xlwings:����4.0��Ϣ
    for i in range(1,100): #�ڴ�������Χ��ȥƥ����Ҫ�޸ĵ���Ϣ
        if sheet2[f'h{i}'].value=="A": #�ж�H���Ƿ�ΪA��AΪ����
            data=copy_line(sheet2,i)#���ƶ�Ӧ�е�����
            print('add:',data)
            for j in range(1,100):#�ڱ���Ĵ�������Χ��ȥƥ��
                if sheet1[f'c{j}'].value==data[0]:#c����Ѱ��data[0]����Name
                    row=lookdown(sheet1,'c',j)
                    while(sheet1[f'c{row+1}'].value==data[0]):#��һ�����Name��ͬ����ͬһ�������������������
                        row=row+1
                    print(row)
                    break
            insert_line(sheet1,row,data) #�ڸ��к����������
        elif sheet2[f'h{i}'].value=="R": #�ж�H���Ƿ�ΪR��RΪ�޸�
            data=copy_line(sheet2,i)#���ƶ�Ӧ�е�����
            print('revise:',data)
#            for j in range(1,100):#�ڱ���Ĵ�������Χ��ȥƥ��
#                if sheet1[f'c{j}'].value==data[0]:#c����Ѱ��data[0]����Name
#                    row_start=j #ͬname�Ĳ�������ʼ��
#                    row_end=lookdown(sheet1,'c',j)
#                    while(sheet1[f'c{row+1}'].value==data[0]):
#                        row_end=row_end+1#ͬname�Ĳ����Ľ�����
#                    break
            rows=row_range(sheet1,data)
            print(rows)
            for j in range(rows[0],row[1]+1):#��ͬһ��������������Χ��ȥƥ����Ϣ
#                paste_line(sheet1,row,data)
                pass
#    wb.save('output1.xls')
        
    
    

if __name__=='__main__':
    Menu()
#    rpt=input("Please input the report path:")
#    rpt_start=int(input("Please input the start line of report:"))
#    data=input("Please input the data source path:")
#    data_start=int(input("Please input the start line of data:"))
#    data_end=int(input("Please input the end line of data:"))
#    data_col1=int(input("Please choose four columns of data (1/4):"))
#    data_col2=int(input("Please choose four columns of data (2/4):"))
#    data_col3=int(input("Please choose four columns of data(3/4):"))
#    data_col4=int(input("Please choose four columns of data (4/4):"))
#    get_data(rpt,data)
#    sort_by_name(rpt)
