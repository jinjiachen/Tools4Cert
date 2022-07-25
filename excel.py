#/bin/python 
#-*-coding:utf-8-*-

import xlrd
import xlwt
#import xlutils
from xlutils.copy import copy
import xlwings as xw
import time

def Menu():
#    choice=input("1.提取数据\n2.修改报告")
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
        get_data(rpt,rpt_start,data,data_start,data_end,data_col1,data_col2,data_col3,data_col4)
    elif choice=='2':
#        app=xw.App(visible=True,add_book=False)
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        data=input("Please input the data source path:") #输入数据源的路径
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        sh=wb.sheets['4.0 Components']
        sh12=wb.sheets['12.0 Revisions']
        wb1=app.books.open(data)
        sh1=wb1.sheets['4.0 Components']
        start=time.time()
        update4(sh,sh1,sh12)
        end=time.time()
        print('operating time:',end-start)
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
#        app.quit()
        app.kill()
#        a=get_name('201100941SHA-001_R3.xls')
    

def get_data(rpt_fn,rpt_start, data_fn,data_start,data_end,data_col1,data_col2,data_col3,data_col4):
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
#        sheet_new.write(i,3,sheet_data.cell_value(i-(rpt_start-data_start),data_col2).replace(',',', ')+'\n('+sheet_data.cell_value(i-(rpt_start-data_start),9)+')',style)
        sheet_new.write(i,3,sheet_data.cell_value(i-(rpt_start-data_start),data_col2).replace(',',', '),style)
        if isinstance(sheet_data.cell_value(i-(rpt_start-data_start),data_col3),float)!=True:
            sheet_new.write(i,4,sheet_data.cell_value(i-(rpt_start-data_start),data_col3).replace(',',', '),style)
        else:
            sheet_new.write(i,4,sheet_data.cell_value(i-(rpt_start-data_start),data_col3),style)
    
        sheet_new.write(i,5,sheet_data.cell_value(i-(rpt_start-data_start),data_col4).replace(',',', '),style)
    xls_new.save('output.xls')	

def get_name(filename): #获取名字列表
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

def get_manufacturer(filename): #获取制造商列表
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

def sort_by_name(filename): #按照名字排序
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
        
def sort_by_manufacturer(filename): #按照制造商排序
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
    
def copy_line(sheet,row): #xlwings:复制指定行
    index=f'B{row}:F{row}' #构造B1:F1字符串索引
    data=sheet[index].value #把B1:F1的内容复制给data
    row_b=row
    row_c=row
    row_d=row
    row_f=row
    while sheet[f'B{row_b}'].value==None: #检查指定行B列(部件名)是否为空，如果为空，向上寻找直到找到非空数值
        row_b=row_b-1
    while sheet[f'C{row_c}'].value==None: #检查指定行C列(部件名)是否为空，如果为空，向上寻找直到找到非空数值
        row_c=row_c-1
    while sheet[f'D{row_d}'].value==None: #检查指定行D列(制造商)是否为空，如果为空，向上寻找直到找到非空数值
        row_d=row_d-1
    while sheet[f'F{row_f}'].value==None: #检查指定行F列(技术参数)是否为空，如果为空，向上寻找直到找到非空数值
        row_f=row_f-1
    data[0]=sheet[f'B{row_b}'].value #把B列的数值更新下，以防为空
    data[1]=sheet[f'C{row_c}'].value #把C列的数值更新下，以防为空
    data[2]=sheet[f'D{row_d}'].value #把D列的数值更新下，以防为空
    data[4]=sheet[f'F{row_f}'].value #把F列的数值更新下，以防为空
    return data

def paste_line(sheet,row,data): #xlwings:指定行粘贴
#    row=str(row)
#    index='C'+str(row)+':'+'F'+str(row)
    index=f'B{row}:F{row}'
    sheet[index].value=data
    sheet[index].api.Font.Color=0xFF00FF
#    sheet[index].api.Font.Bold=True
#    sheet[index].api.Font.Size
#    sheet[index].api.Font.Name

def insert_line(sheet,row,data): #xlwings:在指定行后插入空行并写入数据
    sheet.api.Rows(str(row+1)).Insert()
    paste_line(sheet,str(row+1),data)

def get_row_number(sheet,col,words): #xlwings:查找关键词并返回行数
    for i in range(1,200):
        cell=sheet[f'{col}{i}'].value
        if cell==words:
            return i
            break

def lookdown(sheet,col,row): #xlwings:继续往下寻找，是否有空值，直到找到下一个非空单元格
    while(sheet[f'{col}{row+1}'].value==None and sheet[f'd{row+1}:g{row+1}'].value!=empty(4)): #判断下一行指定列是否为空，并且d到g列不为空，为了防止空行导致出错
#        if row+1<=sheet_total_rows(sheet):
        row=row+1
#        else:
#            break
    return row

def fmt(sheet):#目前主要是合并name列的单元格
    name=get_col_list(sheet,'c',1,sheet_total_rows(sheet)) #获取C列的部件名
    print(name)
    for value in name:
        data=[]
        data.append(value)
        data.append(value)
        print(data)
        rows=row_range(sheet,data)
        print(rows)
        if rows[0]<rows[1]:
            sheet[f'a{rows[0]+1}:c{rows[1]}'].value=''
        sheet[f'c{rows[0]}:c{rows[1]}'].merge()
        sheet[f'b{rows[0]}:b{rows[1]}'].merge()
        sheet[f'a{rows[0]}:a{rows[1]}'].merge()
        
def separate(str,symbol): #字符串和分隔符拆分并重组函数，解决分割不当问题
    new_str=''
    str_list=str.split(symbol) #获取分割的字符串列表
    last_index=len(str_list)-1 #字符串列表长度-1，即为最后一个字符串的索引
    for i in str_list: #遍历字符串列表
#        i=i.replace(' ','') #去除空格
        if i!=str_list[last_index]: #判断是否为最后一个字符串
            i=i.strip() #去除空格
            new_str=new_str+i+symbol+' '
        else:
            i=i.strip() #去除空格
            new_str=new_str+i
    return new_str

    
def str_fmt(str):
#以下为中文的符号的处理
    str=str.replace('，',',')#替换中文逗号
    str=str.replace('（','(')#替换中文括号
    str=str.replace('）',')')#替换中文括号
    str=str.replace('：',':')#替换中文冒号
    str=str.replace('；',';')#替换中文分号

    if ',' in str:
        str=separate(str,',')
        print('正在分割逗号：',str)
    if ':' in str:
        str=separate(str,':')
        print('正在分割冒号：',str)
    if ';' in str:
        str=separate(str,';')
        print('正在分割分号：',str)

    return str

def list_fmt(list):
    for i in range(1,len(list)):
        if isinstance(list[i],str)==True: #只针对字符串进行格式化操作
#            print('正在处理：',list[i])
            list[i]=str_fmt(list[i])
    return list

def row_range(sheet,data): #xlwings:查找相同name的部件的行数范围
    rows=[]
    total_row=sheet_total_rows(sheet)+1
    for i in range(1,total_row):#在报告的此行数范围内去匹配
        if sheet[f'c{i}'].value==data[1]:#c列中寻找data[0]，即Name
            row_start=i #同name的部件的起始行
            rows.append(row_start)#找到对应的关键词，记录开始行
            row_end=lookdown(sheet,'c',i)
            rows.append(row_end)#记录暂定的结束行，如果下方是同一部件，则会被后面的替代，如果不是，这就是最终的行数
            while(sheet[f'c{row_end+1}'].value==data[1]):
#            while(sheet[f'c{row_end+1}'].value==data[1] or sheet[f'c{row_end+1}'].value==None):
                row_end=row_end+1#同name的部件的结束行
                rows[1]=row_end #找到同样的部件名，更新结束行
        if len(rows)==2:
            break
    return rows

#def row_range(sheet,col,words): #xlwings:查找相同name的部件的行数范围
#    rows=[]
#    for i in range(1,200):#在报告的此行数范围内去匹配
#        if sheet[f'col{i}'].value==words:#c列中寻找关键词
#            row_start=i #找到起始行
#            rows.append(row_start)#找到对应的关键词，记录开始行
#            row_end=lookdown(sheet,col,i)
#            rows.append(row_end)#记录暂定的结束行，如果下方是同一部件，则会被后面的替代，如果不是，这就是最终的行数
#            while(sheet[f'col{row_end+1}'].value==words):
#                row_end=row_end+1#同name的部件的结束行
#                rows[1]=row_end #找到同样的部件名，更新结束行
#        if len(rows)==2:
#            break
#    return rows

def update4(sheet1,sheet2,sheet3):#xlwings:更新4.0信息
    row_rev=sheet_total_rows(sheet3)+1
    for i in range(1,sheet_total_rows(sheet2)+1): #在此行数范围内去匹配需要修改的信息
        if sheet2[f'h{i}'].value=="A": #判断H列是否为A，A为新增
            data=copy_line(sheet2,i)#复制对应行的数据
            print('add:',data)
            for j in range(1,sheet_total_rows(sheet1)+1):#在报告的此行数范围内去匹配
                if sheet1[f'c{j}'].value==data[1]:#c列中寻找data[1]，即Name
                    row=lookdown(sheet1,'c',j)
                    while(sheet1[f'c{row+1}'].value==data[1]):#下一个如果Name相同（即同一个部件），则继续向下
                        row=row+1
                    print(row)
                    break
            insert_line(sheet1,row,list_fmt(data)) #在该行后面插入数据
            data_rpt=data
            update12(sheet3,row_rev,data_rpt,data,'A')
#            insert_line(sheet3,row_rev,update12(data_rpt,data,'A'))
#            sheet3[f'e{row_rev}'].value=update12(data_rpt,data,'A')
            row_rev=row_rev+1
        elif sheet2[f'h{i}'].value=="RF": #判断H列是否为RF，RF为修改技术参数
            data=copy_line(sheet2,i)#复制对应行的数据
#            print('revise:',data)
#            rows=row_range(sheet1,'c',data[1]) #返回对应部件相应的行数范围
            rows=row_range(sheet1,data) #返回对应部件相应的行数范围
            print(rows)
            for j in range(rows[0],rows[1]+1):#在同一个部件的行数范围内去匹配信息
                data_rpt=copy_line(sheet1,j)
#                if sheet1[f'd{j}'].value.upper()==data[2].upper() and sheet1[f'e{j}'].value.upper()==data[3].upper(): #匹配制造商与型号，当一致时，进行后面的操作
                if data_rpt[2].upper()==data[2].upper() and data_rpt[3].upper()==data[3].upper(): #匹配制造商与型号，当一致时，进行后面的操作
                    data=list_fmt(data)
                    paste_line(sheet1,j,data) #修改技术参数(technical data), 用了整行复制的方法，但是其实只是修改技术参数那一列，因为部件名称，制造商，型号都是一致的
                    print('revisie technical data from',data_rpt[4],'to',data[4])
                    update12(sheet3,row_rev,data_rpt,data,'RF')
                    row_rev=row_rev+1
        elif sheet2[f'h{i}'].value=="RE": #判断H列是否为RE，RE为修改型号
            data=copy_line(sheet2,i)#复制对应行的数据
#            print('revise:',data)
#            rows=row_range(sheet1,'c',data[1]) #返回对应部件相应的行数范围
            rows=row_range(sheet1,data) #返回对应部件相应的行数范围
            print(rows)
            for j in range(rows[0],rows[1]+1):#在同一个部件的行数范围内去匹配信息
                data_rpt=copy_line(sheet1,j)
                if data_rpt[2].upper()==data[2].upper() and data_rpt[4].upper()==data[4].upper(): #匹配制造商与技术参数，当一致时，进行后面的操作
#                if sheet1[f'd{j}'].value==data[1] and sheet1[f'f{j}'].value==data[3]: #匹配制造商与技术参数，当一致时，进行后面的操作
                    data=list_fmt(data)
                    paste_line(sheet1,j,data) #修改型号(model), 用了整行复制的方法，但是其实只是修改型号那一列，因为部件名称，制造商，技术参数都是一致的
                    print('revise model:',data[3])
                    update12(sheet3,row_rev,data_rpt,data,'RE')
                    row_rev=row_rev+1
        elif sheet2[f'h{i}'].value=="RD": #判断H列是否为RD，RD为修改制造商
            data=copy_line(sheet2,i)#复制对应行的数据
            print('revise:',data)
#            rows=row_range(sheet1,'c',data[1]) #返回对应部件相应的行数范围
            rows=row_range(sheet1,data) #返回对应部件相应的行数范围
            print(rows)
            for j in range(rows[0],rows[1]+1):#在同一个部件的行数范围内去匹配信息
                data_rpt=copy_line(sheet1,j)
#                print(sheet1[f'e{j}'].value==data[2])
#                print(sheet1[f'f{j}'].value)
#                print(data_rpt[3])
                if sheet1[f'e{j}'].value==data[3] and data_rpt[4]==data[4]: #匹配型号与技术参数，当一致时，进行后面的操作
#                if sheet1[f'e{j}'].value==data[2] and sheet1[f'f{j}'].value==data[3]: #匹配型号与技术参数，当一致时，进行后面的操作
                    paste_line(sheet1,j,data) #修改制造商(manufacturer), 用了整行复制的方法，但是其实只是修改制造商那一列，因为部件名称，型号，技术参数都是一致的
                    print('revise manufacturer:',data[2])
                    update12(sheet3,row_rev,data_rpt,data,'RD')
                    row_rev=row_rev+1

    fmt(sheet1)

def update12(sheet12,row,data_rpt,data,cmd):#xlwing:把对应修改信息写入12.0
    if cmd=="RD":#修改制造商
        sentence="Revise the manufacturer of "+data_rpt[1].lower().split('\n')[0]+" "+str(data_rpt[3])+' \nfrom\n\"'+data_rpt[2].split('\n')[0]+'\"\nto\n\"'+data[2].split('\n')[0]+'\".'
    elif cmd=="RE":#修改型号
        sentence='Revise the model name of '+data_rpt[1].lower().split('\n')[0]+" by "+data_rpt[2].split('\n')[0]+'\nfrom\n\"'+str(data_rpt[3])+'\"\nto\n\"'+str(data[3])+'\".'
    elif cmd=="RF":#修改技术参数
        sentence="Revise the technical data of "+data_rpt[1].lower().split('\n')[0]+" "+str(data_rpt[3])+" by "+data_rpt[2].split('\n')[0]+"\nfrom\n\""+data_rpt[4]+"\"\nto\n\""+data[4]+"\"."
    elif cmd=="A":
        sentence='Add alternative '+data[1].lower().split('\n')[0]+' '+str(data[3])+' by '+data[2].split('\n')[0]
    sheet12[f'c{row}'].value='4'
    sheet12[f'd{row}'].value=data_rpt[0]
    sheet12[f'e{row}'].value=sentence
    sheet12[f'c{row}'].api.Font.Color=0xFF00FF
    sheet12[f'd{row}'].api.Font.Color=0xFF00FF
    sheet12[f'e{row}'].api.Font.Color=0xFF00FF
#    wb.save('output1.xls')
        
#def sheet_total_rows(sheet): #返回sheet最大的行数,此方法在连续的时候有效，当有合并单元格的时候就会出现问题
#    rng1=sheet.range('a1').expand('table')
#    rng2=sheet.range('c1').expand('table')
#    rng3=sheet.range('d1').expand('table')
#    return max(rng1.rows.count,rng2.rows.count,rng3.rows.count)

def sheet_total_rows(sheet): #xlwings:返回工作簿的最大行数
    i=0
    empty=[] 
    while i<=6: #这个循环就是构造一个空数列，7个None
        empty.append(None)
        i=i+1
    row=1
    while sheet.range(f'a{row}:g{row}').value!=empty:#判断每一行是否为空数列，直到找到空的对应行数
        row_total=row
        row=row+1
    return row_total

def empty(number):#返回指定数量的空列表，列表值为None
    i=1
    empty=[] 
    while i<=number: #这个循环就是构造一个空数列，x个None
        empty.append(None)
        i=i+1
    return empty



def get_col_list(sheet,col,row_start,row_end): #xlwings:获取指定列的文本信息
    col_values=[]
    for i in sheet[f'{col}{row_start}:{col}{row_end}'].value:
        if i not in col_values:
            if i=='Name':
                pass
            elif i=='Manufacturer/ trademark2':
                pass
            elif i=='Type / model2':
                pass
            elif i=='Technical data and securement means':
                pass
            elif i==None:
                pass
            elif i=='':
                pass
            else:
                col_values.append(i)
    return col_values
    
    

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
