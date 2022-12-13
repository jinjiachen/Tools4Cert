#/bin/python #-*-coding:utf-8-*-

'''
Author: Michael Jin
Date: 2022-04

'''

import xlrd
import xlwt
#import xlutils
from xlutils.copy import copy
import xlwings as xw
import time
import os
import re
from cert import ul_search
from cert import basic_info
from cert import certificate
from cert import filters
#import warnings

def Menu():
    choice=input("1.Extract data\n2.Revise the report\n3.在7.0中自动插入说明书(for GT only)\n4.更新CDR\n5.更新8.0测试总结\n6.提取5.0数据并打印（调试用功能）\n7.在3.0中插入照片\n8针对SEC4&5自动分页功能tmp\n9对sec4.0进行排序\n10同步修改item号\n11.Sec3 sort item\n12自动填充5.0\n13自动核对证书")
    if choice=='1':
        path_rpt=input("Please input the report path:")
        path_data=input("Please input the data source path:")
        path_rpt=path_rpt.replace('"','')
        path_data=path_data.replace('"','')
        data_start=int(input("Please input the start line of data:"))
        data_end=int(input("Please input the end line of data:"))
        col1=input("Please choose four columns of data (1/4):")
        col2=input("Please choose four columns of data (2/4):")
        col3=input("Please choose four columns of data(3/4):")
        col4=input("Please choose four columns of data (4/4):")
        col5=input("是否有单独提供证书，如有，请指出证书号所在列")
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb_rpt=app.books.open(path_rpt)
        wb_data=app.books.open(path_data)
        for sheet in wb_data.sheets:
            print(sheet)
            if sheet.name=='4.0 Components':
                print('find',sheet.name)
                sht_data=sheet
                break
#                print(sht_data.name)
            else:
                sht_data=wb_data.sheets[0]
        data=get_data(sht_data,data_start,data_end,col1,col2,col3,col4,col5)
        generate4(wb_rpt.sheets['4.0 Components'],data)
        wb_rpt.save(path_rpt[:-4]+'_output.xls')
        wb_rpt.close()
        wb_data.close()
        app.kill()
    elif choice=='2':
#        app=xw.App(visible=True,add_book=False)
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
#        rpt=os.path.abspath(rpt)
#        rpt_dir=os.path.dirname(rpt)
#        filename=os.path.basename(rpt)
#        print(rpt_dir)
#        print(os.path.basename(rpt))
        data=input("Please input the data source path:") #输入数据源的路径
        data=data.replace('"','')
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
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        wb1.close()
#        app.quit()
        app.kill()
    elif choice=='3':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
#        rpt=os.path.abspath(rpt)
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht7=wb.sheets['7.0 Illustrations']
        manual_path=input('输入说明书的路径')
        update7(sht7,manual_path)
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        app.kill()
    elif choice=='4':
        app=xw.App(visible=False,add_book=False)
#        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("输入需要更新的报告路径:") #输入要更新的报告的路径
        rpt=rpt.replace('"','')
        template=input("输入CDR新模板的路径:") #输入模板的路径
        template=template.replace('"','')
        wb=app.books.open(rpt)
        if template=='':
            CDR=input('请选择对应的CDR类型：\n1.普通CDR\n2.CDRMM')
            if CDR=='1':
                wb_template=app.books.open(r'D:\Downloads\Tools4Cert-master\template\Certification CDR V5 Form.xls')
            elif CDR=='2':
                wb_template=app.books.open(r'D:\Downloads\Tools4Cert-master\template\Certification CDRMM V5 Form.xls')
        else:
            wb_template=app.books.open(template)
        update_CDR(wb_template,wb)
#        input('pause')#调试用
#        wb.save(rpt[:-4]+'_update.xls')#老报告保存是错误的
        wb_template.save(rpt[:-4]+'_update.xls')#新模板的报告才是需要保存的
        app.kill()
    elif choice=='5':
#        app=xw.App(visible=False,add_book=False)
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("输入需要更新的报告路径:") #输入要更新的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht8=wb.sheets['8.0 Test Summary']
        update8(sht8)
    elif choice=='6':
        app=xw.App(visible=False,add_book=False)
#        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("输入需要提取的报告路径:") #输入要更新的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        uc_all=get_UC(wb)
        for i in uc_all:
            print(i)
        wb.close()
        app.kill()
    elif choice=='7':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
        photo_path=input('输入照片所在路径')
        photo_path=photo_path+'\\'
        update3(sht3,photo_path)
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        app.kill()
    elif choice=='8':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht4=wb.sheets['4.0 Components']
        sht5=wb.sheets['5.0 CEC Comps']
        Page_break(sht4)
        Page_break(sht5)
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        app.kill()
    elif choice=='9':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht4=wb.sheets['4.0 Components']
        sort_by_item(sht4)
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        app.kill()
    elif choice=='10':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
        sht4=wb.sheets['4.0 Components']
        sync_item(sht3,sht4)
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        app.kill()
    elif choice=='11':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
#        get_shapes(sht3)
        line=get_line(sht3)#获取3.0中线的类型
        if line==None:
            print('并未捕获线的类型')
        else:
            print('捕捉到线的类型:',line)
        init_item(sht3,line)
#        init_item(sht3,'AutoShape')
        wb.save(rpt[:-4]+'_output.xls')
        wb.close()
        app.kill()
    elif choice=='12':
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        data=input("Please input the data source path:") #输入数据源的路径
        data=data.replace('"','')
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        wb_data=app.books.open(data)
        sht5_rpt=wb.sheets[4]
        for sheet in wb_data.sheets:
            print(sheet)
            if sheet.name=='5.0 CEC Comps':
                sht5_data=sheet
                break
            else:
                sht5_data=wb_data.sheets[0]
        fill_CEC(sht5_rpt,sht5_data)
    elif choice=='13':
        rpt=input("Please input the report path:") #输入要检查的报告的路径
        rpt=rpt.replace('"','')
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        sht4=wb.sheets['4.0 Components']
        check(sht4,'Yes')

    elif choice=='123':
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
        sht4=wb.sheets['4.0 Components']
        sht5=wb.sheets['5.0 CEC Comps']
        sht7=wb.sheets['7.0 Illustrations']
        sht8=wb.sheets['8.0 Test Summary']
        sht12=wb.sheets['12.0 Revisions']
        wb.save(rpt[:-4]+'_output.xls')
        while True:
            choice=input("1.Extract data\n2.Revise the report\n3.在7.0中自动插入说明书(for GT only)\n4.更新CDR\n5.更新8.0测试总结\n6.提取5.0数据并打印（调试用功能）\n7.在3.0中插入照片\n8针对SEC4&5自动分页功能tmp\n9对sec4.0进行排序\n10同步修改item号\n11.Sec3 sort item\n指令：")
            if choice=='1':
                path_data=input("Please input the data source path:")
                path_data=path_data.replace('"','')
                data_start=int(input("Please input the start line of data:"))
                data_end=int(input("Please input the end line of data:"))
                col1=input("Please choose four columns of data (1/4):")
                col2=input("Please choose four columns of data (2/4):")
                col3=input("Please choose four columns of data(3/4):")
                col4=input("Please choose four columns of data (4/4):")
                col5=input("是否有单独提供证书，如有，请指出证书号所在列")
                wb_data=app.books.open(path_data)
                for sheet in wb_data.sheets:
                    print(sheet)
                    if sheet.name=='4.0 Components':
                        print('find',sheet.name)
                        sht_data=sheet
                        break
#                        print(sht_data)
                    else:
                        sht_data=wb_data.sheets[0]
                data=get_data(sht_data,data_start,data_end,col1,col2,col3,col4,col5)
                generate4(wb.sheets['4.0 Components'],data)
            elif choice=='2':
                data=input("Please input the data source path:") #输入数据源的路径
                data=data.replace('"','')
                app=xw.App(visible=False,add_book=False)
                wb_data=app.books.open(data)
                sht4_data=wb_data.sheets['4.0 Components']
                start=time.time()
                update4(sht4,sht4_data,sht12)
                end=time.time()
                print('operating time:',end-start)
            elif choice=='3':
                manual_path=input('输入说明书的路径')
                update7(sht7,manual_path)
            elif choice=='7':
                photo_path=input('输入照片所在路径')
                photo_path=photo_path+'\\'
                update3(sht3,photo_path)
            elif choice=='8':
                Page_break(sht4)
                Page_break(sht5)
            elif choice=='9':
                sort_by_item(sht4)
            elif choice=='10':
                sync_item(sht3,sht4)
            elif choice=='11':
#                get_shapes(sht3)
                line=get_line(sht3)#获取3.0中线的类型
                if line==None:
                    print('并未捕获线的类型')
                else:
                    print('捕捉到线的类型:',line)
                init_item(sht3,line)
            elif choice=='12':
                data=input("Please input the data source path:") #输入数据源的路径
                data=data.replace('"','')
                wb_data=app.books.open(data)
                for sheet in wb_data.sheets:
                    print(sheet.name)
                    if sheet.name=='5.0 CEC Comps':
                        print('找到sec5.0')
                        sht5_data=sheet
                        break
                    else:
                        sht5_data=wb_data.sheets[0]
                fill_CEC(sht5,sht5_data)
            elif choice=='s':#用于把修改好的内容同步保存到原报告
                wb.save(rpt.replace('_output',''))
                wb.save(rpt[:-4]+'_output.xls')
            elif choice=='exit' or choice=='q':
                break
            elif choice=='r':
                wb.close()
                wb=app.books.open(rpt)
                wb.save(rpt[:-4]+'_output.xls')
                pass
            input('any key to contine!')
            os.system('cls')

#        input('any key to contine!')
        wb.save()
        wb.close()
        app.kill()

        

def get_data_old(rpt_fn,rpt_start, data_fn,data_start,data_end,data_col1,data_col2,data_col3,data_col4):
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

def get_data(sheet,row_start,row_end,column1,column2,column3,column4,column5):#xlwings:获取工作簿指定范围内的数据
    data=[]
    for row in range(row_start,row_end+1):
        rows_value=[]
        rows_value.append(sheet[f'{column1}{row}'].value)
        rows_value.append(sheet[f'{column2}{row}'].value)
        rows_value.append(sheet[f'{column3}{row}'].value)
        rows_value.append(sheet[f'{column4}{row}'].value)
        if column5=='':#没有输入控制号所在列
            pass
        elif sheet[f'{column5}{row}'].value==None:#控制号所在列是否为空
            pass
        elif rows_value[1]==None:
            pass
        else:
            result=re.search('\w\d{5,6}',sheet[f'{column5}{row}'].value)
            print(result)
            if result!=None:
                rows_value[1]=rows_value[1]+'\n('+result.group()+')'
        data.append(rows_value)
    return data


def generate4(sheet,data):#xlwings:自动写入数据，主要针对新报告时4.0数据的写入
    '''
    sheet:报告的SEC4.0
    data:需要写入的数据，一般由get_data获取
    '''
    row=sheet_total_rows(sheet)+1#用sheet_total_rows获取连续数据的最后一行，+1为后一行开始写入数据
    for data in data:
        print(f'正在第{row}行写入数据')
        sheet[f'c{row}:f{row}'].value=list_fmt(data)
        sheet[f'c{row}:f{row}'].api.Font.Color=0xFF00FF
        insert_blank_line(sheet,row)
        row=row+1

#    fmt(sheet)
    last_row=sheet.used_range.last_cell.row
    for col in ['c','d','e','f']:
#        index=row_index=get_index(sheet,col)
        if col=='c':#a,b,c三列的单元格合并是一致的，一起处理
            print(f'正在合并A，B，C列的单元格')
            index=get_index(sheet,col)
            merge_by_index(sheet,col,index)
            merge_by_index(sheet,'a',index)
            merge_by_index(sheet,'b',index)
        else:#对D,E,F列的单元格进行合并操作
            print(f'正在合并{col}列的单元格')
            index=get_index(sheet,col)
            merge_by_index(sheet,col,index)
        print(index)
            

def get_index(sheet,col):#xlwings:此函数服务于合并单元格，记录指定列非空单元格的行数
    rows=[]
    last_row=sheet.used_range.last_cell.row
#    print(last_row)
    for i in range(1,last_row):#在报告的此行数范围内去匹配
        if sheet[f'{col}{i}'].value=='Name':#过滤Name这一行
            pass
        elif sheet[f'{col}{i}'].value=='Manufacturer/ trademark2':#过滤这一行
            pass
        elif sheet[f'{col}{i}'].value=='Type / model2':#过滤这一行
            pass
        elif sheet[f'{col}{i}'].value=='Technical data and securement means':#过滤这一行
            pass
        elif sheet[f'{col}{i}'].value!=None:#指定列是否为空
            rows.append(i)#记录对应的行数
    rows.append(max(row_max(sheet,'c'),row_max(sheet,'d'),row_max(sheet,'e'),row_max(sheet,'f'))+1)#找到C,D,E,F列中最大的行数，+1是为了匹配merge_by_index函数
    return rows

def merge_by_index(sheet,col,index):#xlwings:基于get_index的索引来合并单元格
    for i in range(0,len(index)):#遍历每一个索引
        if i+1>len(index)-1:#超出索引，则退出
            break
        elif index[i+1]-index[i]>1:#比较两个索引之间是否大于1，大于1则合并
            sheet[f'{col}{index[i]}:{col}{index[i+1]-1}'].merge()

def row_max(sheet,col):#xlwings:获取某一列的最大行数
    row=sheet.used_range.last_cell.row#最大行数
    while sheet[f'{col}{row}'].value==None:
        row=row-1
    return row

    
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
#    sheet[index].api.Font.Color=0xFF00FF
#    sheet[index].api.Font.Bold=True
#    sheet[index].api.Font.Size
#    sheet[index].api.Font.Name

def insert_line(sheet,row,data): #xlwings:在指定行后插入空行并写入数据
    sheet.api.Rows(str(row+1)).Insert()
    paste_line(sheet,str(row+1),data)

def insert_blank_line(sheet,row): #xlwings:在指定行后插入空行
    sheet.api.Rows(str(row+1)).Insert()

def insert_blank_lines(sheet,row,numbers): #xlwings:基于insert_blank_line在指定行后插入多个空行
    i=1
    while i<=numbers:
        sheet.api.Rows(str(row+1)).Insert()
        i=i+1

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
#    name=get_col_list(sheet,'c',1,sheet_total_rows(sheet)) #获取C列的部件名
    name=get_col_list(sheet,'c',1,sheet.used_range.last_cell.row) #获取C列的部件名
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
    if str!=None and type(str)=='str':
#    if str!=None:
        str=str.replace('，',',')#替换中文逗号
        str=str.replace('（','(')#替换中文括号
        str=str.replace('）',')')#替换中文括号
        str=str.replace('：',':')#替换中文冒号
        str=str.replace('；',';')#替换中文分号
        str=str.replace('、',',')#替换中文顿号
    
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
    else:
        return str

def list_fmt(list):
    for i in range(1,len(list)):
        if isinstance(list[i],str)==True: #只针对字符串进行格式化操作
#            print('正在处理：',list[i])
            list[i]=str_fmt(list[i])
    return list

def row_range(sheet,data): #xlwings:查找相同name的部件的行数范围
    rows=[]
#    total_row=sheet_total_rows(sheet)+1
    total_row=sheet.used_range.last_cell.row#返回最大的行数
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


def update3(sheet,photo_path): #xlwings:在3.0自动插入照片
    row_height=12.5 #默认行高12.5pt
    last_row=sheet.used_range.last_cell.row #返回最后一行的行号
#    sheet[f'a3:j{last_row}'].clear_contents()#清除A列相关行数的内容
#    sheet[f'a3:j{last_row}'].delete()#删除对应区域的内容，格式保留
    sheet[f'a3:j{last_row}'].api.EntireRow.Delete()#删除对应区域的行数
    while sheet.pictures.count>0:#当sheet中有图片时，删除图片
        sheet.pictures[0].delete()
    number=sheet.pictures.count#当前的图片数量
    row=5
    top=row_height*row #12.5pt初始行高，5为行数
    for root,dirs,files in os.walk(photo_path,topdown=False):#遍历路径下的文件和文件夹，返回root,dirs,files的三元元组
#        files.sort(key=len) #在对文件的长度进行排序
        files.sort(key=mysort)#对文件进行排序
        for file in files:#遍历所有的文件
#            print(files)
            print(photo_path+file)
            sheet.pictures.add(photo_path+file)#插入图片
            if sheet.pictures[number].width>sheet.pictures[number].height:
                sheet.pictures[number].width=354 #单位为pt，72pt=1inch，即288/72=4inch
            else:
                sheet.pictures[number].height=288 #单位为pt，72pt=1inch，即288/72=4inch
            sheet.pictures[number].top=sheet[f'a1:a{row}'].height #用行数来定位
            sheet.pictures[number].left=50.5 #单元格默认列宽50.5
            sheet[f'a{row-2}'].value=f'Photo {number+1} - ' #插入文字描述
            sheet[f'a{row-2}'].characters[:9].font.bold=True #部分字体加粗
            row=row+28 #56行一页，28行一半中间位置
            number=number+1


def update4(sheet1,sheet2,sheet3):#xlwings:更新4.0信息
    '''
    sheet1为报告的sec4.0
    sheet2为数据报告的sec4.0
    sheet3为报告的sec12.0
    '''
    row_rev=sheet_total_rows(sheet3)+1#SEC12的行数,这里不能用used_range来代替，因为used_range会把空行包含进去，包括格式的改变
#    print(sheet2.used_range.last_cell.row)
    for i in range(1,sheet2.used_range.last_cell.row): #在此行数范围内去匹配需要修改的信息
        print(i)
        if sheet2[f'h{i}'].value=="A": #判断H列是否为A，A为新增
            data=copy_line(sheet2,i)#复制对应行的数据
            print('add:',data)
#            for j in range(1,sheet_total_rows(sheet1)+1):#在报告的此行数范围内去匹配
            for j in range(1,sheet1.used_range.last_cell.row):#在报告的此行数范围内去匹配
                if sheet1[f'c{j}'].value==data[1]:#c列中寻找data[1]，即Name
                    row=lookdown(sheet1,'c',j)
                    while(sheet1[f'c{row+1}'].value==data[1]):#下一个如果Name相同（即同一个部件），则继续向下
                        row=row+1
                    print(row)
                    break
            insert_line(sheet1,row,list_fmt(data)) #在该行后面插入数据
            sheet1[f'a{row+1}:f{row+1}'].font.color=0xFF00FF#插入数据后修改字体颜色,这里用row不用j，row才是实际的行数
            data_rpt=data
            update12(sheet3,row_rev,data_rpt,data,'A')
#            insert_line(sheet3,row_rev,update12(data_rpt,data,'A'))
#            sheet3[f'e{row_rev}'].value=update12(data_rpt,data,'A')
            row_rev=row_rev+1
        elif sheet2[f'h{i}'].value=="RF": #判断H列是否为RF，RF为修改技术参数
            data=copy_line(sheet2,i)#复制对应行的数据
            print('revise:',data)
#            rows=row_range(sheet1,'c',data[1]) #返回对应部件相应的行数范围
            rows=row_range(sheet1,data) #返回对应部件相应的行数范围
            print(rows)
            for j in range(rows[0],rows[1]+1):#在同一个部件的行数范围内去匹配信息
                data_rpt=copy_line(sheet1,j)
#                if sheet1[f'd{j}'].value.upper()==data[2].upper() and sheet1[f'e{j}'].value.upper()==data[3].upper(): #匹配制造商与型号，当一致时，进行后面的操作
                if data_rpt[2].upper()==data[2].upper() and data_rpt[3].upper()==data[3].upper(): #匹配制造商与型号，当一致时，进行后面的操作
                    data=list_fmt(data)
                    paste_line(sheet1,j,data) #修改技术参数(technical data), 用了整行复制的方法，但是其实只是修改技术参数那一列，因为部件名称，制造商，型号都是一致的
                    sheet1[f'f{j}'].font.color=0xFF00FF#插入数据后修改字体颜色
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
                    sheet1[f'e{j}'].font.color=0xFF00FF#插入数据后修改字体颜色
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
                    sheet1[f'd{j}'].font.color=0xFF00FF#插入数据后修改字体颜色
                    print('revise manufacturer:',data[2])
                    update12(sheet3,row_rev,data_rpt,data,'RD')
                    row_rev=row_rev+1
        elif sheet2[f'h{i}'].value=="D": #判断H列是否为D，D为删除
            data=copy_line(sheet2,i)#复制对应行的数据
            print('delete:',data)
            rows=row_range(sheet1,data) #返回对应部件相应的行数范围
            print(rows)
            for j in range(rows[0],rows[1]+1):#在同一个部件的行数范围内去匹配信息
                data_rpt=copy_line(sheet1,j)
#                print(data_rpt[2],data[2])
#                print(data_rpt[4],data[4])
                if data_rpt[2]==data[2] and sheet1[f'e{j}'].value==data[3] and data_rpt[4]==data[4]: #匹配部件名，制造商和型号，当一致时，进行后面的操作
                    sheet1[f'c{j}'].api.EntireRow.Delete()#删除改行
                    update12(sheet3,row_rev,data_rpt,data,'D')
                    row_rev=row_rev+1

    fmt(sheet1)

def update7(sheet,manual_path): #xlwings:在7.0自动插入说明书
    letters=[chr(i) for i in range(97,123)] #26个字母的列表
    letters=letters+['a'+chr(i) for i in range(97,123)]#在26个字母的基础上增加aa-az
    letters=letters+['b'+chr(i) for i in range(97,123)]#在52个字母的基础上增加ba-bz
    row_height=12.5 #默认行高
    last_row=sheet.used_range.last_cell.row #返回最后一行的行号
#    sheet[f'a3:j{last_row}'].clear_contents()#清除A列相关行数的内容
    sheet[f'a2:j{last_row}'].delete()#删除对应行数
    while sheet.pictures.count>0:#当sheet中有图片时，删除图片
        sheet.pictures[0].delete()
    number=sheet.pictures.count#当前的图片数量
    row=5
    top=row_height*row #12.5初始行高，5为行数
    for root,dirs,files in os.walk(manual_path,topdown=False):#遍历路径下的文件和文件夹，返回root,dirs,files的三元元组
        files.sort()#对文件进行排序
        files.sort(key=len) #在对文件的长度进行排序
        for file in files:#遍历所有的文件
#            print(files)
            print(manual_path+file)
            sheet.pictures.add(manual_path+file)#插入图片
            sheet.pictures[number].width=450
            sheet.pictures[number].top=top
            sheet[f'a{row-2}'].value=f'Illustration 2{letters[number]} - Manual - page {number+1}' #插入文字描述
            sheet[f'a{row-2}'].characters[:16].font.bold=True #部分字体加粗
            row=row+56
            top=top+row_height*56 #56行为分页的行数
            number=number+1

def update12(sheet12,row,data_rpt,data,cmd):#xlwing:把对应修改信息写入12.0
    if cmd=="RD":#修改制造商
        sentence="Revise the manufacturer of "+data_rpt[1].lower().split('\n')[0]+" "+str(data_rpt[3])+' \nfrom\n\"'+data_rpt[2].split('\n')[0]+'\"\nto\n\"'+data[2].split('\n')[0]+'\".'
    elif cmd=="RE":#修改型号
        if type(data[3])==float:
            print('transfer data[3]')
            data[3]=int(data[3])#data[3]为model列，当型号为纯数字时，转换为整型，防止12.0型号出现浮点的问题
        if type(data_rpt[3])==float:
            print('transfer data_rpt[3]')
            data_rpt[3]=int(data_rpt[3])#data_rpt[3]为报告的model列，当型号为纯数字时，转换为整型，防止12.0出现浮点问题
        sentence='Revise the model name of '+data_rpt[1].lower().split('\n')[0]+" by "+data_rpt[2].split('\n')[0]+'\nfrom\n\"'+str(data_rpt[3])+'\"\nto\n\"'+str(data[3])+'\".'
    elif cmd=="RF":#修改技术参数
        sentence="Revise the technical data of "+data_rpt[1].lower().split('\n')[0]+" "+str(data_rpt[3])+" by "+data_rpt[2].split('\n')[0]+"\nfrom\n\""+data_rpt[4]+"\"\nto\n\""+data[4]+"\"."
    elif cmd=="A":
        if type(data[3])==float:
            data[3]=int(data[3])#data[3]为model列，当型号为纯数字时，转换为整型，防止12.0型号出现浮点的问题
        sentence='Add alternative '+data[1].lower().split('\n')[0]+' '+str(data[3])+' by '+data[2].split('\n')[0]
    elif cmd=="D":
        if type(data[3])==float:
            data[3]=int(data[3])#data[3]为model列，当型号为纯数字时，转换为整型，防止12.0型号出现浮点的问题
        sentence='Delete '+data[1].lower().split('\n')[0]+' '+str(data[3])+' by '+data[2].split('\n')[0]
    sheet12[f'c{row}'].value='4'
    sheet12[f'd{row}'].value=data_rpt[0]
    sheet12[f'e{row}'].value=sentence
    sheet12[f'c{row}'].api.Font.Color=0xFF00FF
    sheet12[f'd{row}'].api.Font.Color=0xFF00FF
    sheet12[f'e{row}'].api.Font.Color=0xFF00FF
#    wb.save('output1.xls')

def update_CDR(workbook,workbook_data):
    '''
    workbook为CDR的模板
    workbook_data为需要更新的报告
    '''
    sheets_name=get_sheets_name(workbook)#获取工作簿中的表名
    workbook.sheets.add('tmp')#增加一个临时的sheet
    for sheet_name in sheets_name:
        if sheet_name=='ATM':
            pass
        elif sheet_name=='Instructions':
            pass
#        elif sheet_name=='1.0 Reference':
#            pass
        else:
            workbook.sheets[sheet_name].delete()
    for i in range(0,12):#1.0到12.0的索引
        workbook_data.sheets[i].copy(after=workbook.sheets[i])#复制1.0到12.0的工作表
#        input('')#调试用
    workbook.sheets['tmp'].delete()#删除临时的sheet

    #以下对一些外部链接做处理
    workbook.sheets['9.0 MLS']['b3'].value='=\'1.0 Reference\'!$B$6'
    workbook.sheets['9.0 MLS']['b4'].value='=\'1.0 Reference\'!$B$7'
    workbook.sheets['9.0 MLS']['b5'].value='=\'1.0 Reference\'!$B$8'
    workbook.sheets['9.0 MLS']['b6'].value='=\'2.0 Description\'!$B$3'
    workbook.sheets['10.0 General']['a36'].value='=Instructions!$P$2'
    workbook.sheets['10.0 General']['a38'].value='=Instructions!$Q$2'
    workbook.sheets['10.0 General']['a39'].value='=Instructions!$R$2'
    workbook.sheets['10.0 General']['a40'].value='=IF(Instructions!$S$2 >"",Instructions!$S$2,"")'
    for row in range(1,workbook.sheets['5.0 CEC Comps'].used_range.last_cell.row):
        if workbook.sheets['5.0 CEC Comps'].range(f'a{row}').value=='Photo #':
            workbook.sheets['5.0 CEC Comps'].range(f'a{row+1}').value='=\''+workbook.sheets['5.0 CEC Comps'].range(f'a{row+1}').formula.split(']')[1]
            workbook.sheets['5.0 CEC Comps'].range(f'b{row+1}').value='=\''+workbook.sheets['5.0 CEC Comps'].range(f'b{row+1}').formula.split(']')[1]
            workbook.sheets['5.0 CEC Comps'].range(f'c{row+1}').value='=\''+workbook.sheets['5.0 CEC Comps'].range(f'c{row+1}').formula.split(']')[1]
            workbook.sheets['5.0 CEC Comps'].range(f'f{row+1}').value='=\''+workbook.sheets['5.0 CEC Comps'].range(f'f{row+1}').formula.split(']')[1]
            workbook.sheets['5.0 CEC Comps'].range(f'i{row+1}').value='=\''+workbook.sheets['5.0 CEC Comps'].range(f'i{row+1}').formula.split(']')[1]
    
def update8(sheet): #xlwings:写入测试总结
    std_ul60335_2_40={
    'Test Description':'UL 60335-1: 2016 Ed. 6\nCSA C22.2#60335-1: 2016 Ed. 2\nUL 60335-2-40:2019 Ed.3\nCSA C22.2#60335-2-40: 2019 Ed. 3\nClause',
    '10':'Power input and current',
    '11':'Heating',
    '13':'Leakage current and electric strength at operating temperature',
    '15':'Moisture resistance',
    '16':'Leakage current and electric strength',
    '17':'Overload protection of transformers and associated circuits',
    '19':'Abnormal operation',
    '20':'Stability and mechanical hazards',
    '21':'Mechanical strength',
    '22':'Construction',
    '23':'Internal wiring',
    '24':'Components',
    '25':'Supply connection and external flexible cords',
    '26':'Terminals for external conductors',
    '27':'Provision for earthing',
    '28':'Screws and connections',
    '30':'Resistance to heat and fire',
    '31':'Resistance to rusting',
}
    
    row=95#固定锚点
    insert_blank_lines(sheet,row,5+len(std_ul60335_2_40)+1)#固定行数5，标准的相应测试章节的行数，+1为空行

    #以下为一些固定内容的填写
    sheet[f'b{row+1}:d{row+1}'].merge()
    sheet[f'b{row+3}:f{row+3}'].merge()
    sheet[f'b{row+4}:f{row+4}'].merge()
    sheet[f'a{row+5}:f{row+5}'].merge()
    sheet[f'a{row+6}:f{row+6}'].merge()
    sheet[f'a{row+1}'].value='Evaluation Period'
    sheet[f'a{row+1}'].color=(192,192,192)
    sheet[f'e{row+1}'].value='Project No.'
    sheet[f'e{row+1}'].color=(192,192,192)
    sheet[f'a{row+2}'].value='Sample Rec. Date'
    sheet[f'a{row+2}'].color=(192,192,192)
    sheet[f'c{row+2}'].value='Condition'
    sheet[f'c{row+2}'].color=(192,192,192)
    sheet[f'e{row+2}'].value='Sample ID.'
    sheet[f'e{row+2}'].color=(192,192,192)
    sheet[f'a{row+3}'].value='Test Location'
    sheet[f'a{row+3}'].color=(192,192,192)
    sheet[f'a{row+4}'].value='Test Procedure'
    sheet[f'a{row+4}'].color=(192,192,192)
    sheet[f'a{row+5}'].value='Determination of the result includes consideration of measurement uncertainty from the test equipment and methods.  The product was tested as indicated below with results in conformance to the relevant test criteria.'
    sheet[f'a{row+6}'].value='The following tests were performed: '
    for i in range(1,len(std_ul60335_2_40)):
        sheet[f'a{row+6+i}:c{row+6+i}'].merge()
        sheet[f'd{row+6+i}:f{row+6+i}'].merge()
        sheet[f'd{row+6+i}:f{row+6+i}'].api.HorizontalAlignment=-4108
        if i==1:
            sheet[f'a{row+6+i}'].value=list(std_ul60335_2_40.keys())[i-1]
            sheet[f'd{row+6+i}'].value=list(std_ul60335_2_40.values())[i-1]
            sheet[f'a{row+6+i}'].color=(192,192,192)
            sheet[f'd{row+6+i}'].color=(192,192,192)
        else:
            sheet[f'd{row+6+i}'].value=list(std_ul60335_2_40.keys())[i-1]
            sheet[f'a{row+6+i}'].value=list(std_ul60335_2_40.values())[i-1]


def get_sheets_name(workbook): #获取工作簿中所有的表名
    sheets_name=[]
    for i in workbook.sheets:
        sheets_name.append(i.name)
    return sheets_name
    
        
#def sheet_total_rows(sheet): #返回sheet最大的行数,此方法在连续的时候有效，当有合并单元格的时候就会出现问题
#    rng1=sheet.range('a1').expand('table')
#    rng2=sheet.range('c1').expand('table')
#    rng3=sheet.range('d1').expand('table')
#    return max(rng1.rows.count,rng2.rows.count,rng3.rows.count)

def sheet_total_rows(sheet): #xlwings:返回工作簿的最大行数,当整行都是合并单元格的时候，则会返回7个None的列表，类似空行，此时会返回错误行数,尝试用used_range函数来替换
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
    
def get_UC(wb):#xlwings: 获取5.0相关信息
    sht1=wb.sheets['1.0 Reference']
    sht5=wb.sheets['5.0 CEC Comps']
    total_row=sht5.used_range.last_cell.row#返回最大的行数
    uc_all=[]
    basic_info={
        'report':sht1['b3'].value
}
    for i in range(1,total_row):#在报告的此行数范围内去匹配
        if sht5[f'a{i}'].value=='Photo #':#a列中寻找Photo
            uc_info={
                'photo_no':sht5[f'a{i+1}'].value,
                'item_no':sht5[f'b{i+1}'].value,
                'name':sht5[f'c{i+1}'].value,
                'manufacturer':sht5[f'f{i+1}'].value,
                'model':sht5[f'i{i+1}'].value,
                'rating':sht5[f'c{i+2}'].value
}
        if sht5[f'a{i}'].value=='WINDING(S) RESISTANCE':#a列中寻找winding
            if 'Compressor'.lower() in uc_info['name'].lower() and 'Hz' in uc_info['rating']: #交流压缩机
                j=3
                while sht5[f'a{i+j}'].value!='VERIFICATION PROCESS':
                    uc_info[f'designation_{j-2}']=sht5[f'a{i+j}'].value
                    uc_info[f'wire_size_{j-2}']=sht5[f'c{i+j}'].value
                    uc_info[f'resistance_{j-2}']=sht5[f'j{i+j}'].value
                    j=j+1
            elif 'motor' in uc_info['name'].lower() and 'Hz' in uc_info['rating']: #交流电机
                j=3
                while sht5[f'a{i+j}'].value!='VERIFICATION PROCESS':
                    uc_info[f'designation_{j-2}']=sht5[f'a{i+j}'].value
                    uc_info[f'wire_size_{j-2}']=sht5[f'c{i+j}'].value
                    uc_info[f'resistance_{j-2}']=sht5[f'j{i+j}'].value
                    j=j+1
            elif 'compressor' in uc_info['name'].lower() and 'dc' in uc_info['rating'].lower(): #DC压缩机
                j=3
                while sht5[f'a{i+j}'].value!='VERIFICATION PROCESS':
                    uc_info[f'designation_{j-2}']=sht5[f'a{i+j}'].value
                    uc_info[f'wire_size_{j-2}']=sht5[f'c{i+j}'].value
                    uc_info[f'resistance_{j-2}']=sht5[f'j{i+j}'].value
                    j=j+1
            elif 'motor' in uc_info['name'].lower() and 'dc' in uc_info['rating'].lower(): #DC电机
                j=3
                while sht5[f'a{i+j}'].value!='VERIFICATION PROCESS':
                    uc_info[f'designation_{j-2}']=sht5[f'a{i+j}'].value
                    uc_info[f'wire_size_{j-2}']=sht5[f'c{i+j}'].value
                    uc_info[f'resistance_{j-2}']=sht5[f'j{i+j}'].value
                    j=j+1
            elif 'power unit'.lower() in uc_info['name'].lower() or 'pwb'.lower() in uc_info['name'].lower() or 'SMPS'.lower() in uc_info['name'].lower() or 'tranformer' in uc_info['name'].lower(): #开关电源
                j=3#WINDING后面两行是格式，跳开
                while sht5[f'a{i+j}'].value!='VERIFICATION PROCESS':#找到VERIFICATION PROCESS这一行，行数-3就是实际的绕组数量
                    uc_info[f'designation_{j-2}']=sht5[f'a{i+j}'].value
                    uc_info[f'wire_size_{j-2}']=sht5[f'c{i+j}'].value
                    uc_info[f'resistance_{j-2}']=sht5[f'j{i+j}'].value
                    j=j+1

#                j=2#WINDING后面两行是格式，跳开
#                while sht5[f'a{i+j}'].value!='VERIFICATION PROCESS':#找到VERIFICATION PROCESS这一行，行数-3就是实际的绕组数量
#                    j=j+1
#                for k in range(1,j-3+1):#j-3为绕组数量，+1是因为range不包含上限
#                    uc_info[f'winding_{k}']=sht5[f'c{i+2+k}'].value
#                    uc_info[f'resistance_{k}']=sht5[f'j{i+2+k}'].value

        if sht5[f'a{i}'].value=='Dielectric Strength':#a列中寻找dielectric strength这一行
            j=1
            while sht5[f'd{i+j}'].value!=None:
                uc_info[f'location_{j}']=sht5[f'd{i+j}'].value
                uc_info[f'voltage_{j}']=sht5[f'h{i+j}'].value
                j=j+1


            uc_all.append(uc_info)
    return {'uc_info':uc_all,'basic_info':basic_info}
    

def Page_break(sheet):#xlwings:自动分页功能
    last_row=sheet.used_range.last_cell.row#工作簿最大的行数
    if sheet.name=='4.0 Components':
        print('正在对sec4进行分页！')
        start=1
        end=1
        while end<=last_row:#在最大行数范围内进行分页
            while sheet[f'a{start}:a{end}'].height<=650:#650为分页的最大行高，超出此行高则分页
                end=end+1#一行行增加，直到范围内最大的行数
                mark=end#记录该行位置
                print(f'mark:{mark}')
            while sheet[f'a{end}'].value==None:#如合并单元格，则不应该在中间分页，往上寻找直到找到合适的分页处
                print(f'end:{end}')
                if end>last_row:#是否超出最大行数，超出则不需要再分页，退出
                    break
                else:
                    end=end-1#如果在最大行数范围内，则往上寻找合适的单元格分页
                    if sheet[f'a{start}:a{end}'].height<=550:#当分页过小时，则不再向上寻找合适单元格
                        end=mark#回到当初记录的位置,此时在此位置进行分页时可行的，但是比较粗犷，下面再进一步优化，在D列制造商出寻找合适位置分页
                        while sheet[f'd{end}'].value==None:#如合并单元格，则不应该在中间分页，往上寻找直到找到合适的分页处
                            end=end-1
                        break
            sheet.api.HPageBreaks.Add(Before=sheet[f'a{end}'].api)#在上方添加分页符
            print(f'在{end}行上方分页')
            start=end#添加分页后的行数为后一页起点

    elif sheet.name=='5.0 CEC Comps':
        print('正在对sec5进行分页！')
        counts=0
        rows=[]
        for i in range(1,last_row):#在报告的此行数范围内去匹配
            if sheet[f'a{i}'].value=='INSULATED COIL ':#a列中寻找INSULATED COIL 
                counts=counts+1#计算找到多少个insulated coil
                rows.append(i)#记录对应的行数
        rows.append(last_row+1)#加入最后一行，把最后的一段考虑进去,+1是由于end取值时做-1处理
        print(rows)
        for i in range(0,len(rows)):
            if i+1>=len(rows):
                break
            elif i==0:
                start=1
                end=rows[i+1]-1
            elif i!=0:
                start=rows[i]-1
                end=rows[i+1]-1
            print(f'正在处理{start}-{end}之间的分页')

            if sheet[f'a{start}:a{end}'].height<=680:
                sheet.api.HPageBreaks.Add(Before=sheet[f'a{end}'].api)#在上方添加分页符
                print(f'无需再分割，在{end}上方分页')

            else:
                scan=start
                while sheet[f'a{start}:a{scan}'].height<=680 and scan<=end:#680为分页的最大行高，超出此行高则分页
                    scan=scan+1
                    if sheet[f'a{scan}'].value=='WINDING(S) RESISTANCE':
                        break
                while sheet[f'a{scan}'].value==None:#如合并单元格，则不应该在中间分页，往上寻找直到找到合适的分页处
#                    if scan>end:#是否超出最大行数，超出则不需要再分页，退出
#                        break
#                    else:
#                        scan=scan-1#如果在最大行数范围内，则往上寻找合适的单元格分页
                    if scan==last_row:
                        break
                    else:
                        scan=scan-1
                sheet.api.HPageBreaks.Add(Before=sheet[f'a{scan}'].api)#在上方添加分页符
                print(start,scan,sheet[f'a{start}:a{scan}'].height)
                print(f'需要分割，在{scan}上方分页')
                start=scan
                if end==last_row:
                    sheet.api.HPageBreaks.Add(Before=sheet[f'a{end+1}'].api)#在上方添加分页符
                else:
                    sheet.api.HPageBreaks.Add(Before=sheet[f'a{end}'].api)#在上方添加分页符

            print('='*10)


def cell_unmerge(sheet):#xlwings:拆分单元格并填充相同数据
    last_row=max(row_max(sheet,'c'),row_max(sheet,'d'),row_max(sheet,'e'),row_max(sheet,'f'))#找到C,D,E,F列中最大的行数，排除最下方notes部分
    for row in range(3,last_row+1):#遍历除了固定格式外的所有行
        for column in ['a','b','c','d','e','f']:
            if sheet[f'{column}{row}'].merge_cells:#是否为合并单元格
                address=sheet[f'{column}{row}'].merge_area.address#获取合并单元格的范围
                sheet[address].unmerge()#拆分单元格
                sheet[address].value=sheet[address].value[0]#拆分后赋予相同的值


def sort_by_item(sheet):#xlwings:按照item进行排序,提取item的序列，按照序列从大到小在B列查找，从最后一行倒序查找，增加效率
    last_row=max(row_max(sheet,'c'),row_max(sheet,'d'),row_max(sheet,'e'),row_max(sheet,'f'))#找到C,D,E,F列中最大的行数，排除最下方notes部分
    items=get_col_list(sheet,'b',3,last_row)#获取item的所有编号
    items=sorted(items,reverse=True)#由大到小排序
#    print(items)
    for item_no in items:#遍历每一个item序号
        row=last_row#从下往上遍历
        while sheet[f'b{row}'].value!=item_no:#从下往上找，知道找到对应的序号
            row=row-1#找不到则行号-1
        if sheet[f'b{row}'].merge_cells:#找到对应item后判断对应的单元格是否为合并单元格，如是则对应区域一起剪切
            address=sheet[f'b{row}'].merge_area.address#获取合并单元格对应的区域范围
            sheet[address].api.EntireRow.Cut()#剪切该区域的完整行
        else:
            sheet[f'b{row}'].api.EntireRow.Cut()#不是合并单元格，直接剪切
        sheet['a3'].api.EntireRow.Insert()#在第三行上方插入剪切的数据

def sync_item(sheet_photo,sheet_components):#xlwings:同步修改后的item号
    '''
    sheet_photo为报告的sec3.0
    sheet_components为报告的sec4.0
    '''
    rows=[]
    for i in range(1,sheet_components.used_range.last_cell.row): #在此行数范围内去匹配需要修改的信息
        rows.append(i)#将范围存储为列表
#    print(rows)
#    print(sorted(rows,reverse=True))
    scan_direction=input('请选择扫描方向（up/down)：up为从下往上扫，down为从上往下扫')
    if scan_direction=='up':
        rows=sorted(rows,reverse=True)
    elif scan_direction=='down':
        rows=sorted(rows,reverse=False)
    print(rows)
    for i in rows:#将列表倒序遍历，顺序遍历会有重复item号问题
#        print(i)
        if sheet_components[f'h{i}'].value==None:#H列为空则pass
            pass
        elif sheet_components[f'b{i}'].value==None:#item列如果有合并单元格，改行不是单元格首行，则pass
            pass
        elif '+' in sheet_components[f'h{i}'].value or '-' in sheet_components[f'h{i}'].value: #判断H列是否有+-符号，有则为需要修改的item
            print(f'正在处理第{i}行的item号')
            old_no=sheet_components[f'b{i}'].value#记录修改前的item号
            new_no=old_no+int(sheet_components[f'h{i}'].value)#计算需要更改后的item号
            sheet_components[f'b{i}'].value=new_no#将item号更新
            change_photo_no(sheet_photo,old_no,new_no,'Line')#同步更新3.0中的序号,默认用line作为关键词去匹配，后期可能需要优化
#            change_photo_no(sheet_photo,old_no,new_no,'AutoShape')#同步更新3.0中的序号,默认用line作为关键词去匹配，后期可能需要优化


def get_shapes(sheet):#xlwings:获取sheet中所有的shape对象
    for shape in sheet.shapes:
        if shape.text==None:
            print(shape.name)
        else:
            print(shape.name+':'+shape.text)

def get_line(sheet):#xlwings:获取3.0中指示线的类型
    for shape in sheet.shapes:
        if 'Line' in shape.name:
            print('get shape.name:',shape.name)
            return shape.name.split(' ')[0] #捕捉到的是其中一个线的具体型号,返回一部分关键字
            break
        elif 'AutoShape' in shape.name:
            print('get shape.name:',shape.name)
            return shape.name.split(' ')[0]#捕捉到的是其中一个线的具体型号,返回一部分关键字
            break


def init_item(sheet,shape_name):#xlwings:对sec3中的item号进行排序
    value=1
    shapes_wanted=[]
#    for shape in sheet.shapes:
#        if shape_name in shape.name:
#            shapes_wanted.append(shape)
#    print(shapes_wanted.sort(key=shape_top))
#    print(shapes_wanted[0])

    for shape in sheet.shapes:
        if shape_name in shape.name:
            shape.text=value
            print(shape.name+f':{value}')
            value+=1

def shape_top(shape):
    print('shape top:',shape.top)
    return str(shape.top)


def change_photo_no(sheet,old_no,new_no,shape_name):#xlwings:更改sec3.0中部件的索引
    '''
    sheet:报告中sec3.0
    old_no:照片中需要更改的索引
    new_no:照片中的新索引
    '''
    for shape in sheet.shapes:#遍历所有的shape对象
        if shape_name in shape.name:#判断是否为部件索引对应的图形
            if shape.text==str(int(old_no)):#找到需要更改的索引
                shape.text=new_no#赋予新的索引值

    
def fill_CEC(sheet_rpt,sheet_data):#xlwings:自动填充5.0信息
#    model_rpt=
#    total_row=sheet_data.used_range.last_cell.row#返回最大的行数
#    for row in range(1,total_row):#在数据的此行数范围内去匹配
#        if sheet_data[f'i{row}'].value==model_rpt:
#            i=row+4
#            while sheet_data[f'a{i}']!='WINDING(S) RESISTANCE':
#                data=copy_line(sheet_data,i)


    for row in range(1,sheet_data.used_range.last_cell.row): #在此行数范围内去匹配需要修改的信息
        print(row)
        if sheet_data[f'l{row}'].value=="A": #判断L列是否为A，A为新增
            manufacturer=sheet_data[f'f{row}'].value
            model=sheet_data[f'i{row}'].value
            start=row+5#数据的起始行数
            row_scan=start
            print(manufacturer)
            print(model)
            while sheet_data[f'a{row_scan}'].value!='WINDING(S) RESISTANCE':
                row_scan+=1
                print(sheet_data[f'a{row_scan}'].value)
            end=row_scan#数据的终止行数
            data=sheet_data[f'a{start}:k{end}'].value#复制范围内的值
            print(f'复制{start}:{end}行的数据',data)
            for row_rpt in range(1,sheet_rpt.used_range.last_cell.row):#在报告的此行数范围内去匹配
#                    print(string_strip(sheet_rpt[f'f{row_rpt}'].value))
#                print(string_strip(manufacturer))
                if string_strip(sheet_rpt[f'f{row_rpt}'].value)==string_strip(manufacturer) and sheet_rpt[f'i{row_rpt}'].value==model:#如果制造商和型号都相同，则认定为找到对应的部件
                    row_insert=row_rpt+4
                    sheet_rpt[f'a{row_rpt+3}'].value=sheet_rpt[f'a{row_rpt+3}'].value.replace(' (refer to illustration _ for assembly drawing) ','')
                    insert_blank_lines(sheet_rpt,row_insert,len(data))#在指定行下方插入对应数据的空白行
                    for i in range(start,end):#遍历数据段的行数
                        row_insert+=1
                        #以下合并单元格调整格式用
                        sheet_rpt[f'a{row_insert}:b{row_insert}'].merge()
                        sheet_rpt[f'c{row_insert}:d{row_insert}'].merge()
                        sheet_rpt[f'e{row_insert}:f{row_insert}'].merge()
                        sheet_rpt[f'g{row_insert}:k{row_insert}'].merge()
                        print(f"在第{row_insert}行写入数据：sheet_data[f'a{i}'].value")
                        #因为合并单元格，所以只对a,c,e,g列的单元格进行赋值即可
                        sheet_rpt[f'a{row_insert}'].value=str_fmt(sheet_data[f'a{i}'].value)#把a列的数据赋值
                        if sheet_data[f'g{i}'].value!=None:
                            ul_no=re.search('\w\d{5,6}',sheet_data[f'g{i}'].value)
                            if ul_no!=None:
                                sheet_rpt[f'c{row_insert}'].value=str_fmt(sheet_data[f'c{i}'].value)+'\n('+ul_no.group()+')'
                                sheet_rpt[f'g{row_insert}'].value=str_fmt(sheet_data[f'g{i}'].value.replace(ul_no.group(),''))
                            else:
                                sheet_rpt[f'c{row_insert}'].value=str_fmt(sheet_data[f'c{i}'].value)
                                sheet_rpt[f'g{row_insert}'].value=str_fmt(sheet_data[f'g{i}'].value)
                        else:
                            sheet_rpt[f'c{row_insert}'].value=str_fmt(sheet_data[f'c{i}'].value)
                            sheet_rpt[f'g{row_insert}'].value=str_fmt(sheet_data[f'g{i}'].value)
                        sheet_rpt[f'e{row_insert}'].value=str_fmt(sheet_data[f'e{i}'].value)
                        sheet_rpt[f'a{row_insert}:k{row_insert}'].font.color=0xFF00FF#对新增的数据颜色区分

def mysort(filename):#xlwings:自定义排序函数
    print(filename.split('_')[0])#提取文件名前面的数字
    return int(filename.split('_')[0])#转换为数字来排序，如果是字符串排序，会出现问题

def string_strip(string):#只保留字符串中的字母
    if string!=None and type(string)=='str':
        string=string.replace(' ','')#替换空格
        string=string.replace(',','')#替换逗号
        string=string.replace('.','')#替换句号
        string=string.replace('-','')#替换连接符
        string=string.replace('_','')#替换下划线
        string=string.replace('，','')#替换中文逗号
        return string.upper()

def check(sheet,ptf='No'):#xlwings:检查报告证书的正确性
    '''
    sheet: SEC4.0
    '''
    for row in range(3,sheet.used_range.last_cell.row): #在此行数范围内去匹配需要修改的信息
        print(f'正在核对第{row}行')
        manufacturer=sheet[f'd{row}'].value
        scan=row#扫描的行数
        while manufacturer==None:#当有合并单元格时，向上扫描，获取制造商信息
            scan=scan-1
            manufacturer=sheet[f'd{scan}'].value
        ul_no=re.search('\w\d{5,6}',manufacturer)#提取黄卡号
        model=str(sheet[f'e{row}'].value)#转化为字符，针对纯数字问题
        mark=sheet[f'g{row}'].value
        if mark=='NR':
            continue
        elif mark=='See 5.0':
            continue
        elif mark==None:
            continue
        elif ul_no==None:
            continue
        else:
            url='https://iq.ulprospector.com/en/_/_results?p=10005,10048,10006,10047&qm=q:'+ul_no.group()
            if ptf=='Yes':
                print(url)
            selector_basic=ul_search(url)#用get方法提交搜索请求，返回搜索结果的response
            links=basic_info(selector_basic)#输出查询的结果并返回详细连接
            if ptf=='Yes':
                print(links)
            if len(links)==0:
                print('invalid cert')
                continue
            else:
                selector_details=ul_search('https://iq.ulprospector.com'+links[0])#此处暂时只对一个链接做处理，后续优化
                models=certificate(selector_details)
                if filters(models,model)=='green':
                    sheet[f'h{row}'].value='ok'
                elif filters(models,model)=='yellow':
                    sheet[f'h{row}'].value='to be check'

    

if __name__=='__main__':
    Menu()
