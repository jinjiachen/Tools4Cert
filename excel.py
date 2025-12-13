#/bin/python #-*-coding:utf-8-*-

'''
Author: Michael Jin
Date: 2022-04

'''

import xlrd
import xlwt
#import xlutils
from xlutils.copy import copy
import time
import os,shutil
import re
from cert import ul_search
from cert import basic_info
from cert import certificate
from cert import filters
import pdfplumber
import random
#import warnings
if os.name=='nt':
    import xlwings as xw
    import win32com.client


def Menu():
    choice=input("1.Extract data\n2.Revise the report\nip7.在7.0中自动插入说明书(for GT only)\n4.更新CDR\n5.更新8.0测试总结\n6.提取5.0数据并打印（调试用功能）\nip3.在3.0中插入照片\n8针对SEC4&5自动分页功能tmp\n9对sec4.0进行排序\nsi同步修改item号\n11.Sec3 sort item\n12自动填充5.0\ncc自动核对证书\naml.增加多重列名\n15.增加基本列名\ntc(to client):生成客户用CDR\ncp3(clear pictures sec.3):清除3.0中的图片\ncp7(clear picture sec.7)清除7.0中图片\nmi(ML info):尝试提取ML信息\ncc5:检查CEC的证书\nef(E-filing) 创建E-filing文件夹模板")
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
#        wb_rpt.save(path_rpt[:-4]+'_output.xls')
        wb_rpt.save(path_rpt[:-5]+'_output.xlsm')
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
        auto_fmt=input('是否需要格式化:')
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
        update4(sh,sh1,sh12,auto_fmt)
        end=time.time()
        print(f'operating time {round(end-start)}s:',)
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        wb1.close()
#        app.quit()
        app.kill()
    elif choice=='ip7':
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
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
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
#                wb_template=app.books.open(r'D:\Downloads\Tools4Cert-master\template\Certification CDR V5 Form.xls')
                wb_template=app.books.open(r'D:\Downloads\Tools4Cert-master\template\Certification CDR V5 Form.xlsm')
            elif CDR=='2':
                wb_template=app.books.open(r'D:\Downloads\Tools4Cert-master\template\Certification CDRMM V5 Form.xls')
        else:
            wb_template=app.books.open(template)
        update_CDR(wb_template,wb)
#        input('pause')#调试用
#        wb.save(rpt[:-4]+'_update.xls')#老报告保存是错误的
#        wb_template.save(rpt[:-4]+'_update.xls')#新模板的报告才是需要保存的
        wb_template.save(rpt[:-5]+'_update.xlsm')
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
    elif choice=='ip3':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
        photo_path=input('输入照片所在路径')
        photo_path=photo_path+'\\'
        new=input('是否删除原有照片（Y/N）？')
        model_description=input('增加型号描述，没有直接回车')
        if new.upper()=='Y':
            update3(sht3,photo_path,start_row='',model=model_description)
        elif new.upper()=='N':
            row=input('请输入插入图片所在的行数')
            update3(sht3,photo_path,start_row=row,model=model_description)
        wb.save(rpt[:-5]+'_output.xlsm')
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
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
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
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        app.kill()
    elif choice=='si':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
        sht4=wb.sheets['4.0 Components']
        line=get_line(sht3)#获取3.0中线的类型
        if line==None:
            print('并未捕获线的类型')
        else:
            print('捕捉到线的类型:',line)
            change_line=input('是否需要手动输入线类型(Y/N):')
            if change_line=='Y':
                line=input('请输入线的类型：')
        sync_item(sht3,sht4,line)
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
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
            change_line=input('是否需要手动输入线类型(Y/N):')
            if change_line=='Y':
                line=input('请输入线的类型：')
        init_item(sht3,line)
#        init_item(sht3,'AutoShape')
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
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
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        wb_data.close()
        app.kill()
    elif choice=='cc':
        rpt=input("Please input the report path:") #输入要检查的报告的路径
        rpt=rpt.replace('"','')
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        sht4=wb.sheets['4.0 Components']
        check(sht4,'No')
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        app.kill()
    elif choice=='cc5':
        rpt=input("Please input the report path:") #输入要检查的报告的路径
        rpt=rpt.replace('"','')
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        sht5=wb.sheets['5.0 CEC Comps']
        check_CEC(sht5,'No')
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        app.kill()
    elif choice=='aml':
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        path=input("Please input the ML application path:") #输入申请表的路径
        path=path.replace('"','')
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        sht9=wb.sheets['9.0 MLS']
        sht12=wb.sheets['12.0 Revisions']
        item=get_ML_item(sht12)
        data=get_ML_info(path,'Yes')
        print(f'报告中已有多重列名ML{item}',)
        modify_ML(sht9,item,data,'A')
#        wb.save(rpt[:-4]+'_output.xls')
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        app.kill()
    elif choice=='mi':
        path=input("Please input the ML application path:") #输入申请表的路径
        path=path.replace('"','')
        data=get_ML_info(path,'Yes')
    elif choice=='15':
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        data=input("Please input the data source path:") #输入数据源的路径
        data=data.replace('"','')
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        wb=app.books.open(rpt)
        sht2=wb.sheets['2.0 Description']
        sht12=wb.sheets['12.0 Revisions']
        wb_data=app.books.open(data)
        sht_data=wb_data.sheets[0]
        add_models(sht2,sht12,sht_data)
        wb.save(rpt[:-5]+'_output.xlsm')
        wb.close()
        wb_data.close()
        app.kill()
    elif choice=='tc':
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        wb=app.books.open(rpt)
        to_client(wb)
        wb.save(rpt[:-5]+'_CDF.xlsm')
        wb.close()
        app.kill()
    elif choice=='cp3':
        clear_pics(sht3)
    elif choice=='cp7':
        clear_pics(sht7)
    elif choice=='ef':
        source_path=r'E:\!E-Filing批处理小程序相关文件\E-filling\文件夹模板\2025\25XXBXXXXSHA_Revision Service_based on 19XXXXXXXSHA_ETL'
        target_path=input('请输入项目文件夹路径:')
        if os.path.exists(target_path):#如果目标路径存在，先删除，否则shutil.copytree会报错
            shutil.rmtree(target_path)
        shutil.copytree(source_path, target_path)
    elif choice=='123':
        app=xw.App(visible=True,add_book=False)
        app.display_alerts=False #取消警告
        app.screen_updating=False#取消屏幕刷新
        rpt=input("Please input the report path:") #输入要修改的报告的路径
        rpt=rpt.replace('"','')
        output_file=output_path(rpt)
        wb=app.books.open(rpt)
        sht3=wb.sheets['3.0 Photos']
        sht4=wb.sheets['4.0 Components']
        sht5=wb.sheets['5.0 CEC Comps']
        sht7=wb.sheets['7.0 Illustrations']
        sht8=wb.sheets['8.0 Test Summary']
        sht9=wb.sheets['9.0 MLS']
        sht12=wb.sheets['12.0 Revisions']
        wb.save(output_file)
        while True:
            choice=input("1.Extract data\n2.Revise the report\nip7.在7.0中自动插入说明书(for GT only)\n4.更新CDR\n5.更新8.0测试总结\n6.提取5.0数据并打印（调试用功能）\nip3.在3.0中插入照片\n8针对SEC4&5自动分页功能tmp\n9对sec4.0进行排序\nsi同步修改item号\n11.Sec3 sort item\n12自动填充5.0\ncc自动核对证书\naml.增加多重列名\ncp3(clear pictures sec.3):清除3.0中的图片\ncp7(clear picture sec.7)清除7.0中图片\n指令：")
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
#                app=xw.App(visible=False,add_book=False)#这个好像多余，打开一个进程即可
                wb_data=app.books.open(data)
                sht4_data=wb_data.sheets['4.0 Components']
                start=time.time()
                update4(sht4,sht4_data,sht12)
                end=time.time()
                print('operating time:',end-start)
                wb_data.close()#关闭打开的文件
                print('已关闭文件：',data)
            elif choice=='ip7':
                manual_path=input('输入说明书的路径')
                update7(sht7,manual_path)
            elif choice=='6':
                uc_all=get_UC(wb)
                for i in uc_all:
                    print(i)
            elif choice=='ip3':
                photo_path=input('输入照片所在路径')
                photo_path=photo_path+'\\'
                new=input('是否删除原有照片（Y/N）？')
                model_description=input('增加型号描述，没有直接回车')
                if new.upper()=='Y':
                    update3(sht3,photo_path,start_row='',model=model_description)
                elif new.upper()=='N':
                    row=input('请输入插入图片所在的行数')
                    update3(sht3,photo_path,start_row=row,model=model_description)
            elif choice=='8':
                Page_break(sht4)
                Page_break(sht5)
            elif choice=='9':
                sort_by_item(sht4)
            elif choice=='si':
                line=get_line(sht3)#获取3.0中线的类型
                if line==None:
                    print('并未捕获线的类型')
                else:
                    print('捕捉到线的类型:',line)
                    change_line=input('是否需要手动输入线类型(Y/N):')
                    if change_line=='Y':
                        line=input('请输入线的类型：')
                sync_item(sht3,sht4,line)
            elif choice=='11':
#                get_shapes(sht3)
                line=get_line(sht3)#获取3.0中线的类型
                if line==None:
                    print('并未捕获线的类型')
                else:
                    print('捕捉到线的类型:',line)
                    change_line=input('是否需要手动输入线类型(Y/N):')
                    if change_line=='Y':
                        line=input('请输入线的类型：')
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
            elif choice=='cc':
                check(sht4,'No')
            elif choice=='cc5':
                check_CEC(sht5,'No')
            elif choice=='aml':
                path=input("Please input the ML application path:") #输入申请表的路径
                path=path.replace('"','')
                item=get_ML_item(sht12)
                data=get_ML_info(path,'Yes')
                print(f'报告中已有多重列名ML{item}',)
                modify_ML(sht9,item,data,'A')
            elif choice=='cp3':
                clear_pics(sht3)
            elif choice=='cp7':
                clear_pics(sht7)
            elif choice=='w':#用于把修改好的内容同步保存到原报告
                wb.save(rpt.replace('_output',''))
                wb.save(output_file)
                print('保存时间为:',time.strftime('%Y-%m-%d %H:%M:%S'))
            elif choice=='exit' or choice=='q':
                wb.close()
                app.kill()
                break
            elif choice=='wq':
                wb.save(output_file)
                wb.save(rpt.replace('_output',''))
                wb.close()
                app.kill()
                break
            elif choice=='r':
                wb.close()
                app.kill()
                print('原报告路径：',rpt)
                app=xw.App(visible=True,add_book=False)
                app.display_alerts=False #取消警告
                app.screen_updating=False#取消屏幕刷新
                wb=app.books.open(rpt)
                sht3=wb.sheets['3.0 Photos']
                sht4=wb.sheets['4.0 Components']
                sht5=wb.sheets['5.0 CEC Comps']
                sht7=wb.sheets['7.0 Illustrations']
                sht8=wb.sheets['8.0 Test Summary']
                sht9=wb.sheets['9.0 MLS']
                sht12=wb.sheets['12.0 Revisions']
#                wb.save(rpt[:-5]+'_output.xlsm')
                wb.save(output_file)
            input('any key to contine!')
            os.system('cls')

        

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

        #检查制造商列是否有黄卡号，如有则进行格式处理
        if rows_value[1]!=None:
            ul_no=re.search('\w{1,2}\d{5,6}',rows_value[1])
#            ul_no=re.search('E\d{5,6}|SA\d{5,6}|MH\d{5,6}',rows_value[1])
            print('制造商列找到黄卡号，正在进行格式处理！')
            if ul_no!=None:
                rows_value[1]=str_fmt(rows_value[1].replace('('+ul_no.group()+')',''))#删除原有的黄卡号信息
                rows_value[1]=rows_value[1].strip().replace('\n','')+'\n('+ul_no.group()+')'#写入新的黄卡号

        #以下第5列是可选的，针对单独给出认证号的情形，将认证号提取出来
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
        #遍历所有元素，把None的空元素替换掉
        for items in data:
            for i in range(0,len(items)):
                if items[i]==None:
                    items[i]='Alternative'
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
    row_max=sheet.used_range.last_cell.row
    for i in range(1,row_max):
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
#    name=get_col_list(sheet,'c',1,sheet.used_range.last_cell.row) #获取C列的部件名
    name=get_col_list(sheet,'b',1,sheet.used_range.last_cell.row) #获取C列的部件名
#    print(name)
    total=len(name)
    count=1
    for value in name:
        data=[]
        data.append(value)
        data.append(value)
        print(f'合并单元格进度：{count}/{total}-->{value}')
        rows=row_range(sheet,data)
#        print(rows)
        if rows[0]<rows[1]:
            sheet[f'a{rows[0]+1}:c{rows[1]}'].value=''
        sheet[f'c{rows[0]}:c{rows[1]}'].merge()
        sheet[f'b{rows[0]}:b{rows[1]}'].merge()
        sheet[f'a{rows[0]}:a{rows[1]}'].merge()
        count+=1
        
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

    
def str_fmt(string,ptf='No'):
#以下为中文的符号的处理
    if ptf=='Yes':
        print('输入得数据类型为：',type(string))
    if string!=None and isinstance(string,str):
        string=string.replace('，',',')#替换中文逗号
        string=string.replace('（','(')#替换中文括号
        string=string.replace('）',')')#替换中文括号
        string=string.replace('：',':')#替换中文冒号
        string=string.replace('；',';')#替换中文分号
        string=string.replace('、',',')#替换中文顿号

        #下面对非整数数字的分隔符做处理
        pattern=re.search('\d+,\d+\s?[AWwmm]',string)#匹配诸如13,6W, 12,9 A这类用逗号分割的数字以及对应的单位，数字和单位间可能有空格
        if pattern!=None:#如果有匹配到，则进行如下处理
            original_value=pattern.group()#提取匹配到的原始内容
            fmt_value=original_value.replace(',','.')#把逗号改为小数点
            fmt_value=fmt_value.replace(' ','')#去除单位前的空格
            string=string.replace(original_value,fmt_value)#将对应内容进行替换
    
        if ',' in string:
            string=separate(string,',')
            print('正在分割逗号：',string)
        if ':' in string:
            string=separate(string,':')
            print('正在分割冒号：',string)
        if ';' in string:
            string=separate(string,';')
            print('正在分割分号：',string)

    
        return string
    else:
        return string

def list_fmt(list):
    for i in range(1,len(list)):
        if isinstance(list[i],str)==True: #只针对字符串进行格式化操作
            list[i]=str_fmt(list[i],'No')
    return list

def row_range(sheet,data): #xlwings:查找相同name or item的部件的行数范围
    rows=[]
#    total_row=sheet_total_rows(sheet)+1
    total_row=sheet.used_range.last_cell.row#返回最大的行数
    for i in range(1,total_row):#在报告的此行数范围内去匹配
#        if sheet[f'c{i}'].value==data[1]:#c列中寻找data[1]，即Name
        if sheet[f'b{i}'].value==data[0]:#b列中寻找data[0]，即item
            row_start=i #同name的部件的起始行
            rows.append(row_start)#找到对应的关键词，记录开始行
            row_end=lookdown(sheet,'c',i)
            rows.append(row_end)#记录暂定的结束行，如果下方是同一部件，则会被后面的替代，如果不是，这就是最终的行数
#            while(sheet[f'c{row_end+1}'].value==data[1]):#查找name的方法
            while(sheet[f'b{row_end+1}'].value==data[0]):#查找item的方法
                row_end=row_end+1#同name or item的部件的结束行
                rows[1]=row_end #找到同样的name or item，更新结束行
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


def update3(sheet,photo_path,start_row='',model=''): #xlwings:在3.0自动插入照片
    row_height=12.5 #默认行高12.5pt
    last_row=sheet.used_range.last_cell.row #返回最后一行的行号
    if start_row=='':
    #    sheet[f'a3:j{last_row}'].clear_contents()#清除A列相关行数的内容
    #    sheet[f'a3:j{last_row}'].delete()#删除对应区域的内容，格式保留
        sheet[f'a3:j{last_row}'].api.EntireRow.Delete()#删除对应区域的行数
        while sheet.pictures.count>0:#当sheet中有图片时，删除图片
            sheet.pictures[0].delete()
        number=sheet.pictures.count#当前的图片数量
        row=5
    else:
        number=sheet.pictures.count#当前的图片数量
        row=int(start_row)
    top=row_height*row #12.5pt初始行高，5为行数
    for root,dirs,files in os.walk(photo_path,topdown=False):#遍历路径下的文件和文件夹，返回root,dirs,files的三元元组
#        files.sort(key=len) #在对文件的长度进行排序
        files.sort(key=mysort)#对文件进行排序
        for file in files:#遍历所有的文件
            filename=file.split('_')[1]
            if model=='':
                description=filename.split('.')[0]
            else:
                description=filename.split('.')[0]+f' ({model})'
#            print(files)
            print(photo_path+file)
            sheet.pictures.add(photo_path+file)#插入图片
            if sheet.pictures[number].width>sheet.pictures[number].height:
                sheet.pictures[number].width=354 #单位为pt，72pt=1inch，即288/72=4inch
            else:
                sheet.pictures[number].height=288 #单位为pt，72pt=1inch，即288/72=4inch
            sheet.pictures[number].top=sheet[f'a1:a{row}'].height #用行数来定位
            sheet.pictures[number].left=50.5 #单元格默认列宽50.5
            sheet[f'a{row-2}'].value=f'Photo {number+1} - {description}' #插入文字描述
            sheet[f'a{row-2}'].characters[:9].font.bold=True #部分字体加粗
            row=row+28 #56行一页，28行一半中间位置
            number=number+1


def update4(sheet1,sheet2,sheet3,auto_fmt='Yes'):#xlwings:更新4.0信息
    '''
    sheet1为报告的sec4.0
    sheet2为数据报告的sec4.0
    sheet3为报告的sec12.0
    auto_fmt(str):是否要格式处理Yes/No
    '''
    row_rev=sheet_total_rows(sheet3)+1#SEC12的行数,这里不能用used_range来代替，因为used_range会把空行包含进去，包括格式的改变
#    print(sheet2.used_range.last_cell.row)
    for i in range(1,sheet2.used_range.last_cell.row+1): #在此行数范围内去匹配需要修改的信息,+1是因为range函数
        print('-'*10+f'正在处理第{i}行'+'-'*10)
        if sheet2[f'h{i}'].value=="A": #判断H列是否为A，A为新增
            data=copy_line(sheet2,i)#复制对应行的数据
            print('add:',data)
#            for j in range(1,sheet_total_rows(sheet1)+1):#在报告的此行数范围内去匹配
            for j in range(1,sheet1.used_range.last_cell.row):#在报告的此行数范围内去匹配
                if sheet1[f'c{j}'].value==data[1]:#c列中寻找data[1]，即Name
#                if sheet1[f'c{j}'].value.strip()==data[1].strip():#c列中寻找data[1]，即Name
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
                if data_rpt[2].upper()==data[2].upper() and str(data_rpt[3]).upper()==str(data[3]).upper(): #匹配制造商与型号，当一致时，进行后面的操作
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
#                print(data_rpt[2]+'###'+data[2])
#                print(sheet1[f'e{j}'].value+'###'+data[3])
#                print(data_rpt[4]+'###'+data[4])
                if data_rpt[2]==data[2] and sheet1[f'e{j}'].value==data[3] and data_rpt[4]==data[4]: #匹配部件名，制造商和型号，当一致时，进行后面的操作
                    sheet1[f'c{j}'].api.EntireRow.Delete()#删除改行
                    update12(sheet3,row_rev,data_rpt,data,'D')
                    row_rev=row_rev+1
    if auto_fmt=='Yes':
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
            if file.endswith(".jpg"):
    #            print(files)
                print(manual_path+'\\'+file)
                sheet.pictures.add(manual_path+'\\'+file)#插入图片
                sheet.pictures[number].width=450
                sheet.pictures[number].top=top
                sheet[f'a{row-2}'].value=f'Illustration 2{letters[number]} - Manual - page {number+1}' #插入文字描述
                sheet[f'a{row-2}'].characters[:16].font.bold=True #部分字体加粗
                row=row+56
                top=top+row_height*56 #56行为分页的行数
                number=number+1
            else:
                print(f'{file}不是jpg文件!!')

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

###删除指定的工作簿，生成客户用CDR
def to_client(workbook):
    '''
    workbook为CDR
    '''
    sheets_name=get_sheets_name(workbook)#获取工作簿中的表名
    for i in ['8.0 Test Summary','9.0 MLS','10.0 General','11.0 Production','12.0 Revisions']:
        workbook.sheets[i].delete()

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
        'report':sht1['b3'].value,
        'applicant':sht1['b6'].value,
        'address':sht1['b7'].value,
        'country':sht1['b8'].value,
        'contact':sht1['b9'].value,
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
            elif 'power unit'.lower() in uc_info['name'].lower() or 'pwb'.lower() in uc_info['name'].lower() or 'SMPS'.lower() in uc_info['name'].lower() or 'transformer' in uc_info['name'].lower(): #开关电源
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
                uc_info[f'time_{j}']=sht5[f'j{i+j}'].value
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
        print(f'正在查找{item_no}')
        row=last_row#从下往上遍历
        while sheet[f'b{row}'].value!=item_no:#从下往上找，知道找到对应的序号
            row=row-1#找不到则行号-1
        if sheet[f'b{row}'].merge_cells:#找到对应item后判断对应的单元格是否为合并单元格，如是则对应区域一起剪切
            address=sheet[f'b{row}'].merge_area.address#获取合并单元格对应的区域范围
            sheet[address].api.EntireRow.Cut()#剪切该区域的完整行
        else:
            sheet[f'b{row}'].api.EntireRow.Cut()#不是合并单元格，直接剪切
        sheet['a3'].api.EntireRow.Insert()#在第三行上方插入剪切的数据

def sync_item(sheet_photo,sheet_components,line):#xlwings:同步修改后的item号
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
            change_photo_no(sheet_photo,old_no,new_no,line)#同步更新3.0中的序号,默认用line作为关键词去匹配，后期可能需要优化
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


def init_item(sheet,shape_name,ptf='No'):#xlwings:对sec3中的item号进行排序
    value=1
    shapes_wanted=[]
    for shape in sheet.shapes:
        if shape_name in shape.name:
            shapes_wanted.append(shape)#提取需要的图形
            if ptf=='Yes':
                print('unsort:'+shape.name+':'+str(shape.text)+':'+str(shape.top))
    shapes_wanted.sort(key=shape_top)#根据高度位置对图形进行排序

#    for shape in sheet.shapes:#遍历每一个shape，给其赋值
    for shape in shapes_wanted:#遍历排序后的shape，给其赋值
        if shape_name in shape.name:
            shape.text=value
            print(shape.name+f':{value}'+':'+str(shape.top))
            value+=1

def shape_top(shape):#返回图形高度位置信息，用来排序
    top=str(shape.top)
    top=top.split('.')[0]
#    print(top)
    return int(top)


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
#    print(filename.split('_')[0])#提取文件名前面的数字
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

def check(sheet,ptf='No'):#xlwings:检查报告4.0证书的正确性
    '''
    sheet: SEC4.0
    '''
    for row in range(3,sheet.used_range.last_cell.row): #在此行数范围内去匹配需要修改的信息
        print(f'\n#####正在核对第{row}行')
        manufacturer=sheet[f'd{row}'].value
        scan=row#扫描的行数
        while manufacturer==None:#当有合并单元格时，向上扫描，获取制造商信息
            scan=scan-1
            manufacturer=sheet[f'd{scan}'].value
        ul_no=re.search('\w{1,2}\d{5,6}',manufacturer)#提取黄卡号
        model=str(sheet[f'e{row}'].value)#转化为字符，针对纯数字问题
        mark=sheet[f'g{row}'].value
        if mark=='NR':#排除NR部件
            continue
        elif mark=='See 5.0':#排除随机测部件
            continue
        elif mark==None:
            continue
        elif ul_no==None:#排除没有控制号的部件
            continue
        else:
            url='https://iq.ulprospector.com/en/_/_results?p=10005,10048,10006,10047&qm=q:'+ul_no.group()
            if ptf=='Yes':
                print(url)
            selector_basic=ul_search(url)#用get方法提交搜索请求，返回搜索结果的response
            items=basic_info(selector_basic)#输出查询的结果并返回详细连接
            time.sleep(random.randint(3,5))
            if ptf=='Yes':
                print(items)
            if len(items)==0:
                print('invalid cert')
                sheet[f'h{row}'].value='invalid cert'
                continue
            else:
                ul_flag=[0]#状态说明：3-查找到精准型号，2-查找到类似型号，1-没查到
                csa_flag=[0]
                for item in items:#遍历所有的查询结果
                    print(f'正在比对{item}')
                    selector_details=ul_search('https://iq.ulprospector.com'+item[3])#查询详细链接中的内容
                    models=certificate(selector_details)
                    if filters(models,model)=='green':
                        if 'Canada' in item[2]:
                            csa_flag.append(3)
                        else:
                            ul_flag.append(3)
                    elif filters(models,model)=='yellow':
                        if 'Canada' in item[2]:
                            csa_flag.append(2)
                        else:
                            ul_flag.append(2)
                    elif filters(models,model)=='red':
                        if 'Canada' in item[2]:
                            csa_flag.append(1)
                        else:
                            ul_flag.append(1)

                if max(ul_flag)==3 and max(csa_flag)==3:
                    sheet[f'h{row}'].value='ok'
                elif max(ul_flag)==3 and max(csa_flag)==2:
                    sheet[f'h{row}'].value='ul is ok, csa to be check'
                elif max(ul_flag)==2 and max(csa_flag)==3:
                    sheet[f'h{row}'].value='csa is ok, ul to be check'
                elif max(ul_flag)==3 and max(csa_flag)==1:
                    sheet[f'h{row}'].value='ul is ok, csa not found'
                elif max(ul_flag)==1 and max(csa_flag)==3:
                    sheet[f'h{row}'].value='csa is ok, ul not found'
                elif max(ul_flag)==2 and max(csa_flag)==2:
                    sheet[f'h{row}'].value='to be check'
                elif max(ul_flag)==2 and max(csa_flag)==1:
                    sheet[f'h{row}'].value='ul to be check,csa not found'
                elif max(ul_flag)==1 and max(csa_flag)==2:
                    sheet[f'h{row}'].value='csa to be check,ul not found'
                elif max(ul_flag)==1 and max(csa_flag)==1:
                    sheet[f'h{row}'].value='not found'
                elif max(ul_flag)>=1 and max(csa_flag)==0:
                    sheet[f'h{row}'].value='no cert for csa'
                elif max(ul_flag)==0 and max(csa_flag)>=1:
                    sheet[f'h{row}'].value='no cert for ul'

def check_CEC(sheet,ptf='No'):#xlwings:检查报告5.0证书的正确性
    '''
    sheet: SEC5.0
    '''
    for row in range(1,sheet.used_range.last_cell.row): #在此行数范围内去匹配需要修改的信息
        print(row)
        if sheet[f'l{row}'].value=="C": #判断L列是否为C，C为检查该部件下的证书
#            manufacturer=sheet_data[f'f{row}'].value
#            model=sheet_data[f'i{row}'].value
            start=row+5#数据的起始行数
            row_scan=start
#            print(manufacturer)
#            print(model)
            while sheet[f'a{row_scan}'].value!='WINDING(S) RESISTANCE':
                row_scan+=1
                print(sheet[f'a{row_scan}'].value)
            end=row_scan#数据的终止行数
    for row in range(start,end): #在此行数范围内去匹配需要修改的信息
        print(f'\n#####正在核对第{row}行')
        manufacturer=sheet[f'c{row}'].value
        scan=row#扫描的行数
        while manufacturer==None:#当有合并单元格时，向上扫描，获取制造商信息
            scan=scan-1
            manufacturer=sheet[f'c{scan}'].value
        ul_no=re.search('\w{1,2}\d{5,6}',manufacturer)#提取黄卡号
        model=str(sheet[f'e{row}'].value)#转化为字符，针对纯数字问题
        if ul_no==None:#排除没有控制号的部件
            continue
        else:
            url='https://iq.ulprospector.com/en/_/_results?p=10005,10048,10006,10047&qm=q:'+ul_no.group()
            if ptf=='Yes':
                print(url)
            selector_basic=ul_search(url)#用get方法提交搜索请求，返回搜索结果的response
            items=basic_info(selector_basic)#输出查询的结果并返回详细连接
            time.sleep(random.randint(3,5))
            if ptf=='Yes':
                print(items)
            if len(items)==0:
                print('invalid cert')
                sheet[f'h{row}'].value='invalid cert'
                continue
            else:
                ul_flag=[0]#状态说明：3-查找到精准型号，2-查找到类似型号，1-没查到
                csa_flag=[0]
                for item in items:#遍历所有的查询结果
                    print(f'正在比对{item}')
                    selector_details=ul_search('https://iq.ulprospector.com'+item[3])#查询详细链接中的内容
                    models=certificate(selector_details)
                    if filters(models,model)=='green':
                        if 'Canada' in item[2]:
                            csa_flag.append(3)
                        else:
                            ul_flag.append(3)
                    elif filters(models,model)=='yellow':
                        if 'Canada' in item[2]:
                            csa_flag.append(2)
                        else:
                            ul_flag.append(2)
                    elif filters(models,model)=='red':
                        if 'Canada' in item[2]:
                            csa_flag.append(1)
                        else:
                            ul_flag.append(1)

                if max(ul_flag)==3 and max(csa_flag)==3:
                    sheet[f'l{row}'].value='ok'
                elif max(ul_flag)==3 and max(csa_flag)==2:
                    sheet[f'l{row}'].value='ul is ok, csa to be check'
                elif max(ul_flag)==2 and max(csa_flag)==3:
                    sheet[f'l{row}'].value='csa is ok, ul to be check'
                elif max(ul_flag)==3 and max(csa_flag)==1:
                    sheet[f'l{row}'].value='ul is ok, csa not found'
                elif max(ul_flag)==1 and max(csa_flag)==3:
                    sheet[f'l{row}'].value='csa is ok, ul not found'
                elif max(ul_flag)==2 and max(csa_flag)==2:
                    sheet[f'l{row}'].value='to be check'
                elif max(ul_flag)==2 and max(csa_flag)==1:
                    sheet[f'l{row}'].value='ul to be check,csa not found'
                elif max(ul_flag)==1 and max(csa_flag)==2:
                    sheet[f'l{row}'].value='csa to be check,ul not found'
                elif max(ul_flag)==1 and max(csa_flag)==1:
                    sheet[f'l{row}'].value='not found'
                elif max(ul_flag)>=1 and max(csa_flag)==0:
                    sheet[f'l{row}'].value='no cert for csa'
                elif max(ul_flag)==0 and max(csa_flag)>=1:
                    sheet[f'l{row}'].value='no cert for ul'


def get_ML_info(path,ptf='No'):#xlwings：获取多重列名的型号
    '''
    path:多重列名申请表的PDF路径
    '''
    print('!!!适用于申请表版本:SFT-ETL-OP-19t (11-November-2021) Mandatory, 其他版本可能会遇到提取信息错乱等问题')
    pdf=pdfplumber.open(path)#打开pdf
    page1=pdf.pages[0]#获取第一页
    text1=page1.extract_text()#提取第一页的文本内容
    res=re.search('Company Name:[\s\S]*Associated',text1)#提取相关内容
    content=res.group().split('\n')#按行分段提取内容为列表
    for line in content:
        if 'Company Name' in line:
            res=re.search(':[\s\S]*:',line)#提取列名厂家啊
            ML_company=res.group().replace('Company Name:','').replace(':','').strip()
        if 'Brand Name' in line:
            res=re.search(': +\w+',line)#提取商标
            Brand=res.group().replace(':','').strip()
        if 'Address' in line:#提取街道信息
            res=re.search(':[\s\S]*:',line)
            street=res.group().replace('Street Address:','').replace(':','').strip()
        if 'City' in line:
            res=re.search(':[\s\S]*City',line)#提取城市
            city=res.group().replace('City','').replace(':','').strip()
        if 'Country' in line:
            res=re.search(':[\s\S]*:',line)#提取国家
            country=res.group().replace('Country:','').replace(':','').strip()
    if ptf=='Yes':
        print('ML company:',ML_company)
        print('Brand name:',Brand)
        print('Address:',street+city)
        print('Country:',country)

    try:
        page2=pdf.pages[1]#获取第二页
        text2=page2.extract_text()#提取第二页的文本内容
        res=re.search('MODELS[\s\S]*A complimentary',text2)#使用正则提取型号相关的部分
        content=res.group()
        content=content.replace('MODELS','')#删除多余信息
        content=content.replace('A complimentary','')#删除多余信息
        content=content.strip()#去除首尾空格
        models_line=content.split('\n')#以行为单位获取内容
        if ptf=='Yes':
            print('获取的所有型号相关信息：',models_line)
        ML_models=[]
        basic_models=[]
        for models in models_line:#遍历每一行
            models=models.strip()#去除首尾多余空格
            if ptf=='Yes':
                print('正在处理:',models)
            res=re.search('^[\s\S]* ',models)#提取多重列名型号
            ML_model=res.group().strip()
            res=re.search(' [\s\S]*$',models)#提取基本列名型号
            basic_model=res.group().strip()
            if ptf=='Yes':
                print('ML model:',ML_model)
                print('basic model:',basic_model)
            ML_models.append(ML_model)
            basic_models.append(basic_model)
        if ptf=='Yes':
            print('='*20)
            print(f'Add new ML for {ML_company}')
            for ML,basic in zip(ML_models,basic_models):
                print(f'ML model {ML} (brand name: {Brand}) for basic model {basic}.')
            print('='*20)
    except:
        MT=input('是否手动输入(Y/N):')
        if MT=='Y':
            ML_model=input('请手动输入多重列名型号:')
            ML_models=ML_model.split('  ')
            basic_model=input('请手动输入基本型号:')
            basic_models=basic_model.split('  ')
            if len(ML_models)==len(basic_models):#比较数量是否一致
                for ML,basic in zip(ML_models,basic_models):
                    print(f'ML model {ML} (brand name: {Brand}) for basic model {basic}.')
                print('='*20)
            else:
                print('数量不一致，请检查')

    return [ML_company,street+city,country,Brand,ML_models,basic_models]
#    print(ML_models,basic_models)

def get_ML_item(sheet):#xlwings:查找使用过的最大列名
    '''
    sheet: sec12的工作簿
    '''
    item_max=0#记录最大的列名号，初始为0
    print(sheet.used_range.last_cell.row)
    for row in range(5,sheet.used_range.last_cell.row+2):#从第五行开始,+1是因为range前闭后开
#        print(f'正在比对第{row}行') #debug
        if sheet[f'c{row}'].value==9.0:
            value=sheet[f'd{row}'].value#获取item列的数值
            if re.search('\d+',str(value))!=None:
                item=re.search('\d+',str(value)).group()#提取数字部分
                item_max=max(item_max,int(item))
    return item_max

def modify_ML(sheet,item,data,act):#xlwings:自动修改多重列名
    '''
    sheet:SEC9
    item:多重列名的序号
    data:由get_ML_info获得的数据列表
    act:具体的行为，如新增，删除等
    '''
    row_max=sheet.used_range.last_cell.row
    if act=='A':#新增列名
        if int(item)<3:#当已有列名小于3时的新增，因为报告模板已有
            row=get_row_number(sheet,'a','MULTIPLE LISTEE '+str(int(item)+1))#item为已有的列名数，+1为新增,定位写入的行数
            sheet[f'b{row}'].value=data[0]#多重列名厂家
            sheet[f'b{row+1}'].value=data[1]#地址
            sheet[f'b{row+2}'].value=data[2]#国家
            sheet[f'b{row+3}'].value=data[3]#商标
            sheet[f'b{row+5}'].value='=B3'
            sheet[f'b{row+6}'].value='=B4'
            sheet[f'b{row+7}'].value='=B5'

            ML_model=''
            for model in data[4]:#型号字符串的拼接处理
                if ML_model=='':
                    ML_model=model
                else:
                    ML_model=ML_model+'\n\n'+model
            sheet[f'a{row+10}'].value=ML_model

            basic_model=''
            for model in data[5]:#型号字符串的拼接处理
                if basic_model=='':
                    basic_model=model
                else:
                    basic_model=basic_model+'\n\n'+model
            sheet[f'c{row+10}'].value=basic_model
        elif int(item)>=3:#当已有列名大于3时的新增，需要自己写入
            insert_row=sheet.used_range.last_cell.row+2
            print(f'在{insert_row}行处开始写入')
            sheet.api.Rows("8:18").Copy(sheet.api.Rows(insert_row))
            sheet[f'a{insert_row}'].value=f'MULTIPLE LISTEE {item+1}'
            sheet[f'a{insert_row+9}'].value=f'MULTIPLE LISTEE {item+1} MODELS'

            sheet[f'b{insert_row}'].value=data[0]#多重列名厂家
            sheet[f'b{insert_row+1}'].value=data[1]#地址
            sheet[f'b{insert_row+2}'].value=data[2]#国家
            sheet[f'b{insert_row+3}'].value=data[3]#商标
            sheet[f'b{insert_row+5}'].value='=B3'
            sheet[f'b{insert_row+6}'].value='=B4'
            sheet[f'b{insert_row+7}'].value='=B5'

            ML_model=''
            for model in data[4]:#型号字符串的拼接处理
                if ML_model=='':
                    ML_model=model
                else:
                    ML_model=ML_model+'\n\n'+model
            sheet[f'a{insert_row+10}'].value=ML_model

            basic_model=''
            for model in data[5]:#型号字符串的拼接处理
                if basic_model=='':
                    basic_model=model
                else:
                    basic_model=basic_model+'\n\n'+model
            sheet[f'c{insert_row+10}'].value=basic_model
            pass

def output_path(file_path,ptf='No'):
    '''
    file_path:文件的具体路径
    '''
    path_split=os.path.split(file_path)
    output_path=os.path.join(path_split[0],'output')
    if os.path.exists(output_path):
        pass
    else:
        if ptf=='Yes':
            print('正在创建output文件夹')
        os.mkdir(output_path)
    if ptf=='Yes':
        print(output_file)
    output_file=os.path.join(output_path,path_split[1][:-5]+'_output.xlsm')
    return output_file


###增加基本列名
def add_models(sheet_rpt2,sheet_rpt12,sheet_data):
    '''
    sheet_rpt2:报告的sec2.0对应的sheet
    sheet_rpt12:报告的sec12.0对应的sheet
    sheet_data:数据对应的sheet,包含基本型号，列名型号以及商标这三列信息
    '''
    last_row=sheet_data.used_range.last_cell.row#sheet中最大的行数
    print('最大行数',last_row)
    models=[]
    basic_model=[]
    new_model=[]
    brand=[]
    similarity=[]
    for row in range(1,last_row+1):
        print(f'正在比对第{row}行')
        if sheet_data[f'd{row}'].value=="A": #判断D列是否为A，A为新增
            #获取基本信息
            basic_model.append(sheet_data[f'a{row}'].value)#获取基本型号
            new_model.append(sheet_data[f'b{row}'].value)#获取增加的列名型号
            brand.append(sheet_data[f'c{row}'].value)#获取商标
    for basic,new in zip(basic_model,new_model):
        similarity.append(f'{new} is identical with {basic} except for the model name.')
    #写入报告
    add_cell_text(sheet_rpt2,'B4',"\n"+", ".join(brand))#在报告sec2中增加商标
    add_cell_text(sheet_rpt2,'B7',"\n"+", ".join(new_model))#在报告sec2中增加型号
    print(f'正在增加{new_model}')
    add_cell_text(sheet_rpt2,'B8',"\n"+"\n".join(similarity))#增加相似性描述
    

###单元格增加内容
def add_cell_text(sheet,cell,text):
    '''
    sheet:要修改的sheet
    cell(str):修改的单元格，如'A1'
    text(str):增加的文本
    '''
    old_text=sheet[cell].value
#    print(old_text)
#    print(type(old_text))
#    print(type(text))
    new_text=old_text+text
    start=len(old_text)+1
    length=len(text)
    sheet[cell].value=new_text
    sheet[cell].api.GetCharacters(Start=start,Length=length).Font.Color=0xFF00FF







###清除指定sheet中的照片
def clear_pics(sheet): #xlwings:删除照片
    for shape in sheet.shapes:#删除所有的shape对象
        shape.delete()
    while sheet.pictures.count>0:#当sheet中有图片时，删除图片
        sheet.pictures[0].delete()
    last_row=sheet.used_range.last_cell.row #返回最后一行的行号
    sheet[f'a3:j{last_row}'].api.EntireRow.Delete()#删除对应区域的行数

if __name__=='__main__':
    Menu()
