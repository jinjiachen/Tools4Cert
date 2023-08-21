
#/bin/python #-*-coding:utf-8-*-

'''
Author: Michael Jin
Date: 2023-02
'''

import xlwings as xw
import pandas as pd

def read_excel():
#    file=input('请输入excel文件路径')
#    file=file.replace('"','')
    file=r'Y:\Lab-Share\Equipment list-SH\Equipment List-SH (202307).xls.xlsx'
    EC_num=input('请输入要查询的设备号')
    sheet=pd.read_excel(io=file)
    res=sheet[sheet['REG. NO.']==EC_num]
#    print(res)
#    print(type(res))
#    print('len:',len(res))

    if len(res)==0:
        print('无此设备！')
    else:
        #提取需要的信息
        name=res['EQUIPMENT NAME'].values[0]
        PL=res['PL'].values[0]
        location=res['location'].values[0]
        cal_start=res['CAL.DATE'].values[0]
        cal_end=res['CAL.NEXT DUE'].values[0]
        status=res['STATUS'].values[0]
        print('#'*20)
        print(f'设备号：{EC_num}\n设备名称：{name}\n地址：{PL}: {location}\n计量日期：{cal_start} to {cal_end}\n状态：{status}')
        print('#'*20+'\n')


if __name__=='__main__':
    while True:
        read_excel()
