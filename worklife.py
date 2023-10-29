#!/bin/python
#coding=utf8
'''
Author: Michael Jin
date:2023-10-29
'''
import uiautomator2 as u2
import time
import os
import base64
import requests
import redis
from configparser import ConfigParser


###读取配置文件
def load_config():#加载配置文件
    conf=ConfigParser()
    if os.name=='nt':
#        path='K:/config.ini'
        path=r'D:\Downloads\PortableGit-2.36.1-64-bit.7z\bin\Quat\config.ini'
    elif os.name=='posix':
        path='/usr/local/src/Quat/config.ini'
    else:
        print('no config file was found!')

    conf.read(path,encoding="utf-8")
    return conf


###连接手机
def u2_connect(conf):
    try:
        print('正在尝试有线连接!')
        d=u2.connect()
        print(d.info)
    except:
        print('正在尝试无线连接!')
        addr=conf.get('adb','ip')
        cmd=f'adb connect {addr}'
        print(cmd)
        if os.name=='posix':
            os.system(cmd)
        elif os.name=='nt':
            os.system(f'D:\Downloads\scrcpy-win64-v2.1\\{cmd}')
        d=u2.connect(addr)
        print(d.info)
    return d

###通知功能
def notify(method,title,content):
    '''
    功能：推送消息
    title:消息的标题
    content:消息的内容

    '''
    conf=load_config()#读取配置文件
    token=conf.get('pushplus','token')#获取配置文件中的token

    if method=='get':
        if isinstance(content,str):
            url=f'http://www.pushplus.plus/send?token={token}&title={title}&content={content}'#构建get请求的地址
        elif isinstance(content,list):
            pass
    
    
        res=requests.get(url)#发送get请求
        print(res.status_code)
        print(res.url)
    elif method=='post':
        url='http://www.pushplus.plus/send/'
        data={
                'token':f'{token}',
                'title':f'{title}',
                'content':f'{content}'
                }
        res=requests.post(url,data=data)
        print(res.status_code)
        print(data)

###打卡
def check_in(d):
    if d.info.get('screenOn')==False:#熄屏状态
        d.unlock()
        unlock=conf.get('adb','unlock')#解锁密码
        if os.name=='posix':
            os.system('adb shell input text {}'.format(unlock))
        elif os.name=='nt':
            os.system('D:\Downloads\scrcpy-win64-v2.1\\adb shell input text {}'.format(unlock))

    if check_running(d,'com.cdp.epPortal'):
        print('检测到后台运行，正在停止该app')
        d.app_stop('com.cdp.epPortal')

    #打开app开始操作
    d.app_start('com.cdp.epPortal')
    #开始打卡
    d(text='移动打卡').click()
    time.sleep(5)
    d.press('back')
    d(text='移动打卡').click()
    #等待加载完成
    while True:
        print('loading')
        if d(description='上海钦州北路1198号').exists():
            print('加载完成')
            break
    while True:
        if d(description='第1次打卡').exists():
            notify('post','worklife','未打过卡，正在进行第一次打卡')
#        d(description='第1次打卡').click()
            break
        elif d(description='第2次打卡').exists():
            notify('post','worklife','打过1次卡，正在进行第二次打卡')
#        d(description='第2次打卡').click()
            break


###检查某个app是否在后台运行
def check_running(d,name):
    '''
    d(obj):u2连接对象
    name(str):app名称
    '''
    running_apps=d.app_list_running()
#    print(running_apps)
    for app in running_apps:
        print(f'正在比对{app}')
        if name==app:
            return True


def main(conf):
    token=conf.get('redis','token')
    r = redis.Redis(
        host='redis-16873.c294.ap-northeast-1-2.ec2.cloud.redislabs.com',
        port=16873,
        password=token)
    ps = r.pubsub()
    ps.subscribe('myChannel')
    for item in ps.listen():  # keep listening, and print the message in the channel
        if item['type'] == 'message':
            signals = item['data'].decode('utf-8')
            if signals == 'exit':
                break
            elif signals=='check_in':
                check_in(d)

if __name__=='__main__':
    conf=load_config()
    d=u2_connect(conf)
    main(conf)
