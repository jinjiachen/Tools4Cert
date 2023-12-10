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


###检查ui2服务是否运行
def check_status(d):
    if d.service('uiautomator').running():
        print('servise running')
    else:
        print('starting servise')
        d.service('uiautomator').start()

###死循环解锁屏幕，确保解锁成功
def wakeup(d,conf):
    '''
    d(obj):u2对象
    conf:load_conf返回结果
    '''
    while True:
        #检查屏幕状态，如果息屏，点亮并解锁
        if d.info.get('screenOn')==False:#熄屏状态
            d.unlock()
            unlock=conf.get('adb','unlock')#解锁密码
            if os.name=='posix':
                os.system('adb shell input text {}'.format(unlock))
            elif os.name=='nt':
                os.system('D:\Downloads\scrcpy-win64-v2.1\\adb shell input text {}'.format(unlock))
        elif d.info.get('screenOn')==True:#熄屏状态
            break

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
def check_in(d,r,click='NO'):
    '''
    d(obj):uiautomator2对象
    r(obj):redis对象
    '''
    wakeup(d,conf)#解锁屏幕
    if check_running(d,'com.cdp.epPortal'):
        print('检测到后台运行，正在停止该app')
        d.app_stop('com.cdp.epPortal')

    #打开app开始操作
    d.app_start('com.cdp.epPortal')
    #开始打卡
    d(text='移动打卡').click()
    time.sleep(15)
    d.press('back')
    d(text='移动打卡').click()
    #等待加载完成
    while True:
        print('loading')
        if d(description='上海钦州北路1198号').exists():
            print('加载完成')
            break
    #开始查询或打卡
    while True:
        now=time.strftime("%H:%M:%S")
        if d(description='第1次打卡').exists():
            if click=='YES':
                d(description='第1次打卡').click()
#                if d(description='打卡成功').exists():
#                    r.set('worklife','[res]:第一次打卡成功！')
#                    notify('post',f'worklife-{now}',f'{now}\n第一次打卡成功!')
#                else:
#                    r.set('worklife','[res]:第一次打卡失败！')
#                    notify('post',f'worklife-{now}',f'{now}\n第一次打卡失败!')
                d.press('back')
                d(text='移动打卡').click()
                if d(description='第1次打卡').exists():
                    notify('post',f'worklife-{now}',f'{now}\n第一次打卡成功!')
            else:
                r.set('worklife','[res]:未过一次卡！')
                notify('post',f'worklife-{now}',f'{now}\n未打过卡!')
            break
        if d(description='第2次打卡').exists():
            if click=='YES':
                d(description='第2次打卡').click()
#                if d(description='打卡成功').exists():
#                if d(text='打卡成功').exists():
#                    r.set('worklife','[res]:第二次打卡成功！')
#                    notify('post',f'worklife-{now}',f'{now}\n第二次打卡成功!')
#                else:
#                    r.set('worklife','[res]:第二次打卡失败！')
#                    notify('post',f'worklife-{now}',f'{now}\n第二次打卡失败!')
                d.press('back')
                d(text='移动打卡').click()
                if d(description='第2次打卡').exists():
                    notify('post',f'worklife-{now}',f'{now}\n第二次打卡成功!')
            else:
                r.set('worklife','[res]:打过一次卡！')
                notify('post',f'worklife-{now}',f'{now}\n打过一次卡!')
            break
    d.app_stop('com.cdp.epPortal')
    d.screen_off()


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
    ps.subscribe('worklife')
    for item in ps.listen():  # keep listening, and print the message in the channel
        now=time.strftime("%H:%M:%S")
        print(f'listining: {now}')
        if item['type'] == 'message':
            signals = item['data'].decode('utf-8')
            if signals == 'exit':
                break
            elif signals=='test':
                check_in(d,r)
            elif signals=='check_in':
                check_in(d,r,'YES')

if __name__=='__main__':
    conf=load_config()
    d=u2_connect(conf)
    while True:
        try:
            main(conf)
        except:
            check_status(d)
            continue
