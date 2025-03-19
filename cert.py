from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from lxml import etree
import os
import pdb
import urllib
import requests,httpx
requests.packages.urllib3.disable_warnings
import warnings
warnings.filterwarnings("ignore")
from configparser import ConfigParser
import base64

def Menu():
    choice=input("请选择认证的类型：\n1.UL认证\n2.TUV莱茵认证\n3.VDE认证\n4.CSA认证\n5.TUV南德\n11.UL查询（模拟请求)")
    if choice=='1':
        UL(driver)
    elif choice=='2':
        TUV(driver)
    elif choice=='11':
        while True:
            ul_no=input('请输入需要查询的关键字:')
#            url='https://iq.ulprospector.com/en/_/_results?p=10005,10048,10006,10047&qm=q:'+ul_no
#            url='https://iq.ulprospector.com/en/_?p=10005,10048,10006,10047&qm=q:'+ul_no
#            url='https://iq.ulprospector.com/en/_?qm=q:relay'
#            url='https://iq.ulprospector.com/en/_?p=10005,10048,10006,10047&qm=q:relay'
            url='https://www.ul.com'
#            url='https://iq.ulprospector.com/en'
            res_basic=ul_search(url)
            items=basic_info(res_basic)
            no=input('请选择对应的部件序号:')
            res_details=ul_search('https://iq.ulprospector.com'+items[int(no)][3])
            models=certificate(res_details)
            model=input('想要查找的型号：')
            filters(models,model)
            input('PRESS ANYTHING TO CONINUE!')
            if os.name=='nt':
                os.system('cls')
            elif os.name=='posix':
                os.system('clear')


def Driver():
    #Chrom的配置
    options = webdriver.ChromeOptions()
#    options.add_argument("--proxy-server=http://192.168.2.108:8889")
    #options.add_argument("--no-proxy-server")
#    options.add_argument("--headless")
    #options.add_argument('user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"')
    #options.add_argument('user-agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36"')
    #options.add_argument('user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36')
    options.add_argument('log-level=3') #INFO = 0 WARNING = 1 LOG_ERROR = 2 LOG_FATAL = 3 default is 0
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    #Chrome的驱动和路径
    path="C:\Program Files\Google\Chrome\Application\chrome.exe"
    #driver=webdriver.Chrome(chrome_options=options,executable_path=path)
    #driver=webdriver.Chrome(path,chrome_options=options)
    driver=webdriver.Chrome(chrome_options=options)
    #driver.set_page_load_timeout(10)
    #driver.set_script_timeout(10)
    print('starting')
    return driver


def UL(driver):
    driver.get('https://iq.ulprospector.com/en')
    
    #user=driver.find_element_by_id('email')#老方法，不推荐使用
    #password=driver.find_element_by_id('loginPass')#老方法，不推荐使用
    #login=driver.find_element_by_id('main-login') #老方法，不推荐使用
    #search=driver.find_element_by_id('q2') #老方法，不推荐使用
    
    user=driver.find_element(By.ID,'email')
    user.send_keys('jun.xia@intertek.com')
    check=driver.find_element(By.XPATH,'//input[@type="submit"]')
    check.click()
    time.sleep(30)
    password=driver.find_element(By.ID,'password')
    password.send_keys('Ul-123456')
    login=driver.find_element(By.ID,'continue')
    login.submit()
    print('login')
    
    #方法一：因为是get方法，直接网址来进行访问
    #driver.get('https://iq.ulprospector.com/en/_?p=10005,10048,10006,10047&qm=q:E231292')
    
    #方法二：模拟点击来进行访问
    #search=driver.find_element(By.ID,'q2')
    #search.send_keys('E231292')
    #search.send_keys('s7312')
    #search.submit()
    
    #print(driver.find_element(By.XPATH,'//td[@class="entry"]').text)
    #time.sleep(1)
    #html=driver.page_source
    #print(html)
    #selector=etree.HTML(html)
    #print(selector)
    #company=selector.xpath('//tr[@class=" "]/td[2]/div/span/text()')
    #Control=selector.xpath('//td[@class="entry"]/a/span/text()')
    #base_url='https://iq.ulprospector.com'
    #link=selector.xpath('//td[@class="entry"]/a/@href')
    #CCN=selector.xpath('//tr[@class=" "]/td[4]/div/span/text()')
    #print(company)
    #print(Control)
    #print(CCN)
    #print(base_url+link[0])
    
    #driver.get(base_url+link[2])
    #time.sleep(1)
    #html=driver.page_source
    #print(html)
    #selector=etree.HTML(html)
    #name=selector.xpath('//nameline/text()')
    #addressline=selector.xpath('//addressline/text()')
    #print('\n')
    #print(name)
    #print(addressline)
    
    
    #cookies=driver.get_cookies()
    #url=driver.current_url()
    #print(driver.page_source)
    
    while True:
        E=input('Please input the control number:')
        model=input('Please input the model name:')
        if os.name=='nt':
            os.system('cls')
        elif os.name=='posix':
            os.system('clear')
        if E=='exit':
            break
        else:
            driver.get('https://iq.ulprospector.com/en/_?p=10005,10048,10006,10047&qm=q:'+E)
            time.sleep(1)
            html=driver.page_source
            selector=etree.HTML(html)
            company=selector.xpath('//tr[@class=" "]/td[2]/div/span/text()')
            Control=selector.xpath('//td[@class="entry"]/a/span/text()')
            base_url='https://iq.ulprospector.com'
            link=selector.xpath('//td[@class="entry"]/a/@href')
            CCN=selector.xpath('//tr[@class=" "]/td[4]/div/span/text()')
            if len(company)==0:
                print('控制号无效!!')
                continue
            for i in range(0,len(company)):
                driver.get(base_url+link[i])
                time.sleep(1)
                html=driver.page_source
                selector=etree.HTML(html)
                models=selector.xpath('//prodid/text()')
                if len(models)==0:
                    models=selector.xpath('//a[@style="text-decoration: none;"]/text()')
                elif len(models)==0:
                    models=selector.xpath('//prodid/b/text()')
                print('-'*20+str(i)+'-'*20)
                print(company[i])
                print(Control[i])
                print(CCN[i])
                print(base_url+link[i])
                for i in models:
                    if model==i or model.upper()==i:
                        print('找到型号：',i)
                        break
                    elif model.split('-')[0] in i or model.split('-')[0].upper() in i:
                        print('-'*20)
                        print('找到相似型号：',i)
#                    else:
#                        print('没有找到对应型号')

                print('\n')

#            choice=input('请选择对应的序号：')
#            driver.get(base_url+link[int(choice)])
#            time.sleep(1)
#            html=driver.page_source
#            selector=etree.HTML(html)
#    #        name=selector.xpath('//nameline/text()')
#    #        control_no=selector.xpath('//table[@width="100%"]/tbody/tr[3]/td[2]/text()')
#    #        addressline=selector.xpath('//addressline/text()')
#    #        city=selector.xpath('//city/text()')
#    #        province=selector.xpath('//province/text()')
#    #        postalcode=selector.xpath('//postalcode/text()')
#    #        country=selector.xpath('//country/text()')
#            models=selector.xpath('//prodid/text()')
#            if len(models)==0:
#                models=selector.xpath('//a[@style="text-decoration: none;"]/text()')
#            elif len(models)==0:
#                models=selector.xpath('//prodid/b/text()')
#            model=input('Please input the model name:')
#    #        print('\n')
#    #        print(name[0]+control_no[0])
#    #        print(addressline[0]+city[0]+province[0]+postalcode[0]+country[0])
##            print(models)
#            for i in models:
#                if model==i or model.upper()==i:
#                    print('找到型号：',i)
#                    break
#                if model in i or model.upper() in i:
#                    print('-'*20)
#                    print('找到相似型号：',i)
    
#    time.sleep(2)
#    driver.close()
#    driver.quit()

def TUV(driver):
    while True:
        input('按回车继续')
        if os.name=='nt':
            os.system('cls')
        elif os.name=='posix':
            os.system('clear')
        cert=input('请输入TUV证书号:')
        words=input('请输入型号：')
        if cert=='exit':
            break
        else:
        #    pdb.set_trace()
            url='https://www.certipedia.com/certificates/'+cert+'?locale=en'
            selector=get_html(url)
        #    res=urllib.request.urlopen(url)
        #    html=res.read()
        #    selector=etree.HTML(html)
            pages=selector.xpath('//div[@class="tuv-pagination__pages"]/a/text()')
            company=selector.xpath('//div[@class="certificate-holder-address"]/strong/text()')
            name=selector.xpath('//div[@class="model-designation"]/../text()')
            products=[]
            for page in pages:
                url='https://www.certipedia.com/certificates/'+cert+'?locale=en&page_number='+page
                selector=get_html(url)
                models=selector.xpath('//div[@class="model-designation"]/p/text()')
                for model in models:
                    if 'Model Designation' in model:
                        pass
                    else:
                        products.append(model)
            print("Certificate Holder:"," ".join(company))
            print("Certified Product:",name[0].strip())
            for product in products:
                if words==product or words.upper()==product:
                    print('找到精准型号：',product.strip())
                    print('-'*20)
                    break
                elif words.split('-')[0] in product or words.split('-')[0].upper() in product:
                    print('找到相似型号如下：\n',product.strip())
                    print('-'*20)
#            print("\n".join(products))
#            print(pages)

def get_html(url):#通过selenium获取html后转换为lxml的对象
    start=time.time()
    driver.get(url)
#    time.sleep(1)
    html=driver.page_source#获取html源码
    selector=etree.HTML(html)#转化为lxml
    end=time.time()
    print(end-start)
    return selector


def ul_search(url):#通过模拟请求方式查询黄卡号
    conf=load_config()
    ul_cookies=conf.get('ul','cookie')
    ul_cookies=base64.b64decode(ul_cookies).decode('ascii')
#    print(ul_cookies)
    header={
            'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'accept-encoding':'gzip, deflate, br, zstd',
            'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8',
            'Cache-Control':'max-age=0',
#            'Connection':'keep-alive',
            'cookie':ul_cookies,
            'Host':'iq.ulprospector.com',
            'Refer':'https://iq.ulprospector.com/en',
            'Priority':'u=0,i',
            'sec-ch-ua':'"Google Chrome";v="129", "Not=A?Brand";v="8", "Chromium";v="129"',
            'sec-ch-ua-mobile:':'?0',
            'sec-ch-ua-platform':'"Linux"',
            'sec-fetch-dest':'document',
            'sec-fetch-mode':'navigate',
            'sec-fetch-site':'none',
            'sec-fetch-site':'same-origin',
            'sec-fetch-user':'?1',
            'sec-gpc':'1',
            'Upgrade-Insecure-Requests':'1',
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
            }

    data={
            'p':'10005,10048,10006,10047',
            'qm':'q:E485395'
            }

    new_header={
            'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Encoding':'gzip, deflate, br, zstd',
            'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8',
            'Cache-Control':'max-age=0',
            'cookie':ul_cookies,
            'priority':'u=0, i',
            'referer':'https://iq.ulprospector.com/en',
            'sec-ch-ua':'"Not A(Brand";v="8", "Chromium";v="132", "Google Chrome";v="132"',
            'sec-ch-ua-arch':'"x86"',
            'sec-ch-ua-full-version':'"132.0.6834.197"',
            'sec-ch-ua-full-version-list':'"Not A(Brand";v="8.0.0.0", "Chromium";v="132.0.6834.197", "Google Chrome";v="132.0.6834.197"',
            'sec-ch-ua-mobile':'?0',
            'sec-ch-ua-model':'""',
            'sec-ch-ua-platform':'"Windows"',
            'sec-ch-ua-platform-version':'"10.0.0"',
            'sec-fetch-dest':'document',
            'sec-fetch-mode':'navigate',
            'sec-fetch-site':'same-origin',
            'sec-fetch-user':'?1',
            'sec-gpc':'1',
            'upgrade-insecure-requests':'1',
            'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
            }

    
    while True:
        try:
    #    res=requests.get(url,headers=header,params=data,verify=False)#方法1
#            res=requests.get(url,headers=header,timeout=5,verify=False)#方法2
#            print(res)
#            html=res.text
#            selector=etree.HTML(html)#转化为lxml
#            return selector
            with httpx.Client(http2=False) as client:
                response=client.get(url)
                print('http2 code:',response.status_code)
#                print(response.read())
#                print(response.text)
#                html=res.text
#                selector=etree.HTML(html)#转化为lxml
#                return selector
            break
        except:
            print('连接超时，正在重新连接！')
            break

def basic_info(selector):#针对搜索结果进行处理并输出所需要的信息
    links=selector.xpath('//tbody/tr/td[1]/a/@href')#详细信息的链接
    number=selector.xpath('//tbody/tr/td[1]/a/span/text()')#黄卡号
    company=selector.xpath('//tbody/tr/td[2]/div/span/text()')#公司名称
    description=selector.xpath('//tbody/tr/td[4]/div/span/text()')#部件的描述信息
    res=[]

    print('-'*10+'以下是查询结果'+'-'*10)
    for index in range(0,len(links)):#格式化输出
        res.append([number[index],company[index],description[index],links[index]])#保存一组结果
        print(str(index)+'\t',number[index]+'\t',company[index]+'\t',description[index]+'\t')
    return res


def certificate(selector,ptf='No'):#查找证书中所有的型号
    models=[]
    if len(models)==0:
        print('mode 1')
        models=selector.xpath('//prodid/text()')
    if len(models)==0 or models[0]==', ' or models[0]=='\n':#增加逗号和回车的判断是为了应对抓取全为逗号或回车的情况下，继续向下抓取型号
        print('mode 2')
        models=selector.xpath('//prodid/b/text()')
    if len(models)==0 or models[0]==', ' or models[0]=='\n':#增加逗号和回车的判断是为了应对抓取全为逗号或回车的情况下，继续向下抓取型号
        print('mode 3')
        models=selector.xpath('//prodid/a/text()')
    if len(models)==0 or models[0]==', ' or models[0]=='\n':#增加逗号和回车的判断是为了应对抓取全为逗号或回车的情况下，继续向下抓取型号
        print('mode 4')
        models=selector.xpath('//prodid/b/a/text()')
    if ptf=='Yes':
        print(models)
    return models


def filters(models,wanted):#对所有型号进行过滤，是否有查找的型号
    '''
    models：一个列表
    wanted:字符串，想要查找的型号
    '''
    flag='red'#默认没有找到，当找到对应型号时，修改flag

    #精准查找型号
    for i in models:
#        if model==i or model.upper()==i:
        if wanted.upper()==i.upper():#统一大写来进行比较
            print(f'{i} vs {wanted},找到完全一致型号')
            flag='green'
            break
        elif wanted.upper() in i.upper():
            print(f'{i} vs {wanted},找到相似型号')
            flag='green'
            break
    if flag=='green':
        return flag
    else:
        for i in models:
            if wanted.split('-')[0] in i or wanted.split('-')[0].upper() in i:
                print('-'*20)
                print(f'{i} vs {wanted},找到部分相似型号')
                flag='yellow'
                break
        if flag=='yellow':
            return flag
        else:
#            print(f'{i} vs {wanted},没有找到型号')
            print(f'{wanted},没有找到型号')
            return flag


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



if __name__=='__main__':
    driver=Driver()
    Menu()
    driver.close()
    driver.quit()
#    driver=Driver()
#    print('engine start!')
#    while True:
#    start=time.time()
#    UL(driver)    
#        TUV(driver)
#    end=time.time()
#    print(end-start)

