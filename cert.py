from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from lxml import etree
import os
import pdb
import urllib
import requests
requests.packages.urllib3.disable_warnings
import warnings
warnings.filterwarnings("ignore")

def Menu():
    choice=input("请选择认证的类型：\n1.UL认证\n2.TUV莱茵认证\n3.VDE认证\n4.CSA认证\n5.TUV南德\n11.UL查询（模拟请求)")
    if choice=='1':
        UL(driver)
    elif choice=='2':
        TUV(driver)
    elif choice=='11':
        while True:
            ul_no=input('请输入需要查询的关键字:')
            url='https://iq.ulprospector.com/en/_/_results?p=10005,10048,10006,10047&qm=q:'+ul_no
            res_basic=ul_search(url)
            links=basic_info(res_basic)
            no=input('请选择对应的部件序号:')
            res_details=ul_search('https://iq.ulprospector.com'+links[int(no)])
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
    options.add_argument("--headless")
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
    user.send_keys('shelway.wu@intertek.com')
    password=driver.find_element(By.ID,'loginPass')
    password.send_keys('Ul123456')
    check=driver.find_element(By.XPATH,'//input[@id="cnConsent"]')
    check.click()
    login=driver.find_element(By.ID,'main-login')
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
#    url='https://iq.ulprospector.com/en/_/_results?p=10005,10048,10006,10047&qm=q:'+ul_no
    header={
            'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8',
            'Cache-Control':'max-age=0',
            'Connection':'keep-alive',
            'cookie':r'OptanonAlertBoxClosed=2022-06-28T01:31:48.056Z; chdc_prod=1; ASP.NET_SessionId=cip5yrlw1zbh0blm15oydov5; SERVERID=iis03; ii=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IjI4NjAyNDkxNzg4NzQ4NTkxMzc0NTczMTM5MDM0MTEyIiwiZSI6InNoZWx3YXkud3VAaW50ZXJ0ZWsuY29tIiwibiI6IlNoZWx3YXkgV3UiLCJ1cmkiOiJodHRwczovL2NvcmUudWxwcm9zcGVjdG9yLmNvbS91c2Vycy8yODYwMjQ5MTc4ODc0ODU5MTM3NDU3MzEzOTAzNDExMiIsInNpZCI6ImM2NTJmNWU4LTAxZTgtNGUzZC1iNDFiLTkzMGU0ZjY5MTNjYyIsIm5iZiI6MTY3NTc1OTcxMCwiZXhwIjoxNjc4MzUxNzEwLCJpYXQiOjE2NzU3NTk3MTAsImlzcyI6IlVMUHJvc3BlY3RvciIsImF1ZCI6Imh0dHA6Ly93d3cudWxwcm9zcGVjdG9yLmNvbSJ9.unsXQwFnWVvhIPzN0y-fqst_wDAh9lVEau5eFilVFz0; __cfruid=73a182391e73501b4f3d60e8ee4b41ea0966d488-1675910715; ii_sess=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzaWQiOiJjNjUyZjVlOC0wMWU4LTRlM2QtYjQxYi05MzBlNGY2OTEzY2MiLCJpZCI6IjI4NjAyNDkxNzg4NzQ4NTkxMzc0NTczMTM5MDM0MTEyIiwibmJmIjoxNjc1OTExNDg4LCJleHAiOjE2NzU5MTE2NjgsImlhdCI6MTY3NTkxMTQ4OCwiaXNzIjoiVUxQcm9zcGVjdG9yIiwiYXVkIjoiaHR0cDovL3d3dy51bHByb3NwZWN0b3IuY29tIn0.oJa00w9ZTAMfQ6b8leVMaLYh3s5758zT4qExs6v_TnA; OptanonConsent=isGpcEnabled=1&datestamp=Thu+Feb+09+2023+10:58:19+GMT+0800+(%E4%B8%AD%E5%9B%BD%E6%A0%87%E5%87%86%E6%97%B6%E9%97%B4)&version=202212.1.0&hosts=&groups=C0001:1,C0002:0,C0003:0,C0004:0&consentId=24e9e77b-e2ed-42ee-9370-c6dbdd4dbe83&geolocation=CN;SH&isIABGlobal=false&interactionCount=0&landingPath=NotLandingPage&AwaitingReconsent=false; pro_iq=projwt=h4KhJbOm7z3KEy6aP7WJIaRDkKe5VAIKlEYbfzLsqCCQ3qzccogOeO9eCXQOpkNcBKKfDcsYO-brXwconhUxsgg4Lz2aj0eOSFePceFrtyDhlC8VptqHe2ygElTa9vsdnGlnvvCuD4keQuQFxonH25EYCUXU591nSs25JEoRfQ7WF5ncWZlXPXGZPjvWEb89uXxy26aI9De8Q95-CPXOV4WHtdldeRWee-d6_GFzGeglD5g1HOJZA9nkEnDIlqP2mGiynSsZsiWMrSoT6CQ0_wkpEDb1jvADJkyYtxNdXOxLc_z7JHVouVAuSLAFzU88-4NHPlREfUo76TBHZBpmDt9Zs0PwJwOKun0PIJoaF2JQnY0xDvomr5hrzvG0JLp_qJo7ecSJr_UexcQ5b58VOZZOfTc1',
            'Host':'iq.ulprospector.com',
            'Refer':'https://iq.ulprospector.com/en',
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1474.0'
            }

    data={
            'p':'10005,10048,10006,10047',
            'qm':'q:E485395'
            }

#以下两个方法2选1，get方法直接用url比较方便
    while True:
        try:
    #    res=requests.get(url,headers=header,params=data,verify=False)#方法1
            res=requests.get(url,headers=header,timeout=5,verify=False)#方法2
            html=res.text
            selector=etree.HTML(html)#转化为lxml
            return selector
            break
        except:
            print('连接超时，正在重新连接！')

def basic_info(selector):#针对搜索结果进行处理并输出所需要的信息
    links=selector.xpath('//tbody/tr/td[1]/a/@href')#详细信息的链接
    number=selector.xpath('//tbody/tr/td[1]/a/span/text()')#黄卡号
    company=selector.xpath('//tbody/tr/td[2]/div/span/text()')#公司名称
    description=selector.xpath('//tbody/tr/td[4]/div/span/text()')#部件的描述信息

    print('-'*10+'以下是查询结果'+'-'*10)
    for index in range(0,len(links)):#格式化输出
        print(str(index)+'\t',number[index]+'\t',company[index]+'\t',description[index]+'\t')
    return links


def certificate(selector):#查找证书中所有的型号
    models=[]
    if len(models)==0:
        print('mode 1')
        models=selector.xpath('//prodid/text()')
    if len(models)==0:
        print('mode 2')
        models=selector.xpath('//prodid/b/text()')
    if len(models)==0:
        print('mode 3')
        models=selector.xpath('//prodid/a/text()')
#        models=selector.xpath('//a[@style="text-decoration: none;"]/text()')
    print(models)
    return models


def filters(models,model):#对所有型号进行过滤，是否有查找的型号
    '''
    models：一个列表
    model:字符串，想要查找的型号
    '''
    for i in models:
        if model==i or model.upper()==i:
            print('找到型号：',i)
            return 'green'
            break
        elif model.split('-')[0] in i or model.split('-')[0].upper() in i:
            print('-'*20)
            print('找到相似型号：',i)
            return 'yellow'
#        else:
#            print('没有找到对应型号')



if __name__=='__main__':
#    driver=Driver()
    Menu()
#    driver.close()
#    driver.quit()
#    driver=Driver()
#    print('engine start!')
#    while True:
#    start=time.time()
#    UL(driver)    
#        TUV(driver)
#    end=time.time()
#    print(end-start)

