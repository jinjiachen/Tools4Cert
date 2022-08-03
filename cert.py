from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from lxml import etree
import os
import pdb
import urllib

def Menu():
    choice=input("请选择认证的类型：\n1.UL认证\n2.TUV莱茵认证\n3.VDE认证\n4.CSA认证\n5.TUV南德")
    if choice=='1':
        UL(driver)
    elif choice=='2':
        TUV(driver)


def Driver():
    #Chrom的配置
    options = webdriver.ChromeOptions()
    options.add_argument("--proxy-server=http://192.168.2.108:8889")
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

def get_html(url):
    start=time.time()
    driver.get(url)
#    time.sleep(1)
    html=driver.page_source
    selector=etree.HTML(html)
    end=time.time()
    print(end-start)
    return selector

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

