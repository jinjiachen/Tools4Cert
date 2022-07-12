#coding=utf8

from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from lxml import etree

#Chrom的配置
options = webdriver.ChromeOptions()
options.add_argument("--proxy-server=http://192.168.2.108:8889")
#options.add_argument("--no-proxy-server")
#options.add_argument("--headless")
#options.add_argument('user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"')
#options.add_argument('user-agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36"')
#options.add_argument('user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

#Chrome的驱动和路径
path="C:\Program Files\Google\Chrome\Application\chrome.exe"
#driver=webdriver.Chrome(chrome_options=options,executable_path=path)
#driver=webdriver.Chrome(path,chrome_options=options)
driver=webdriver.Chrome(chrome_options=options)
#driver.set_page_load_timeout(10)
#driver.set_script_timeout(10)


time.sleep(3)
print('start already')
#driver.get('https://www.baidu.com')
driver.get('https://iq.ulprospector.com/en/')
#user=driver.find_element_by_id('email')
user=driver.find_element(By.ID,'email')
user.send_keys('shelway.wu@intertek.com')
#password=driver.find_element_by_id('loginPass')
password=driver.find_element(By.ID,'loginPass')
password.send_keys('Ul123456')
#login=driver.find_element_by_id('main-login')
login=driver.find_element(By.ID,'main-login')
login.submit()
print('login')
#search=driver.find_element_by_id('q2')
search=driver.find_element(By.ID,'q2')
search.send_keys('E231292')
search.submit()
#print(driver.find_element(By.XPATH,'//td[@class="entry"]').text)
html=driver.page_source
selector=etree.HTML(html)
txt=selector.xpath('//tr[@class=" "]/td[2]/div/span/text()')
#cookies=driver.get_cookies()
#url=driver.current_url()
print(txt)
#print(driver.page_source)
#while True:
#    E=input('Please input the control number:')
#    if E=='exit':
#        break
#    else:
#        driver.get('https://iq.ulprospector.com/en/_?p=10005,10048,10006,10047&qm=q:'+E)
#        print('done')
#driver.get('https://www.ul.com')
#driver.get('https://www.taobao.com')
time.sleep(10)
driver.close()
driver.quit()


