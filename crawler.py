from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import time
import json

url = 'https://proteacher.moe.edu.tw/#login'
year_list = list(range(95, 111))

ACCOUNT = 'utcollege'
PASSWORD = 'i-love-tim99'

Wait_time = 5
anchor = None
data = []
try:
    driver = webdriver.Chrome()
except Exception as e:
    print(e)
    os.system('pause')

driver.get(url)

acc = (By.XPATH,('//*[@id="root"]/div/div[3]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/input'))
acc_input = WebDriverWait(driver,Wait_time).until(EC.presence_of_element_located(acc),"找不到指定的元素")
acc_input.send_keys(ACCOUNT)

pwd = (By.XPATH,'//*[@id="root"]/div/div[3]/div/div/div/div/div/div[2]/div/div/div[2]/div[3]/input')
pwd_input = WebDriverWait(driver,Wait_time).until(EC.presence_of_element_located(pwd),"找不到指定的元素")
pwd_input.send_keys(PASSWORD)

submit = (By.XPATH,'//*[@id="root"]/div/div[3]/div/div/div/div/div/div[2]/div/div/div[2]/div[4]/button')
sub_input = WebDriverWait(driver,Wait_time).until(EC.presence_of_element_located(submit),"找不到指定的元素")
sub_input.click()
time.sleep(3)
driver.refresh()

start = time.time()
for year in range(95,111):
    while True:
        if anchor == None:
            try:
                driver.get('https://proteacher.moe.edu.tw/api/v1/certificates?type=begin&year='+str(year)+'&limit=50')
                driver.implicitly_wait(1)
                pre = driver.find_element_by_tag_name("pre").text
            except:
                driver.refresh()    
            data = json.loads(pre)
        else:
            try:
                driver.get('https://proteacher.moe.edu.tw/api/v1/certificates?type=begin&year='+str(year)+'&limit=50&anchor='+str(anchor))
                driver.implicitly_wait(1)
                pre = driver.find_element_by_tag_name("pre").text
            except:
                driver.refresh()           
            data = json.loads(pre)
        if 'anchor' in data:
            anchor = data['anchor']
        else:
            anchor = None
            break
        # 
end = time.time()       
print(end - start)





