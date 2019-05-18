#!/usr/bin/env python
# coding: utf-8

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from bs4 import BeautifulSoup as sp
from selenium import webdriver
from datetime import timedelta
from datetime import datetime
import pandas as pd
import time
import re

rota_path = r'windows_rota_updated.xlsx'

daytime = datetime.now()
cur_t = daytime.strftime('%H')
cur_t = int(cur_t)
#cur_t = 4
n_1 = daytime - timedelta(days=1)

cur_d = int(daytime.strftime('%d'))

pre_d = int(n_1.strftime('%d'))
#pre_d = 23
df = pd.read_excel(rota_path, index_col=cur_d)

m = list(range(8,11))
m_g = list(range(11,15))
m_g_e = list(range(15,17))
g_e = list(range(17,19))
e = list(range(19,22))
n = [22,23]
n1 = list(range(0,8))

if cur_t in m:
    user = df.loc['M']
elif cur_t in m_g:
    user = df.loc[['M','G']]
elif cur_t in m_g_e:
    user = df.loc[['M','G','E']]
elif cur_t in g_e:
    user = df.loc[['G','E']]
elif cur_t in e:
    user = df.loc['E'] 
elif cur_t in n:
    user = df.loc['N']
elif cur_t in n1:
    df = pd.read_excel(rota_path, index_col=pre_d)
    user = df.loc['N']
else:
    print('Item not found')
user_name = user.Name.tolist()
print(user_name)

#Portal section starts here
olm_id = 'A1yizb18'
password = 'noida@35'

driver = webdriver.Chrome("/home/klenam/Downloads/WebDrivers/chromedriver_linux64/chromedriver")
driver.get("https://airtel.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_fixed_query%3Dactive%3Dtrue%5Eassignment_groupDYNAMICd6435e965f510100a9ad2572f2b47744%26%3D%26sysparm_query%3DstateNOT%2520IN6,7,8%5EEQ%26sysparm_userpref_module%3Db55fbec4c0a800090088e83d7ff500de%26sysparm_clear_stack%3Dtrue")
driver.maximize_window()
driver.implicitly_wait(10)

email_field = driver.find_element_by_id("username")
email_field.send_keys("A1yizb18")
email_field.send_keys(Keys.RETURN)

driver.implicitly_wait(60)
b = driver.find_element_by_xpath('//*[@id="bySelection"]/div[2]/div/span')
b.click()

driver.implicitly_wait(60)
el_field = driver.find_element_by_id("userNameInput")
el_field.send_keys(olm_id)
text_field = driver.find_element_by_id("passwordInput")
text_field.send_keys(password)
text_field.send_keys(Keys.RETURN)

driver.implicitly_wait(60)
#find and select Incidents tab and click on it
incidents = driver.find_element_by_xpath('/html/body/div/div/div/nav/div/div[3]/div/div/magellan-favorites-list/div[2]/div[3]/div/a').click()

driver.implicitly_wait(60)
#Switch to iframe on webpage
driver.switch_to.frame("gsft_main")
time.sleep(10)

#driver.implicitly_wait(60)
iframe = driver.page_source
src = sp(iframe, 'lxml')

#list of rows numbers we are selecting
ListOfRows = []

#parse iframe and get the length of the number table rows
reg_id = re.compile(r"^row_incident.+$")
soup = src.find_all('tr', {"id": reg_id})

for tr in soup:
    x = int(soup.index(tr)) + 1
    souptr = sp(tr.encode("utf-8"), 'lxml')
    name = re.compile(r'^\w+ \w+\(\w+\)$')
    text = souptr.find_all('a', {"class":"linked"})
    
    cnfrm = False
    for links in text:
        if name.search(links.text) != None:
            cnfrm = True            
        else:
            continue
    if cnfrm != True:
        ListOfRows.append(int(x))
    else:
        continue

#Funtion here takes a row number(i) and username
def AssignUserName(i, username):
    #select Assigned to column
    table_row  = driver.find_element_by_xpath('//table[@id="incident_table"]/tbody/tr[{0}]/td[7]'.format(str(i)))
    table_row.click()
    time.sleep(2)
    
    #double click on Assigned to coulmn
    actionChains = ActionChains(driver)
    actionChains.double_click(table_row).perform()  
    driver.implicitly_wait(10)
    
    #open popup menu and send text
    popup = driver.find_element_by_css_selector("#sys_display\.LIST_EDIT_incident\.assigned_to").send_keys(username)
    time.sleep(3)

    #click green tick to send
    #cmmd = driver.find_element_by_css_selector("#cell_edit_ok")
    
    #close the popup menu
    cmmd = driver.find_element_by_css_selector("#cell_edit_cancel").click()  

#pass-in username for each table row using this algorithim below

cnt = 0
for i in range(len(ListOfRows)):
    try:
        if cnt > int(len(user_name) -  1):
            cnt = 0
            AssignUserName(ListOfRows[i], user_name[cnt])
        else:
            AssignUserName(ListOfRows[i], user_name[cnt])
        cnt += 1
    except NoSuchElementException as problem:
        print(problem,'\nOn row',str(ListOfRows[i]), "\n")
        continue
print("\nSuccessful\n")


# In[ ]:





# In[ ]:




