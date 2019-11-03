#Authors: Fred, Aditya, Joshua

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from bs4 import BeautifulSoup as sp
from selenium import webdriver
import pandas as pd
import numpy as np
import time
import re

#Step ONE(1) +Downloading the reports file
olm_id = '******'
password = '******'

driver = webdriver.Chrome("/home/klenam/Downloads/WebDrivers/chromedriver_linux64/chromedriver")
driver.get("https://airtel.service-now.com")
driver.maximize_window()
driver.implicitly_wait(30)

#Locates username login field ,passes in username and submits
UsrName_field = driver.find_element_by_id("username")
UsrName_field.send_keys(olm_id)
UsrName_field.send_keys(Keys.RETURN)

driver.implicitly_wait(60)
b = driver.find_element_by_xpath('//*[@id="bySelection"]/div[2]/div/span')
b.click()

#Fill in Login Form
driver.implicitly_wait(60)
el_field = driver.find_element_by_id("userNameInput")
el_field.send_keys(olm_id)
text_field = driver.find_element_by_id("passwordInput")
text_field.send_keys(password)
text_field.send_keys(Keys.RETURN)
time.sleep(10)

#click Reports Tab
reports_tab = driver.find_element_by_css_selector("#gsft_nav > div > magellan-favorites-list > ul > li.sn-widget.ng-scope.selected.ui-sortable-handle > div > div:nth-child(1) > a")
reports_tab.click()

#switch to iframe
driver.implicitly_wait(60)
driver.switch_to.frame("gsft_main") #//*[@id="gsft_main"]
time.sleep(3)
SMP_UID = driver.find_element_by_xpath('//*[@id="report-list"]/tbody/tr[2]/td[4]/a')
SMP_UID.click()

#parsing iframe with BeautifulSoup
iframe_src = driver.page_source
soup= sp(iframe_src, "lxml")
inc_tab = soup.find_all("table", {"class":"data_list_table list_table table table-hover " })

driver.implicitly_wait(30)
time.sleep(3)
tableRow = soup.find('tr', {'id':re.compile(r"hdr_\w{32}") })['id']

ID = str(tableRow[4:])
action_chains = ActionChains(driver) 

time.sleep(10)

#locate NumberTab column by randomly generated IDs
NumberTab = driver.find_element_by_xpath('//*[@id="hdr_' + str(ID) +'"]/th[1]/span').click()
action_chains.context_click(NumberTab)
action_chains.perform() 
driver.implicitly_wait(30)

#find export in popup menu and click
export = driver.find_element_by_xpath('//*[@id="context_list_header' + str(ID) + '"]/div[10]').click()
driver.implicitly_wait(30)

#find Excel in export popup menu and click
exportToExcel = driver.find_element_by_xpath("//div[contains(text(), 'Excel (.xlsx)')]").click()
time.sleep(10)

#click on download
DownloadExcel = driver.find_element_by_xpath('//*[@id="download_button"]').click()

#Step TWO(2) +Cleaning downloaded file
time.sleep(10)
df = pd.read_excel('/home/klenam/Downloads/sc_req_item.xlsx')
df['Requested for'] = df['Requested for'].str.split('(', n=1, expand=True)
df1 = df.drop(['Business Justification', 'State'], axis=1) 
new = df1["Requested for"].str.split(" ", n = 1, expand = True)
df1["First Name"]= new[0]
df1["Last Name"]= new[1] 
df1['Last Name'] = df1['Last Name'].str.title()

df2 = df1.sort_values(by=['Req Subtype'])
df3 = df2.reset_index(drop=True)
df4 = df3.loc[df3['Req For'] == 'ID Creation']
df5 = df4.loc[df3['Stage'] == 'Implementation']
df_olm = df5.loc[df4['UID Type']== 'OLMID']
df_oth = df5.loc[df4['UID Type']== 'OTHER']

df_olm_isibm = df_olm[df_olm['Email'].str.contains("ibm")] 
df_olm_iscus = df_olm[~df_olm['Email'].str.contains("ibm")]
df_other_isibm = df_oth[df_oth['Email'].str.contains("ibm")] 
df_other_iscus = df_oth[~df_oth['Email'].str.contains("ibm")]



try:
    df_olm_iscus['Taging'] = df_olm_iscus.apply(lambda x:'744/C/%s//%s_%s_%s_%s_I_%s' % (x['User ID'],x['First Name'],x['Last Name'],x['User ID'],x['Req Subtype'],x['Number']),axis=1)
except ValueError:
    pass
try:
    df_olm_isibm['Taging'] = df_olm_isibm.apply(lambda x:'744/I/%s//%s_%s_%s_%s_I_%s' % (x['User ID'],x['First Name'],x['Last Name'],x['User ID'],x['Req Subtype'],x['Number']),axis=1)
except ValueError:
    pass
try:
    df_other_iscus['Taging'] = df_other_iscus.apply(lambda x:'744/C/%s//%s_%s_%s_%s_I_%s' % (x['User ID'],x['First Name'],x['Last Name'],x['User ID'],x['Req Subtype'],x['Number']),axis=1)
except ValueError:
    pass
try:
    df_other_isibm['Taging'] = df_other_isibm.apply(lambda x:'744/F/%s//%s_%s_%s_%s_I_%s' % (x['User ID'],x['First Name'],x['Last Name'],x['User ID'],x['Req Subtype'],x['Number']),axis=1)
except ValueError:
    pass

frames = [df_olm_iscus, df_olm_isibm, df_other_iscus, df_other_isibm]
df_orignal = pd.concat(frames, sort=False)
df_orignal.index = np.arange(1,len(df_orignal)+1)

#Save and export to ID_creation_Details.xlsx file
t = df_orignal.to_excel('./save_to/ID_creation_Details.xlsx')

data = pd.read_excel('./save_to/ID_cretion_Details.xlsx')
data

#Done




