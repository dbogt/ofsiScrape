# -*- coding: utf-8 -*-
"""
Created on Wed Jul  8 17:30:50 2020

@author: Bogdan Tudose
"""

#%% Import Packages
from selenium import webdriver
import pandas as pd
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By

#%% Open Website w Selenium
url = "https://www.osfi-bsif.gc.ca/Eng/wt-ow/Pages/FINDAT.aspx"
driver = webdriver.Chrome('C:\webdrivers\chromedriver.exe')
driver.maximize_window()
driver.get(url)

#%% Select Date
seq = driver.find_elements(By.TAG_NAME, 'iframe')
iframe = driver.find_elements(By.TAG_NAME, 'iframe')[0]
driver.switch_to.frame(iframe)

#Finding the Date Dropdown
xpath = '//*[@id="DTIWebPartManager_gwpDTIBankControl1_DTIBankControl1_dtiReportCriteria_monthlyDatesDropDownList"]'
select_from = Select(driver.find_element(By.XPATH, xpath))
options = select_from.options
allValues = []
for option in options:
    print(option.text,option.get_attribute('value'))
    allValues.append(option.get_attribute('value'))

#%% Generic Function
def loadNewMonth(selectValue):    
    #Switch back to the main tab
    driver.switch_to.window(driver.window_handles[0])
    iframe = driver.find_elements(By.TAG_NAME, 'iframe')[0]
    driver.switch_to.frame(iframe)
    
    #Grab the Date dropdown and select custom month
    # dropdowns = driver.find_elements(By.TAG_NAME, 'select')
    
    xpath = '//*[@id="DTIWebPartManager_gwpDTIBankControl1_DTIBankControl1_dtiReportCriteria_monthlyDatesDropDownList"]'
    select_from = Select(driver.find_element(By.XPATH, xpath))
    select_from.select_by_value(selectValue) #last month

    submit = driver.find_element(By.NAME,'DTIWebPartManager$gwpDTIBankControl1$DTIBankControl1$submitButton')
    submit.click()
    
    #Switch to new open window/tab
    driver.switch_to.window(driver.window_handles[1])
    newURL = driver.current_url
    dfs = pd.read_html(newURL)
    
    #Add the data frames to the dictionary
    allMonths[selectValue] = dfs

#%% Select Custom Months
allMonths = {}
months = ["1 - 2020", "2 - 2020", "3 - 2020"]
# for month in months:
#     loadNewMonth(month)

print(allValues)

for month in allValues[0:40]:
    loadNewMonth(month)

#%% Extract all balance sheets
balanceSheets = {}
for month in allMonths.keys():
    mDFs = allMonths[month]
    df = mDFs[1]
    df.rename(columns=df.iloc[0], inplace = True)
    df.drop([0], axis=0, inplace = True) #delete row 0, axis=1 for columns

    df.loc[df['Foreign Currency'] == df['Section I - Assets'], 'Foreign Currency'] = ""
    df.loc[df['Total Currency'] == df['Section I - Assets'], 'Total Currency'] = ""
    balanceSheets[month] = df

#%% Extract all balance sheets
table2s = {}
for month in allMonths.keys():
    mDFs = allMonths[month]
    df = mDFs[7]
    # df.rename(columns=df.iloc[0], inplace = True)
    # df.drop([0], axis=0, inplace = True) #delete row 0, axis=1 for columns

    # df.loc[df['Foreign Currency'] == df['Section II - Liabilities'], 'Foreign Currency'] = ""
    # df.loc[df['Total'] == df['Section II - Liabilities'], 'Total Currency'] = ""
    table2s[month] = df

#%% Save Excel File of Balance Sheets
with pd.ExcelWriter('Output/ofsi.xlsx') as writer:
    # table1.to_excel(writer, sheet_name="Assets", index=False)
    # table2.to_excel(writer, sheet_name="B/S", index=False)
    
    for month in balanceSheets.keys():
        monthDFs = allMonths[month]
        saveDF = monthDFs[1]
        saveDF.to_excel(writer, sheet_name=month, index=False)
    writer.save()
