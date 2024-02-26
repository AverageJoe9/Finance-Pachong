from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from lxml import etree
import pandas as pd
import numpy as np
from time import sleep
import os

# 设置公司
companies = np.array(pd.read_excel('./Companies.xlsx', sheet_name='Sheet1')).tolist()

if not os.path.exists('temporary'):
    os.makedirs('temporary')
driver = webdriver.Edge(service=Service('others/msedgedriver.exe'))
sleep(1)
driver.maximize_window()
sleep(1)


def processRows(list, type, companyname, whnum, rows, webheader):
    for el in list:
        list2 = el.getchildren()  # 获取行
        list3 = list2[0].getchildren()  # 第0行先处理
        name = ((list3[0].getchildren())[0].getchildren())[-1].text.strip()
        for i in range(1, whnum):
            c = (list3[i].getchildren())[0].text.strip() if len(list3[i].getchildren()) != 0 else list3[i].text.strip()
            rows.append((companyname, type, name, webheader[i], c))
        if len(list2[1].getchildren()) != 0:  # 再看第1行的组
            list4 = list2[1].getchildren()
            rows = processRows(list4, type, companyname, whnum, rows, webheader)
    return rows


def processCompany(companyname, companyurl):
    driver.get(companyurl)
    sleep(3)
    header = ('Company', 'Type', 'Account', 'Period', 'Amount')
    rows = []
    for i in ['IS Annual', 'IS Quarterly', 'BS Annual', 'BS Quarterly', 'CF Annual', 'CF Quarterly']:
        if i == 'IS Annual':
            AorQ = 'Annual'
        elif i == 'IS Quarterly':
            AorQ = 'Quarterly'
            driver.find_element(by=By.XPATH,
                                value='//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button').click()  # to Quarterly
            sleep(3)
        elif i == 'BS Annual':
            AorQ = 'Annual'
            driver.find_element(by=By.XPATH,
                                value='//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[2]/a').click()  # to BS
            sleep(3)
        elif i == 'BS Quarterly':
            AorQ = 'Quarterly'
            driver.find_element(by=By.XPATH,
                                value='//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button').click()  # to Quarterly
            sleep(3)
        elif i == 'CF Annual':
            AorQ = 'Annual'
            driver.find_element(by=By.XPATH,
                                value='//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[3]/a').click()  # to CF
            sleep(3)
        elif i == 'CF Quarterly':
            AorQ = 'Quarterly'
            driver.find_element(by=By.XPATH,
                                value='//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button').click()  # to Quarterly
            sleep(3)
        else:
            AorQ = '-'
        if driver.find_element(by=By.XPATH,
                               value='//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button/div/span').text == 'Expand All':
            driver.find_element(by=By.XPATH,
                                value='//*[@id="Col1-1-Financials-Proxy"]/section/div[2]/button').click()  # Expand All
            sleep(1)
        html = driver.page_source
        sleep(1)
        webheader = ['Breakdown']
        webtable = etree.HTML(html).xpath('//*[@id="Col1-1-Financials-Proxy"]/section/div[3]/div[1]/div')[
            0].getchildren()
        whlist = webtable[0].getchildren()[
            0].getchildren()
        i = 0
        for el in whlist:
            if i == 0:
                i += 1
                continue
            webheader.append(el.getchildren()[0].text.strip())
        whnum = len(webheader)
        list = webtable[1].getchildren()
        rows = processRows(list, AorQ, companyname, whnum, rows, webheader)

    df = pd.DataFrame(rows, columns=header)
    df.to_excel('./temporary/{}.xlsx'.format(companyname), sheet_name='Sheet1', index=False)


for companyname, companyurl in companies:
    try:
        processCompany(companyname, companyurl)
    except Exception as result:
        print('')
        print('{} failed'.format(companyname))
        print(result)
        print('-' * 100)
    else:
        print('')
        print('{} finished'.format(companyname))
        print('-' * 100)
driver.close()

# 合并表
try:
    dflist = []
    for filename in os.listdir('temporary'):
        df = pd.read_excel('./temporary/{}'.format(filename), sheet_name='Sheet1', header=0)
        dflist.append(df)
    pd.concat(dflist).to_excel('others/FinancePachong.xlsx', sheet_name='Sheet1', index=False)
except Exception as result:
    print('')
    print('merging failed')
    print(result)
else:
    print('')
    print('merging finished')
