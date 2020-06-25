# -*- coding: utf-8 -*-
import os
import sys
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from lxml import etree
from time import sleep
from numpy import record
import xlsxwriter
import os
import time
from columnate import lists

def init(drpath):
    return ""
    
def hmp5E(url,browser,uname,passww):
    wait=WebDriverWait(browser,10)
    startTime1 = time.time()
    browser.get(url)
    browser.refresh()
    browser.set_window_size(1600,900)

    element=browser.find_element_by_class_name("top-login")
    element.click()

    account = browser.find_element_by_name("account")
    account.click()
    account.send_keys(uname)

    passw = browser.find_element_by_name("password")
    passw.click()
    passw.send_keys(passww)

    login = browser.find_element_by_id("J_LoginSubmit")
    login.click()
    sleep(3)

    recordp1 = browser.find_element_by_id("person-info")
    ActionChains(browser).move_to_element(recordp1).perform()
    recordp2 = browser.find_element_by_class_name("player")
    recordp2.click()
    sleep(1)
    browser.execute_script("var q=document.documentElement.scrollTop=100000")

    while True:
        sleep(0.5)
        recordx = browser.find_element_by_xpath("//div[@id='match-tb']/table/tbody")
        print(recordx.text)
        print(is_disable("loadMatch"))
        if is_disable("loadMatch"):
            load_next()
        else:
            break

    html1 = browser.execute_script("return document.documentElement.outerHTML")
    element = etree.HTML(html1.replace('\n', '').replace('\r', ''))
    text = element.xpath("//div[@id='match-tb']/table/tbody/tr[position()>1]/td[position()>2 and position()<11]")
    text2 = element.xpath("//div[@id='match-tb']/table/tbody/tr[position()>1]/td[1]/span")
    text3 = element.xpath("//div[@id='match-tb']/table/tbody/tr[position()>1]/td[2]/span")
    
    workbook = xlsxwriter.Workbook(os.path.join(os.path.expanduser("~"), 'Desktop')+"/5e战绩.xlsx") 
    worksheet = workbook.add_worksheet()               
    title = [U'赛季',U'比赛类型',U'比赛时间',U'比赛耗时 ',U'地图',U'比分',U'杀敌',U'赛果 ',U'贡献值RWS',U'技术得分Rating']
    worksheet.write_row('A1',title)

    idx=0
    c_lists=0
    lists = [[] for i in range(int(len(text)/8))]
    list_in = []
    for i in text:
        idx=idx+1       
        if(idx%8==0):
            lists[c_lists].append(i.text)
            c_lists=c_lists+1
        else:
            if(i.text==None):
                i.text='胜'
            else: 
                pass
            lists[c_lists].append(i.text)
        
    c_text2=0
    for i in text2:
        lists[c_text2].insert(0,i.text)
        c_text2+=1
        
    c_text3=0    
    for i in text3:
        lists[c_text3].insert(1,i.text)
        c_text3+=1
    
    print("*********************************************")
    print(lists)
    print("列表获取")
    print("列表写入excel中")
    c_xlsx=1
    c_unpack_lists=0
    for i in range(int(len(text)/8)):
        c_xlsx+=1
        num0 = str(c_xlsx)
        row = 'A' + num0
        data = lists[c_unpack_lists]
        worksheet.write_row(row, data)
        c_unpack_lists+=1
    
    excel_end(worksheet,workbook)
    endTime1 = time.time()
    print("保存至桌面")
    print("耗时：",end=' ')
    print(endTime1-startTime1)
    
def load_next():
    loadMatch = browser.find_element_by_id("loadMatch")
    loadMatch.click()

def is_disable(l_id):
    try:
        elm = browser.find_element_by_id(l_id)
        if elm.is_displayed():
            return True
        else:
            return False
    except:
        return False
    
def excel_end(worksheet,workbook):
    worksheet.set_column('C:C',26)
    worksheet.set_column('E:E',11)
    worksheet.set_column('J:J',14)
    workbook.close()
    
if __name__ == '__main__':
    browser= webdriver.Chrome(executable_path=os.path.join(os.path.expanduser("~"), 'Desktop')+"\chromedriver.exe")
    print("将爬取www.5ewin.com中的个人战绩数据至excel")
    uname = str(input('输入账号（邮箱）:'))
    passww = str(input('输入密码:'))
    hmp5E("https://www.5ewin.com/",browser,uname,passww)