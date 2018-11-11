import requests
from bs4 import BeautifulSoup as soup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
import math

def read_excel():
    df_out=pd.read_excel(open('D:\Python\Projects_for_Viper\Apple2\Working_URL.xlsx','rb'))
    return df_out

def read_excel2():
    df_out=pd.read_excel(open('D:\Python\Projects_for_Viper\Apple2\Working_URL.xlsx','rb'))
    url_line = df_out.columns.values
    return df_out, url_line
    
def read_txt():
    url=[]
    f=open("D:\\Python\\Projects_for_Viper\\Apple2\\URLS.txt",'r')
    f1=f.readlines()
    for x in f1:
        url.append(x)
    return url
    
def init(myurl):
	driver = webdriver.Chrome(executable_path='D:\Python\Projects_for_Viper\Apple\chromedriver.exe')
	driver.get(myurl)
	return driver

def findInfo(driver):
    Info = []
    xpath1 = '//*[@id="tabs_dimensionCapacity"]/fieldset/ul/li[1]/div[2]/div[1]/div/div[1]/span[2]'
    xpath2 = '//*[@id="tabs_dimensionCapacity"]/fieldset/ul/li[2]/div[2]/div[1]/div/div[1]/span[2]'
    xpath3 = '//*[@id="tabs_dimensionCapacity"]/fieldset/ul/li[3]/div[2]/div[1]/div/div[1]/span[2]'
    xpath = {xpath1,xpath2,xpath3}
    for x in xpath:
        element = WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.XPATH, x))
        )
        content = driver.find_elements_by_xpath(x)
        for a in content:
            Info.append(a.text)
    return Info

def region_Change(url):
	a=url.split('ttps://www.apple.com')[1]
	region=a.split('shop/buy-iphone/')[0]
	if region == '/':
		region = 'us'
	else:
		region = region[1:-1]
	return region

def generate_final_1(urls):
    start = datetime.datetime.now()
    Final = []
    G64B=[]
    G256B=[]
    G512B=[]
    Time=[]
    region=[]
    void=[]
    print("In total of ",len(urls)," website to be processed")
    counter = 0
    for url in urls:
        counter+=1
        print("In process of ", counter, " website")
        time = datetime.datetime.now()
        Time.append(time)
        driver = init(url)
        region.append(region_Change(url))
        content=findInfo(driver)
        G64B.append(content[0])
        G256B.append(content[1])
        G512B.append(content[2])
        driver.close()
        print(counter, " done")
    col = urls
    data=['Time','Region','64GB','256GB','512GB','']
    Final=[Time,region,G64B,G256B,G512B,void]
    sheet = pd.DataFrame(Final,index=data,columns=col)
	end=datetime.datetime.now()
	print('elapse: %s'%str(end-start))
    return sheet
    
def export_to_excel(table):
	writer = pd.ExcelWriter('D:\Python\Projects_for_Viper\Apple2\Working_URL.xlsx',engine='xlsxwriter')
	Workbook=writer.book
	table.to_excel(writer,sheet_name='Sheet1')
	worksheet=writer.sheets['Sheet1']
	writer.save()

if __name__ == '__main__':
    urls=read_txt()
    sheet = generate_final_1(urls)
    old=read_excel()
    sheet=old.append(sheet)
    export_to_excel(sheet)
    
    
    