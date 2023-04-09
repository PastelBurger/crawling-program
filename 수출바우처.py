
import requests
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup as bs
from pandas import DataFrame
from selenium import webdriver
import time
import openpyxl
import datetime

# 함수

# < 수출바우처 공지 - -----------------------------------------------------------------------------------시작 >

driver = webdriver.Chrome()
driver.implicitly_wait(3)
driver.get('https://www.exportvoucher.com/portal/board/boardList?bbs_id=1&active_menu_cd=EZ005004000')
time.sleep(1)


def function_exportvoucher() :
    page = driver.page_source
    soup = bs ( page , "html.parser" )
    # 제목 생성
    a =  soup.select('body>div>section>div>article>form>table>tbody>tr>td>a')
    s1=[]
    for tag in a :
        s1.append(tag.text)

    a= pd.DataFrame(s1)

    # 번호 생성
    b =  soup.select('body>div>section>div>article>form>table>tbody>tr>td:nth-child(1)')

    s1=[]
    for tag in b :
        s1.append(tag.text)

    b= pd.DataFrame(s1)

    # 등록일 생성
    c =  soup.select('body>div>section>div>article>form>table>tbody>tr>td:nth-child(3)')
    s1=[]
    for tag in c :
        s1.append(tag.text)
    c= pd.DataFrame(s1)
    time_format = "%Y-%m-%d"
    s1=[]
    for tag in c[0]:
        s1.append ( datetime.datetime.strptime ( tag , time_format ) )
    c=pd.DataFrame(s1)

    # 조회수 생성
    d =  soup.select('body>div>section>div>article>form>table>tbody>tr>td:nth-child(4)')
    s1=[]
    for tag in d :
        s1.append(tag.text)
    d = pd.DataFrame(s1)

    df= pd.merge(a,b, left_index=True, right_index=True, how='outer')
    df['Uploader']= ''
    df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
    df['Link']= 'https://www.exportvoucher.com/portal/board/boardList'
    df['Site']= '수출바우처_공지'
    df['Deadline']= ''
    df= pd.merge(df,d, left_index=True, right_index=True, how='outer')

    df.columns=['Headline', "Number", 'Uploader', "Date1", "Link", "Site", 'Deadline', 'Condition']
    df=df[["Site", "Number", 'Headline', "Link", 'Uploader',"Date1", 'Deadline', 'Condition' ]]
    return df


# 1번
Export_voucher=[]
Export_voucher = function_exportvoucher() # 1번

# 2번
driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[2]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 3번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[4]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 4번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[5]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 5번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[6]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 6번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[7]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 7번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[8]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 8번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[9]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 9번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[10]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

# 10번

driver.find_element_by_xpath('//*[@id="contents"]/form/div[3]/div/a[11]').click()
driver.implicitly_wait(10)
a = function_exportvoucher()
Export_voucher=pd.concat([Export_voucher,a])

s1=Export_voucher.drop_duplicates(['Headline', 'Date1'])
s1= s1.reset_index(drop=True)

#조회수 숫자로 변경
s2=[]
for tag in s1.Condition:
    s2.append(int(tag))
s2=pd.DataFrame(s2)
s2.columns=['Condition']
s1.Condition=s2.Condition

#중요 삽입
s2=s1.Number
s2=pd.DataFrame(s2)

for tag in range(len(s2)) :
    if s2['Number'][tag].isdecimal() == False :
        s2['Number'][tag] = '중요'

s2.columns=['Number']
s1.Number=s2.Number

Export_voucher = s1

# < 수출바우처 공지 - -----------------------------------------------------------------------------------시작 >
