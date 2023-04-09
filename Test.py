
import requests
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup as bs
from pandas import DataFrame
from selenium import webdriver
import time
import openpyxl
import datetime
from datetime import timedelta


# < IP실시간뉴스 --------------------------------------------------------------------------------- 시작 1>


options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(3)
driver.get('http://www.ip-navi.or.kr/ipnavi/ref/boardList.navi?boardCode=B00017')
time.sleep(1)

def function_IP_news() :
    page = driver.page_source
    soup = bs(page, "html.parser")
    #헤드라인
    a = soup.select('body > div.container > div.board_wrap > table > tbody > tr> td.align_L > a')
    s1=[]
    for tag in a :
        s1.append(tag.text)
    s2=[]
    for tag in s1 :
        s2.append(tag.replace('\n','').replace('\t','').replace('  ','').replace('\xa0','').replace('\r',''))

    a= pd.DataFrame(s2)

    #링크
    b = soup.select('body > div.container > div.board_wrap > table > tbody > tr> td.align_L > a')
    s1=[]
    for tag in b :
        s1.append(tag["href"])

    s1=pd.DataFrame(s1)
    v_split = s1[0].str.split('(')
    s1= pd.DataFrame(v_split.str.get(1))

    v_split = s1[0].str.split(',')
    s1= pd.DataFrame(v_split.str.get(0))
    s2=[]
    for tag in s1[0] :
        s2.append(tag[1:-1])

    b= pd.DataFrame(s2)

    # Uploader
    s1=[]
    for tag in range(len(a)) :
        s1.append("IP_NAVI_NEWS")

    c= pd.DataFrame(s1)

    # 날짜
    d=soup.select('body > div.container > div.board_wrap > table > tbody > tr > td:nth-child(3)')
    d= pd.DataFrame(d)
    s1=[]
    time_format= "%Y-%m-%d"
    for tag in d[0]:
          s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    d= pd.DataFrame(s1)

    # 조회수
    s1=[]
    for tag in range(len(a)) :
        s1.append("None")

    e= pd.DataFrame(s1)


    # 분류
    s1=[]
    for tag in range(len(a)) :
        s1.append("None")

    f= pd.DataFrame(s1)

    #번호
    g=soup.select('body > div.container > div.board_wrap > table > tbody > tr > td:nth-child(1)')
    g= pd.DataFrame(g)

    #헤드라인 a, 링크 b, 업로더 c, 날짜 d, 조회수 e, 분류 f, 번호 g
    df=[]
    df= pd.merge(g,f, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,a, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,b, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,e, left_index=True, right_index=True, how='outer')
    df = pd.merge ( df , d , left_index=True , right_index=True , how='outer' )
    return df


from selenium.webdriver.common.keys import Keys

# 1번
IP_news=[]
IP_news = function_IP_news() # 1번

# 2번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[1]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 3번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[4]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 4번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[5]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 5번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[6]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 6번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[7]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 7번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[8]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 8번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[9]').send_keys(Keys.ENTER)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 9번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[9]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

# 10번

driver.find_element_by_xpath('//*[@id="divPagination"]/a[9]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_news()
IP_news=pd.concat([IP_news,a])

IP_news.columns=['번호', '분류', "제목", '링크',"작성자", "조회", "등록일"]
# IP_news.to_excel('./Text1.xlsx',sheet_name='전체')

# < IP_Desk_news --------------------------------------------------------------------------------- 종료 2>


options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(3)
driver.get('http://www.ip-navi.or.kr/ipnavi/ref/boardList.navi?boardCode=B00018')
time.sleep(1)


def function_IP_Desk_news() :
    page = driver.page_source
    soup = bs ( page , "html.parser" )

    # 헤드라인
    a = []
    a = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr > td.align_L>a' )
    s1 = []
    for tag in a :
        s1.append ( tag.text )

    s2 = []
    for tag in s1 :
        s2.append (
            tag.replace ( '\n' , '' ).replace ( '\t' , '' ).replace ( '  ' , '' ).replace ( '\xa0' , '' ).replace (
                '\r' , '' ) )

    a = pd.DataFrame ( s2 )

    # 링크
    b = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr > td.align_L >a ' )
    s1 = []
    for tag in b :
        s1.append ( tag["href"] )

    s1 = pd.DataFrame ( s1 )
    v_split = s1[0].str.split ( ',' )
    s1 = pd.DataFrame ( v_split.str.get ( 1 ) )
    v_split = s1[0].str.split ( ')' )
    s1 = pd.DataFrame ( v_split.str.get ( 0 ) )
    s2 = []
    for tag in s1[0] :
        s2.append ( tag[1 :-1] )
    s1 = []
    for tag in s2 :
        s1.append (
            "https://www.ip-navi.or.kr/ipnavi/ref/boardDetail.navi;jsessionid=2D70B0DCB138F2CF8F99F936475D17F6?boardCode=B00018&boardSeq=" + tag )

    b = pd.DataFrame ( s1 )

    # Uploader
    s1 = []
    for tag in range ( len ( a ) ) :
        s1.append ( "IP_DESK_NEWS" )
    c = pd.DataFrame ( s1 )
    # 날짜
    d = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr > td:nth-child(4)' )
    d = pd.DataFrame ( d )
    s1 = []
    time_format = "%Y-%m-%d"
    for tag in d[0] :
        s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    d = pd.DataFrame ( s1 )
    # 조회수
    e = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr > td:nth-child(3)' )
    e = pd.DataFrame ( e )
    s1 = []
    for tag in e[0] :
        s1.append ( int ( tag ) )

    e = pd.DataFrame ( s1 )
    s1 = []
    for tag in e[0] :
        s1.append ( int ( tag ) )

    e = pd.DataFrame ( s1 )

    # 분류
    s1 = []
    for tag in range ( len ( a ) ) :
        s1.append ( "NONE" )
    f = pd.DataFrame ( s1 )
    # 번호
    g = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr > td:nth-child(1)' )
    g = pd.DataFrame ( g )

    # 헤드라인 a, 링크 b, 업로더 c, 날짜 d, 조회수 e, 분류 f, 번호 g
    df = []
    df = g
    df['분류'] = 'IP_DESK_지재권뉴스'
    df = pd.merge ( df , a , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , b , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , c , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , e , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , d , left_index=True , right_index=True , how='outer' )
    return df


#1
IP_Desk_news=[]
IP_Desk_news = function_IP_Desk_news() # 1번

# 2번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[1]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_Desk_news()
IP_Desk_news=pd.concat([IP_Desk_news,a])

# 3번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[4]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_Desk_news()
IP_Desk_news=pd.concat([IP_Desk_news,a])

# 4번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[5]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_Desk_news()
IP_Desk_news=pd.concat([IP_Desk_news,a])

# 5번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[6]').send_keys(Keys.ENTER)
driver.implicitly_wait(3)
a = function_IP_Desk_news()
IP_Desk_news=pd.concat([IP_Desk_news,a])


IP_Desk_news.columns=['번호', '분류', "제목", '링크',"작성자", "조회", "등록일"]
# IP_Desk_news.to_excel('./Text2.xlsx',sheet_name='전체')