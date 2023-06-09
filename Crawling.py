
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

def func_to_date(list) :
    time_format = "%Y-%m-%d"
    ss=[]
    for tag in list :
        ss.append ( datetime.datetime.strptime(tag, time_format ) )
    return ss

def func_create_Condition(list) :
    C_time = datetime.datetime.now ()
    for tag in list:
        if C_time <= tag :
            Condition.append("공고중")
        else:
            Condition.append("공고종료")
    return Condition



# <IP_NAVI_사업공고&입찰공고------------------------------------------------------------------------------------시작 1>

driver = webdriver.Chrome()
driver.implicitly_wait(3)
driver.get('http://www.ip-navi.or.kr/ipnavi/board/boardList.navi?boardCode=B00001')
time.sleep(1)

page = driver.page_source
soup = bs(page, "html.parser")


def function_IP_NAVI_사업공고() :

    page = driver.page_source
    soup = bs(page, "html.parser")

    # 헤드라인
    a = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr> td.align_L > a' )
    s1 = []
    for tag in a :
        s1.append ( tag.text )
    s2 = []
    for tag in s1 :
        s2.append (
            tag.replace ( '\n' , '' ).replace ( '\t' , '' ).replace ( '  ' , '' ).replace ( '\xa0' , '' ).replace ( '\r' ,
                                                                                                                    '' ) )

    a = pd.DataFrame ( s2 )

    # 링크
    b = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr> td.align_L > a' )
    s1 = []
    for tag in b :
        s1.append ( tag["href"] )

    s1 = pd.DataFrame ( s1 )
    v_split = s1[0].str.split ( '(' )
    s1 = pd.DataFrame ( v_split.str.get ( 1 ) )

    v_split = s1[0].str.split ( ',' )
    s1 = pd.DataFrame ( v_split.str.get ( 1 ) )
    s2 = []
    for tag in s1[0] :
        s2.append ( tag[1 :-2] )
    s3=[]
    for tag in s2:
        s3.append ( "http://www.ip-navi.or.kr/ipnavi/board/boardDetail.navi?boardCode=B00001&boardSeq="+tag )

    b = pd.DataFrame ( s3 )

    # Uploader
    s1 = []
    for tag in range ( len ( a ) ) :
        s1.append ( "IP_NAVI_사업공고" )

    c = pd.DataFrame ( s1 )

    # 날짜
    d = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr> td:nth-child(4)' )
    d = pd.DataFrame ( d )
    s1 = []
    time_format = "%Y-%m-%d"
    for tag in d[0] :
        s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    d = pd.DataFrame ( s1 )

    # 조회수

    e = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr> td:nth-child(3)' )
    s1 = []
    for tag in e :
        s1.append ( tag.text )

    s2 = []
    for tag in s1 :
        s2.append ( int(tag) )

    e = pd.DataFrame ( s2 )

    # 분류
    s1 = []
    for tag in range ( len ( a ) ) :
        s1.append ( "None" )

    f = pd.DataFrame ( s1 )

    # 번호
    g = soup.select ( '#frm > div > div.board_wrap > table > tbody > tr > td:nth-child(1)' )
    g = pd.DataFrame ( g )


    # 헤드라인 a, 링크 b, 업로더 c, 날짜 d, 조회수 e, 분류 f, 번호 g
    # Site :  c , Number: g, Headline: a, Link : b, Uploader : f, Date1 : d, Deadline : f, Condition : f

    df = []
    df = pd.merge ( c , g , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , a , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , b , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , f , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , d , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , f , left_index=True , right_index=True , how='outer' )
    df = pd.merge ( df , f , left_index=True , right_index=True , how='outer' )
    return df

# 1번 ~ 2번
IP_NAVI_사업공고 = function_IP_NAVI_사업공고() # 1번
time.sleep(1)
element = driver.find_element_by_xpath('//*[@id="divPagination"]/a[1]')
driver.execute_script("arguments[0].click();", element)
time.sleep(1)
a = function_IP_NAVI_사업공고()
IP_NAVI_사업공고=pd.concat([IP_NAVI_사업공고,a]) #2번 합치기
# 3번
element = driver.find_element_by_xpath('//*[@id="divPagination"]/a[4]')
driver.execute_script("arguments[0].click();", element)
driver.implicitly_wait(3)
time.sleep(1)
a = function_IP_NAVI_사업공고()
IP_NAVI_사업공고=pd.concat([IP_NAVI_사업공고,a])

# 4번
element = driver.find_element_by_xpath('//*[@id="divPagination"]/a[5]')
driver.execute_script("arguments[0].click();", element)
driver.implicitly_wait(3)
time.sleep(1)
a = function_IP_NAVI_사업공고()
IP_NAVI_사업공고=pd.concat([IP_NAVI_사업공고,a])
# 5번
element = driver.find_element_by_xpath('//*[@id="divPagination"]/a[6]')
driver.execute_script("arguments[0].click();", element)
driver.implicitly_wait(3)
time.sleep(1)
a = function_IP_NAVI_사업공고()
IP_NAVI_사업공고=pd.concat([IP_NAVI_사업공고,a])


IP_NAVI_사업공고.columns=["Site","Number",'Headline',"Link",'Uploader',"Date1", 'Deadline', 'Condition']



# < IP_NAVI_사업공고&입찰공고 - -----------------------------------------------------------------------------------완료 >


# < RIPC_입찰공고 - -----------------------------------------------------------------------------------시작 2>

driver = webdriver.Chrome()
driver.implicitly_wait(3)
driver.get('https://pms.ripc.org/pms/biz/applicant/notice/list.do')
time.sleep(1)

page = driver.page_source
soup = bs(page, "html.parser")
Data1 = soup.select(' div.table_area > table > tbody > tr > td')

def function_RIPC_사업공고() :
    page = driver.page_source
    soup = bs(page, "html.parser")
    Data1=soup.select(' div.table_area > table > tbody > tr > td')

    # 링크 생성
    d = [tag for tag in Data1 if Data1.index(tag) % 6 == 3]
    link = []
    for tag in d:
        link.append("https://pms.ripc.org" + tag['onclick'][15:-1])
    link = pd.DataFrame(link)
    # 다른항목 생성
    a = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 6==0)
    b = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 6==1)
    c = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 6==2)
    d = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 6==3)
    e = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 6==4)
    f = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 6==5)
    # 날짜나누기
    e.columns = ['Date']
    v_split = e.Date.str.split('~')
    v_Date1 = pd.DataFrame(v_split.str.get(0))
    v_Date2 = pd.DataFrame(v_split.str.get(1))
    time_format = "%Y-%m-%d %H:%M "
    s1 = []
    s2 = []
    time_format = "%Y-%m-%d %H:%M "
    for tag in v_Date1["Date"] :
        s1.append(tag.replace('\n','').replace('\t',''))
    s1= pd.DataFrame ( s1 )
    for tag in s1[0]:
        s2.append ( datetime.datetime.strptime ( tag , time_format ) )
    v_Date1=pd.DataFrame(s2)
    s1 = []
    s2 = []
    time_format = " %Y-%m-%d %H:%M"
    for tag in v_Date2["Date"] :
        s1.append(tag.replace('\n','').replace('\t',''))
    s1 = pd.DataFrame ( s1 )
    for tag in s1[0]:
        s2.append ( datetime.datetime.strptime ( tag , time_format ) )
    v_Date2=pd.DataFrame(s2)

    df= pd.merge(a,b, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,d, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,v_Date1, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,v_Date2, left_index=True, right_index=True, how='outer')
    df= pd.merge(df,f, left_index=True, right_index=True, how='outer')
    df = pd.merge(df, link, left_index=True, right_index=True, how='outer')
    df['Site'] = 'IP_RIPC_입찰공고'
    return df

# 1번 ~ 2번
RIPC_사업공고 = function_RIPC_사업공고() # 1번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[1]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a]) #2번 합치기
# 3번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[4]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])

# 4번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[5]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])
# 5번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[6]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])
# 6번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[7]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])
# 7번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[8]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])
# 8번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[9]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])
# 9번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[9]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])
# 10번
driver.find_element_by_xpath('//*[@id="divPagination"]/a[9]').click()
driver.implicitly_wait(3)
time.sleep(1)
a = function_RIPC_사업공고()
RIPC_사업공고=pd.concat([RIPC_사업공고,a])

RIPC_사업공고.columns=['AA', 'Uploader', "Number", "Headline", "Date1", "Deadline", 'Condition', "Link","Site"]
RIPC_사업공고=RIPC_사업공고[["Site","Number",'Headline',"Link",'Uploader',"Date1", 'Deadline', 'Condition' ]]


# End_result.to_excel('./Test.xlsx', sheet_name='Sheet1')

# < RIPC_입찰공고 - -----------------------------------------------------------------------------------완료 >

# < 전략개발원_사업공고 - -----------------------------------------------------------------------------------시작 3>

url="https://biz.kista.re.kr/ippro/com/iprndMain/selectBusinessAnnounceList.do?bbsType=bs"
page = requests.get(url)
soup = bs(page.text, "html.parser")

Data1=soup.select('tbody >tr >td')


a = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 5==1)
b = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 5==2)
c = pd.DataFrame(tag for tag in Data1 if Data1.index(tag) % 5==3)
d = pd.DataFrame(tag for tag in Data1 if Data1.index(tag) % 5==4)

Dead_s =[]
Deadline =[]

for tag in d[2] :
    Dead_s.append(tag.strip(' ~ '))

Dead_s = pd.DataFrame(Dead_s)
Deadline=pd.DataFrame(func_to_date(Dead_s[0]))
Date1=pd.DataFrame(func_to_date(d[0]))
Condition=[]
Condition = pd.DataFrame(func_create_Condition(Deadline[0]))


df=[]

df= pd.merge(a,b, left_index=True, right_index=True, how='outer')
df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
df= pd.merge(df,Date1, left_index=True, right_index=True, how='outer')
df['link']= url
df['Site']= '전략개발원_사업공고'
df= pd.merge(df,Deadline, left_index=True, right_index=True, how='outer')
df= pd.merge(df,Condition, left_index=True, right_index=True, how='outer')

df.columns=['Uploader','Headline', "Number", "Date1", "Link", "Site", 'Deadline', 'Condition']
전략개발원_사업공고=df[["Site","Number",'Headline',"Link",'Uploader',"Date1", 'Deadline', 'Condition' ]]


# < 전략개발원_사업공고 - -----------------------------------------------------------------------------------완료 >

# < 전략개발원_협력기관공고 - -----------------------------------------------------------------------------------시작 4>

url="https://biz.kista.re.kr/ippro/com/iprndMain/selectBusinessAnnounceList.do?bbsType=ac"
page = requests.get(url)
soup = bs(page.text, "html.parser")

Data1=soup.select('tbody >tr >td')


a = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 5==1)
b = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 5==2)
c = pd.DataFrame(tag for tag in Data1 if Data1.index(tag) % 5==3)
d = pd.DataFrame(tag for tag in Data1 if Data1.index(tag) % 5==4)

# 현재 날짜랑 비교해서 컨디션 기입
Dead_s=[]
Deadline=[]
Condition=[]


for tag in d[2] :
    Dead_s.append(tag.strip(' ~ '))

Dead_s = pd.DataFrame(Dead_s)
Deadline=pd.DataFrame(func_to_date(Dead_s[0]))
Date1=pd.DataFrame(func_to_date(d[0]))
Condition = pd.DataFrame(func_create_Condition(Deadline[0]))


df=[]

df= pd.merge(a,b, left_index=True, right_index=True, how='outer')
df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
df= pd.merge(df,Date1, left_index=True, right_index=True, how='outer')
df['link']= url
df['Site']= '전략개발원_협력기관모집공고'
df= pd.merge(df,Deadline, left_index=True, right_index=True, how='outer')
df= pd.merge(df,Condition, left_index=True, right_index=True, how='outer')

df.columns=['Uploader','Headline', "Number", "Date1", "Link", "Site", 'Deadline', 'Condition']

전략개발원_협력기관모집공고=df[["Site","Number",'Headline',"Link",'Uploader',"Date1", 'Deadline', 'Condition' ]]




# < 전략개발원_협력기관공고 - -----------------------------------------------------------------------------------종료 >

# # < 지식재산바우처_스타트업 --------------------------------------------------------------------------------------- 시작 5>
#
# url="https://biz.kista.re.kr/ipvoucher/notiStatupList.do"
# page = requests.get(url)
# soup = bs(page.text, "html.parser")
#
# Data1=soup.select('body>div>form>div>div>div>p')
#
#
# a = pd.DataFrame(tag.text for tag in Data1 if Data1.index(tag) % 3==1)
# b = pd.DataFrame(tag for tag in Data1 if Data1.index(tag) % 3==2)
# c = pd.DataFrame(tag for tag in Data1 if Data1.index(tag) % 3==0)
#
# # 날짜나누기
# v_split = b[1].str.split('~')
# v_Date1 = pd.DataFrame(v_split.str.get(0))
# v_Date2 = pd.DataFrame(v_split.str.get(1))
#
# s1 = []
# s2 = []
# time_format = " %Y.%m.%d %H:%M "
# for tag in v_Date1[1] :
#       s1.append(tag.replace('\n','').replace('\t',''))
#
# s1= pd.DataFrame ( s1 )
# for tag in s1[0]:
#       s2.append ( datetime.datetime.strptime ( tag , time_format ) )
#
# v_Date1=pd.DataFrame(s2)
# s1 = []
# s2 = []
# time_format = " %Y.%m.%d %H:%M"
# for tag in v_Date2[1] :
#      s1.append(tag.replace('\n','').replace('\t',''))
#
# s1 = pd.DataFrame ( s1 )
# for tag in s1[0]:
#      s2.append ( datetime.datetime.strptime ( tag , time_format ) )
#
# v_Date2=pd.DataFrame(s2)
# Condition=[]
# Condition = func_create_Condition(v_Date2[0])
# Condition=pd.DataFrame(Condition)
#
# df= pd.merge(a,c, left_index=True, right_index=True, how='outer')
# df['Uploader'] = ""
# df= pd.merge(df,v_Date1, left_index=True, right_index=True, how='outer')
# df['link']= url
# df['Site']= '지식재산바우처_스타트업'
# df= pd.merge(df,v_Date2, left_index=True, right_index=True, how='outer')
# df= pd.merge(df,Condition, left_index=True, right_index=True, how='outer')
#
# df.columns=['Headline', "Number", 'Uploader', "Date1", "Link", "Site", 'Deadline', 'Condition']
#
# 지식재산바우처_스타트업=df[["Site", "Number", 'Headline', "Link", 'Uploader',"Date1", 'Deadline', 'Condition' ]]
#
#
#
#
# # < 지식재산바우처_스타트업 ------------------------------------------------------------------------------------ 종료 >
#
# # < 지식재산바우처_IP서비스기관모집 --------------------------------------------------------------------------------- 시작 6>
#
# url="https://biz.kista.re.kr/ipvoucher/ipSrvc/notiIpSrvcList.do"
#
# page = requests.get(url)
#
# soup = bs(page.text, "html.parser" )
#
# #넘버
# a=soup.select('#notiIpSrvcListForm > div > div.startup_list > div > div.startup_txt > p:nth-child(1)')
# a= pd.DataFrame(a)
#
# #헤드라인
# b=soup.select('#notiIpSrvcListForm > div > div.startup_list > div > div.startup_txt > p:nth-child(2)')
# ss=[]
#
# for tag in b :
#     ss.append(tag.text)
#
# b= pd.DataFrame(ss)
#
# #날짜
# c=[]
# c=soup.select('#notiIpSrvcListForm > div > div.startup_list > div > div.startup_txt > p:nth-child(3)')
# c= pd.DataFrame(c)
#
# v_split = c[2].str.split('~')
# v_Date1 = pd.DataFrame(v_split.str.get(0))
# v_Date2 = pd.DataFrame(v_split.str.get(1))
#
# s1 = []
# s2 = []
# time_format = "%Y.%m.%d %H:%M"
#
# for tag in v_Date1[2] :
#       s1.append(tag.replace('\n','').replace('\t','').replace('  ','').replace('\xa0','').replace('\r',''))
#
# for tag in s1:
#       s2.append ( datetime.datetime.strptime ( tag , time_format ) )
#
# c = pd.DataFrame(s2)
#
# s1 = []
# s2 = []
# time_format = " %Y.%m.%d %H:%M"
#
# for tag in v_Date2[2] :
#       s1.append(tag.replace('\n','').replace('\t','').replace('  ','').replace('\xa0','').replace('\r',''))
#
# for tag in s1:
#       s2.append ( datetime.datetime.strptime ( tag , time_format ) )
#
# d = pd.DataFrame(s2)
#
# # Condition
# e=[]
# e=soup.select('#notiIpSrvcListForm > div > div.startup_list > div > div.startup_icon.close')
# s1=[]
# s2=[]
#
# for tag in e :
#     s1.append(tag.text)
#
# for tag in s1 :
#       s2.append(tag.replace('\n','').replace('\t','').replace('  ','').replace('\xa0','').replace('\r',''))
#
# e = pd.DataFrame(s2)
#
# #notiIpSrvcListForm > div > div.startup_list > div:nth-child(1) > div.startup_icon.close
#
# df= pd.merge(b,a, left_index=True, right_index=True, how='outer')
# df['Uploader'] = ""
# df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
# df['link']= url
# df['Site']= '지식재산바우처_IP서비스기관모집'
# df= pd.merge(df,d, left_index=True, right_index=True, how='outer')
# df= pd.merge(df,e, left_index=True, right_index=True, how='outer')
#
# df.columns=['Headline', "Number", 'Uploader', "Date1", "Link", "Site", 'Deadline', 'Condition']
#
# 지식재산바우처_IP서비스기관모집=df[["Site", "Number", 'Headline', "Link", 'Uploader',"Date1", 'Deadline', 'Condition' ]]
#
#
#
#
# # < 지식재산바우처_IP서비스기관모집 --------------------------------------------------------------------------------- 종료 >

# < KAUTM ----------------------------------------------------------------------------------------------------- 시작 7>

url="http://www.kautm.net/bbs/?so_table=tlo_news&category=business"

page = requests.get(url)

soup = bs(page.content.decode('utf-8','replace'), "html.parser" )

#헤드라인
a=soup.select('body > div > section > section.sec-centent > div.page-conts-wrap > div > div.boardListArea > div.srboardList > div.listTable > table > tbody > tr > td.title > a > span.board-subject')
s1=[]
for tag in a :
    s1.append(tag.text)

a= pd.DataFrame(s1)

#링크
b=soup.select('body > div > section > section.sec-centent > div.page-conts-wrap > div > div.boardListArea > div.srboardList > div.listTable > table > tbody > tr > td.title > a')
s1=[]
for tag in b :
    s1.append('www.kautm.net'+tag["href"])

b= pd.DataFrame(s1)

# Uploader, 날짜, 조회수
c=soup.select('body > div > section > section.sec-centent > div.page-conts-wrap > div > div.boardListArea > div.srboardList > div.listTable > table > tbody > tr > td.mob-none:nth-child(3)')
s1=[]
s2=[]
s3=[]
s4=[]

for tag in c :
    s1.append(tag.text)

c= pd.DataFrame(s1)

v_split = c[0].str.split('\n')
#uploader
c = pd.DataFrame(v_split.str.get(0))
#조회수
e = pd.DataFrame(v_split.str.get(2))
#날짜
d = pd.DataFrame(v_split.str.get(1))
s1=[]
time_format= "%Y-%m-%d"
for tag in d[0]:
      s1.append ( datetime.datetime.strptime ( tag , time_format ) )

d = pd.DataFrame(s1)

# Number
f=soup.select('body > div > section > section.sec-centent > div.page-conts-wrap > div > div.boardListArea > div.srboardList > div.listTable > table > tbody > tr > td:nth-child(1)')
s1=[]
for tag in f :
    s1.append(tag.text)

f = pd.DataFrame(s1)


df= pd.merge(a,f, left_index=True, right_index=True, how='outer')
df= pd.merge(df,c, left_index=True, right_index=True, how='outer')
df= pd.merge(df,d, left_index=True, right_index=True, how='outer')
df= pd.merge(df,b, left_index=True, right_index=True, how='outer')
df['Site']= 'KAUTM'
df['Deadline']= ''
df['Condition']= ''

df.columns=['Headline', "Number", 'Uploader', "Date1", "Link", "Site", 'Deadline', 'Condition']

KAUTM=df[["Site", "Number", 'Headline', "Link", 'Uploader',"Date1", 'Deadline', 'Condition' ]]



# < KAUTM --------------------------------------------------------------------------------- 종료 7>
# < 수출바우처 공지 - -----------------------------------------------------------------------------------시작 8>

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
    #조회수생성
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

# < 수출바우처 공지 - -----------------------------------------------------------------------------------종료 8>

# < 저장 ------------------------------------------------------------------------------------------------------ 저장 >

ct=datetime.datetime.now()
ctstr=ct.strftime("%Y년 %m월 %d일 %H시%M분%S초")

End_result=[]
# End_result=pd.concat([IP_NAVI_사업공고,IP_NAVI_입찰공고])
# End_result=pd.concat([End_result,RIPC_사업공고])
End_result=pd.concat([RIPC_사업공고,IP_NAVI_사업공고])
End_result=pd.concat([End_result,전략개발원_사업공고])
End_result=pd.concat([End_result,전략개발원_협력기관모집공고])
# End_result=pd.concat([End_result,지식재산바우처_스타트업])
# End_result=pd.concat([End_result,지식재산바우처_IP서비스기관모집])
End_result=pd.concat([End_result,KAUTM])
End_result=pd.concat([End_result,Export_voucher])

#시간순으로 정렬
전략개발원_협력기관모집공고=전략개발원_협력기관모집공고.sort_values('Date1',ascending=False)
전략개발원_사업공고=전략개발원_사업공고.sort_values('Date1',ascending=False)
RIPC_사업공고=RIPC_사업공고.sort_values('Date1',ascending=False)
IP_NAVI_사업공고=IP_NAVI_사업공고.sort_values('Date1',ascending=False)
# IP_NAVI_입찰공고=IP_NAVI_입찰공고.sort_values('Date1',ascending=False)
# 지식재산바우처_스타트업=지식재산바우처_스타트업.sort_values('Date1',ascending=False)
# 지식재산바우처_IP서비스기관모집=지식재산바우처_IP서비스기관모집.sort_values('Date1',ascending=False)
End_result=End_result.sort_values('Date1',ascending=False)
KAUTM=KAUTM.sort_values('Date1',ascending=False)
Export_voucher=Export_voucher.sort_values('Date1',ascending=False)

#중복제거
End_result=End_result.drop_duplicates(["Headline","Uploader"], keep='first',  ignore_index=True)
RIPC_사업공고=RIPC_사업공고.drop_duplicates(["Headline","Uploader"], keep='first',  ignore_index=True)

writer=pd.ExcelWriter('C:/작업서류G/작업서류/0.Crawling/정부사업크롤링('+ctstr+').xlsx', engine='openpyxl')

End_result.to_excel(writer, sheet_name='전체')
전략개발원_협력기관모집공고.to_excel(writer, sheet_name='전략개발원_협력기관모집공고')
전략개발원_사업공고.to_excel(writer, sheet_name='전략개발원_사업공고')
RIPC_사업공고.to_excel(writer, sheet_name='RIPC_사업공고')
IP_NAVI_사업공고.to_excel(writer, sheet_name='IP_NAVI_사업공고')
# IP_NAVI_입찰공고.to_excel(writer, sheet_name='IP_NAVI_입찰공고')
# 지식재산바우처_스타트업.to_excel(writer, sheet_name='지식재산바우처_스타트업')
# 지식재산바우처_IP서비스기관모집.to_excel(writer, sheet_name='지식재산바우처_IP서비스기관모집')
KAUTM.to_excel(writer, sheet_name='KAUTM')
Export_voucher.to_excel(writer, sheet_name='Export_voucher')


writer.save()
# 엑셀 간격 조정
from openpyxl import load_workbook
wb = load_workbook('C:/작업서류G/작업서류/0.Crawling/정부사업크롤링('+ctstr+').xlsx')
for tag in wb.sheetnames:
    ws = wb[tag]
    ws.column_dimensions['A'].width=3
    ws.column_dimensions['B'].width=20
    ws.column_dimensions['C'].width=20
    ws.column_dimensions['D'].width=80
    ws.column_dimensions['E'].width=20
    ws.column_dimensions['F'].width=10
    ws.column_dimensions['G'].width=20
    ws.column_dimensions['H'].width=20
    ws.column_dimensions['I'].width=10

wb.save('C:/작업서류G/작업서류/0.Crawling/정부사업크롤링('+ctstr+').xlsx')

# ------------------------------------------------------------------------------
import subprocess


