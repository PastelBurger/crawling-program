import requests
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup as bs
from pandas import DataFrame
from selenium import webdriver
import time
import openpyxl
import datetime
import matplotlib.pyplot as plt


url="C:/작업서류G/작업서류/3.정부사업관련/8. IP 사업수행/10. IP-나래_시프트포/동향조사_Test/"

file_name= "0.Base.xlsx"

df=pd.read_excel(url+file_name)

# ------------------------------------------------------------------------------------------------1. 연도별 출원 동향 그래프
df_1=df[["대표출원인",'출원일']]

s1 = []
time_format ="%Y.%m.%d"
for tag in df_1['출원일']:
      s1.append ( datetime.datetime.strptime ( tag , time_format ) )

df_1['출원일']= s1


df_1['출원연도']=df_1['출원일'].dt.strftime('%Y')

s1 = []
for tag in df_1['출원연도']:
      s1.append ( int(tag) )

df_1['출원연도']=s1

mn=min(s1)
mx=max(s1)

s1=[]
for tag in range(mn,mx+1) :
      ss=df_1['출원연도'][(df_1['출원연도'] == tag)].count ()
      s1.append(ss)

df_1_1=pd.DataFrame(zip(range(mn,mx+1),s1))
df_1_1.columns = ['출원연도', '개수']

df_1_11=pd.DataFrame.copy(df_1_1)
df_1_11.loc[len(df_1_11)]=['합계',sum(df_1_11['개수'])]

df_1_11.to_excel(url+'1.연도별출원동향(엑셀).xlsx')

# s1 = []
# for tag in df_1_1['출원연도']:
#       s1.append ( str(tag) )
#
# df_1_1['출원연도']=s1


# 한글 폰트 사용을 위해서 세팅
from matplotlib import font_manager, rc
font_path = "C:/Windows/Fonts/NGULIM.TTF"
font = font_manager.FontProperties(fname=font_path).get_name()
rc('font', family=font)

#-------------------------- 그래프그리기

plt.figure(figsize=(10, 5))
bar = plt.bar(df_1_1['출원연도'],df_1_1['개수'], color='#B69F5E')
plt.grid(True, axis='y',alpha=0.5, linestyle='--')
plt.xticks(df_1_1['출원연도'], rotation=45)
plt.gca().spines['right'].set_visible(False) #오른쪽 테두리 제거
plt.gca().spines['top'].set_visible(False) #위 테두리 제거
plt.gca().spines['left'].set_visible(False) #왼쪽 테두리 제거
# plt.gca().set_facecolor('#E6D4A1') #배경색
# plt.yticks(ticks= []) #y축 tick 제거

# 숫자 넣는 부분
for rect in bar:
    height = rect.get_height()
    plt.text(rect.get_x() + rect.get_width()/2.0, height,  height, ha='center', va='bottom', size = 10)

plt.savefig(url+'1.전체출원동향.jpg',dpi=300)

# ----------------------------------------------------------------------------------------1, 2 연도별/국가별 출원 동향 그래프

df_1['국가코드']=df['국가코드']
nt= list(set(df_1['국가코드']))
num=len(df_1['국가코드'])

ss=[]
s1=[]
for tag in nt :
      ss=df_1['국가코드'][(df_1['국가코드'] == tag)].count ()
      s1.append(ss)

df_1_2= pd.DataFrame(nt)
df_1_2['개수']= s1
df_1_2.columns = ['국가', '개수'] # -------- 국가별 개수

ss=[]
s1=[]
s3=pd.DataFrame(range(mn,mx+1))
s3.columns = ['출원연도']
for tag2 in nt :
    ss = []
    s1 = []
    for tag in range(mn,mx+1) :
        ss=df_1['출원연도'][(df_1['출원연도'] == tag) & (df_1['국가코드'] == tag2) ].count ()
        s1.append(ss)
        s2 = pd.DataFrame ( zip ( range ( mn , mx + 1 ) , s1 ) )
        s2.columns = ['출원연도' , '개수']
        s3[tag2]=s2['개수']

df_1_3=s3 # ------- 국가별 연도별 출원 개수

df_1_2.to_excel(url+'2.국가별출원합계(엑셀).xlsx')
df_1_3.to_excel(url+'2.국가별출원동향(엑셀).xlsx')


#-------------------------- 그래프그리기

for tag in nt :
    plt.figure(figsize=(5, 5))
    plt.plot(df_1_1['출원연도'],df_1_1['개수'],'--', color='#B69F5E')
    plt.plot(df_1_3['출원연도'],df_1_3[tag], color='#B69F5E')
    plt.scatter( df_1_3['출원연도'] , df_1_3[tag] , color='#B69F5E',s=5)
    plt.fill_between(df_1_3['출원연도'],df_1_3[tag], 0, alpha=0.3, linewidth=0, edgecolor='grey', facecolor='grey', antialiased=True)
    plt.grid(True, axis='y',alpha=0.5, linestyle='--')
    # plt.xticks(df_1_1['출원연도'], rotation=45)
    # plt.gca().spines['right'].set_visible(False) #오른쪽 테두리 제거
    # plt.gca().spines['top'].set_visible(False) #위 테두리 제거
    # plt.gca().spines['left'].set_visible(False) #왼쪽 테두리 제거
    plt.savefig ( url + '2.국가별출원동향_' + tag + '.jpg' , dpi=300 )


# -----------------------------------------------------------------------------------------------------------3. 메인 IPC

# ipc=list(set(df['메인 IPC']))
# dt=len(df['메인 IPC'])
# s1=[]
# ss=[]
# for tag in ipc :
#       ss=df['메인 IPC'][(df['메인 IPC'] == tag)].count ()
#       s1.append(ss)

ipc=pd.DataFrame(df['메인 IPC'])
ipc.columns=['ipc']
# ipc['개수']= s1
# ipc.columns=['ipc','개수']

s1=[]
for tag in ipc['ipc'] :
    s1.append(tag[0:4])

ipc['대분류']= s1
df_2=ipc   # ----------IPC 분류 개수
# -----------------------------------
s1=[]
s1=list(set(df_2['대분류']))
s2=[]
ss=[]
for tag in s1 :
      ss=df_2['대분류'][(df_2['대분류'] == tag)].count ()
      s2.append(ss)

s3=pd.DataFrame(s1)
s3['개수']= s2
s3.columns=['대분류','개수']
s3=s3.sort_values('개수',ascending=False)
df_23=pd.DataFrame.copy(s3)

df_23.to_excel(url+'3.IPC대분류개수(엑셀).xlsx')

s3=s3.reset_index()
t=[]
t = sum(s3['개수'][range(4,len(s3['개수']))])
s3['개수'][4] =t

s4= s3.loc[0:4,:]
ss=[]
for tag in s4['개수'] :
    ss.append(round(tag/sum(s4['개수'])*100,2))

s4['퍼센트']=ss

df_2_1=s4
df_2_1['대분류'][4]='기타'
df_2_1=df_2_1.sort_values('개수',ascending=False)
df_2_1=df_2_1.reset_index()
df_2_1=df_2_1.loc[:, ['대분류','개수','퍼센트']]
df_2_1.columns=['대분류','개수','퍼센트']

# ------------------------------------
df24=df_2[df_2['대분류']==df_2_1['대분류'][0]]

ipc=list(set(df24['ipc']))
dt=len(df['메인 IPC'])
s1=[]
ss=[]
for tag in ipc :
      ss=df['메인 IPC'][(df['메인 IPC'] == tag)].count ()
      s1.append(ss)

ipc=pd.DataFrame(ipc)
ipc.columns=['ipc']
ipc['개수']= s1
ipc.columns=['ipc','개수']
s1=ipc
s1=s1.sort_values('개수',ascending=False)
s1=s1.reset_index()
ss=[]
for tag in s1['개수'] :
    ss.append(round(tag/sum(s1['개수'])*100,2))

s1['퍼센트']=ss
s2=s1.loc[0:3,:]
df_2_2=s2 # ----  iPC세부분류



#-------------------------- 그래프그리기
# 1. IPC 대분류 그리기
# 그래프 외곽선
wedgeprops = {
    'edgecolor': 'black',
    'linestyle': '-',
    'linewidth': 0.5,
    'width': 0.7
}
color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']

explode = np.zeros(len(df_2_1['개수']))
explode[np.argmax(df_2_1['개수'])] = 0.1  # 가장 큰 값에 대한 explode 값 설정

plt.figure(figsize=(5, 5))
plt.pie(df_2_1['개수'], labels=df_2_1['대분류'], startangle=90, explode=explode, autopct='%.1f%%',counterclock=False,wedgeprops=wedgeprops,colors=color,shadow=True)
# plt.show()

plt.savefig(url+'3.IPC분류.jpg',dpi=300)

# 2. IPC 세부 분류 그리기
# df_2_2=sort_values('개수')
plt.figure(figsize=(5, 5))
bar = plt.bar(df_2_2['ipc'],df_2_2['개수'], color= color)
plt.grid(True, axis='y',alpha=0.5, linestyle='--')
# # plt.xticks(df_1_1['출원연도'], rotation=45)
plt.gca().spines['right'].set_visible(False) #오른쪽 테두리 제거
plt.gca().spines['top'].set_visible(False) #위 테두리 제거
# plt.gca().spines['left'].set_visible(False) #왼쪽 테두리 제거


# 숫자 넣는 부분
t=0
for rect in bar:
    height = rect.get_height()
    plt.text(rect.get_x() + rect.get_width()/2.0, height, str(df_2_2['퍼센트'][t])+'%', ha='center', va='bottom', size = 10)
    t=t+1

# plt.show()

plt.savefig(url+'3.1_IPC소분류.jpg',dpi=300)

# -----------------------------------------------------------------------------------------------------------4. 경쟁사 뽑기

df_4=df[["대표출원인",'출원일']]

s1 = []
time_format ="%Y.%m.%d"
for tag in df_4['출원일']:
      s1.append ( datetime.datetime.strptime ( tag , time_format ) )

df_4.loc[:,'출원일']= s1
df_4.loc[:,'출원연도']=df_4['출원일'].dt.strftime('%Y')

s1 = []
s1= counts = df_4.groupby('대표출원인')['출원연도'].value_counts().unstack(fill_value=0)
s1['합계1'] = s1.iloc[:, 1:].sum(axis=1)
s1= s1.sort_values(by='합계1', ascending=False)
s1=s1[['합계1'] + list(s1.columns[:-1])]
year_totals = s1.iloc[:, 1:].sum()
s1.loc['합계2'] = year_totals

# 합계의 개수별로 상위 10개의 대표출원인 선택
top_10 = s1.nlargest(10, '합계1')
top_30 = s1.nlargest(30, '합계1')

# 선택된 10개의 대표출원인의 연도별 개수와 마지막 행 포함하는 테이블 생성
df_4_1= s1.loc[top_10.index, :]
df_4_1.loc['합계2'] =  s1.iloc[-1]

df_4_1.to_excel(url+'4_1.상위10개 출원인.xlsx', index=True)
top_30.to_excel(url+'4_1_1.상위30개 출원인.xlsx', index=True)

# ----------------------------------------------------------------------그래프
# 그래프 설정
plt.figure(figsize=(10, 6))  # 그래프 크기 설정

# 대표출원인 리스트 생성
대표출원인 = df_4_1.index.tolist()
출원연도 = df_4_1.columns[1:].tolist()

# color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
# 그래프 그리기
for 출원인 in 대표출원인 :
    plt.figure ( figsize=(10 , 6) )  # 그래프 크기 설정
    개수_합계2 = df_4_1.loc['합계2' , 출원연도]
    plt.plot ( 출원연도 , 개수_합계2 , linestyle='--' , marker='o' , label='합계2', color='#C58940' )
    # 대표출원인 열만 선택
    개수 = df_4_1.loc[출원인 , 출원연도]
    # 대표출원인은 실선 꺾은선 그래프로 그리기
    plt.plot ( 출원연도 , 개수 , linestyle='-' , marker='o' , label=출원인 , color='#D4A657' )
    # 음영 처리
    plt.fill_between ( 출원연도 , 개수 , 0 , color='gray' , alpha=0.3 )
    # 그래프 레이블, 타이틀 설정
    plt.xticks ( rotation=45 )
    plt.grid ( True , axis='y' , alpha=0.5 , linestyle='--' )
    plt.savefig ( url + '4.1 출원인별그래프_' + 출원인 + '.jpg' , dpi=300 )

# 그래프 그리기
for 출원인 in 대표출원인 :
    plt.figure ( figsize=(10 , 6) )  # 그래프 크기 설정
    개수 = df_4_1.loc[출원인 , 출원연도]
    # 대표출원인은 실선 꺾은선 그래프로 그리기
    plt.plot ( 출원연도 , 개수 , linestyle='-' , marker='o' , label=출원인, color='#C58940' )
    # 음영 처리
    plt.fill_between ( 출원연도 , 개수 , 0 , color='gray' , alpha=0.3 )
    # 그래프 레이블, 타이틀 설정
    plt.xticks ( rotation=45 )
    plt.grid ( True , axis='y' , alpha=0.5 , linestyle='--' )
    plt.savefig ( url + '4.1 출원인별그래프_' + 출원인 +'(전체없음)'+'.jpg' , dpi=300 )

from matplotlib import font_manager, rc
font_path = "C:/Windows/Fonts/NGULIM.TTF"
font = font_manager.FontProperties(fname=font_path).get_name()
rc('font', family=font)


# 대표출원인 리스트 생성
대표출원인 = df_4_1.index.tolist()

# 대표출원인별 합계 추출
합계 = df_4_1.loc[:, '합계1']

# 합계를 기준으로 대표출원인 정렬
정렬된_대표출원인 = sorted(대표출원인, key=lambda x: 합계[x], reverse=True)
정렬된_합계 = [합계[x] for x in 정렬된_대표출원인]

# 그래프 설정
plt.figure(figsize=(10, 6))  # 그래프 크기 설정

# 가로 막대 그래프 그리기
plt.barh(정렬된_대표출원인[::-1], 정렬된_합계[::-1],color='#C58940')  # 역순으로 그래프 그리기
plt.grid ( True , axis='x' , alpha=0.5 , linestyle='--' )
# 좌우측 여백 추가


# y축 범례 폰트 사이즈 설정
max_label_length = max([len(label) for label in 정렬된_대표출원인])
font_size = 17 - (max_label_length // 2)
plt.yticks(fontsize=font_size)

plt.savefig ( url + '4.2 상위 출원인 10개 전체 그래프_'+'.jpg' , dpi=300 )


# ---------------------------------------------------------------------- 출원인 국가별 개수


df_4_2=df[["대표출원인",'국가코드']]


s1 = []
s1= counts = df_4_2.groupby('대표출원인')['국가코드'].value_counts().unstack(fill_value=0)
s1['합계1'] = s1.iloc[:, 1:].sum(axis=1)
s1= s1.sort_values(by='합계1', ascending=False)
s1=s1[['합계1'] + list(s1.columns[:-1])]

s1.to_excel(url+'4_2.출원인 국가별 개수.xlsx', index=True)

# 합계의 개수별로 상위 10개의 대표출원인 선택
top_10 = s1.nlargest(10, '합계1')

# 데이터 설정
countries = top_10.columns[1:]  # 국가 코드 리스트
num_countries = len(countries)  # 국가 코드 개수
total_applications = top_10['합계1']  # 대표출원인별 합계
representatives = top_10.index.tolist()  # 대표출원인 리스트

# 그래프 설정
plt.figure(figsize=(10, 6))  # 그래프 크기 설정
colors=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
bar_colors = colors[:num_countries]  # 개수에 맞게 색상 리스트 할당

# 가로 바 그래프 그리기
plt.barh(representatives[::-1], total_applications[::-1], color=bar_colors)

plt.grid ( True , axis='x' , alpha=0.5 , linestyle='--' )


# y축 범례 폰트 사이즈 설정
max_label_length = max([len(label) for label in representatives])
font_size = 17 - (max_label_length // 2)
plt.yticks(fontsize=font_size)

plt.savefig ( url + '4.2_1 상위 출원인 10개 전체 그래프_국가표시'+'.jpg' , dpi=300 )


