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
# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\11. 아이티엔제이\가검색\0.Base.xlsx"
url="C:/작업서류SG/작업서류/3.정부사업관련/8_IP_사업수행/17.IP나래_시제/실험용/"
# "C:\작업서류SG\작업서류\3.정부사업관련\8_IP_사업수행\17.IP나래_시제\실험용\0.Base.xlsx"
# "C:\작업서류SG\작업서류\3.정부사업관련\8_IP_사업수행\11. 아이티엔제이\가검색\0.Base.xlsx"
# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\10_IP-나래_시프트포\1. 준비자료\7회차준비자료\세부분석그래프\0.Base.xlsx"
# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\11. 아이티엔제이\가검색"
# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\11. 아이티엔제이\가검색\0.Base_1.xlsx"
file_name= "0.Base.xlsx"
df=pd.read_excel(url+file_name)

df_7 = df[["대표출원인", '출원일','국가코드']]
s1 = []
time_format = "%Y.%m.%d"
for tag in df_7['출원일']:
    s1.append(datetime.datetime.strptime(tag, time_format))

df_7.loc[:, '출원일'] = s1
df_7['출원일'] = pd.to_datetime(df_7['출원일'], format=time_format)
df_7.loc[:, '출원연도'] = df_7['출원일'].dt.year

s1 = []
s1 = counts = df_7.groupby('대표출원인')['출원연도'].value_counts().unstack(fill_value=0)
s1['합계1'] = s1.iloc[:, :].sum(axis=1)
s1 = s1.sort_values(by='합계1', ascending=False)
s1 = s1[['합계1'] + list(s1.columns[:-1])]
year_totals = s1.iloc[:, 1:].sum()
s1.loc['합계2'] = year_totals

# 합계의 개수별로 상위 10개의 대표출원인 선택
top_10 = s1.nlargest(10, '합계1')
s2=pd.DataFrame(top_10.index)
sss=0

for tag in s2["대표출원인"] :
    sss+=1
    tag_rows = df_7[df_7['대표출원인'] == tag]
    s3 = counts = tag_rows.groupby('국가코드')['출원연도'].value_counts().unstack(fill_value=0)

    # 데이터프레임의 첫 번째 인덱스와 마지막 인덱스 추출
    first_index = s3.columns[0]
    last_index = s3.columns[-1]
    index_to_add = [int(i) for i in range(first_index, last_index)]
    for col in index_to_add:
        if col not in s3.columns:
            s3[col] = 0

    s3 = s3.sort_index(axis=1)

    s3['합계1'] = s3.iloc[:, :].sum(axis=1)
    s3 = s3.sort_values(by='합계1', ascending=False)
    s3 = s3[['합계1'] + list(s3.columns[:-1])]
    year_totals = s3.iloc[:, :].sum()
    s3.loc['합계2'] = year_totals

    # ----------------------------------------------------------------------그래프
    # 그래프 설정
    plt.figure(figsize=(10, 6))  # 그래프 크기 설정

    # 국가코드 리스트 생성
    국가코드 = s3.index.tolist()
    출원연도 = s3.columns[1:].tolist()

    # color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
    # 그래프 그리기
    ss = 1
    for 국가 in 국가코드:
        plt.figure(figsize=(10, 6))  # 그래프 크기 설정
        개수_합계2 = s3.loc['합계2', 출원연도]
        plt.plot(출원연도, 개수_합계2, linestyle='--', marker='o', label='합계2', color='#C58940')
        # 대표출원인 열만 선택
        개수 = s3.loc[국가, 출원연도]
        # 대표출원인은 실선 꺾은선 그래프로 그리기
        plt.plot(출원연도, 개수, linestyle='-', marker='o', label=국가, color='#D4A657')
        # 음영 처리
        plt.fill_between(출원연도, 개수, 0, color='gray', alpha=0.3)
        # 그래프 레이블, 타이틀 설정
        plt.xticks(출원연도, rotation=45)
        plt.yticks(range(int(min(개수_합계2)), int(max(개수_합계2)) + 1))
        plt.grid(True, axis='y', alpha=0.5, linestyle='--')
        plt.savefig(url + '7_1_'+str(sss)+'_'+tag+'_' + 국가 + '.jpg', dpi=300)
        plt.close()
        ss += 1



