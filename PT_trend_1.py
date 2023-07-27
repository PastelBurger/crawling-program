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
import os
# 함수 불러오기
from Functions import pt_trend
from Functions import pre_pt_trend
from Functions import df_check
from Functions import delete_duplicate
from Functions import pt_middle_ratio
from Functions import pt_applicant_middle_ratio


# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\11. 아이티엔제이\가검색\0.Base.xlsx"
url="C:/작업서류G/작업서류/3.정부사업관련/8_IP_사업수행/10_IP-나래_시프트포/1. 준비자료/7회차준비자료/세부분석그래프/"
# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\10_IP-나래_시프트포\1. 준비자료\7회차준비자료\세부분석그래프\0.Base.xlsx"
# "C:\작업서류G\작업서류\3.정부사업관련\8_IP_사업수행\11. 아이티엔제이\가검색"

file_name= "0.Base.xlsx"
df=pd.read_excel(url+file_name)
# 분석 진행할 것인지 결정(1 진행, 0 미진행)
Switch=1

print("체크전")
# 파일 잘 되어있는지 확인

df_check(df)


print("체크후")

# Pre 그래프(스위치가 0이면 진행), 중복제거
if Switch == 0:
    pre_pt_trend ( url , df )
    delete_duplicate(url,file_name)

# 전체 그래프(스위치가 1이면 진행)
if Switch == 1:
    pt_trend(url, df)

# 중분류 그래프

if '중분류' in df.columns and Switch==1 :
    pt_middle_ratio(url, df)
    pt_applicant_middle_ratio(url, df)
    # 소분류 컬럼이 존재하는 경우
    midium_categories = df['중분류'].tolist()
    midium_categories = list(set(midium_categories))
    # 중분류_폴더생성
    for data in midium_categories:
        folder_path = os.path.join('폴더경로', url + data)  # 폴더 경로 설정
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)  # 폴더 생성
    for tag in midium_categories:
        url_midium=url + tag +'/'
        df_midium=df[df['중분류'] == tag]
        df_midium.to_excel ( url_midium + '0_'+tag+'_엑셀데이터.xlsx' , index=True )
        pt_trend ( url_midium , df_midium )

if '소분류' in df.columns  and Switch==1:
    # 소분류 컬럼이 존재하는 경우
    small_categories = df['소분류'].tolist()
    small_categories = list(set(small_categories))
    # 소분류_폴더생성
    for data in small_categories:
        folder_path = os.path.join('폴더경로', url + data)  # 폴더 경로 설정
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)  # 폴더 생성
    for tag in small_categories :
        url_small = url + tag + '/'
        df_small = df[df['소분류'] == tag]
        df_small.to_excel ( url_small + '0_' + tag + '_엑셀데이터.xlsx' , index=True )
        pt_trend ( url_small , df_small )

    # #    코드 수정
    #
    # df_4_2=df[["대표출원인",'국가코드']]
    #
    #
    # s1 = []
    # s1= counts = df_4_2.groupby('대표출원인')['국가코드'].value_counts().unstack(fill_value=0)
    # s1['합계1'] = s1.iloc[:, 1:].sum(axis=1)
    # s1= s1.sort_values(by='합계1', ascending=False)
    # s1=s1[['합계1'] + list(s1.columns[:-1])]
    #
    # s1.to_excel(url+'4_2_0.출원인 국가별 개수.xlsx', index=True)
    #
    # # 합계의 개수별로 상위 10개의 대표출원인 선택
    # top_10 = s1.nlargest(10, '합계1')
    #
    # # 데이터 설정
    # top_10_1= top_10.drop(top_10.columns[0], axis=1)
    # countries = top_10_1.columns[:]  # 국가 코드 리스트
    # application =top_10_1.index
    #
    # # 그래프 설정
    # plt.figure(figsize=(8, 6))  # 그래프 크기 설정
    # colors=['#4C3D3D','#83764F','#C07F00','#FFD95A','#FFF7D4']
    #
    # s1=[]
    # s2= pd.Series(0, index=application)
    # s3=0
    #
    # for tag in countries :
    #     s1=top_10_1[tag]
    #     plt.barh(application, s1, left=s2, color=colors[s3])
    #     s2+=s1
    #     s3+=1
    #
    # # y축 범례 폰트 사이즈 설정
    # max_label_length = max([len(label) for label in application])
    # font_size = 17 - (max_label_length // 2)
    # plt.yticks(fontsize=font_size)
    # # 그래프 상하 뒤집기
    # plt.gca().invert_yaxis()
    # # 그래프 배경에 점선 추가
    # plt.grid ( True , axis='x' , alpha=0.5 , linestyle='--' )
    # plt.subplots_adjust(right=0.8)
    #
    # plt.savefig ( url + '4_2_1 상위 출원인 10개 전체 그래프_국가표시'+'.jpg' , dpi=300 )
