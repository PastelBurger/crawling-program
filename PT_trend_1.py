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

pre_pt_tren
url="C:/작업서류G/작업서류/3.정부사업관련/8_IP_사업수행/10_IP-나래_시프트포/동향조사_Test/"
file_name= "0.Base.xlsx"
df=pd.read_excel(url+file_name)
# 분석 진행할 것인지 결정(1 진행, 0 미진행)
Switch=0

if '대표출원인' not in df.columns:
    print("대표출원인 컬럼이 없습니다.")
    exit()  # 프로그램 종료

if any(df['대표출원인'].isin(['', '-'])):
    if '' in df['대표출원인']:
        print("대표출원인에 빈칸이 있습니다.")
    if '-' in df['대표출원인']:
        print("대표출원인에 '-'가 있습니다.")
    exit()  # 프로그램 종료

if 'JPPAJ' in df['국가코드'].values:
    print("JPPAJ가 있습니다.")
    exit()  # 프로그램 종료

if any(df['제1출원인국적'].isin(['', '-'])):
    if '' in df['제1출원인국적']:
        print("제1출원인국적에 빈칸이 있습니다.")
    if '-' in df['제1출원인국적']:
        print("제1출원인국적에 '-'가 있습니다.")
    exit()  # 프로그램 종료


# Pre 그래프(스위치가 1이면 진행)

if Switch == 0:
    pre_pt_trend ( url , df )

# 전체 그래프(스위치가 1이면 진행)
if Switch == 1:
    pt_trend(url, df)

# 중분류 그래프

if '중분류' in df.columns and Switch==1 :
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
        pt_trend ( url_small , df_small )








#중분류 그래프 생성

for tag in midium_categories:
    url_midium=url + tag +'/'
    df_midium=df[df['중분류'] == tag]
    pt_trend ( url_midium , df_midium )



