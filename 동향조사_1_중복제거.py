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

# 출원번호가 중복된 경우 랜덤하게 하나의 행만 남기고 나머지 행 삭제
duplicate_mask = df.duplicated(subset='출원번호')
duplicate_indices = df[duplicate_mask].index

# 중복된 행들 중에서 랜덤하게 하나의 행 선택
random_index = np.random.choice(duplicate_indices)

# 중복된 행들 중에서 선택한 행을 제외한 나머지 행 삭제
df.drop(index=duplicate_indices[duplicate_indices != random_index], inplace=True)

df.to_excel(url+'base1.xlsx')