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


# -----------------------------------------------------------------------------------------------------------------------------이하는 함수

def pt_applicant_middle_ratio(url , df) :
    df_4_2 = df[["대표출원인" , '중분류']]
    s1 = []
    s1 = counts = df_4_2.groupby ( '대표출원인' )['중분류'].value_counts ().unstack ( fill_value=0 )
    s1['합계1'] = s1.iloc[: , :].sum ( axis=1 )
    s1 = s1.sort_values ( by='합계1' , ascending=False )
    s1 = s1[['합계1'] + list ( s1.columns[:-1] )]

    s1.to_excel ( url + '4_2_2.출원인 중분류별 개수.xlsx' , index=True )

    # 합계의 개수별로 상위 10개의 대표출원인 선택
    top_10 = s1.nlargest ( 10 , '합계1' )

    # 데이터 설정
    top_10_1 = top_10.drop ( top_10.columns[0] , axis=1 )
    countries = top_10_1.columns[:]  # 국가 코드 리스트
    application = top_10_1.index

    # 그래프 설정
    plt.figure ( figsize=(8 , 6) )  # 그래프 크기 설정
    colors = ['#4C3D3D' , '#83764F' , '#C07F00' , '#FFD95A' , '#FFF7D4']

    s1 = []
    s2 = pd.Series ( 0 , index=application )
    s3 = 0

    for tag in countries :
        s1 = top_10_1[tag]
        plt.barh ( application , s1 , left=s2 , color=colors[s3] )
        s2 += s1
        s3 += 1

    # y축 범례 폰트 사이즈 설정
    max_label_length = max ( [len ( label ) for label in application] )
    font_size = 17 - (max_label_length // 2)
    plt.yticks ( fontsize=font_size )
    # 그래프 상하 뒤집기
    plt.gca ().invert_yaxis ()
    # 그래프 배경에 점선 추가
    plt.grid ( True , axis='x' , alpha=0.5 , linestyle='--' )
    plt.subplots_adjust ( right=0.8 )

    plt.savefig ( url + '4_2_2 상위 출원인 10개 전체 그래프_중분류표시' + '.jpg' , dpi=300 )
    plt.close()

def pt_middle_ratio (url, df) :
    # '국가코드'별로 '중분류'의 개수를 세기
    counts = df.groupby(['중분류']).size().reset_index(name='개수')

    ## 중분류 비율 그래프 1 -텍스트 포함
    df_us = counts.copy()
    # 데이터 설정
    labels = df_us['중분류']
    sizes = df_us[['중분류','개수']]
    sizes = sizes.set_index('중분류')['개수']
    # 색상 지정
    colors = ['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']  # 색상 지정
    line_color = 'Black'  # 선의 색상 설정
    # 원 그래프 그리기
    plt.figure(figsize=(5, 5))
    wedges, text, autotext = plt.pie(sizes, labels=labels, colors=colors, autopct=lambda p: '{:.1f}%\n({:.0f})'.format(p, p * sum(sizes) / 100),startangle=90,
                                     wedgeprops={'edgecolor': line_color, 'linewidth': 0.5, 'linestyle': 'solid'})

    # 원그래프 모양 조정 (동그란 원으로 만들기)
    plt.axis('equal')
    # 텍스트 스타일 설정
    plt.setp(wedges, edgecolor=line_color)
    plt.setp(text, color=line_color)
    plt.setp(autotext, color=line_color)
    plt.savefig ( url + '6_1 중분류 비율_.jpg' , dpi=300 )
    plt.close()

    ## 중분류 비율 그래프 2 -텍스트 미포함
    # 데이터 설정
    labels = df_us['중분류']
    sizes = df_us[['중분류','개수']]
    sizes = sizes.set_index('중분류')['개수']
    # 색상 지정
    colors = ['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']  # 색상 지정
    line_color = 'Black'  # 선의 색상 설정
    # 원 그래프 그리기
    plt.figure(figsize=(5, 5))
    wedges, text = plt.pie(sizes, labels=labels, colors=colors, startangle=90,wedgeprops={'edgecolor': line_color, 'linewidth': 0.5, 'linestyle': 'solid'})

    # 원그래프 모양 조정 (동그란 원으로 만들기)
    plt.axis('equal')
    # 텍스트 스타일 설정
    plt.setp(wedges, edgecolor=line_color)
    plt.setp(text, color=line_color)
    plt.setp(autotext, color=line_color)
    plt.savefig ( url + '6_2 중분류 비율(텍스트 없음)_.jpg' , dpi=300 )
    plt.close()

    ## 엑셀저장
    s1= counts.copy()
    total_count = s1['개수'].sum()
    s1['퍼센트'] = (s1['개수'] / total_count) * 100
    s1['개수와 퍼센트'] = s1.apply(lambda row: f"{row['개수']} ({row['퍼센트']:.1f}%)", axis=1)
    s1.to_excel(url + '6_2 중분류 비율.xlsx', index=False)


def pt_trend (url, df) :

    # 한글 폰트 사용을 위해서 세팅
    from matplotlib import font_manager, rc
    font_path = "C:/Windows/Fonts/NGULIM.TTF"
    font = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font)


    # ------------------------------------------------------------------------------------------------1. 연도별 출원 동향 그래프
    df_1=df[["대표출원인",'출원일']]

    s1 = []
    time_format ="%Y.%m.%d"
    for tag in df_1['출원일']:
          s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    df_1['출원일']= s1

    df_1['출원일'] = pd.to_datetime(df_1['출원일'], format=time_format)
    df_1.loc[:, '출원연도'] = df_1['출원일'].dt.year

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
    plt.close()

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
        plt.close()


    # -----------------------------------------------------------------------------------------------------------3. 메인 IPC
    #
    # # ipc=list(set(df['메인 IPC']))
    # # dt=len(df['메인 IPC'])
    # # s1=[]
    # # ss=[]
    # # for tag in ipc :
    # #       ss=df['메인 IPC'][(df['메인 IPC'] == tag)].count ()
    # #       s1.append(ss)
    #
    # ipc=pd.DataFrame(df['메인 IPC'])
    # ipc.columns=['ipc']
    # # ipc['개수']= s1
    # # ipc.columns=['ipc','개수']
    #
    # s1=[]
    # for tag in ipc['ipc'] :
    #     s1.append(tag[0:4])
    #
    # ipc['대분류']= s1
    # df_2=ipc   # ----------IPC 분류 개수
    # # -----------------------------------
    # s1=[]
    # s1=list(set(df_2['대분류']))
    # s2=[]
    # ss=[]
    # for tag in s1 :
    #       ss=df_2['대분류'][(df_2['대분류'] == tag)].count ()
    #       s2.append(ss)
    #
    # s3=pd.DataFrame(s1)
    # s3['개수']= s2
    # s3.columns=['대분류','개수']
    # s3=s3.sort_values('개수',ascending=False)
    # df_23=pd.DataFrame.copy(s3)
    #
    # df_23.to_excel(url+'3.IPC대분류개수(엑셀).xlsx')
    #
    # s3=s3.reset_index()
    # t=[]
    # t = sum(s3['개수'][range(4,len(s3['개수']))])
    # s3['개수'][4] =t
    #
    # s4= s3.loc[0:4,:]
    # ss=[]
    # for tag in s4['개수'] :
    #     ss.append(round(tag/sum(s4['개수'])*100,2))
    #
    # s4['퍼센트']=ss
    #
    # df_2_1=s4
    # df_2_1['대분류'][4]='기타'
    # df_2_1=df_2_1.sort_values('개수',ascending=False)
    # df_2_1=df_2_1.reset_index()
    # df_2_1=df_2_1.loc[:, ['대분류','개수','퍼센트']]
    # df_2_1.columns=['대분류','개수','퍼센트']
    #
    # # ------------------------------------
    # df24=df_2[df_2['대분류']==df_2_1['대분류'][1]]
    #
    # ipc=list(set(df24['ipc']))
    # dt=len(df['메인 IPC'])
    # s1=[]
    # ss=[]
    # for tag in ipc :
    #       ss=df['메인 IPC'][(df['메인 IPC'] == tag)].count ()
    #       s1.append(ss)
    #
    # ipc=pd.DataFrame(ipc)
    # ipc.columns=['ipc']
    # ipc['개수']= s1
    # ipc.columns=['ipc','개수']
    # s1=ipc
    # s1=s1.sort_values('개수',ascending=False)
    # s1=s1.reset_index()
    # ss=[]
    # for tag in s1['개수'] :
    #     ss.append(round(tag/sum(s1['개수'])*100,2))
    #
    # s1['퍼센트']=ss
    # s2=s1.loc[0:3,:]
    # df_2_2=s2 # ----  iPC세부분류
    #
    #
    #
    # #-------------------------- 그래프그리기
    # # 1. IPC 대분류 그리기
    # # 그래프 외곽선
    # wedgeprops = {
    #     'edgecolor': 'black',
    #     'linestyle': '-',
    #     'linewidth': 0.5,
    #     'width': 0.7
    # }
    # color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
    #
    # explode = np.zeros(len(df_2_1['개수']))
    # explode[np.argmax(df_2_1['개수'])] = 0.1  # 가장 큰 값에 대한 explode 값 설정
    #
    # plt.figure(figsize=(5, 5))
    # plt.pie(df_2_1['개수'], labels=df_2_1['대분류'], startangle=90, explode=explode, autopct='%.1f%%',counterclock=False,wedgeprops=wedgeprops,colors=color,shadow=True)
    # # plt.show()
    #
    # plt.savefig(url+'3.IPC분류.jpg',dpi=300)
    #
    # # 2. IPC 세부 분류 그리기
    # # df_2_2=sort_values('개수')
    # plt.figure(figsize=(5, 5))
    # bar = plt.bar(df_2_2['ipc'],df_2_2['개수'], color= color)
    # plt.grid(True, axis='y',alpha=0.5, linestyle='--')
    # # # plt.xticks(df_1_1['출원연도'], rotation=45)
    # plt.gca().spines['right'].set_visible(False) #오른쪽 테두리 제거
    # plt.gca().spines['top'].set_visible(False) #위 테두리 제거
    # # plt.gca().spines['left'].set_visible(False) #왼쪽 테두리 제거
    #
    #
    # # 숫자 넣는 부분
    # t=0
    # for rect in bar:
    #     height = rect.get_height()
    #     plt.text(rect.get_x() + rect.get_width()/2.0, height, str(df_2_2['퍼센트'][t])+'%', ha='center', va='bottom', size = 10)
    #     t=t+1
    #
    # # plt.show()
    #
    # plt.savefig(url+'3.1_IPC소분류.jpg',dpi=300)

    # -----------------------------------------------------------------------------------------------------------4. 경쟁사 뽑기
    # ----------------------------------------------------------------------------------------------4.1 출원인별 그래프
    df_4=df[["대표출원인",'출원일']]

    s1 = []
    time_format ="%Y.%m.%d"
    for tag in df_4['출원일']:
          s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    df_4.loc[:,'출원일']= s1
    df_4['출원일'] = pd.to_datetime(df_4['출원일'], format=time_format)
    df_4.loc[:,'출원연도']=df_4['출원일'].dt.year

    s1 = []
    s1= counts = df_4.groupby('대표출원인')['출원연도'].value_counts().unstack(fill_value=0)
    s1['합계1'] = s1.iloc[:, :].sum(axis=1)
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

    df_4_1.to_excel(url+'4_1_0.상위10개 출원인.xlsx', index=True)
    top_30.to_excel(url+'4_1_1.상위30개 출원인.xlsx', index=True)

    # ----------------------------------------------------------------------그래프
    # 그래프 설정
    plt.figure(figsize=(10, 6))  # 그래프 크기 설정

    # 대표출원인 리스트 생성
    대표출원인 = df_4_1.index.tolist()
    출원연도 = df_4_1.columns[1:].tolist()

    # color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
    # 그래프 그리기
    ss=1
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
        plt.savefig ( url + '4_1_' +str(ss)+출원인 + '.jpg' , dpi=300 )
        plt.close()
        ss+=1

    # 그래프 그리기
    ss=1
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
        plt.savefig ( url + '4_1 _' +str(ss)+ 출원인 +'(전체없음)'+'.jpg' , dpi=300 )
        plt.close()
        ss+=1

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

    plt.savefig ( url + '4_2_0 상위 출원인 10개 전체 그래프'+'.jpg' , dpi=300 )
    plt.close()


    # ----------------------------------------------------------------------------------------------4.2 출원인 국가별 개수


    df_4_2=df[["대표출원인",'국가코드']]


    s1 = []
    s1= counts = df_4_2.groupby('대표출원인')['국가코드'].value_counts().unstack(fill_value=0)
    s1['합계1'] = s1.iloc[:, :].sum(axis=1)
    s1= s1.sort_values(by='합계1', ascending=False)
    s1=s1[['합계1'] + list(s1.columns[:-1])]

    s1.to_excel(url+'4_2_0.출원인 국가별 개수.xlsx', index=True)

    # 합계의 개수별로 상위 10개의 대표출원인 선택
    top_10 = s1.nlargest(10, '합계1')

    # 데이터 설정
    top_10_1= top_10.drop(top_10.columns[0], axis=1)
    countries = top_10_1.columns[:]  # 국가 코드 리스트
    application =top_10_1.index

    # 그래프 설정
    plt.figure(figsize=(8, 6))  # 그래프 크기 설정
    colors=['#4C3D3D','#83764F','#C07F00','#FFD95A','#FFF7D4']

    s1=[]
    s2= pd.Series(0, index=application)
    s3=0

    for tag in countries :
        s1=top_10_1[tag]
        plt.barh(application, s1, left=s2, color=colors[s3])
        s2+=s1
        s3+=1

    # y축 범례 폰트 사이즈 설정
    max_label_length = max([len(label) for label in application])
    font_size = 17 - (max_label_length // 2)
    plt.yticks(fontsize=font_size)
    # 그래프 상하 뒤집기
    plt.gca().invert_yaxis()
    # 그래프 배경에 점선 추가
    plt.grid ( True , axis='x' , alpha=0.5 , linestyle='--' )
    plt.subplots_adjust(right=0.8)

    plt.savefig ( url + '4_2_1 상위 출원인 10개 전체 그래프_국가표시'+'.jpg' , dpi=300 )
    plt.close()


    # ----------------------------------------------------------------------------------------------4.3 국가별 출원인 분석

    country_codes = df['국가코드'].unique()

    for unique_tag in country_codes:
        df_country=df[df['국가코드'] == unique_tag]

        df_4=df_country[["대표출원인",'출원일']]

        s1 = []
        time_format ="%Y.%m.%d"
        for tag in df_4['출원일']:
              s1.append ( datetime.datetime.strptime ( tag , time_format ) )

        df_4.loc[:,'출원일']= s1
        df_4['출원일'] = pd.to_datetime(df_4['출원일'], format=time_format)
        df_4.loc[:, '출원연도'] = df_4['출원일'].dt.year

        s1 = []
        s1= counts = df_4.groupby('대표출원인')['출원연도'].value_counts().unstack(fill_value=0)
        s1['합계1'] = s1.iloc[:, :].sum(axis=1)
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

        df_4_1.to_excel(url+'4_1_0_'+unique_tag+'.상위10개 출원인.xlsx', index=True)
        top_30.to_excel(url+'4_1_1_'+unique_tag+'.상위30개 출원인.xlsx', index=True)

        # ----------------------------------------------------------------------그래프
        # 그래프 설정
        plt.figure(figsize=(10, 6))  # 그래프 크기 설정

        # 대표출원인 리스트 생성
        대표출원인 = df_4_1.index.tolist()
        출원연도 = df_4_1.columns[1:].tolist()

        # color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
        # 그래프 그리기
        ss=1
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
            plt.savefig ( url + '4_1_'+unique_tag+'_' + str(ss)+출원인 + '.jpg' , dpi=300 )
            plt.close()
            ss+=1

        # 그래프 그리기
        ss=1
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
            plt.savefig ( url + '4_1_'+unique_tag+'_' +str(ss)+ 출원인 +'(전체없음)'+'.jpg' , dpi=300 )
            plt.close()
            ss+=1

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

        plt.savefig ( url + '4_2_1_'+unique_tag+' 상위 출원인 10개 전체 그래프_'+'.jpg' , dpi=300 )
        plt.close()


        # -----------------------------------------------------------------------------------------------------------5. 내외국인


    # 합계의 개수별로 상위 10개의 대표출원인 선택
    top_10 = s1.nlargest(10, '합계1')
    top_30 = s1.nlargest(30, '합계1')

    # 선택된 10개의 대표출원인의 연도별 개수와 마지막 행 포함하는 테이블 생성
    df_4_1 = s1.loc[top_10.index, :]
    df_4_1.loc['합계2'] = s1.iloc[-1]


    # ----------------------------5_1. 내외국인비율

    df_5=df[["대표출원인",'국가코드', '출원일','제1출원인국적' ]]
    s1 = []
    time_format ="%Y.%m.%d"
    for tag in df_5['출원일']:
          s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    df_5['출원일']= s1
    df_5['출원일'] = pd.to_datetime(df_5['출원일'], format=time_format)
    df_5['출원연도']=df_5['출원일'].dt.year

    def get_nationality(row):
        if row['국가코드'] == 'EP':
            if row['제1출원인국적'] in ['DE', 'FR', 'IE', 'CH', 'DK','NL','ES','PT','PL', 'IT']:
                return '내국인'
            else:
                return '외국인'
        else:
            if row['국가코드'] == row['제1출원인국적']:
                return '내국인'
            else:
                return '외국인'

    df_5['내외국인']=df_5.apply(get_nationality, axis=1)

    # 국가코드별 내외국인 비율 계산
    nation_nationality_counts = df_5.groupby('국가코드')['내외국인'].value_counts()
    nation_nationality_counts=pd.DataFrame(nation_nationality_counts)
    nation_nationality_counts.to_excel(url+'5.1 내외국인 개수.xlsx', index=True)
    #첫번째 인덱스 추가
    s1=nation_nationality_counts.index.get_level_values(0).unique().tolist()

    for tag in s1:
        # US에 대한 데이터 추출
        df_us = nation_nationality_counts.loc[tag]
        # 데이터 설정
        labels = df_us.index
        sizes = df_us.values
        sizes = [item for sublist in sizes for item in sublist]
        # 색상 지정
        colors = ['#C58940', '#FAEAB1']  # 내외국인에 대한 색상 지정
        line_color = 'Black'  # 선의 색상 설정
        # 원 그래프 그리기
        plt.figure(figsize=(5, 5))
        wedges, text, autotext = plt.pie(sizes, labels=labels, colors=colors, autopct=lambda p: '{:.1f}%\n({:.0f})'.format(p, p * sum(sizes) / 100),startangle=90,
                                         wedgeprops={'edgecolor': line_color, 'linewidth': 0.5, 'linestyle': 'solid'})
        # 원그래프 모양 조정 (동그란 원으로 만들기)
        plt.axis('equal')
        # 텍스트 스타일 설정
        plt.setp(wedges, edgecolor=line_color)
        plt.setp(text, color=line_color)
        plt.setp(autotext, color=line_color)
        plt.savefig ( url + '5_1_'+ tag +'_내외국인 비율.jpg' , dpi=300 )
        plt.close()

        # ----------------------------5_2. 외국인 비율

    foreigners = df_5.loc[df_5['내외국인'] == '외국인']

    # '국가코드'별로 '제1출원인국적'의 개수를 세기
    counts = foreigners.groupby(['국가코드', '제1출원인국적']).size().reset_index(name='개수')
    nation_counts = counts.groupby('국가코드').size().reset_index(name='개수')
    # 가장 작은 개수가 4 이하인 경우 해당 수를 반환하고, 그렇지 않은 경우에는 4를 반환
    min_count = nation_counts['개수'].min()
    result = min_count if min_count <= 4 else 4

    # '국가코드'별로 상위 4개의 '제1출원인국적' 유지하고, 나머지 국가들의 개수 세기

    top4_countries = counts.groupby('국가코드').apply(lambda x: x.nlargest(result, '개수')).reset_index(drop=True)
    others = counts.groupby('국가코드').apply(lambda x: x.nsmallest(len(x) - result, '개수')).reset_index(drop=True)
    others['제1출원인국적'] = '기타'
    others_counts = others.groupby(['국가코드', '제1출원인국적'])['개수'].sum().reset_index()

    # '기타'로 합쳐진 데이터 추가하기
    counts_result = pd.concat([top4_countries, others_counts], ignore_index=True)

    #첫번째 인덱스 추가
    s1=nation_nationality_counts.index.get_level_values(0).unique().tolist()


    for tag in s1:
        df_us = counts_result[counts_result['국가코드'] == tag]
        # 데이터 설정
        labels = df_us['제1출원인국적']
        sizes = df_us[['제1출원인국적','개수']]
        sizes = sizes.set_index('제1출원인국적')['개수']
        # 색상 지정
        colors = ['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']  # 색상 지정
        line_color = 'Black'  # 선의 색상 설정
        # 원 그래프 그리기
        plt.figure(figsize=(5, 5))
        wedges, text, autotext = plt.pie(sizes, labels=labels, colors=colors, autopct=lambda p: '{:.1f}%\n({:.0f})'.format(p, p * sum(sizes) / 100),startangle=90,
                                         wedgeprops={'edgecolor': line_color, 'linewidth': 0.5, 'linestyle': 'solid'})
        # 원그래프 모양 조정 (동그란 원으로 만들기)
        plt.axis('equal')
        # 텍스트 스타일 설정
        plt.setp(wedges, edgecolor=line_color)
        plt.setp(text, color=line_color)
        plt.setp(autotext, color=line_color)
        plt.savefig ( url + '5_1_'+ tag +'_외국인 비율.jpg' , dpi=300 )
        plt.close()


    # -------------------------------5_2. 내외국인 출원동향

    country_codes = df_5['국가코드'].unique().tolist()

    df_5['출원연도'] = df_5['출원연도'].astype(int)

    domestic = df_5[df_5['내외국인']=='내국인']
    foreign = df_5[df_5['내외국인']=='외국인']
    domestic_counts = domestic.groupby(['국가코드', '출원연도']).size().reset_index(name='개수')
    foreign_counts = foreign.groupby(['국가코드', '출원연도']).size().reset_index(name='개수')
    domestic_counts=domestic_counts.sort_values('출원연도')
    foreign_counts=foreign_counts.sort_values('출원연도')

    for tag in country_codes:
        domestic_data = domestic_counts[domestic_counts['국가코드'] == tag]
        foreign_data = foreign_counts[foreign_counts['국가코드'] == tag]
        domestic_series = domestic_data.set_index('출원연도')['개수']
        foreign_series = foreign_data.set_index('출원연도')['개수']
        #-------------------------- 그래프그리기
        plt.figure(figsize=(8, 6))
        plt.bar(foreign_series.index,foreign_series, color='#FAEAB1',width=0.3,  label='외국인',edgecolor='#E5BA73',linewidth=0.2)
        plt.plot(domestic_series.index,domestic_series,'--', color='#B69F5E', label='내국인', marker='o', markersize=3)
        plt.grid(True, axis='y',alpha=0.5, linestyle='--')
        plt.xticks( rotation=45)
        plt.legend(loc='upper left')
        plt.savefig ( url + '5_1_' + tag +  '_내외국인 연도별 동향.jpg' , dpi=300 )
        plt.close()


def pre_pt_trend (url, df) :
    # ------------------------------------------------------------------------------------------------1. 연도별 출원 동향 그래프
    df_1=df[["대표출원인",'출원일']]

    s1 = []
    time_format ="%Y.%m.%d"
    for tag in df_1['출원일']:
          s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    df_1['출원일']= s1

    df_1['출원일'] = pd.to_datetime(df_1['출원일'], format=time_format)
    df_1.loc[:, '출원연도'] = df_1['출원일'].dt.year

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

    df_1_11.to_excel(url+'0.pre_연도별출원동향(엑셀).xlsx')

    # s1 = []
    # for tag in df_1_1['출원연도']:
    #       s1.append ( str(tag) )
    #
    # df_1_1['출원연도']=s1


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

    plt.savefig(url+'0_pre_전체출원동향.jpg',dpi=300)
    plt.close()

    # -----------------------------------------------------------------------------------------------------------4. 경쟁사 뽑기
    # ----------------------------------------------------------------------------------------------4.1 출원인별 그래프
    df_4=df[["대표출원인",'출원일']]

    s1 = []
    time_format ="%Y.%m.%d"
    for tag in df_4['출원일']:
          s1.append ( datetime.datetime.strptime ( tag , time_format ) )

    df_4.loc[:,'출원일']= s1
    df_4['출원일'] = pd.to_datetime(df_4['출원일'], format=time_format)
    df_4.loc[:, '출원연도'] = df_4['출원일'].dt.year

    s1 = []
    s1= counts = df_4.groupby('대표출원인')['출원연도'].value_counts().unstack(fill_value=0)
    s1['합계1'] = s1.iloc[:, :].sum(axis=1)
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

    df_4_1.to_excel(url+'0_4_1_0.상위10개 출원인.xlsx', index=True)
    top_30.to_excel(url+'0_4_1_1.상위30개 출원인.xlsx', index=True)

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

    plt.savefig ( url + '0_4_2_0 상위 출원인 10개 전체 그래프'+'.jpg' , dpi=300 )
    plt.close()

def df_check (df) :
    if '대표출원인' not in df.columns :
        print ( "대표출원인 컬럼이 없습니다." )
        exit ()  # 프로그램 종료

    if 'JPPAJ' in df['국가코드'].values :
        print ( "JPPAJ가 있습니다." )
        exit ()  # 프로그램 종료

    if any ( df['대표출원인'].isin ( ['' , '-'] ) ) :
        if '' in df['대표출원인'] :
            print ( "대표출원인에 빈칸이 있습니다." )
            exit ()  # 프로그램 종료
        if '-' in df['대표출원인'] :
            print ( "대표출원인에 '-'가 있습니다." )
            exit ()  # 프로그램 종료

    if any ( df['제1출원인국적'].isin ( ['' , '-'] ) ) :
        if '' in df['제1출원인국적'] :
            print ( "제1출원인국적에 빈칸이 있습니다." )
            exit ()  # 프로그램 종료
        if '-' in df['제1출원인국적'] :
            print ( "제1출원인국적에 '-'가 있습니다." )
            exit ()  # 프로그램 종료

def delete_duplicate (url, file_name) :

    df=pd.read_excel(url+file_name)

    # 출원번호가 중복된 경우 랜덤하게 하나의 행만 남기고 나머지 행 삭제
    duplicate_mask = df.duplicated(subset='출원번호')
    duplicate_indices = df[duplicate_mask].index

    # 중복된 행들 중에서 랜덤하게 하나의 행 선택
    random_index = np.random.choice(duplicate_indices)

    # 중복된 행들 중에서 선택한 행을 제외한 나머지 행 삭제
    df.drop(index=duplicate_indices[duplicate_indices != random_index], inplace=True)

    df.to_excel(url+'base1.xlsx')






# color=['#C58940','#D4A657','#E5BA73','#FAEAB1','#FAF8F1']
#
#
# for tag in s1:
#     print(tag)
#
# d
#
#
#
#
#
# nation_ratios = {}