import win32com.client
import pandas as pd
import time

outlook=win32com.client.Dispatch("Outlook.Application")

url="C:/작업서류G/작업서류/1.기술특례상장/1.클리노믹스\메일전송/"
name="주소_김범준.xlsx"
ex=pd.read_excel(url+name)

s1=ex['메일주소']
s2=ex['성명']
s3=ex['변리사']
s4=ex['회사']
s5=[] # 변리사 이메일
s6=[] # 변리사 전화번호
for tag in s3:
    if tag=="이창민":
        s5.append("cmlee@sjpip.co.kr")
        s6.append ("010 4262 7227")
    if tag=="김범준":
        s5.append("bjkim@sjpip.co.kr")
        s6.append ("010 3889 6525")
    if tag=="한상은":
        s5.append("hanse@sjpip.co.kr")
        s6.append ("010 7133 0957")
    if tag=="송민정":
        s5.append("mjsong@sjpip.co.kr")
        s6.append ("010 3169 6583")



f1='클리노믹스_동향조사_에스제이파트너스_2023.01.pdf'
f2='SJP 소개서_최종.pdf'
# attach_files=[url+f1,url+f2]
# for tag in attach_files:
#     send_mail.Attachments.Add(tag)


# -------------------------------------------------------
s_mail = outlook.CreateItem (0)
ss=1
for tag in range(0,len(s1)) :
    s_mail = outlook.CreateItem ( 0 )
    attach_files = [url + f1 , url + f2]
    for tag1 in attach_files :
        s_mail.Attachments.Add ( tag1 )
    s_mail.CC = s5[tag]+";sjp@sjpip.co.kr"
    s_mail.To = s1[tag]
    s_mail.Subject = "[SJP특허/"+s4[tag]+"]기술특례상장 기업 분석 자료 송부의 건"
    a="<html><body style=\"font-family:'맑은 고딕(본문 한글)'; font-size:10pt;\">"
    b="<body>안녕하세요, "+s2[tag]+"님<br/> 에스제이 파트너스 "+s3[tag]+" 변리사 입니다.<br/><br/>"
    c="당사에서는 2023년을 맞이하여, 새로운 프로젝트를 시작하였습니다.<br/><br/>기술특례상장한 기업들의 특허를 분석하고, 그 인사이트를 당사의 고객들에게 전달하고자 하는 것인데요.<br/><br/>첫번째로, 주식회사 클리노믹스의 특허를 분석해 보았습니다.<br/><br/>당사의 자료가 기업 활동에 도움이 되길 바라며, 향후에도 지속적으로 기업활동에 도움이 되는 자료들을 전달 드릴 수 있도록 하겠습니다.<br/><br/>기타 문의사항 있으시다면 언제든지 편하게 연락 부탁드립니다.<br/><br/>감사합니다.<br/><br/>"
    d=s3[tag]+"드림<br/><br/>"
    e="<hr><div style=\"font-family:'맑은 고딕(본문 한글)';font-size:10pt;\"><b>대표 변리사 "+s3[tag]+"</b></div><div style=\"font-family:'맑은 고딕(본문 한글)';font-size:10pt;color: #B69F5E;\"><b>특허법률사무소 에스제이파트너스</b></div><div style=\"font-family:'맑은 고딕(본문 한글)';font-size:10pt;\"><b>특허</b>: 서울시 강남구 언주로 706, 시그너스빌딩 4층 / <b>법무</b>: 서울시 중구 퇴계로 100 (회현동 2가 88) 스테이트타워남산 7층</div><div style=\"font-family:'Tahoma';font-size:10pt;\">TEL: 02 2039 6507 | HP: "+s6[tag]+" | FAX: 02 2039 6508 | EMAIL: "+s5[tag]+" | WEB SITE: <b>특허</b>: www.sjpip.co.kr <b>법무</b>: www.sjplaw.co.kr </div><hr><div style=\"font-family:'맑은 고딕(본문 한글)';font-size:9pt;\">지정된 수신인만이 본 e-mail과 첨부자료를 이용할 수 있습니다. 또한 본 e-mail과 첨부자료는 임의로 공개되어서는 안되며, 어떠한 경우에도 무단 이용, 공개 및 배포가 허용되지 않습니다. 귀하가 지정된 수신인이 아니라면 즉시 발송인에게 회신 e-mail로 이와 같은 사실을 알려주신 후 본 e-mail과 첨부자료를 완전히 삭제하여 주시기 바랍니다. 자세한 내용은 웹사이트를 방문하여 주시기 바랍니다. www.sjpip.co.kr</div></body></html>"
    body=a+b+c+d+e
    s_mail.HTMLBody = body
    s_mail.send
    print(ss)
    ss=ss+1






# send_mail.Display(True)

