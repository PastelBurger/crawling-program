
# 파일명을 변수에 저장합니다.
file_name = "Crawling.py"

# 파일을 읽어옵니다.
with open(file_name, 'r', encoding='utf-8') as file:
    # 파일 내용을 읽어와서 문자열로 저장합니다.
    code = file.read()

# 코드를 실행합니다.
exec(code)


# 파일명을 변수에 저장합니다.
file_name = "News.py"

# 파일을 읽어옵니다.
with open(file_name, 'r', encoding='utf-8') as file:
    # 파일 내용을 읽어와서 문자열로 저장합니다.
    code = file.read()

# 코드를 실행합니다.
exec(code)

