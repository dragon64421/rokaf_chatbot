import sqlite3

sql_filename = input("Enter db filename : ") # 수정할 db 파일 이름
csv_filename = input("Enter csv filename : ") # 소스 데이터 (.csv)파일 이름
table_name = input("Enter table name : ") # 수정할 db의 테이블 이름
sizeOfData = int(input("Enter size of data : ")) # 입력할 데이터의 개수

conn = sqlite3.connect(sql_filename, isolation_level = None) # DB 파일 open
c = conn.cursor() # 커서 연결

csv = open(csv_filename).read().split('\n')[1:] # csv 파일 open 및 parsing

count = 0 # 실습때 편의상 100개의 데이터만 입력하도록 함.

for line in csv:
    if count >= sizeOfData: # 실습때 편의상 일부 데이터만 입력하도록 함.
        break
    count += 1
    line = line.split(',') # 라인 parsing 
    command = "insert into " + table_name + " values(" # sql command 작성
    for item in line: # sql command 작성
        try: # 숫자와 문자열의 작성 구분
            item = float(item)
        except:
            item = "'"+item+"'"
        command += str(item) + ","
    command = command.rstrip(',')
    command += ")"
    try:
        c.execute(command) # sql 실행
        print(command)
    except:
        print('failed!')
