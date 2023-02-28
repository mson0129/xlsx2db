from xlsx2db import XLSX2DB

xlsx2db = XLSX2DB(section="DEFAULT")    # section은 설정 파일(xlsx2db.ini)의 섹션명
output: bool = xlsx2db.convert_xlsx2db(
    path="example.xlsx",                # path는 엑셀 파일의 경로
    table="example"                     # table은 DB 테이블의 이름
)
print(output)                           # True인 경우 정상 종료
