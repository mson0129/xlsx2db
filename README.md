# XLSX2DB

[![MIT License](https://img.shields.io/github/license/mson0129/xlsx2db)](https://www.mit.edu/~amini/LICENSE.md)
![Repo Size](https://img.shields.io/github/repo-size/mson0129/xlsx2db)
![Last Commit](https://img.shields.io/github/last-commit/mson0129/xlsx2db)
![Release Version](https://img.shields.io/github/v/release/mson0129/xlsx2db)

[![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fmson0129%2Fxlsx2db&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://hits.seeyoufarm.com)

XML기반 엑셀 파일(.xlsx)을 읽어 DB 테이블에 입력합니다.
MSSQL, MySQL, MariaDB를 지원합니다.

## 사용법

### XLSX 파일 준비

DB에 입력할 데이터가 있는 파일입니다.
첫 번째 행에는 컬럼명을 입력하고, 나머지 행에는 데이터를 입력합니다.
example.xlsx를 복사 후 수정하여 작성하시면 편리합니다.


### xlsx2db.ini 파일 준비

DB 접속정보를 입력하는 파일입니다.

여러 개의 DB를 사용하고자 하는 경우 ini 파일에 섹션을 여러개 추가하여 사용하셔도 됩니다.

xlsx2db.ini
```
[섹션 이름]
dbms_name = DBMS의 이름입니다. mssql | mysql | mariadb 중 하나의 값을 사용합니다. 대소문자 구분하지 않습니다.
dbms_server = DBMS 서버의 주소입니다. 도메인 형식, IP주소 형식 모두 가능합니다.
dbms_username = DBMS 사용자 이름입니다.
dbms_password = DBMS 사용자의 암호입니다.
database_name = 기본 데이터 베이스 이름입니다.
```

### 실행
새로운 파일을 생성하여 아래와 같이 코드 작성후 사용할 수 있습니다.

example.py
```
from xlsx2db import XLSX2DB

xlsx2db = XLSX2DB(section="DEFAULT")    # section은 설정 파일(xlsx2db.ini)의 섹션명
output: bool = xlsx2db.convert_xlsx2db(
    path="example.xlsx",                # path는 엑셀 파일의 경로
    table="example"                     # table은 DB 테이블의 이름
)
print(output)                           # True인 경우 정상 종료
```

위의 코드 작성이 어려운 경우 xlsx2db.py의 하단 코드를 수정하여 사용하여도 됩니다.
