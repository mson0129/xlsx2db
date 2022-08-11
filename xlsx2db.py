"""
# XLSX2DB
XLSX 파일을 DB 테이블에 입력합니다.

## Contributor
- 손남규(snk@fasoo.com)
"""

# 내부 모듈
import datetime
import configparser
import pprint   # 개발용
pp = pprint.PrettyPrinter(indent=4, sort_dicts=False)

# 외부 모듈
import openpyxl
import pymssql

class XLSX2DB:
    """
    # XSLX2DB
    XLSX 파일을 읽어 DB에 INSERT 합니다.
    """
    __debug: bool = False                   # 개발용
    dbms_name: str                          # DBMS 종류
    dbms_server: str                        # DBMS 서버 주소
    dbms_username: str                      # DBMS 사용자 이름
    dbms_password: str                      # DBMS 사용자 암호
    database_name: str                      # 데이터베이스 이름
    table_name: str                         # 테이블 이름

    def __init__(self) -> None:
        config = configparser.ConfigParser()
        config.read("xlsx2db.ini")
        self.dbms_name = config["DEFAULT"]["dbms_name"]
        self.dbms_server = config["DEFAULT"]["dbms_server"]
        self.dbms_username = config["DEFAULT"]["dbms_username"]
        self.dbms_password = config["DEFAULT"]["dbms_password"]
        self.database_name = config["DEFAULT"]["database_name"]
        self.table_name = config["DEFAULT"]["table_name"]

    def convert_xlsx2db(self, path: str) -> bool:
        """
        # convert_xlsx2db()
        XLSX 파일을 읽어, 1행에 있는 컬럼명을 기준으로 DB에 데이터를 삽입합니다.

        ## Args:
        1. path: str = "example.xlsx"                       # 엑셀 파일 경로.

        ## Returns:
        output: bool = True # 정상실행일 경우 True
        """
        output: bool = True

        # Part I. Loading Excel File
        workbook = openpyxl.load_workbook(path)
        worksheet = workbook.active

        print("{dt}: Start of loading the file...".format(dt = datetime.datetime.now().strftime("%Y-%m-%d T %H:%M:%S")))
        records: list = []
        i: int = 0
        for row in worksheet.iter_rows(values_only=True):
            if i == 0:
                # Names of column(컬럼명)
                column_names = list(row)
            else:
                # Values of record(레코드값)
                records.append(tuple([list(row)[i] for i in range(len(column_names)) if column_names[i] is not None]))
            i += 1
        print("{dt}: End of loading the file...".format(dt = datetime.datetime.now().strftime("%Y-%m-%d T %H:%M:%S")))

        if self.__debug:
            print("column_names = ")
            pp.pprint(column_names)
            print("records = ")
            pp.pprint(records)

        # Part II. Execute Query
        try:
            if self.dbms_name.lower() == "mssql":
                connection = pymssql.connect(server = self.dbms_server, user = self.dbms_username, password = self.dbms_password, database = self.database_name)
                cursor = connection.cursor()
                
                columns = ", ".join(column_names)
                values_placeholder = ", ".join(["%s"]*len([column_name for column_name in column_names if column_name is not None]))
                query = f"""INSERT INTO {self.table_name} ({columns}) VALUES ({values_placeholder})"""

                if self.__debug:
                    print("query = ")
                    print(query)

                cursor.executemany(query, records)
                connection.commit()
                connection.close()
        except Exception as e:
            print("Database Error: ", e)
            output = False

        return output


    def convert_xlsx2dictlist(self, path: str) -> list:
        """
        # convert_xlsx2list()
        XLSX 파일을 읽어, 1행에 있는 컬럼명을 기준으로 DB에 삽입하기 위한 리스트를 생성합니다.

        ## Args:
        1. path: str = "example.xlsx"                       # 엑셀 파일 경로.

        ## Returns:
        output: list = [{"column1": "value1", ...}, ...]    # DB에 삽입하기 위한 Dictionary로 이루어진 리스트. 리스트 내 개별 원소 값이 레코드 하나에 대응함.
        """
        output: list = []

        workbook = openpyxl.load_workbook(path)
        worksheet = workbook.active

        print("{dt}: Start of loading the file...".format(dt = datetime.datetime.now().strftime("%Y-%m-%d T %H:%M:%S")))
        i: int = 0
        for row in worksheet.iter_rows(values_only=True):
            if i == 0:
                # Names of column(컬럼명)
                column_names = list(row)
            else:
                # Values of record(레코드값)
                output.append({column_names[i]: list(row)[i] for i in range(len(column_names)) if column_names[i] is not None})
            i += 1
        print("{dt}: End of loading the file...".format(dt = datetime.datetime.now().strftime("%Y-%m-%d T %H:%M:%S")))

        return output


    def convert_dictlist2sql(self, data: list) -> str:
        """
        # convert_dictlist2sql()

        ## Args:
        1. data: list = [{"column1": "value1", ...}, ...]   # DB에 삽입하기 위한 Dictionary로 이루어진 리스트. 리스트 내 개별 원소 값이 레코드 하나에 대응함.

        ## Returns:
        output: str = "INSERT INTO example (column1, ...) VALUES ('value1', ...); ..."
        """
        output: str = ""

        for row in data:
            columns = ', '.join(list(row.keys()))
            values = ','.join(["'" + v + "'" for v in list(row.values())])

            pp.pprint(values)

            output += f"INSERT INTO {self.table_name} ({columns}) VALUES ({values});\n"

        return output


    def execute_query(self, query: str) -> bool:
        """
        # execute_query()
        쿼리를 실행합니다.

        ## Args:
        1. query: str = "INSERT INTO example (column1, ...) VALUES ('value1', ...); ..."

        ## Returns:
        output: bool = True # 정상실행일 경우 True
        """
        output: bool = True

        try:
            if self.dbms_name.lower() == "mssql":
                connection = pymssql.connect(server = self.dbms_server, user = self.dbms_username, password = self.dbms_password, database = self.database_name)
                cursor = connection.cursor()
                cursor.execute(query)
                connection.commit()
                connection.close()
        except:
            output = False

        return output


if __name__ == "__main__":
    xlsx2db = XLSX2DB()
    path: str = "example.xlsx"
    output: bool = xlsx2db.convert_xlsx2db(path=path)
    print(output)