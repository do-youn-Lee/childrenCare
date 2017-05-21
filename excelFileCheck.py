# 1. 테이블에 데이터가 있는지 확인
# 2. 테이블 별 Data테이블에 포함하지 않는 값 확인 -> 리턴
# 3. DB connect
# 4. DB select

import sqlite3
from sqlite3 import Error
from pyexcelerate import Workbook


class ExcelFileCheck(object):

    def __init__(self):
        self.database = "ChildCare.db"

    def create_connection(self, db_file):
        try:
            conn = sqlite3.connect(db_file)
            return conn
        except Error as e:
            print(e)

        return None

    def select_task_by_priority(self, conn, sql):
        sql = sql
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        return rows

    # 테이블 자료 확인
    def dbtable_zero_check(self, args):
        # args 테이블 이름
        sql_table_zero = 'SELECT COUNT(*) FROM ' + args
        conn = self.create_connection(self.database)
        with conn:
            result = self.select_task_by_priority(conn, sql_table_zero)
        conn.close()
        if result[0][0] == 0:
            return True
        else:
            return False

    # Data 테이블 데이터 유무 확인. dbtable_zero_check method
    def dbtable_not_in_check(self, args):

        sql_database_table_check = ''

        if args == '지원구분':
            if self.dbtable_zero_check('Data'):
                return False
            sql_database_table_check = 'SELECT T1.품명 FROM Data T1 WHERE T1.거래처 LIKE \'%어린이집\' AND T1.품명 NOT IN '
            sql_database_table_check += '(SELECT T2.품목 FROM 지원구분 T2)'
        elif args == '권역배정':
            if self.dbtable_zero_check('Data'):
                return False
            sql_database_table_check = 'SELECT T3.어린이집 FROM 권역배정 T3 WHERE T3.어린이집||\'어린이집\' NOT IN '
            sql_database_table_check += '(SELECT T1.거래처 FROM Data T1 WHERE T1.거래처 LIKE \'%어린이집\')'
        elif args == 'Data':
            sql_database_table_check = 'SELECT T1.품명 FROM Data T1 WHERE T1.거래처 LIKE \'%어린이집\' AND T1.품명 NOT IN '
            sql_database_table_check += '(SELECT T2.품목 FROM 지원구분 T2)'
        conn = self.create_connection(self.database)
        with conn:
            result = self.select_task_by_priority(conn, sql_database_table_check)
        conn.close()
        return result

    # Data 테이블에 어린이집 데이터만 파일로 출력
    def db_data_table_expert(self):
        sql_select_data_table = 'SELECT * FROM Data WHERE 거래처 LIKE \'%어린이집\''
        conn = self.create_connection(self.database)
        with conn:
            result = self.select_task_by_priority(conn, sql_select_data_table)
        conn.close()

        # pyexcelerate 엑셀파일에 쓰기
        wb = Workbook()
        wb.new_sheet("Data", data=result)
        wb.save('어린이집Data.xlsx')