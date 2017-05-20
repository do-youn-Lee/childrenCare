import sqlite3
from sqlite3 import Error

def create_connection(db_file):
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return None

def create_table(conn, create_table_sql):
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)


def main():
    database = "ChildCare.db"

    ######### Data Table ############################################
    sql_create_Data_table = """ CREATE TABLE IF NOT EXISTS Data (
                                    거래처  TEXT NOT NULL,
                                    거래날짜 DATE NOT NULL,
                                    구분   TEXT NOT NULL,
                                    품명   TEXT NOT NULL,
                                    규격   TEXT,
                                    수량   TEXT,
                                    단가   TEXT,
                                    금액   TEXT
                                ); """


    ######### 1권역배정 Table #########################################
    sql_create_Share_table = """CREATE TABLE IF NOT EXISTS 권역배정 (
                                    어린이집    TEXT NOT NULL,
                                    주소        TEXT,
                                    원장        TEXT,
                                    아동현원    TEXT,
                                    분기지원총액 TEXT NOT NULL,
                                    쌀지원액    TEXT,
                                    제주산농산물 TEXT,
                                    전화번호    TEXT
                             );"""

    ######### 지원구분 Table ##########################################
    sql_create_Support_table = """CREATE TABLE IF NOT EXISTS 지원구분 (
                                    품목   TEXT NOT NULL,
                                    구분내용 TEXT NOT NULL
                               );"""


    conn = create_connection(database)

    if conn is not None:
        create_table(conn, sql_create_Data_table)
        create_table(conn, sql_create_Share_table)
        create_table(conn, sql_create_Support_table)
        conn.close()
    else:
        print("Error! cannot create the database connection.")
        conn.close()


if __name__ == '__main__':
    main()