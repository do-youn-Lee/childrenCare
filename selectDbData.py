import sqlite3
from sqlite3 import Error


class SelectDatabase(object):

    def __init__(self):

        self.sql_rec1 = '''
            SELECT T1.거래처, DATE('1899-12-30', '+'||T1.거래날짜||' DAYS') AS 거래날짜, T1.품명, T1.규격, 
                SUM(T1.수량) AS 수량, SUM(T1.단가) AS 단가, SUM(T1.금액) AS 금액, 
                T2.구분내용 AS 지원구분, T3.분기지원총액 , T3.쌀지원액, T3.제주산농산물
            FROM Data T1, 지원구분 T2, 권역배정 T3
            WHERE T1.품명 = T2.품목
            AND REPLACE(T1.거래처, '어린이집', '') = T3.어린이집
            AND STRFTIME('%Y%m%d',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS')) BETWEEN ? And ?
            AND T1.거래처 LIKE ?
            AND T1.구분 = ?
            GROUP BY T1.거래처, T1.거래날짜, T1.품명, T1.규격, T2.구분내용, T3.쌀지원액, T3.제주산농산물, T3.분기지원총액
            '''

        self.sql_rec2 = '''
            SELECT T1.거래처, T2.구분내용 AS 지원구분, DATE('1899-12-30', '+'||T1.거래날짜||' DAYS') AS 거래날짜, 
                    MIN(T1.품명) AS 품명, SUM(T1.금액) AS 금액, T3.원장, T3.주소, T3.전화번호,
                    T3.분기지원총액, T3.쌀지원액, T3.제주산농산물
            FROM Data T1, 지원구분 T2, 권역배정 T3
            WHERE T1.품명 = T2.품목
            AND REPLACE(T1.거래처, '어린이집', '') = T3.어린이집
            AND STRFTIME('%Y%m%d',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS')) BETWEEN ? And ?
            AND T1.거래처 LIKE ?
            AND T1.구분 =  ?
            AND T2.구분내용 <> '자가결제'
            GROUP BY T1.거래처, T1.거래날짜, T2.구분내용, T3.원장, T3.주소, T3.전화번호, 
                    T3.분기지원총액, T3.쌀지원액, T3.제주산농산물
            '''

        self.sql_rec3 = '''
            SELECT T1.거래처, DATE('1899-12-30', '+'||T1.거래날짜||' DAYS') AS 거래날짜, T1.품명, T1.규격, 
                    SUM(T1.수량) AS 수량, SUM(T1.단가) AS 단가, SUM(T1.금액) AS 금액, T2.구분내용 AS 지원구분
            FROM Data T1, 지원구분 T2
            WHERE T1.품명 = T2.품목
            AND STRFTIME('%Y%m%d',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS')) BETWEEN ? And ?
            AND T1.거래처 LIKE ?
            AND T1.구분 =  ?
            AND T2.구분내용 = '자가결제'
            GROUP BY T1.거래처, T1.거래날짜, T1.품명, T1.규격, T2.구분내용
            '''

        self.sql_sumPart = '''
            SELECT T3.어린이집, T3.분기지원총액, T3.쌀지원액, T3.제주산농산물, 
                    IFNULL(TT1.month1, 0), IFNULL(TT1.month2, 0), IFNULL(TT1.month3, 0), IFNULL(TT1.합계, 0)                    
            FROM
                (SELECT 어린이집||'어린이집' AS '어린이집', 분기지원총액, 쌀지원액, 제주산농산물 FROM 권역배정) T3
            LEFT OUTER JOIN
                (SELECT T1.거래처,                    
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= ? THEN T1.금액 END), 0) AS 'month1',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= ? THEN T1.금액 END), 0) AS 'month2',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= ? THEN T1.금액 END), 0) AS 'month3',
                    IFNULL(CASE WHEN SUM(T1.금액)=0 THEN 0 ELSE SUM(T1.금액) END, 0) AS '합계'
                    FROM Data T1, 지원구분 T2
                    WHERE T1.품명 = T2.품목                    
                    AND T2.구분내용 <> ?
                    AND T1.거래처 LIKE ?
                    AND T1.구분 = ?
                    AND STRFTIME('%Y',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= ?
                    AND T2.구분내용 <> '자가결제'
                GROUP BY T1.거래처) TT1
            ON T3.어린이집 = TT1.거래처
            '''

        self.sql_sumYear = '''
            SELECT T3.어린이집||'어린이집', T3.분기지원총액, T3.쌀지원액, T3.제주산농산물, 
                    IFNULL(TT1.'03월', 0), IFNULL(TT1.'04월', 0), IFNULL(TT1.'05월', 0), IFNULL(TT1.'06월', 0), 
                    IFNULL(TT1.'07월', 0), IFNULL(TT1.'08월', 0), IFNULL(TT1.'09월', 0), IFNULL(TT1.'10월', 0), 
                    IFNULL(TT1.'11월', 0), IFNULL(TT1.'12월', 0), IFNULL(TT1.합계, 0)
            FROM
                (SELECT 어린이집||'어린이집' AS 어린이집, 분기지원총액, 쌀지원액, 제주산농산물 FROM 권역배정) T3
            LEFT OUTER JOIN
                (SELECT T1.거래처,
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '03' THEN T1.금액 END), 0) AS '03월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '04' THEN T1.금액 END), 0) AS '04월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '05' THEN T1.금액 END), 0) AS '05월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '06' THEN T1.금액 END), 0) AS '06월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '07' THEN T1.금액 END), 0) AS '07월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '08' THEN T1.금액 END), 0) AS '08월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '09' THEN T1.금액 END), 0) AS '09월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '10' THEN T1.금액 END), 0) AS '10월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '11' THEN T1.금액 END), 0) AS '11월',
                    IFNULL(SUM(CASE WHEN STRFTIME('%m',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= '12' THEN T1.금액 END), 0) AS '12월',
                    IFNULL(CASE WHEN SUM(T1.금액)=0 THEN 0 ELSE SUM(T1.금액) END, 0) AS '합계'
                FROM Data T1, 지원구분 T2
                WHERE T1.품명 = T2.품목
                AND T2.구분내용 <> ?                
                AND T1.거래처 LIKE ?
                AND T1.구분 =  ?                
                AND STRFTIME('%Y',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS'))= ?
                AND T2.구분내용 <> '자가결제'
                GROUP BY T1.거래처) TT1
            ON T3.어린이집 = TT1.거래처
            '''

        self.sql_cRec = '''
            SELECT T1.거래처, T3.원장, T3.주소, T3.전화번호, T3.아동현원, T3.분기지원총액, T3.쌀지원액, T3.제주산농산물,
                CASE WHEN IFNULL(SUM(T1.금액),0)=0 THEN 0 ELSE SUM(T1.금액) END AS 납품액합계
            FROM Data T1, 지원구분 T2, 권역배정 T3
            WHERE T1.품명 = T2.품목
            AND REPLACE(T1.거래처, '어린이집', '') = T3.어린이집
            AND STRFTIME('%Y%m%d',DATE('1899-12-30', '+'||T1.거래날짜||' DAYS')) BETWEEN ? And ?
            AND T1.거래처 LIKE ?
            AND T1.구분 =  ?
            AND T2.구분내용 <> '자가결제'
            GROUP BY T1.거래처, T3.원장, T3.주소, T3.전화번호, T3.아동현원, T3.쌀지원액, T3.제주산농산물, T3.분기지원총액
            '''

    def create_connection(self, db_file):
        try:
            conn = sqlite3.connect(db_file)
            return conn
        except Error as e:
            print(e)

        return None

    def select_task_by_priority(self, conn, *args):
        sql = ''
        # args[0] 으로 분기 시작 월만 넘어오게 됨
        part_month2 = ''    # 분기 2번째 달
        part_month3 = ''    # 분기 3번째 달

        # 버튼별 sql
        if args[4] == '거래명세서1':
            sql = self.sql_rec1
        elif args[4] == '거래명세서2':
            sql = self.sql_rec2
        elif args[4] == '거래명세서3':
            sql = self.sql_rec3
        elif args[4] == '집행현황(분기)':
            sql = self.sql_sumPart
            # 분기 구분에 따른 월 변수
            if args[0] == '03':
                part_month2 = '04'
                part_month3 = '05'
            elif args[0] == '06':
                part_month2 = '07'
                part_month3 = '08'
            elif args[0] == '09':
                part_month2 = '10'
                part_month3 = '11'
            elif args[0] == '12':
                part_month2 = '01'
                part_month3 = '02'
        elif args[4] == '집행현황(년)':
            sql = self.sql_sumYear
        elif args[4] == '완료인수증':
            sql = self.sql_cRec

        cur = conn.cursor()

        if args[4] == '집행현황(분기)':
            cur.execute(sql, (args[0], part_month2, part_month3, args[1], args[2], args[3], args[5]))
        else:
            # from_date, to_date, custom_name, sale_divide, button_name
            cur.execute(sql, (args[0], args[1], args[2], args[3]))

        rows = cur.fetchall()
        return rows

    def select_database(self, *args):
        database = "ChildCare.db"
        # create a database connection
        conn = self.create_connection(database)
        with conn:
            if args[4] == '집행현황(분기)':
                result = self.select_task_by_priority(conn, args[0], args[1], '%' + args[2], args[3], args[4], args[5])
            else:   # conn, fromDate, toDate, '%'+customNm, saleDivide
                result = self.select_task_by_priority(conn, args[0], args[1], '%' + args[2], args[3], args[4])
        conn.close()
        return result
