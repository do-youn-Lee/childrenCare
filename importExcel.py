import sqlite3
from sqlite3 import Error
import xlrd

class Import_Excel(object):

    def __init__(self, UpLoadFileName, UpLoadKind):
        self.UpLoadFileName = UpLoadFileName    # 경로를 포함함 파일 풀네임
        self.SheetName = UpLoadKind    #   업로드 파일 종류(Data, 1권역변역, 지원구분)
        self.DbName = 'ChildCare.db'    # mainWindow.py 에서 생성(table까지 다 만들어져 있음)

        if self.SheetName == 'Data':
            self.sql = '''
                        INSERT INTO Data (거래처,거래날짜,구분,품명,규격,수량,단가,금액) 
                        VALUES(?,?,?,?,?,?,?,?)
                        '''
        elif self.SheetName == '권역배정':
            self.sql = '''
                        INSERT INTO 권역배정 (어린이집,주소,원장,아동현원,분기지원총액,쌀지원액,제주산농산물,전화번호) 
                        VALUES(?,?,?,?,?,?,?,?)
                        '''
        elif self.SheetName == '지원구분':
            self.sql = '''
                        INSERT INTO 지원구분 (품목,구분내용)
                        VALUES(?,?)
                        '''

    def createConnection(self, db_file):
        try:
            conn = sqlite3.connect(db_file)
            return conn
        except Error as e:
            print(e)

        return None

    def createProject(self,conn, sql, project):
        cur = conn.cursor()
        cur.execute(sql, project)
        return cur.lastrowid

    def importExcel(self):
        workbook = xlrd.open_workbook(self.UpLoadFileName)
        worksheet_names = workbook.sheet_names()

        # 업로드 엑셀 파일에 해당 시트가 있는지 확인
        if not self.SheetName in worksheet_names:
            return False

        worksheet = workbook.sheet_by_name(self.SheetName)
        nrows = worksheet.nrows     # 시트 전체 행 수
        conn = self.createConnection(self.DbName)

        sql = 'DELETE FROM '+ self.SheetName
        self.createProject(conn, sql, '')

        with conn:
            for row_num in range(1, nrows):
                if self.SheetName == 'Data':
                    project = (worksheet.cell_value(row_num, 0),
                               worksheet.cell_value(row_num, 2),
                               worksheet.cell_value(row_num, 3),
                               worksheet.cell_value(row_num, 5),
                               worksheet.cell_value(row_num, 6),
                               worksheet.cell_value(row_num, 7),
                               worksheet.cell_value(row_num, 8),
                               worksheet.cell_value(row_num, 9)
                               );
                elif self.SheetName == '권역배정':
                    project = (worksheet.cell_value(row_num, 2),
                               worksheet.cell_value(row_num, 4),
                               worksheet.cell_value(row_num, 5),
                               worksheet.cell_value(row_num, 7),
                               worksheet.cell_value(row_num, 8),
                               worksheet.cell_value(row_num, 9),
                               worksheet.cell_value(row_num, 10),
                               worksheet.cell_value(row_num, 11)
                               );
                elif self.SheetName == '지원구분':
                    project = (worksheet.cell_value(row_num, 0),
                               worksheet.cell_value(row_num, 2)
                               );

                self.createProject(conn, self.sql, project)

        conn.close()

        return True