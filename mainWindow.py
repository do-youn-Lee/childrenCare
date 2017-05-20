# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import sqlite3CreateTable
import importExcel
from datetime import datetime
import calendar
import selectDbData
import expertExcel_rec1
import expertExcel_rec2
import expertExcel_rec3
import expertExcel_sumPart
import expertExcel_sumYear
import expertExcel_cRec
import excelFileCheck


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1280, 720)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 1111, 651))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.tableWidget = QtWidgets.QTableWidget(self.gridLayoutWidget)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.gridLayout.addWidget(self.tableWidget, 0, 0, 1, 1)
        self.pushButton_rec2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_rec2.setGeometry(QtCore.QRect(1130, 280, 141, 51))
        self.pushButton_rec2.setObjectName("pushButton_rec2")
        self.pushButton_rec3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_rec3.setGeometry(QtCore.QRect(1130, 340, 141, 51))
        self.pushButton_rec3.setObjectName("pushButton_rec3")
        self.pushButton_sumPart = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_sumPart.setGeometry(QtCore.QRect(1130, 490, 141, 51))
        self.pushButton_sumPart.setObjectName("pushButton_sumPart")
        self.pushButton_sumYear = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_sumYear.setGeometry(QtCore.QRect(1130, 550, 141, 51))
        self.pushButton_sumYear.setObjectName("pushButton_sumYear")
        self.pushButton_cRec = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_cRec.setGeometry(QtCore.QRect(1130, 610, 141, 51))
        self.pushButton_cRec.setObjectName("pushButton_cRec")
        self.pushButton_rec1 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_rec1.setGeometry(QtCore.QRect(1130, 220, 141, 51))
        self.pushButton_rec1.setObjectName("pushButton_rec1")
        self.pushButton_upFile = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_upFile.setGeometry(QtCore.QRect(1130, 40, 141, 51))
        self.pushButton_upFile.setAutoDefault(False)
        self.pushButton_upFile.setObjectName("pushButton_upFile")
        self.comboBox_fileUpload = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_fileUpload.setGeometry(QtCore.QRect(1130, 10, 141, 22))
        self.comboBox_fileUpload.setObjectName("comboBox_fileUpload")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(1130, 120, 71, 21))
        self.label.setScaledContents(False)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.textEdit_custom = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_custom.setGeometry(QtCore.QRect(1130, 140, 71, 21))
        font = QtGui.QFont()
        font.setFamily("굴림")
        font.setPointSize(10)
        self.textEdit_custom.setFont(font)
        self.textEdit_custom.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_custom.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_custom.setTabChangesFocus(True)
        self.textEdit_custom.setObjectName("textEdit_custom")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(1210, 120, 61, 21))
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.textEdit_saleDiv = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_saleDiv.setGeometry(QtCore.QRect(1210, 140, 61, 21))
        font = QtGui.QFont()
        font.setFamily("굴림")
        font.setPointSize(11)
        self.textEdit_saleDiv.setFont(font)
        self.textEdit_saleDiv.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_saleDiv.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_saleDiv.setTabChangesFocus(True)
        self.textEdit_saleDiv.setObjectName("textEdit_saleDiv")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(1130, 100, 141, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setGeometry(QtCore.QRect(1130, 410, 141, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(1130, 170, 141, 21))
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.textEdit_date1 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_date1.setGeometry(QtCore.QRect(1130, 190, 61, 21))
        font = QtGui.QFont()
        font.setFamily("굴림")
        font.setPointSize(9)
        self.textEdit_date1.setFont(font)
        self.textEdit_date1.setInputMethodHints(QtCore.Qt.ImhDate|QtCore.Qt.ImhMultiLine)
        self.textEdit_date1.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_date1.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_date1.setTabChangesFocus(True)
        self.textEdit_date1.setObjectName("textEdit_date1")
        self.textEdit_date2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_date2.setGeometry(QtCore.QRect(1210, 190, 61, 21))
        font = QtGui.QFont()
        font.setFamily("굴림")
        font.setPointSize(9)
        self.textEdit_date2.setFont(font)
        self.textEdit_date2.setInputMethodHints(QtCore.Qt.ImhDate|QtCore.Qt.ImhMultiLine)
        self.textEdit_date2.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_date2.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textEdit_date2.setTabChangesFocus(True)
        self.textEdit_date2.setObjectName("textEdit_date2")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(1130, 440, 71, 21))
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(1200, 440, 71, 21))
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")
        self.comboBox_2 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_2.setGeometry(QtCore.QRect(1130, 460, 76, 22))
        self.comboBox_2.setObjectName("comboBox_2")
        self.comboBox_3 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_3.setGeometry(QtCore.QRect(1210, 460, 61, 22))
        self.comboBox_3.setObjectName("comboBox_3")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(1190, 190, 16, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.comboBox_fileUpload, self.pushButton_upFile)
        MainWindow.setTabOrder(self.pushButton_upFile, self.textEdit_custom)
        MainWindow.setTabOrder(self.textEdit_custom, self.textEdit_saleDiv)
        MainWindow.setTabOrder(self.textEdit_saleDiv, self.textEdit_date1)
        MainWindow.setTabOrder(self.textEdit_date1, self.textEdit_date2)
        MainWindow.setTabOrder(self.textEdit_date2, self.pushButton_rec1)
        MainWindow.setTabOrder(self.pushButton_rec1, self.pushButton_rec2)
        MainWindow.setTabOrder(self.pushButton_rec2, self.pushButton_rec3)
        MainWindow.setTabOrder(self.pushButton_rec3, self.comboBox_2)
        MainWindow.setTabOrder(self.comboBox_2, self.comboBox_3)
        MainWindow.setTabOrder(self.comboBox_3, self.pushButton_sumPart)
        MainWindow.setTabOrder(self.pushButton_sumPart, self.pushButton_sumYear)
        MainWindow.setTabOrder(self.pushButton_sumYear, self.pushButton_cRec)
        MainWindow.setTabOrder(self.pushButton_cRec, self.pushButton_cRec)
        # MainWindow.setWindowState(QtCore.Qt.WindowMaximized)

        # #### sqlite connect #####
        sqlite3CreateTable.main()
        ##########################

        self.pushButton_upFile.clicked.connect(self.file_upload)
        self.pushButton_rec1.clicked.connect(self.rec1_output)
        self.pushButton_rec2.clicked.connect(self.rec2_output)
        self.pushButton_rec3.clicked.connect(self.rec3_output)
        self.pushButton_sumPart.clicked.connect(self.sumPart_output)
        self.pushButton_sumYear.clicked.connect(self.sumYear_output)
        self.pushButton_cRec.clicked.connect(self.cRec_output)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "어린이집처리 v1.4"))
        self.pushButton_rec2.setText(_translate("MainWindow", "거래명세서2"))
        self.pushButton_rec3.setText(_translate("MainWindow", "거래명세서3"))
        self.pushButton_sumPart.setText(_translate("MainWindow", "집행현황(분기)"))
        self.pushButton_sumYear.setText(_translate("MainWindow", "집행현황(년)"))
        self.pushButton_cRec.setText(_translate("MainWindow", "완료인수증"))
        self.pushButton_rec1.setText(_translate("MainWindow", "거래명세서1"))
        self.pushButton_upFile.setText(_translate("MainWindow", "파일 업로드"))
        self.label.setText(_translate("MainWindow", "거래처"))
        self.label_2.setText(_translate("MainWindow", "구분"))
        self.label_3.setText(_translate("MainWindow", "처리기간"))
        self.label_4.setText(_translate("MainWindow", "구분"))
        self.label_5.setText(_translate("MainWindow", "분기"))
        self.label_6.setText(_translate("MainWindow", "~~"))
        self.comboBox_fileUpload.addItems(["Data", "지원구분", "권역배정"])
        self.comboBox_2.addItems(["총괄", "쌀", "제주산"])
        self.comboBox_3.addItems(["1분기", "2분기", "3분기", "4분기"])
        self.textEdit_custom.setText("어린이집")
        self.textEdit_saleDiv.setText("판 매")

        dateYearMonth = datetime.today().strftime("%Y%m")  # // 년월 표시 201704
        endOfMonth = calendar.monthrange(datetime.today().year, datetime.today().month)[1]  # // 현재 '월'의 마지막 날
        self.textEdit_date1.setText(dateYearMonth + '01')
        # self.textEdit_date1.setText('20170301')
        self.textEdit_date2.setText(dateYearMonth + str(endOfMonth))
        # self.textEdit_date2.setText('20170531')

    def file_upload(self):
        msgbox_file_upload = QMessageBox()
        msgbox_file_upload.setGeometry((MainWindow.geometry()))

        comboboxSelect = self.comboBox_fileUpload.currentText()

        ######## fileOpenDialogWindow.py 띄우기 ################
        # from fileOpenDialogWindow import Ui_FileUpload_Dialog
        # window = Ui_FileUpload_Dialog(comboboxSelect)
        # dialog = QtWidgets.QDialog()
        # window.setupUi(dialog)
        # dialog.exec_()
        ##################### 미완료 창닫기 처리 전 불가 #######

        uploadFileName = QFileDialog.getOpenFileName(None, 'Open Excel File', '', 'Excel File (*.xlsx *.xlsm *.xls)')

        # 업로드 파일 시트 확인
        if uploadFileName:
            importExecute = importExcel.Import_Excel(uploadFileName[0], comboboxSelect)

            if not importExecute.importExcel():
                QMessageBox.information(msgbox_file_upload, 'Sheet 체크', '업로드 파일에 해당 시트가 없습니다')
                return

            # excelFileCheck
            excel_check = excelFileCheck.ExcelFileCheck()
            out_range_data = excel_check.dbtable_not_in_check(comboboxSelect)

            if out_range_data:
                QMessageBox.information(msgbox_file_upload, "엑셀 파일 체크", 'Data와 업로드 시트에 일치하지 않는 자료가 있습니다')

                # 포함하지 않는 자료 출력
                self.tableWidget.setRowCount(len(out_range_data))
                self.tableWidget.setColumnCount(1)
                for rows in range(len(out_range_data)):
                    out_range_data_string = ''.join(out_range_data[rows])
                    self.tableWidget.setItem(rows, 0, QtWidgets.QTableWidgetItem(out_range_data_string))

                return
            else:
                if comboboxSelect == 'Data':
                    excel_check.db_data_table_expert()
                # 업로드 확인 메세지박스
                QMessageBox.information(msgbox_file_upload, "업로드 확인", "업로드가 완료되었습니다")

    def rec1_output(self):
        custom_name = self.textEdit_custom.toPlainText()
        sale_divide = self.textEdit_saleDiv.toPlainText()
        from_date = self.textEdit_date1.toPlainText()
        to_date = self.textEdit_date2.toPlainText()
        button_name = self.pushButton_rec1.text()

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(from_date, to_date, custom_name, sale_divide, button_name)
        # //  results return ( list )
        # //    결과 table row 수 : len(results)
        # //    결과 table col 수 : len(results[0])

        if results:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(results))
            self.tableWidget.setColumnCount(len(results[0]))

            # ########### 결과값 테이블에 쓰기 ##############
            for result_rows in range(len(results)):
                for result_cols in range(len(results[0])):
                    self.tableWidget.setItem(result_rows, result_cols,
                                             QtWidgets.QTableWidgetItem(str(results[result_rows][result_cols])))

            # ########### 결과값 엑셀 쓰기 #########
            expertExcel_rec1.expert_excel_rec1(results)

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_recipt1 = QMessageBox()
            msgbox_recipt1.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_recipt1, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def rec2_output(self):
        custom_name = self.textEdit_custom.toPlainText()
        sale_divide = self.textEdit_saleDiv.toPlainText()
        from_date = self.textEdit_date1.toPlainText()
        to_date = self.textEdit_date2.toPlainText()
        button_name = self.pushButton_rec2.text()

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(from_date, to_date, custom_name, sale_divide, button_name)

        if results:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(results))
            self.tableWidget.setColumnCount(len(results[0]))

            # ########### 결과값 테이블에 쓰기 ##############
            for result_rows in range(len(results)):
                for result_cols in range(len(results[0])):
                    self.tableWidget.setItem(result_rows, result_cols,
                                             QtWidgets.QTableWidgetItem(str(results[result_rows][result_cols])))

            # ########### 결과값 엑셀 쓰기 #########
            expertExcel_rec2.expert_excel_rec2(results)

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_recipt2 = QMessageBox()
            msgbox_recipt2.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_recipt2, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def rec3_output(self):
        custom_name = self.textEdit_custom.toPlainText()
        sale_divide = self.textEdit_saleDiv.toPlainText()
        from_date = self.textEdit_date1.toPlainText()
        to_date = self.textEdit_date2.toPlainText()
        button_name = self.pushButton_rec3.text()

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(from_date, to_date, custom_name, sale_divide, button_name)

        if results:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(results))
            self.tableWidget.setColumnCount(len(results[0]))

            # ########### 결과값 테이블에 쓰기 ##############
            for result_rows in range(len(results)):
                for result_cols in range(len(results[0])):
                    self.tableWidget.setItem(result_rows, result_cols,
                                             QtWidgets.QTableWidgetItem(str(results[result_rows][result_cols])))

            # ########### 결과값 엑셀 쓰기 #########
            expertExcel_rec3.expert_excel_rec3(results)

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_recipt3 = QMessageBox()
            msgbox_recipt3.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_recipt3, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def sumPart_output(self):
        custom_name = self.textEdit_custom.toPlainText()
        sale_divide = self.textEdit_saleDiv.toPlainText()
        from_date = self.textEdit_date1.toPlainText()
        to_date = self.textEdit_date2.toPlainText()
        button_name = self.pushButton_sumPart.text()

        if from_date[:4] != to_date[:4]:
            msgbox_query_year = QMessageBox()
            msgbox_query_year.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_query_year, '조회 년도 비교', '입력된 년도가 다릅니다')
            return

        # 구분 처리: 쿼리에서 <> 로 처리하게 되므로 반대로 넣음
        part_dist = ''
        if self.comboBox_2.currentText() == '총괄':
            part_dist = '총괄'
        elif self.comboBox_2.currentText() == '쌀':
            part_dist = '제주산30%'
        elif self.comboBox_2.currentText() == '제주산':
            part_dist = '국내산70%'

        # 분기 처리: 해당 분기 첫 월을 넘김
        part_for_year = ''
        if self.comboBox_3.currentText() == '1분기':
            part_for_year = '03'
        elif self.comboBox_3.currentText() == '2분기':
            part_for_year = '06'
        elif self.comboBox_3.currentText() == '3분기':
            part_for_year = '09'
        elif self.comboBox_3.currentText() == '4분기':
            part_for_year = '12'

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(part_for_year, part_dist, custom_name, sale_divide, button_name, from_date[:4])

        if results:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(results))
            self.tableWidget.setColumnCount(len(results[0]))

            # ########### 결과값 테이블에 쓰기 ##############
            for result_rows in range(len(results)):
                for result_cols in range(len(results[0])):
                    self.tableWidget.setItem(result_rows, result_cols,
                                             QtWidgets.QTableWidgetItem(str(results[result_rows][result_cols])))

            # ########### 결과값 엑셀 쓰기 #########
            expertExcel_sumPart.expert_excel_sumPart(results, part_dist, self.comboBox_3.currentText(), from_date[:4])

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_sumPart = QMessageBox()
            msgbox_sumPart.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sumPart, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def sumYear_output(self):
        custom_name = self.textEdit_custom.toPlainText()
        sale_divide = self.textEdit_saleDiv.toPlainText()
        from_date = self.textEdit_date1.toPlainText()
        to_date = self.textEdit_date2.toPlainText()
        button_name = self.pushButton_sumYear.text()

        if from_date[:4] != to_date[:4]:
            msgbox_query_year = QMessageBox()
            msgbox_query_year.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_query_year, '조회 년도 비교', '입력된 년도가 다릅니다')
            return

        # 구분 처리: 쿼리에서 <> 로 처리하게 되므로 반대로 넣음
        part_dist = ''
        if self.comboBox_2.currentText() == '총괄':
            part_dist = '총괄'
        elif self.comboBox_2.currentText() == '쌀':
            part_dist = '제주산30%'
        elif self.comboBox_2.currentText() == '제주산':
            part_dist = '국내산70%'

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(part_dist, sale_divide, custom_name, from_date[:4], button_name)

        if results:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(results))
            self.tableWidget.setColumnCount(len(results[0]))

            # ########### 결과값 테이블에 쓰기 ##############
            for result_rows in range(len(results)):
                for result_cols in range(len(results[0])):
                    self.tableWidget.setItem(result_rows, result_cols,
                                             QtWidgets.QTableWidgetItem(str(results[result_rows][result_cols])))

            # ########### 결과값 엑셀 쓰기 #########
            expertExcel_sumYear.expert_excel_sumYear(results, part_dist, self.comboBox_3.currentText(), from_date[:4])

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_sumYear = QMessageBox()
            msgbox_sumYear.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sumYear, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def cRec_output(self):
        custom_name = self.textEdit_custom.toPlainText()
        sale_divide = self.textEdit_saleDiv.toPlainText()
        from_date = self.textEdit_date1.toPlainText()
        to_date = self.textEdit_date2.toPlainText()
        button_name = self.pushButton_cRec.text()

        if from_date[:4] != to_date[:4]:
            msgbox_query_year = QMessageBox()
            msgbox_query_year.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_query_year, '조회 년도 비교', '입력된 년도가 다릅니다')
            return

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(from_date, to_date, custom_name, sale_divide, button_name)

        if results:
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(len(results))
            self.tableWidget.setColumnCount(len(results[0]))

            # ########### 결과값 테이블에 쓰기 ##############
            for result_rows in range(len(results)):
                for result_cols in range(len(results[0])):
                    self.tableWidget.setItem(result_rows, result_cols,
                                             QtWidgets.QTableWidgetItem(str(results[result_rows][result_cols])))

            # ########### 결과값 엑셀 쓰기 #########
            expertExcel_cRec.expert_excel_cRec(results, from_date, to_date, custom_name, sale_divide, button_name)

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_cRec = QMessageBox()
            msgbox_cRec.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_cRec, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())