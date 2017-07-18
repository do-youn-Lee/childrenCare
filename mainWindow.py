# -*- coding: utf-8 -*-

import sys
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
import clipboard


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1280, 720)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.horizontalLayout.addWidget(self.tableWidget)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.verticalLayout_4.setContentsMargins(-1, 0, -1, -1)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setSpacing(2)
        self.verticalLayout.setObjectName("verticalLayout")
        self.comboBox_fileUpload = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_fileUpload.setObjectName("comboBox_fileUpload")
        self.verticalLayout.addWidget(self.comboBox_fileUpload)
        self.pushButton_upFile = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_upFile.sizePolicy().hasHeightForWidth())
        self.pushButton_upFile.setSizePolicy(sizePolicy)
        self.pushButton_upFile.setObjectName("pushButton_upFile")
        self.verticalLayout.addWidget(self.pushButton_upFile)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem1)
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.verticalLayout.addWidget(self.line_2)
        self.verticalLayout_4.addLayout(self.verticalLayout)
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("굴림")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 1, 0, 1, 1)
        self.dateEdit_to = QtWidgets.QDateEdit(self.centralwidget)
        self.dateEdit_to.setObjectName("dateEdit_to")
        self.gridLayout.addWidget(self.dateEdit_to, 3, 1, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 2, 0, 1, 1)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem2, 2, 1, 1, 1)
        self.label_saleDiv = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_saleDiv.sizePolicy().hasHeightForWidth())
        self.label_saleDiv.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("굴림")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.label_saleDiv.setFont(font)
        self.label_saleDiv.setAlignment(QtCore.Qt.AlignCenter)
        self.label_saleDiv.setObjectName("label_saleDiv")
        self.gridLayout.addWidget(self.label_saleDiv, 1, 1, 1, 1)
        self.dateEdit_from = QtWidgets.QDateEdit(self.centralwidget)
        self.dateEdit_from.setObjectName("dateEdit_from")
        self.gridLayout.addWidget(self.dateEdit_from, 3, 0, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 0, 0, 1, 1)
        self.label_custom = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_custom.setFont(font)
        self.label_custom.setObjectName("label_custom")
        self.gridLayout.addWidget(self.label_custom, 0, 1, 1, 1)
        self.verticalLayout_4.addLayout(self.gridLayout)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setSpacing(3)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.pushButton_rec1 = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_rec1.sizePolicy().hasHeightForWidth())
        self.pushButton_rec1.setSizePolicy(sizePolicy)
        self.pushButton_rec1.setObjectName("pushButton_rec1")
        self.verticalLayout_2.addWidget(self.pushButton_rec1)
        self.pushButton_rec2 = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_rec2.sizePolicy().hasHeightForWidth())
        self.pushButton_rec2.setSizePolicy(sizePolicy)
        self.pushButton_rec2.setObjectName("pushButton_rec2")
        self.verticalLayout_2.addWidget(self.pushButton_rec2)
        self.pushButton_rec3 = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_rec3.sizePolicy().hasHeightForWidth())
        self.pushButton_rec3.setSizePolicy(sizePolicy)
        self.pushButton_rec3.setObjectName("pushButton_rec3")
        self.verticalLayout_2.addWidget(self.pushButton_rec3)
        self.verticalLayout_4.addLayout(self.verticalLayout_2)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout_4.addWidget(self.line)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setSpacing(5)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.comboBox_2 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_2.setObjectName("comboBox_2")
        self.verticalLayout_3.addWidget(self.comboBox_2)
        self.comboBox_3 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_3.setObjectName("comboBox_3")
        self.verticalLayout_3.addWidget(self.comboBox_3)
        self.pushButton_sumPart = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_sumPart.sizePolicy().hasHeightForWidth())
        self.pushButton_sumPart.setSizePolicy(sizePolicy)
        self.pushButton_sumPart.setObjectName("pushButton_sumPart")
        self.verticalLayout_3.addWidget(self.pushButton_sumPart)
        self.pushButton_sumYear = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_sumYear.sizePolicy().hasHeightForWidth())
        self.pushButton_sumYear.setSizePolicy(sizePolicy)
        self.pushButton_sumYear.setObjectName("pushButton_sumYear")
        self.verticalLayout_3.addWidget(self.pushButton_sumYear)
        self.pushButton_cRec = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_cRec.sizePolicy().hasHeightForWidth())
        self.pushButton_cRec.setSizePolicy(sizePolicy)
        self.pushButton_cRec.setObjectName("pushButton_cRec")
        self.verticalLayout_3.addWidget(self.pushButton_cRec)
        self.verticalLayout_4.addLayout(self.verticalLayout_3)
        self.horizontalLayout.addLayout(self.verticalLayout_4)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1280, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.comboBox_fileUpload, self.pushButton_upFile)
        MainWindow.setTabOrder(self.pushButton_upFile, self.dateEdit_from)
        MainWindow.setTabOrder(self.dateEdit_from, self.dateEdit_to)
        MainWindow.setTabOrder(self.dateEdit_to, self.pushButton_rec1)
        MainWindow.setTabOrder(self.pushButton_rec1, self.pushButton_rec2)
        MainWindow.setTabOrder(self.pushButton_rec2, self.pushButton_rec3)
        MainWindow.setTabOrder(self.pushButton_rec3, self.comboBox_2)
        MainWindow.setTabOrder(self.comboBox_2, self.comboBox_3)
        MainWindow.setTabOrder(self.comboBox_3, self.pushButton_sumPart)
        MainWindow.setTabOrder(self.pushButton_sumPart, self.pushButton_sumYear)
        MainWindow.setTabOrder(self.pushButton_sumYear, self.pushButton_cRec)
        MainWindow.setTabOrder(self.pushButton_cRec, self.tableWidget)

        # #### sqlite connect #####
        sqlite3CreateTable.main()

        self.pushButton_upFile.clicked.connect(self.file_upload)
        self.pushButton_rec1.clicked.connect(self.rec1_output)
        self.pushButton_rec2.clicked.connect(self.rec2_output)
        self.pushButton_rec3.clicked.connect(self.rec3_output)
        self.pushButton_sumPart.clicked.connect(self.sumPart_output)
        self.pushButton_sumYear.clicked.connect(self.sumYear_output)
        self.pushButton_cRec.clicked.connect(self.cRec_output)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "어린이집처리 v2.3"))
        self.pushButton_upFile.setText(_translate("MainWindow", "파일 업로드"))
        self.label.setText(_translate("MainWindow", "구분:"))
        self.label_3.setText(_translate("MainWindow", "처리기간"))
        self.label_saleDiv.setText(_translate("MainWindow", "판 매"))
        self.label_4.setText(_translate("MainWindow", "거래처:"))
        self.label_custom.setText(_translate("MainWindow", "어린이집"))
        self.pushButton_rec1.setText(_translate("MainWindow", "거래명세서1"))
        self.pushButton_rec2.setText(_translate("MainWindow", "거래명세서2"))
        self.pushButton_rec3.setText(_translate("MainWindow", "거래명세서3"))
        self.pushButton_sumPart.setText(_translate("MainWindow", "집행현황(분기)"))
        self.pushButton_sumYear.setText(_translate("MainWindow", "집행현황(년)"))
        self.pushButton_cRec.setText(_translate("MainWindow", "완료인수증"))
        self.comboBox_fileUpload.addItems(["Data", "지원구분", "권역배정"])
        self.comboBox_2.addItems(["총괄", "쌀", "제주산"])
        self.comboBox_3.addItems(["1분기", "2분기", "3분기", "4분기"])

        dateYearMonth = datetime.today().strftime("%Y-%m")  # // 년월 표시 201704
        from_date = datetime.strptime(dateYearMonth + '-01', '%Y-%m-%d')
        endOfMonth = calendar.monthrange(datetime.today().year, datetime.today().month)[1]  # // 현재 '월'의 마지막 날
        to_date = datetime.strptime(dateYearMonth + '-' + str(endOfMonth), '%Y-%m-%d')

        # 날짜 세팅
        self.dateEdit_from.setDate(from_date)
        self.dateEdit_to.setDate(to_date)

        # progress Bar
        # self.progressBar = QProgressBar()
        # self.progressBar.setRange(0, 10000)
        # self.progressBar.setValue(0)
        # self.progressBar.setGeometry(500, 200, 401, 31)
        # self.progressBar.hide()

    def file_upload(self):
        msgbox_file_upload = QMessageBox()
        msgbox_file_upload.setGeometry(MainWindow.geometry())

        comboboxSelect = self.comboBox_fileUpload.currentText()

        ######## fileOpenDialogWindow.py 띄우기 ################
        # from fileOpenDialogWindow import Ui_FileUpload_Dialog
        # window = Ui_FileUpload_Dialog(comboboxSelect)
        # dialog = QtWidgets.QDialog()
        # window.setupUi(dialog)
        # dialog.exec_()
        ##################### 미완료 창닫기 처리 전 불가 #######

        try:
            self.uploadFileName = QFileDialog.getOpenFileName(None, 'Open Excel File', '', 'Excel File (*.xlsx *.xlsm *.xls)')
        except Exception as e:
            print(e)

        # 업로드 파일 시트 확인
        if self.uploadFileName:
            importExecute = importExcel.Import_Excel(self.uploadFileName[0], comboboxSelect)

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
                self.table_selected_item = ''  # clipboard 복사용 저장 변수
                for rows in range(len(out_range_data)):
                    out_range_data_string = ''.join(out_range_data[rows])
                    self.tableWidget.setItem(rows, 0, QtWidgets.QTableWidgetItem(out_range_data_string))
                    self.table_selected_item += out_range_data_string + '\n'    # clipboard 복사용 데이터 저장

                # 선택 아이템 Ctrl+C to clipboard
                clipboard.copy(self.table_selected_item)

                if comboboxSelect == 'Data':
                    # Data 엑셀에서 '어린이집' 만 별도의 엑셀 파일로 출력
                    excel_check.db_data_table_expert()

                return
            else:
                # 업로드 확인 메세지박스
                QMessageBox.information(msgbox_file_upload, "업로드 확인", "업로드가 완료되었습니다")

    def rec1_output(self):
        custom_name = self.label_custom.text()
        sale_divide = self.label_saleDiv.text()
        from_date = self.dateEdit_from.text()[:4] + self.dateEdit_from.text()[5:7] + self.dateEdit_from.text()[8:10]
        to_date = self.dateEdit_to.text()[:4] + self.dateEdit_to.text()[5:7] + self.dateEdit_to.text()[8:10]
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
        custom_name = self.label_custom.text()
        sale_divide = self.label_saleDiv.text()
        from_date = self.dateEdit_from.text()[:4] + self.dateEdit_from.text()[5:7] + self.dateEdit_from.text()[8:10]
        to_date = self.dateEdit_to.text()[:4] + self.dateEdit_to.text()[5:7] + self.dateEdit_to.text()[8:10]
        button_name = self.pushButton_rec2.text()
        part_for_year = self.comboBox_3.currentText()

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
            expertExcel_rec2.expert_excel_rec2(results, part_for_year)

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_recipt2 = QMessageBox()
            msgbox_recipt2.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_recipt2, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def rec3_output(self):
        custom_name = self.label_custom.text()
        sale_divide = self.label_saleDiv.text()
        from_date = self.dateEdit_from.text()[:4] + self.dateEdit_from.text()[5:7] + self.dateEdit_from.text()[8:10]
        to_date = self.dateEdit_to.text()[:4] + self.dateEdit_to.text()[5:7] + self.dateEdit_to.text()[8:10]
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
        custom_name = self.label_custom.text()
        sale_divide = self.label_saleDiv.text()
        from_date = self.dateEdit_from.text()[:4] + self.dateEdit_from.text()[5:7] + self.dateEdit_from.text()[8:10]
        to_date = self.dateEdit_to.text()[:4] + self.dateEdit_to.text()[5:7] + self.dateEdit_to.text()[8:10]
        button_name = self.pushButton_sumPart.text()

        if from_date[:4] != to_date[:4]:
            msgbox_query_year = QMessageBox()
            msgbox_query_year.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_query_year, '조회 년도 비교', '입력된 년도가 다릅니다')
            return

        # 구분 처리: 쿼리에서 <> 로 처리하게 되므로 반대로 넣음
        part_dist_dis = ''
        part_dist = ''
        if self.comboBox_2.currentText() == '총괄':
            part_dist_dis = '총괄'
            part_dist = '총괄'
        elif self.comboBox_2.currentText() == '쌀':
            part_dist_dis = '제주산30%'
            part_dist = '국내산70%'
        elif self.comboBox_2.currentText() == '제주산':
            part_dist_dis = '국내산70%'
            part_dist = '제주산30%'

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
        results = select_db_data.select_database(part_for_year, part_dist_dis, custom_name, sale_divide, button_name, from_date[:4])

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
        custom_name = self.label_custom.text()
        sale_divide = self.label_saleDiv.text()
        from_date = self.dateEdit_from.text()[:4] + self.dateEdit_from.text()[5:7] + self.dateEdit_from.text()[8:10]
        to_date = self.dateEdit_to.text()[:4] + self.dateEdit_to.text()[5:7] + self.dateEdit_to.text()[8:10]
        button_name = self.pushButton_sumYear.text()

        if from_date[:4] != to_date[:4]:
            msgbox_query_year = QMessageBox()
            msgbox_query_year.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_query_year, '조회 년도 비교', '입력된 년도가 다릅니다')
            return

        # 구분 처리: 쿼리에서 <> 로 처리하게 되므로 반대로 넣음
        part_dist_dis = ''
        part_dist = ''
        if self.comboBox_2.currentText() == '총괄':
            part_dist_dis = '총괄'
            part_dist = '총괄'
        elif self.comboBox_2.currentText() == '쌀':
            part_dist_dis = '제주산30%'
            part_dist = '국내산70%'
        elif self.comboBox_2.currentText() == '제주산':
            part_dist_dis = '국내산70%'
            part_dist = '제주산30%'

        select_db_data = selectDbData.SelectDatabase()
        results = select_db_data.select_database(part_dist_dis, sale_divide, custom_name, from_date[:4], button_name)

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
            expertExcel_sumYear.expert_excel_sumYear(results, part_dist, from_date[:4])

            # ############### 결과 완료 메세지 박스 ##################
            msgbox_sumYear = QMessageBox()
            msgbox_sumYear.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sumYear, button_name, button_name + " 출력 완료되었습니다")
        else:   # 검색된 값이 없을 때
            msgbox_sql_result = QMessageBox()
            msgbox_sql_result.setGeometry(MainWindow.geometry())
            QMessageBox.information(msgbox_sql_result, 'SQL 조회 결과', '조회된 데이터가 없습니다')

    def cRec_output(self):
        custom_name = self.label_custom.text()
        sale_divide = self.label_saleDiv.text()
        from_date = self.dateEdit_from.text()[:4] + self.dateEdit_from.text()[5:7] + self.dateEdit_from.text()[8:10]
        to_date = self.dateEdit_to.text()[:4] + self.dateEdit_to.text()[5:7] + self.dateEdit_to.text()[8:10]
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
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

