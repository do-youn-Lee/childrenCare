import xlsxwriter
import datetime


def expert_excel_rec3(results):
    # 출력할 엑셀 파일 이름
    now = datetime.datetime.now()
    file_name = '어린이집_거래명세서3(미지원)_'+ now.strftime('%Y%m%d_%H%M%S') + '.xlsx'
    now_date = str(now.year) + '년  ' + str(now.month) + '월  ' + str(now.day) + '일'  # 인수 날짜 세팅: 오늘 날짜

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.worksheets()

    sheet_start_row = 8 # 결과 출력 시작 행 설정
    for result_rows in range(len(results)):
        if results[result_rows - 1][0] != results[result_rows][0]:  # 이전 이름과 다르면 새로 생성
            result_name = results[result_rows][0]
            worksheet_name = result_name[:result_name.find('어린이집')] # 뒤에 '어린이집' 은 빼고 시트이름
            worksheet = workbook.add_worksheet(worksheet_name)
            worksheet.set_paper(9)  # 용지 A4 세팅
            worksheet.print_area('A1:G45') # 출력 범위 지정
            worksheet.set_print_scale(90)  # 출력 비율 90%
            # worksheet.set_margins(left=0.5, right=0.5)
            worksheet.center_horizontally()
            sheet_start_row = 8 # 시트를 새로 만들면 다시 초기화

            # ########################################### 시트 상단 ##################################################
            # 전체 열 크기 지정
            column_width = [9.75, 18.13, 7.13, 6 , 9.38, 9.63, 24.13]
            for cell_index in range(len(column_width)):
                worksheet.set_column(cell_index, cell_index, column_width[cell_index])

            # 출력 셀 선 그리기
            result_layout_format = workbook.add_format({'border':1, 'num_format': '#,##0'})
            result_layout_format_left = workbook.add_format({'left':2, 'top':1, 'bottom':1, 'right':1,
                                                             'num_format':'#,##0'})
            result_layout_format_right = workbook.add_format({'left': 1, 'top': 1, 'bottom': 1, 'right': 2,
                                                              'num_format': '#,##0'})
            for result_layout1 in range(8,41):
                for result_layout2 in range(0,7):
                    if result_layout2 == 0:
                        worksheet.write_blank(result_layout1, result_layout2, None, result_layout_format_left)
                    elif result_layout2 == 6:
                        worksheet.write_blank(result_layout1, result_layout2, None, result_layout_format_right)
                    else:
                        worksheet.write_blank(result_layout1, result_layout2, None, result_layout_format)

            # 상단 셀 내용 넣기
            line1_format = workbook.add_format({'align':'center', 'bold':True, 'font_size':24, 'top':2, 'left':2, 'right':2, 'bottom':1})
            mergeA2_A5_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'bold': True, 'font_size': 11, 'top': 1, 'left': 2, 'right': 1, 'bottom': 1})
            mergeB2_D5_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_size': 12, 'border': 1})
            cellF2_F5_format = workbook.add_format({'align':'center', 'valign':'vcenter', 'bold':True, 'font_size': 12, 'border': 1})
            cellG2_G5_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'bold': True, 'font_size': 12, 'top':1, 'left':1, 'right':2, 'bottom':1})
            mergeF6_G7_format = workbook.add_format({'align': 'center', 'valign':'vcenter', 'bold': True, 'font_size': 12, 'top':1, 'left':1, 'right':2, 'bottom':2, 'num_format':'#,##0'})

            worksheet.merge_range('A1:G1', '거래 명세서', line1_format)
            worksheet.merge_range('A2:A5', '어린이집', mergeA2_A5_format)
            worksheet.merge_range('B2:D5', results[result_rows][0], mergeB2_D5_format)
            worksheet.merge_range('E2:E5', '공급자', mergeB2_D5_format)
            worksheet.write('F2', '업체명', cellF2_F5_format)
            worksheet.write('F3', '대표자', cellF2_F5_format)
            worksheet.write('F4', '소 재 지', cellF2_F5_format)
            worksheet.write('F5', '연 락 처', cellF2_F5_format)
            worksheet.write('G2', '굿츠친환경영농조합법인', cellG2_G5_format)
            worksheet.write('G3', '방현철', cellG2_G5_format)
            worksheet.write('G4', '제주시 신산마을동길12', cellG2_G5_format)
            worksheet.write('G5', '064-712-2445', cellG2_G5_format)
            worksheet.merge_range('A6:D7', '아래와 같이 거래하였습니다.', mergeA2_A5_format)
            worksheet.merge_range('E6:E7', '합계액', cellF2_F5_format)
            worksheet.merge_range('F6:G7', '=SUM(F9:F100)', mergeF6_G7_format)

            # 제목행 값 넣기
            line8_format = workbook.add_format({'align':'center', 'bold':True, 'font_size':12, 'border':2})
            cell_title = ['납품일', '품목', '규격', '수량', '단가', '합계', '비고']
            for cell_write in range(len(cell_title)):   # 제목행(10행) 넣기
                worksheet.write(7, cell_write, cell_title[cell_write], line8_format)

        # ########################################  내용 출력 ####################################################
        # 출력 셀 선 그리기용 format
        cell_border_left = workbook.add_format({'left': 2, 'top': 1, 'bottom': 1, 'right': 1, 'num_format': '#,##0', 'font_size':10})
        cell_border_center = workbook.add_format({'border': 1, 'num_format': '#,##0', 'font_size':10})
        cell_border_right = workbook.add_format({'left': 1, 'top': 1, 'bottom': 1, 'right': 2, 'font_size':10})

        # '날짜', '거래처명', '단가', '누계' 셀 입력 처리
        worksheet.write(sheet_start_row, 0, results[result_rows][1], cell_border_left)
        worksheet.write(sheet_start_row, 1, results[result_rows][2], cell_border_center)
        worksheet.write(sheet_start_row, 2, results[result_rows][3], cell_border_center)
        worksheet.write(sheet_start_row, 3, results[result_rows][4], cell_border_center)
        worksheet.write(sheet_start_row, 4, results[result_rows][5], cell_border_center)
        worksheet.write(sheet_start_row, 5, results[result_rows][6], cell_border_center)
        worksheet.write(sheet_start_row, 6, '', cell_border_right)

        sheet_start_row += 1

    workbook.close()