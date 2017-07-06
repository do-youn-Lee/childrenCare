import xlsxwriter
import datetime


def expert_excel_rec2(results, part):
    # 출력할 엑셀 파일 이름
    now = datetime.datetime.now()
    file_name = '어린이집_거래명세서2(어린이집연합회)_'+ now.strftime('%Y%m%d_%H%M%S') + '.xlsx'
    now_date = str(now.year) + '년  ' + str(now.month) + '월  ' + str(now.day) + '일'  # 인수 날짜 세팅: 오늘 날짜

    part_for_year = ''
    if part == '1분기':
        part_for_year = '1분기:3월~5월'
    elif part == '2분기':
        part_for_year = '2분기:6월~8월'
    elif part == '3분기':
        part_for_year = '3분기:9월~11월'
    elif part == '4분기':
        part_for_year = '4분기:12월'

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.worksheets()

    sheet_start_row = 9 # 결과 출력 시작 행 설정
    for result_rows in range(len(results)):
        if results[result_rows - 1][0] != results[result_rows][0]:  # 이전 이름과 다르면 새로 생성
            result_name = results[result_rows][0]
            worksheet_name = result_name[:result_name.find('어린이집')] # 뒤에 '어린이집' 은 빼고 시트이름
            worksheet = workbook.add_worksheet(worksheet_name)
            worksheet.set_paper(9)  # 용지 A4 세팅
            worksheet.print_area('A2:H44') # 출력 범위 지정
            worksheet.set_print_scale(90)  # 출력 비율 90%
            # worksheet.set_margins(left=0.5, right=0.5)
            worksheet.center_horizontally()
            sheet_start_row = 9 # 시트를 새로 만들면 다시 초기화

            # ########################################### 시트 상단 ##################################################
            # 전체 열 크기 지정
            column_width = [11, 10, 9, 11, 10, 9, 11, 10]
            for cell_index in range(len(column_width)):
                worksheet.set_column(cell_index, cell_index, column_width[cell_index])

            # 출력 셀 선 그리기
            result_layout_format = workbook.add_format({'border':1, 'num_format': '#,##0'})
            result_layout_format_left = workbook.add_format({'left':2, 'top':1, 'bottom':1, 'right':1,
                                                             'num_format':'#,##0'})
            result_layout_format_right = workbook.add_format({'left': 1, 'top': 1, 'bottom': 1, 'right': 2,
                                                              'num_format': '#,##0'})
            for result_layout1 in range(9,35):
                for result_layout2 in range(0,8):
                    if result_layout2 % 4 == 0:
                        worksheet.write_blank(result_layout1, result_layout2, None, result_layout_format_left)
                    elif result_layout2 == 7:
                        worksheet.write_blank(result_layout1, result_layout2, None, result_layout_format_right)
                    else:
                        worksheet.write_blank(result_layout1, result_layout2, None, result_layout_format)

            # 상단 셀 내용 넣기
            merge_line2_title = workbook.add_format({'align':'center', 'bold':True, 'font_size':18, 'top':2, 'left':2,
                                                     'right':2, 'bottom':1})
            merge_line3_8_title = workbook.add_format({'align':'center', 'valign':'vcenter', 'bold':True,
                                                       'font_size':12, 'left': 2, 'right': 1, 'top': 1,'bottom': 1})
            merge_line3_7_value = workbook.add_format({'align':'center', 'valign':'vcenter', 'font_size':11, 'top':1,
                                                       'left':1, 'right':2, 'bottom':1})
            merge_line8_value = workbook.add_format({'align':'center', 'valign':'vcenter', 'font_size':12, 'border':1,
                                                     'num_format':'#,##0'})
            merge_line8_value_right = workbook.add_format({'align':'center','valign': 'vcenter', 'font_size': 12, 'top':1,
                                                       'left':1, 'right':2, 'bottom':1, 'num_format': '#,##0'})
            merge_line36_title_left = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,
                                                       'font_size': 12, 'left':2, 'right':1, 'top':2, 'bottom':1})
            merge_line36_title_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,
                                                           'font_size': 12, 'left': 1, 'right': 1, 'top': 2,
                                                             'bottom': 1})
            merge_line36_value_right = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 12,
                                                            'left':1, 'right':2, 'top': 2,'bottom': 1,
                                                            'num_format':'#,##0'})
            merge_line37_title_left = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,
                                                           'font_size': 12, 'left': 2, 'right': 1, 'top': 1,
                                                           'bottom': 2})
            merge_line37_value_right = workbook.add_format({'align':'center', 'valign':'vcenter', 'font_size':12,
                                                            'left':1, 'right':2, 'top':1, 'bottom':2,
                                                            'num_format':'#,##0'})
            merge_line38_43_title = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True,
                                                         'font_size': 12, 'left':2, 'right':2})
            merge_line44_title = workbook.add_format({'align': 'center', 'bold': True, 'font_size': 18, 'left':2,
                                                     'right':2, 'bottom':2})

            worksheet.merge_range('A2:H2', part_for_year+') 친환경 우리농산물 거래 명세서', merge_line2_title)
            worksheet.merge_range('A3:B6', '어린이집명', merge_line3_8_title)
            worksheet.merge_range('C3:D6', '', merge_line3_7_value)
            worksheet.merge_range('F3:H3', results[result_rows][5], merge_line3_7_value)
            worksheet.merge_range('F4:H4', results[result_rows][5], merge_line3_7_value)
            worksheet.merge_range('F5:H5', results[result_rows][6], merge_line3_7_value)
            worksheet.merge_range('F6:H6', results[result_rows][7], merge_line3_7_value)
            worksheet.merge_range('A7:B8', '납품업체명', merge_line3_8_title)
            worksheet.merge_range('C7:D8', '굿츠친환경영농조합법인', merge_line3_7_value)
            worksheet.merge_range('E7:E8', '배정액', merge_line3_8_title)
            worksheet.merge_range('F7:F8', '=' + results[result_rows][8], merge_line8_value)
            worksheet.merge_range('C36:D36', '=SUM(C10:C35)', merge_line36_value_right)
            worksheet.merge_range('G36:H36', '=SUM(F10:F35)', merge_line36_value_right)
            worksheet.merge_range('C37:D37', '=H7-C36', merge_line37_value_right)
            worksheet.merge_range('G37:H37', '=H8-G36', merge_line37_value_right)
            worksheet.merge_range('A38:H38', '', merge_line38_43_title)
            worksheet.merge_range('A40:H40', '', merge_line38_43_title)
            worksheet.merge_range('A43:H43', '', merge_line38_43_title)
            worksheet.merge_range('A37:B37', '남은 금액', merge_line37_title_left)
            worksheet.merge_range('E37:F37', '제주산 미달액', merge_line37_title_left)
            worksheet.merge_range('A39:H39', '어린이집 친환경우리농수산 급식지원에 대하여 상기와 같이 품목 및 수량을 정히 인수함.', merge_line38_43_title)
            worksheet.merge_range('A41:H41', now_date, merge_line38_43_title)    # 날짜: ex) 2107년  5월  31일
            worksheet.merge_range('A42:H42', '인수인: ' + results[result_rows][5] + ' 원장 (직인)', merge_line38_43_title)
            worksheet.merge_range('A44:H44','제주시어린이집연합회 귀하', merge_line44_title)

            worksheet.write('C3', results[result_rows][0], merge_line3_7_value)
            worksheet.write('E3', '대표자', merge_line3_8_title)
            worksheet.write('E4', '원장', merge_line3_8_title)
            worksheet.write('E5', '소 재 지', merge_line3_8_title)
            worksheet.write('E6', '연 락 처', merge_line3_8_title)
            worksheet.write('E7', '배정액', merge_line3_8_title)
            worksheet.write('G7', '쌀잡곡(70%)', merge_line3_8_title)
            worksheet.write('G8', '제주산(30%)', merge_line3_8_title)
            worksheet.write('H7', '=' + results[result_rows][9], merge_line8_value_right)
            worksheet.write('H8', '=' + results[result_rows][10], merge_line8_value_right)
            worksheet.write('A36', '쌀잡곡', merge_line36_title_left)
            worksheet.write('B36', '납품총액', merge_line36_title_center)
            worksheet.write('E36', '제주산', merge_line36_title_left)
            worksheet.write('F36', '납품총액', merge_line36_title_center)

            # 제목행 값 넣기
            merge_format_line10_title = workbook.add_format({'align':'center', 'bold':True, 'font_size':12, 'border':2})
            cell_title = ['납품일', '쌀잡곡', '금액', '누계', '제주산', '금액', '누계', '서명']
            for cell_write in range(len(cell_title)):   # 제목행(10행) 넣기
                worksheet.write(8, cell_write, cell_title[cell_write], merge_format_line10_title)

        # ########################################  내용 출력 ####################################################
        # 출력 셀 선 그리기
        cell_border_left = workbook.add_format({'left': 2, 'top': 1, 'bottom': 1, 'right': 1, 'num_format': '#,##0', 'font_size':10})
        cell_border_center = workbook.add_format({'border': 1, 'num_format': '#,##0', 'font_size':10})
        cell_border_right = workbook.add_format({'left': 1, 'top': 1, 'bottom': 1, 'right': 2, 'num_format': '#,##0', 'font_size':10})
        cell_format1 = cell_border_center
        cell_format3 = cell_border_right

        # 제주산 행 처리: 출력 인덱스 조정(col +=3, row -= 1, 누계셀 값)
        sheet_start_column = 0
        sheet_sum_cell = '=SUM($C$10:C' + str(sheet_start_row + 1) + ')'
        if results[result_rows][1] == '제주산30%':
            sheet_start_column = 3
            sheet_sum_cell = '=SUM($F$10:F' + str(sheet_start_row + 1) + ')'
            cell_format1 = cell_border_left
            cell_format3 = cell_border_center
            if results[result_rows - 1][2] == results[result_rows][2]:
                sheet_sum_cell = '=SUM($F$10:F' + str(sheet_start_row) + ')'
                sheet_start_row -= 1

        # 거래처명 처리: 거래처명 '()'와 '/' 이후 지우고 + '외'
        custom_rename = str(results[result_rows][3])
        if str(results[result_rows][3]).find('/'):
            custom_name = str(results[result_rows][3])
            custom_name_slashindex = str(results[result_rows][3]).find('/')
            custom_rename = custom_name[:custom_name_slashindex] + '외'
            if str(results[result_rows][3]).find('('):
                custom_rename = custom_rename[:custom_rename.find('(')] + ' 외'

        # '날짜', '거래처명', '단가', '누계' 셀 입력 처리
        worksheet.write(sheet_start_row, 0, results[result_rows][2], cell_border_left)
        worksheet.write(sheet_start_row, sheet_start_column + 1, custom_rename, cell_format1)
        worksheet.write(sheet_start_row, sheet_start_column + 2, results[result_rows][4], cell_border_center)
        worksheet.write(sheet_start_row, sheet_start_column + 3, sheet_sum_cell, cell_format3)

        sheet_start_row += 1

    workbook.close()