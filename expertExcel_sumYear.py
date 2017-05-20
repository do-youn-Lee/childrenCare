import xlsxwriter
import datetime


def expert_excel_sumYear(results, *args):
    # args: part_dist, self.comboBox_3.currentText(), from_date[:4]
    # 출력할 엑셀 파일 이름
    now = datetime.datetime.now()
    now_date = str(now.year) + '년 ' + str(now.month) + '월' + str(now.day) + '일' # 기준일 날짜 세팅: 오늘 날짜 + 말일
    file_name = '어린이집_집행현황('+args[2]+'년)_'+args[0]+'_'+now.strftime('%Y%m%d_%H%M%S') + '.xlsx'

    workbook = xlsxwriter.Workbook(file_name)

    worksheet_name = '집행현황(' + args[2] + '년)_' + args[0]  # 뒤에 '어린이집' 은 빼고 시트이름
    worksheet = workbook.add_worksheet(worksheet_name)

    # ########################################### 시트 상단 ##################################################
    # 전체 열 크기 지정
    column_width = [14, 10, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 10, 8]
    for cell_index in range(len(column_width)):
        worksheet.set_column(cell_index, cell_index, column_width[cell_index])

    line1_format = workbook.add_format({'bold':True, 'font_size':15, 'align':'center'})
    line3_format_front = workbook.add_format({'bold':True, 'font_size':15, 'align':'center'})
    line3_format_tail = workbook.add_format({'font_size': 11, 'align': 'right'})
    line_title_format = workbook.add_format({'font_size':11, 'align':'center', 'valign':'vcenter', 'bold':True, 'border':2, 'bg_color':'#92CDDB'})

    worksheet.merge_range('A1:O1',args[2] + '년 월별 친환경우리농산물 급식지원금액 집행현황('+ args[0] + ')', line1_format)
    worksheet.merge_range('A2:O2', '')
    worksheet.merge_range('A3:D3', '업체명: 굿츠친환경영농조합법인', line3_format_front)
    worksheet.merge_range('M3:O3', '기준일: ' + now_date + '까지 ', line3_format_tail)
    worksheet.merge_range('A4:A5', '어린이집', line_title_format)
    worksheet.merge_range('B4:B5', '지원금액', line_title_format)
    worksheet.merge_range('C4:M4', '집 행 액', line_title_format)
    worksheet.write('C5', '3월', line_title_format)
    worksheet.write('D5', '4월', line_title_format)
    worksheet.write('E5', '5월', line_title_format)
    worksheet.write('F5', '6월', line_title_format)
    worksheet.write('G5', '7월', line_title_format)
    worksheet.write('H5', '8월', line_title_format)
    worksheet.write('I5', '9월', line_title_format)
    worksheet.write('J5', '10월', line_title_format)
    worksheet.write('K5', '11월', line_title_format)
    worksheet.write('L5', '12월', line_title_format)
    worksheet.write('M5', '계', line_title_format)
    worksheet.merge_range('N4:N5', '잔액', line_title_format)
    worksheet.merge_range('O4:O5', '비고', line_title_format)

    sheet_start_row = 5     # 결과 출력 시작 행 설정
    # ########################################  내용 출력 ####################################################
    for result_rows in range(len(results)):
        # 출력 셀 선 그리기용 format
        cell_border_center = workbook.add_format({'border':1, 'num_format': '#,##0', 'font_size':10})
        cell_border6 = cell_border_center

        part_dist_sum = 0.0
        if args[0] == '총괄':
            part_dist_sum = float(results[result_rows][1])
        elif args[0] == '국내산70%':
            part_dist_sum = float(results[result_rows][2])
        elif args[0] == '제주산30%':
            part_dist_sum = float(results[result_rows][3])

        worksheet.write(sheet_start_row, 0, results[result_rows][0], cell_border_center)
        worksheet.write(sheet_start_row, 1, part_dist_sum, cell_border_center)
        worksheet.write(sheet_start_row, 2, results[result_rows][4], cell_border_center)
        worksheet.write(sheet_start_row, 3, results[result_rows][5], cell_border_center)
        worksheet.write(sheet_start_row, 4, results[result_rows][6], cell_border_center)
        worksheet.write(sheet_start_row, 5, results[result_rows][7], cell_border_center)
        worksheet.write(sheet_start_row, 6, results[result_rows][8], cell_border_center)
        worksheet.write(sheet_start_row, 7, results[result_rows][9], cell_border_center)
        worksheet.write(sheet_start_row, 8, results[result_rows][10], cell_border_center)
        worksheet.write(sheet_start_row, 9, results[result_rows][11], cell_border_center)
        worksheet.write(sheet_start_row, 10, results[result_rows][12], cell_border_center)
        worksheet.write(sheet_start_row, 11, results[result_rows][13], cell_border_center)
        worksheet.write(sheet_start_row, 12, results[result_rows][14], cell_border_center)
        change = part_dist_sum - results[result_rows][14]
        if change < 0.0:
            cell_border6 = workbook.add_format({'font_color':'red', 'border':1, 'num_format': '#,##0'})
        worksheet.write(sheet_start_row, 13, change, cell_border6)
        worksheet.write(sheet_start_row, 14, '', cell_border_center)

        sheet_start_row += 1

    workbook.close()