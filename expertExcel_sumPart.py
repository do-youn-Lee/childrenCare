import xlsxwriter
import datetime


def expert_excel_sumPart(results, *args):
    # 출력할 엑셀 파일 이름
    now = datetime.datetime.now()
    now_date = str(now.year) + '년 ' + str(now.month) + '월'  # 기준일 날짜 세팅: 오늘 날짜 + 말일
    file_name = '어린이집_집행현황('+args[1]+')_'+args[0]+'_'+now.strftime('%Y%m%d_%H%M%S') + '.xlsx'

    # 분기 처리: 해당 분기 첫 월을 넘김
    month1 = ''
    month2 = ''
    month3 = ''
    if args[1] == '1분기':
        month1 = '3월'
        month2 = '4월'
        month3 = '5월'
    elif args[1] == '2분기':
        month1 = '6월'
        month2 = '7월'
        month3 = '8월'
    elif args[1] == '3분기':
        month1 = '9월'
        month2 = '10월'
        month3 = '11월'
    elif args[1] == '4분기':
        month1 = '12월'
        month2 = '1월'
        month3 = '2월'

    workbook = xlsxwriter.Workbook(file_name)

    worksheet_name = '집행현황(' + args[1] + ')_' + args[0]  # 뒤에 '어린이집' 은 빼고 시트이름
    worksheet = workbook.add_worksheet(worksheet_name)

    # ########################################### 시트 상단 ##################################################
    # 전체 열 크기 지정
    column_width = [19, 13, 12, 12, 12, 13, 13, 12]
    for cell_index in range(len(column_width)):
        worksheet.set_column(cell_index, cell_index, column_width[cell_index])

    line1_format = workbook.add_format({'bold':True, 'font_size':15, 'align':'center'})
    line3_format_front = workbook.add_format({'bold':True, 'font_size':15, 'align':'center'})
    line3_format_tail = workbook.add_format({'font_size': 11, 'align': 'right'})
    line_title_format = workbook.add_format({'font_size':11, 'align':'center', 'valign':'vcenter', 'bold':True, 'border':2, 'bg_color':'#92CDDB'})

    worksheet.merge_range('A1:H1',args[2] + '년 '+ args[1] + ' 친환경우리농산물 급식지원금액 집행현황('+ args[0] + ')', line1_format)
    worksheet.merge_range('A2:H2', '')
    worksheet.merge_range('A3:C3', '업체명: 굿츠친환경영농조합법인', line3_format_front)
    worksheet.merge_range('F3:H3', '기준일: ' + now_date + '말 일 ', line3_format_tail)
    worksheet.merge_range('A4:A5', '어린이집', line_title_format)
    worksheet.merge_range('B4:B5', args[1] + ' 배정액', line_title_format)
    worksheet.merge_range('C4:F4', '집 행 액', line_title_format)
    worksheet.write('C5', month1, line_title_format)
    worksheet.write('D5', month2, line_title_format)
    worksheet.write('E5', month3, line_title_format)
    worksheet.write('F5', '계', line_title_format)
    worksheet.merge_range('G4:G5', '잔액', line_title_format)
    worksheet.merge_range('H4:H5', '비고', line_title_format)

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
        change = part_dist_sum - results[result_rows][7]
        if change < 0.0:
            cell_border6 = workbook.add_format({'font_color':'red', 'border':1, 'num_format': '#,##0', 'font_size':10})
        worksheet.write(sheet_start_row, 6, change, cell_border6)
        worksheet.write(sheet_start_row, 7, '', cell_border_center)

        sheet_start_row += 1

    workbook.close()