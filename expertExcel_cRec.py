import xlsxwriter
import datetime


def expert_excel_cRec(results, *args):
    start_month = args[0][4:6]
    end_month = args[1][4:6]
    part_of_year = ''
    if start_month in ('03', '04', '05'):
        part_of_year = '1분기'
        start_month = '03'
        end_month = '05'
    elif start_month in ('06', '07', '08'):
        part_of_year = '2분기'
        start_month = '06'
        end_month = '08'
    elif start_month in ('09', '10', '11'):
        part_of_year = '3분기'
        start_month = '09'
        end_month = '11'
    elif start_month in ('12', '01', '02'):
        part_of_year = '4분기'
        start_month = '12'
        end_month = '01'

    now = datetime.datetime.now()
    # 출력할 엑셀 파일 이름
    file_name = '어린이집_완료인수증('+args[0][:4]+'년 '+part_of_year+')_'+ now.strftime('%Y%m%d_%H%M%S') + '.xlsx'
    now_date = str(now.year) + '년  ' + str(now.month) + '월  ' + str(now.day) + '일'  # 인수 날짜 세팅: 오늘 날짜

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.worksheets()

    for result_rows in range(len(results)):
        if results[result_rows - 1][0] != results[result_rows][0]:  # 이전 이름과 다르면 새로 생성
            result_name = results[result_rows][0]
            worksheet_name = result_name[:result_name.find('어린이집')] # 뒤에 '어린이집' 은 빼고 시트이름
            worksheet = workbook.add_worksheet(worksheet_name)
            worksheet.set_paper(9)  # 용지 A4 세팅
            worksheet.print_area('A2:F18') # 출력 범위 지정
            # worksheet.set_print_scale(97)  # 출력 비율
            # worksheet.set_margins(left=0.5, right=0.5)
            worksheet.center_horizontally()

            # ########################################### 시트 상단 ##################################################
            # 전체 열 크기 지정
            column_width = [19, 17, 11, 11, 9, 9]
            for cell_index_col in range(len(column_width)):
                worksheet.set_column(cell_index_col, cell_index_col, column_width[cell_index_col])

            # 전체 행 크기 지정
            worksheet.set_default_row(40)

            # 상단 셀 내용 넣기
            line2_format = workbook.add_format({'align':'center', 'valign':'vcenter', 'bold':True, 'font_size':16, 'border':1})
            mergeA3A4_format = workbook.add_format({'align':'center', 'valign':'vcenter', 'bold':True, 'font_size': 15, 'border':1})
            mergeA3A4_format.set_text_wrap()
            line11_12_format = workbook.add_format({'valign': 'vcenter', 'font_size': 15, 'left':1, 'right':1})
            line14_format = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'font_size': 15, 'left':1, 'right':1})
            line18_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold':True, 'font_size': 17, 'left': 1, 'right': 1, 'bottom':1})

            worksheet.merge_range('A2:F2',part_of_year + ': '+ start_month + '월' + ' ~ ' + end_month + '월) 친환경우리농산물 급식지원 물품 완료 인수증', line2_format)
            worksheet.merge_range('A3:A4', '어린이집 명칭', mergeA3A4_format)
            worksheet.write('D3', '대표자', mergeA3A4_format)
            worksheet.write('D4', '원장', mergeA3A4_format)
            worksheet.write('A5', '소 재 지', mergeA3A4_format)
            worksheet.write('A6', '연 락 처', mergeA3A4_format)
            worksheet.write('A7', '급식원아 인원', mergeA3A4_format)
            worksheet.merge_range('C7:D7', '품목', mergeA3A4_format)
            worksheet.merge_range('E7:F7', '친환경쌀 및\n우리농산물', mergeA3A4_format)
            worksheet.merge_range('A8:A9', part_of_year + ' 배정액', mergeA3A4_format)
            worksheet.merge_range('C8:D8', '쌀납품액(70%)', mergeA3A4_format)
            worksheet.merge_range('C9:D9', '잡곡 및 농산물\n납품액(30%)', mergeA3A4_format)
            worksheet.merge_range('A10:D10', part_of_year + ' 납품액 합계', mergeA3A4_format)
            worksheet.merge_range('A11:F11', ' 어린이집 친환경우리농산물 급식지원에 대하여 상기와 같이', line11_12_format)
            worksheet.merge_range('A12:F12', ' 품목 및 수량을 정히 인수합니다.', line11_12_format)
            worksheet.merge_range('A13:F13', '', line11_12_format)  # 공란
            worksheet.merge_range('A14:F14', now_date + ' ', line14_format)  # 날짜
            worksheet.merge_range('A15:F15', '', line14_format)  # 공란
            worksheet.merge_range('A17:F17', '', line14_format)  # 공란
            worksheet.merge_range('A18:F18', '제주특별자치도지사 귀하', line18_format)

        # ########################################  내용 출력 ####################################################
        mergeB3_C3C4_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 13, 'border': 1, 'num_format':'#,##0'})
        mergeB5_B6_format = workbook.add_format({'valign': 'vcenter', 'font_size': 13, 'border': 1})
        mergeB8_B9_format = workbook.add_format({'align':'right', 'valign': 'vcenter', 'font_size': 13, 'border': 1, 'num_format':'#,##0'})
        line16_format = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'font_size': 15, 'left': 1, 'right': 1})

        results4 = str(format(int(float(results[result_rows][4])), ','))
        results5 = str(format(int(float(results[result_rows][5])), ','))
        results6 = str(format(int(float(results[result_rows][6])), ','))
        results7 = str(format(int(float(results[result_rows][7])), ','))
        results8 = str(format(int(results[result_rows][8]), ','))

        worksheet.merge_range('B3:C4', results[result_rows][0], mergeB3_C3C4_format)  # 어린이집 명칭 value
        worksheet.merge_range('E3:F3', results[result_rows][1], mergeB3_C3C4_format)  # 대표자 value
        worksheet.merge_range('E4:F4', results[result_rows][1], mergeB3_C3C4_format)  # 원장 value
        worksheet.merge_range('B5:F5', results[result_rows][2], mergeB5_B6_format)  # 소재지 value
        worksheet.merge_range('B6:F6', results[result_rows][3], mergeB5_B6_format)  # 연락처 value
        worksheet.write('B7', results4 + '명', mergeB3_C3C4_format)   # 급식인원 원아 value
        worksheet.merge_range('B8:B9', results5 + '원', mergeB8_B9_format)  # 분기 매정액 value
        worksheet.merge_range('E8:F8', results6 + '원', mergeB8_B9_format)  # 쌀납품액 value
        worksheet.merge_range('E9:F9', results7 + '원', mergeB8_B9_format)  # 잡곡 value
        worksheet.merge_range('E10:F10', results5 + '원', mergeB8_B9_format)    # 분기 납품액 합계 value
        worksheet.merge_range('A16:F16', '인수인: ' + results[result_rows][0] + ' 원장 (직인) ', line16_format)

    workbook.close()