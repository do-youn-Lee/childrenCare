import xlsxwriter
import datetime


def expert_excel_rec1(results):
    now = datetime.datetime.now()
    file_name = '어린이집_거래명세서1_'+ now.strftime('%Y%m%d_%H%M%S') + '.xlsx'

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.worksheets()

    # #################################################################
    sheet_rows = 4
    for result_rows in range(len(results)):
        if results[result_rows - 1][0] != results[result_rows][0]:
            result_name = results[result_rows][0]
            worksheet_name = result_name[:result_name.find('어린이집')]
            worksheet = workbook.add_worksheet(worksheet_name)
            sheet_rows = 4
            # ############################ 시트 상단 ###################################
            column_width = [12.13, 14.63, 5.88, 5.88, 10.63, 10.63, 16.75, 5.88,
                            5.88, 10.63, 10.63, 18.75, 5.88, 5.88, 12.13, 12.13, 18.88]
            for cell_index in range(len(column_width)):
                worksheet.set_column(cell_index, cell_index, column_width[cell_index])

            merge_format_A1 = workbook.add_format({'align':'center','bold': True,
                                             'font_size':20, 'bg_color':'#92CDDB', 'border':6})
            merge_format_line2_title = workbook.add_format({'align': 'right', 'bold': True,
                                                      'font_size': 12, 'bg_color':'#DBEEF4', 'border':6})
            merge_format_line2_value = workbook.add_format({'align': 'left', 'bold': True,
                                                      'font_size': 12, 'font_color':'green', 'bg_color': '#DBEEF4',
                                                      'border':6})
            merge_format_line2_value.set_num_format('#,##0')
            merge_format_line3_title = workbook.add_format({'align': 'right', 'bold': True,
                                                      'font_size': 12, 'bg_color': '#B7DDE8', 'border':6})
            merge_format_line3_value = workbook.add_format({'align': 'left', 'bold': True,
                                                      'font_size': 12, 'bg_color': '#B7DDE8', 'border':6})
            merge_format_line3_value.set_num_format('#,##0')
            merge_format_line4_title = workbook.add_format({'align': 'center', 'bold': True,
                                                      'font_size': 12, 'font_color':'white', 'bg_color':'#31869B',
                                                      'border':6})

            worksheet.merge_range('A1:Q1',results[result_rows][0],merge_format_A1)
            worksheet.write('A2','분기지원총액',merge_format_line2_title)
            worksheet.merge_range('B2:C2', '='+results[result_rows][8], merge_format_line2_value)
            worksheet.merge_range('D2:E2','쌀지원액',merge_format_line2_title)
            worksheet.write('F2','='+results[result_rows][9],merge_format_line2_value)
            worksheet.write('G2','제주산',merge_format_line2_title)
            worksheet.merge_range('H2:I2','='+results[result_rows][10],merge_format_line2_value)
            worksheet.write('J2','',merge_format_line2_value)
            worksheet.write('K2','남은지원액(쌀)',merge_format_line2_title)
            worksheet.merge_range('L2:M2','=F2-K3',merge_format_line2_value)
            worksheet.merge_range('N2:O2','남은지원액(제주산)',merge_format_line2_title)
            worksheet.merge_range('P2:Q2','=H2-F3',merge_format_line2_value)
            worksheet.write('A3','총판매액',merge_format_line3_title)
            worksheet.merge_range('B3:C3','=F3+K3+P3',merge_format_line3_value)
            worksheet.merge_range('D3:E3','제주산 납품 총액',merge_format_line3_title)
            worksheet.merge_range('F3:H3','=SUM(K5:K1048576)',merge_format_line3_value)
            worksheet.merge_range('I3:J3','백미 납품 총액',merge_format_line3_title)
            worksheet.merge_range('K3:M3','=SUM(F5:F1048576)',merge_format_line3_value)
            worksheet.merge_range('N3:O3','미지원 납품 총액',merge_format_line3_title)
            worksheet.merge_range('P3:Q3','=SUM(P5:P1048576)',merge_format_line3_value)

            cell_title = ['날짜', '백미', '규격', '수량', '단가', '금액',
                          '제주산농산물', '규격', '수량', '단가', '금액',
                          '지원 제외 품목', '규격', '수량', '단가', '금액', '비고']
            for cell_write in range(len(cell_title)):
                worksheet.write(3, cell_write, cell_title[cell_write], merge_format_line4_title)

        # ################################  내용 출력 #####################################
        for line_col in range(17):
            result_line = workbook.add_format({'border': 1})
            worksheet.write_blank(sheet_rows, line_col, None, result_line)

        sheet_columns = 0
        if results[result_rows][7] == '제주산30%':
            sheet_columns += 5
        if results[result_rows][7] == '자가결제':
            sheet_columns += 10

        cell_comma = workbook.add_format({'border': 1, 'font_size':10, 'num_format':'##,#0'})

        worksheet.write(sheet_rows, 0, results[result_rows][1], cell_comma)
        worksheet.write(sheet_rows, sheet_columns + 1, results[result_rows][2], cell_comma)
        worksheet.write(sheet_rows, sheet_columns + 2, results[result_rows][3], cell_comma)
        worksheet.write(sheet_rows, sheet_columns + 3, results[result_rows][4], cell_comma)
        worksheet.write(sheet_rows, sheet_columns + 4, results[result_rows][5], cell_comma)
        worksheet.write(sheet_rows, sheet_columns + 5, results[result_rows][6], cell_comma)

        sheet_rows += 1
    workbook.close()