import openpyxl

# 엑셀 파일 열기
wb = openpyxl.load_workbook('test.xlsx')

# 시트 선택
sheet_name = '사원명부'
sheet = wb[sheet_name]

# VLOOKUP 수식 적용
lookup_value = "KR-004"  # VLOOKUP 함수의 검색 대상 값
table_array = 'A2:G10'  # VLOOKUP 함수의 검색 범위
col_index_num = 2  # VLOOKUP 함수의 반환할 열 번호

# VLOOKUP 함수 적용 예시
sheet["J1"] = f'=VLOOKUP("{lookup_value}", {table_array}, 2, False)'
sheet["J2"] = f'=VLOOKUP("{lookup_value}", {table_array}, 5, False)'

# 엑셀 파일 저장
wb.save('result.xlsx')

# 엑셀 파일 닫기
wb.close()