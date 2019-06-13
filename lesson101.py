import openpyxl

# 엑셀 파일 열기
wb = openpyxl.load_workbook('./xlsx/2017년 광고비 - 삼성전자.xlsx')
wb

# 모든 시트 이름들 얻기
# wb.get_sheet_names() # ['Sheet1'], 중단 경고

# 시트 이름으로 시트 얻기
sheet = wb['Sheet1']
sheet # <Worksheet "Sheet1">

# 활성화 시트 얻기
sheet = wb.active
sheet # <Worksheet "Sheet1">

# cell에 접근
sheet['A2'].value # '2017-01'
sheet['B1'].value # 'name'
sheet.cell(row=1,column=3).value # 'magazine' cell(row=n, column=m)

# 범위 접근
muti_cells = sheet['E2':'F14']
muti_cells  

muti_cells = sheet['E2':'F14']
for row in muti_cells:
    print(row[0].value, row[1].value)
    
# 모든 row 살펴보기
for row in sheet.rows:
    print([col.value for col in row])
