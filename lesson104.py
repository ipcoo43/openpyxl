import pandas as pd

# 데이터 가져오기
df_01 = pd.read_csv('https://goo.gl/VwsTBR', parse_dates=['년/월/일'], thousands=',', index_col='년/월/일') # 삼성전자
df_02 = pd.read_csv('https://goo.gl/2udsQq', parse_dates=['년/월/일'], thousands=',', index_col='년/월/일') # SK하이닉스
df_03 = pd.read_csv('https://goo.gl/TbBGhJ', parse_dates=['년/월/일'], thousands=',', index_col='년/월/일') # 셀트리온

print(df_01.head(1))
print(df_02.head(1))
print(df_03.head(1))

# 종가 취합하기
df_close = pd.DataFrame()
df_close['삼성전자'] = df_01['종가']
df_close['SK하이닉스'] = df_02['종가']
df_close['셀트리온'] = df_03['종가']

print(df_close.head(3))

# 엑셀로 저장
df_close.to_excel('./xlsx/종목별종가_104.xlsx', sheet_name='종목별종가')

# 별도의 시트에 종가 저장하기
writer = pd.ExcelWriter('./xlsx/종목별종가_시트별_104.xlsx')
df_01['종가'].to_excel(writer,'삼성전자')
df_02['종가'].to_excel(writer,'SK하이닉스')
df_03['종가'].to_excel(writer,'셀트리온')
writer.save()