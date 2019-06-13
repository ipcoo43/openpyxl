import pandas as pd

# 데이터 가져오기
df_01 = pd.read_csv('https://goo.gl/VwsTBR', parse_dates=['년/월/일'], thousands=',', index_col='년/월/일') # 삼성전자
df_02 = pd.read_csv('https://goo.gl/2udsQq', parse_dates=['년/월/일'], thousands=',', index_col='년/월/일') # SK하이닉스
df_03 = pd.read_csv('https://goo.gl/TbBGhJ', parse_dates=['년/월/일'], thousands=',', index_col='년/월/일') # 셀트리온

print(df_01.head(1))
print(df_02.head(1))
print(df_03.head(1))

# 특정 주기 데이터만 추출하기
# pd.date_range()을 사용하여 특정 기간의 DatetimeIndex를 만들고
# 매주 월요일, 매주 1일, 매월 말일, 매월 마지막 영업일

# 매주 월요일
idx = pd.date_range('2017-01-01', '2017-12-31', freq='W-MON')
df_mon = pd.DataFrame(index=idx)

df_mon[['종가']] = df_01[['종가']] # 삼성전자 종가(매주 월요일)
print(df_mon.head(5))

# 매월 1일
# 2017년, MS (매월 시작일)
idx = pd.date_range('2017-01-01', '2017-12-31', freq='MS')
df_mon = pd.DataFrame(index=idx)
print(df_mon.head(5))

# 매월 말일
# 2017년, M (매월 말일)
inx = pd.date_range('2017-01-01', '2017-12-31', freq='M')
df_mon = pd.DataFrame(index=inx)
print(df_mon.head(5))

# 삼성전자 종가(매월 말일)
df_mon['종가'] = df_01['종가'] # 삼성전자 종가 (매월 말일)
print(df_mon.head(10))

# 매월 마지막 영업일
# 단순하게 매월 마직막 날은 휴장일 일 수 있다 (여기서 영업일은 주말을 제외한 날을 가르킴)

# 2017년, BM (매월 마지막 영업일)
inx = pd.date_range('2017-01-01', '2017-12-31', freq='BM') # 삼성전자 종가 (매월 말일 주말제외)
df_mon = pd.DataFrame(index=inx)
print(df_mon.head(5))

df_mon['종가'] = df_01['종가'] # 삼성전자 종가 (매월 마지막 영업일)
print(df_mon.head(10))