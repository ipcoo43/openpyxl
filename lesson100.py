'''
파이썬 + 엑셀 자동화 도구

[ xlwings ]
 http://xlwings.org/
 문서 https://goo.gl/WGsFVc (PDF)
 엑셀 자동화 스크립팅, 엑셀의 VBA를 Python으로 대체
 Windows, OSX 지원, 설치된 엑셀 필요

[ openpyxl ]
 https://openpyxl.readthedocs.io
 2010 (.xlsx) 포맷 읽고 쓰기 지원
 설치된 엑셀 필요 없음

[ xlrd & xlwt ]
 http://xlrd.readthedocs.io , http://xlwt.readthedocs.io
 엑셀 파일 읽기/쓰기 (가볍고 빠르다, 기능이 적다)
 설치된 엑셀 필요 없음
 
[ xlsxwriter ]
 https://xlsxwriter.readthedocs.org/
 2010 (.xlsx) 포맷 쓰기, 차트 가능
 설치된 엑셀 필요 없음

[ PyWin32 ]
 윈도우의 COM 기술(정확히는 OLE Automation) 사용
 엑셀 뿐만 아니라 다양한 오피스 제품과 상호작용이 가능
 Windows 지원, 설치된 엑셀 필요
 
파이썬에서 엑셀을 다루는 라이브러리는 꽤 다양한데, 크게 엑셀 파일을 읽고 쓰는 라이브러리(xlrd & xlwt, openpyxl, xlsxwriter)와 엑셀과 상호작용 하는 부류(xlwings, PyWin32)로 나눌 수 있다.

[ 파이썬+엑셀 추천 ]
 엑셀 파일 읽고 쓰기 (엑셀 설치 없이): openpyxl 추천
 엑셀 매크로 자동화 (설치된 엑셀): xlwings 추천
 엑셀 파일을 생성(서버) : xlsxwriter
'''