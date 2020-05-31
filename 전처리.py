# %%
#pandas 임포트
import pandas as pd

#데이터가 있는 엑셀 파일 불러오기
preprocess = pd.read_excel(r"C:\Users\신충현\Desktop\python\공정처\project\All_Task.xlsx",sheet_name = "원본")

#PERSON열의 빈 셀 채우기
preprocess["PERSON"] = preprocess["PERSON"].fillna("none")

#print(preprocess)

# %%
#데이터 START,STATUS 순서로 우선순위를 매겨 정렬
preprocess.sort_values(by = ['START','STATUS'],axis=0,ascending=True,inplace=True)
#print(preprocess)

# %%
#print(len(preprocess)) #행의 개수

# %%
#START_MONTH 새로운 열 생성 (월 표현)
for i in range(1,len(preprocess)):
    preprocess.loc[i,'START_MONTH'] = preprocess.loc[i,'START'][0:2]
#print(preprocess)

#엑셀 데이터에 가장 최근에 입력된 월 인식
present_month = preprocess.iloc[len(preprocess)-1,10]

#지난 달과 지지난달 변수 생성 (순환구조로 오류 없음)
months = []
for i in range(1,13):
    if len(str(i)) == 1:
        m = "0" + str(i)
        months.append(m)
    else:
        months.append(str(i))

last_month = months[months.index(present_month)-1]
last2_month = months[months.index(present_month)-2]

#print(last_month)
#print(last2_month)

# %%
#월별 DataFrame 생성 (현재,지난,지지난 달)
present_month_Data = preprocess.loc[preprocess['START_MONTH'] == present_month]
#print(present_month_Data)

last_month_Data = preprocess.loc[preprocess['START_MONTH'] == last_month]
#print(last_month_Data)

last2_month_Data = preprocess.loc[preprocess['START_MONTH'] == last2_month]
# %%
#월별 DataFrame sheet 생성
from openpyxl import load_workbook

path = r"C:\Users\신충현\Desktop\python\공정처\project\All_Task.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

last2_month_Data.to_excel(writer, sheet_name = "{}_month".format(last2_month),index = False)
last_month_Data.to_excel(writer, sheet_name = "{}_month".format(last_month),index = False)
present_month_Data.to_excel(writer, sheet_name = "{}_month".format(present_month), index = False)
writer.save()
writer.close()
# %%
