# %%
#pandas 임포트
import pandas as pd
import numpy as np

#데이터가 있는 엑셀 파일 불러오기
preprocess = pd.read_excel(r"C:\Users\신충현\Desktop\project\All_Task.xlsx",sheet_name = "원본")

#PERSON열의 빈 셀 채우기
preprocess["PERSON"] = preprocess["PERSON"].fillna("none")

#WORK HOURS열 빈 셀 채우기(CATEGORY의 평균 WORK HOURS 대입)
pre_hour = preprocess.loc[preprocess['STATUS'] == 'Done']
pre_hour.dropna(subset = ['WORKED HOURS'], inplace = True)
kind_cat = set(preprocess['CATEGORY'])

kind_cat_dic = {}
for i in kind_cat:
    cat = pre_hour.loc[pre_hour['CATEGORY'] == i]
    cat_count = len(cat)
    cat_hour = 0
    for p in range(1,cat_count-1):
        cat_hour += float(cat.iloc[p,4])
    cat_average = cat_hour/cat_count
    kind_cat_dic[i] = round(cat_average,2)

preprocess_one = preprocess.loc[:,['CATEGORY','WORKED HOURS']]
fill_cat = lambda d: d.fillna(kind_cat_dic[d.name])
preprocess_perpect = preprocess_one.groupby('CATEGORY').apply(fill_cat)

#CATEGORY별 평균 시간 딕셔너리 적용
preprocess['WORKED HOURS'] = preprocess_perpect['WORKED HOURS']


# %%
#데이터 START,STATUS 순서로 우선순위를 매겨 정렬
preprocess.sort_values(by = ['START','STATUS'],axis=0,ascending=True,inplace=True)


# %%
#START_MONTH 새로운 열 생성 (월 표현)
for i in range(1,len(preprocess)):
    preprocess.loc[i,'START_MONTH'] = preprocess.loc[i,'START'][0:2]

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

# %%
#월별 DataFrame 생성 (현재,지난,지지난 달)
present_month_Data = preprocess.loc[preprocess['START_MONTH'] == present_month]

last_month_Data = preprocess.loc[preprocess['START_MONTH'] == last_month]

last2_month_Data = preprocess.loc[preprocess['START_MONTH'] == last2_month]

#월별 DataFrame sheet 생성
from openpyxl import load_workbook

path = r"C:\Users\신충현\Desktop\project\All_Task.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

preprocess.to_excel(writer, sheet_name = "all_data", index = False)
last2_month_Data.to_excel(writer, sheet_name = "{}_month".format(last2_month),index = False)
last_month_Data.to_excel(writer, sheet_name = "{}_month".format(last_month),index = False)
present_month_Data.to_excel(writer, sheet_name = "{}_month".format(present_month), index = False)

writer.save()
writer.close()