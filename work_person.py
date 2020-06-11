import openpyxl
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import *
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

wb= openpyxl.load_workbook(r"C:\Users\신충현\Desktop\project\All_Task.xlsx")
sheet1 = wb.get_sheet_by_name("04_month")

class Person:
    def __init__(self):
        p_who = []       
        for i in list(range(2,sheet1.max_row-1)):
            p_who.append(sheet1.cell(i,2).value)          
        str_who = ",".join(p_who)
        list_who = str_who.split(",")
        self.set_who = set(list_who)        

    def ratio_Cat(self,name): 
        self.name = name
        name_category = {}
        for i in list(range(2,sheet1.max_row-1)):
            if name in sheet1.cell(i,2).value:
                if  sheet1.cell(i,4).value not in name_category:
                     name_category[sheet1.cell(i,4).value] = float(sheet1.cell(i,5).value)
                else:
                    name_category[sheet1.cell(i,4).value] += float(sheet1.cell(i,5).value)
            else:
                continue
        
        #총 분야의 소요시간 합계
        sum_value = 0
        for i in name_category.values():
            sum_value += i
        
        #시간을 percentage 단위로 변환
        for key,element in name_category.items():
            name_category[key] = round((element/sum_value)*100,2)

        #시간에 대해 최댓값을 갖는 3가지 분야만 도출
        final_category = {}
        for i in range(3):
            value_list = list(name_category.values())
            key_list = list(name_category.keys())
            max_key = key_list[value_list.index(max(value_list))]
            final_category[max_key] = max(value_list)
            del name_category[max_key]


        #남은 부분 계산
        three_sum = 0
        for i in final_category.values():
            three_sum += i
        residual = 100 - three_sum

        #남은 부분 추가
        final_category["residual"] = round(residual,2)

        #self.cat_max = final_category

        #입력된 사람의 분야별 소요시간(최대 3)
        #return ",".join("{}:{}".format(key,value) for key, value in final_category.items())
        return ",".join("{}:{}".format(key,value) for key, value in final_category.items())

    def chart_pie(self,name):
        self.name = name
        name_category = {}
        for i in list(range(2,sheet1.max_row-1)):
            if name in sheet1.cell(i,2).value:
                if  sheet1.cell(i,4).value not in name_category:
                     name_category[sheet1.cell(i,4).value] = float(sheet1.cell(i,5).value)
                else:
                    name_category[sheet1.cell(i,4).value] += float(sheet1.cell(i,5).value)
            else:
                continue
        
        #총 분야의 소요시간 합계
        sum_value = 0
        for i in name_category.values():
            sum_value += i
        
        #시간을 percentage 단위로 변환
        for key,element in name_category.items():
            name_category[key] = round((element/sum_value)*100,2)

        #시간에 대해 최댓값을 갖는 3가지 분야만 도출
        final_category = {}
        for i in range(3):
            value_list = list(name_category.values())
            key_list = list(name_category.keys())
            max_key = key_list[value_list.index(max(value_list))]
            final_category[max_key] = max(value_list)
            del name_category[max_key]


        #남은 부분 계산
        three_sum = 0
        for i in final_category.values():
            three_sum += i
        residual = 100 - three_sum

        #남은 부분 추가
        final_category["residual"] = round(residual,2)
        labels = tuple(final_category.keys())
        sizes = list(final_category.values())
        explode = (0, 0.1, 0, 0)

        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
                shadow=True, startangle=90)
        ax1.axis('equal') 
        
        plt.savefig('{}_04.png'.format(self.name),dpi=300)
        a = '{}_04.png'.format(self.name)
        return a        

