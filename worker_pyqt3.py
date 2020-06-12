#%%
import openpyxl
import os

from wordcloud import WordCloud, STOPWORDS
     

import nltk
import re
from nltk.stem.snowball import SnowballStemmer
from nltk.corpus import stopwords

import matplotlib.pyplot as plt
from PyQt5.QtWidgets import *
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import urllib.request
import time

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QVBoxLayout
from PyQt5.QtGui import *
from PyQt5.QtGui import QPixmap


wb= openpyxl.load_workbook(r"C:\Users\신충현\Desktop\project\All_Task.xlsx")
sheet1 = wb.get_sheet_by_name("04_month")

from work_person import Person

final_who = Person()
#%%
class MyApp(QWidget,Person):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        
        self.lbl = QLabel(self)
        self.lbl.move(50, 100)
        
        self.lbl_book_img = QLabel(self)
        self.lbl_book_img.move(600,150)

        #self.lbl_book_tilte = QLabel(self)
        #self.lbl_book_tilte.move(1000,150)
        
        self.lbl_img = QLabel(self)
        self.lbl_img.move(50,150)

        self.lbl_word = QLabel(self)
        self.lbl_word.move(400,150)


    

        cb = QComboBox(self)
        for i in final_who.set_who:
            cb.addItem(i)
        cb.move(50, 50)

        cb.activated[str].connect(self.Book_img_Person)
        cb.activated[str].connect(self.onActivated)
        cb.activated[str].connect(self.imageload)
        cb.activated[str].connect(self.img_WordCloud)
        #cb.activated[str].connect(self.Book_Title)
        

        self.setWindowTitle('QComboBox')
        self.resize(500,500)
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def onActivated(self, text):
        self.lbl.setText(final_who.ratio_Cat(text))
        self.lbl.adjustSize()


    def imageload(self,text):
        pixmap = QPixmap(final_who.chart_pie(text))
        pixmap = pixmap.scaled(250,250)
        self.lbl_img.setPixmap(pixmap)
        self.lbl_img.adjustSize()

    def img_WordCloud(self,text):
        pixmap_word = QPixmap(final_who.WordCloud(text))
        pixmap_word = pixmap_word.scaledToHeight(150)
        self.lbl_word.setPixmap(pixmap_word)
        self.lbl_word.adjustSize()

    """def Book_Title(self, text):
        self.lbl_book_tilte.setText(final_who.Book(text))
        self.lbl_book_tilte.adjustSize()"""  

    def Book_img_Person(self,text):
        pixmap_book = QPixmap(final_who.Book_img(text))
        pixmap_book = pixmap_book.scaled(100,200)
        self.lbl_book_img.setPixmap(pixmap_book)
        self.lbl_book_img.adjustSize()      



app = QApplication(sys.argv)
ex = MyApp()
sys.exit(app.exec_())

