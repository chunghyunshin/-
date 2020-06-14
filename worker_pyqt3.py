#%%
import openpyxl
import os

from wordcloud import WordCloud, STOPWORDS
     

import nltk
import re
from nltk.stem.snowball import SnowballStemmer
from nltk.corpus import stopwords

import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import urllib.request
import time

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QVBoxLayout
from PyQt5.QtGui import QPixmap

from work_person import Person, Kind_person
who = Kind_person()

#%%
class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        
        self.lbl = QLabel(self)
        self.lbl.move(50, 100)
        
        self.lbl_book_img = QLabel(self)
        self.lbl_book_img.move(700,700)

        self.lbl_book_tilte = QLabel(self)
        self.lbl_book_tilte.move(50,700)
        
        self.lbl_img = QLabel(self)
        self.lbl_img.move(50,150)

        self.lbl_word = QLabel(self)
        self.lbl_word.move(700,150)


        cb = QComboBox(self)
        for i in who.set_who:
            cb.addItem(i)
        cb.move(50, 50)

        cb.activated[str].connect(self.Book_img_Person)
        cb.activated[str].connect(self.onActivated)
        cb.activated[str].connect(self.imageload)
        cb.activated[str].connect(self.img_WordCloud)
        cb.activated[str].connect(self.Book_Title)
        
        self.setWindowTitle('도서추천 자동화시스템')
        self.resize(500,500)
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def onActivated(self, text):
        self.lbl.setText(Person(text).ratio_Cat())
        self.lbl.adjustSize()


    def imageload(self,text):
        pixmap = QPixmap(Person(text).chart_pie())
        pixmap = pixmap.scaled(500,500)
        self.lbl_img.setPixmap(pixmap)
        self.lbl_img.adjustSize()

    def img_WordCloud(self,text):
        pixmap_word = QPixmap(Person(text).WordCloud())
        pixmap_word = pixmap_word.scaledToHeight(350)
        self.lbl_word.setPixmap(pixmap_word)
        self.lbl_word.adjustSize()

    def Book_Title(self, text):
        self.lbl_book_tilte.setText(Person(text).WordSearch())
        self.lbl_book_tilte.adjustSize()

    def Book_img_Person(self, text):
        url = Person(text).Book_img()
        image = urllib.request.urlopen(url).read()
        
        #self._timer_painter = QTimer(self)
        #self._timer_painter.start(2000)
        #self._timer_painter.timeout.connect(self.show_currencies)       
        
        self.pixmap_book = QPixmap()
        self.pixmap_book.loadFromData(image)
        self.pixmap_book = self.pixmap_book.scaled(300,400)
        self.lbl_book_img.setPixmap(self.pixmap_book)
        self.lbl_book_img.adjustSize()

    """def Book_img_Person(self,text):
        pixmap_book = QPixmap(final_who.Book_img(text)) 
        while pipmax_book.isNull():
            bb = final_who.Book_img(text)
            pixmap_book = QPixmap(bb)
        loop = QEvetLoop()
        QTimer.singleShot(10*1000, loop.quit)
        loop.exec_()
        #QtCore.QTimer.singleShot()                
        pixmap_book = pixmap_book.scaled(100,200)
        self.lbl_book_img.setPixmap(pixmap_book)
        self.lbl_book_img.adjustSize()"""      



app = QApplication(sys.argv)
ex = MyApp()
sys.exit(app.exec_())

# %%
