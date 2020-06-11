#%%
import sys
import openpyxl
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
#import urllib.request
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
#from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QVBoxLayout
from PyQt5.QtGui import *
#from PyQt5.QtCore import *
#from PyQt5 import uic

wb= openpyxl.load_workbook(r"C:\Users\신충현\Desktop\project\All_Task.xlsx")
sheet1 = wb.get_sheet_by_name("04_month")

from work_person import Person


#%%
class MyApp(QWidget,Person):

    def __init__(self, name):
        super.__init__()
        Person.__init__(self, name)
        self.initUI()
        
        

    def initUI(self):
        
        self.lbl = QLabel(self)
        self.lbl.move(50, 150)
        
        self.lbl_img = QLabel(self)
        self.lbl_img.resize(50,50)
        self.lbl_img.move(50,200)
        
        #self.fig = plt.Figure()
        #self.canvas = FigureCanvas(self.fig)
        
        #layout = QVBoxLayout()
        #layout.addWidget(self.canvas)

        cb = QComboBox(self)
        for i in self.final_who.set_who:
            cb.addItem(i)
        cb.move(50, 50)

        #self.layout = layout

        cb.activated[str].connect(self.onActivated)
        cb.activated[str].connect(self.imageload)


        self.setWindowTitle('QComboBox')
        self.resize(500,500)
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def onActivated(self, text):
        
        self.lbl.setText(self.final_who.ratio_cat())
        self.lbl.adjustSize()


    
    #def onComboBoxChanged(self, text):
        #final_who.chart_pie(text)

    def imageload(self):
           
        pixmap = QPixmap(self.final_who.chart_pie())
        self.lbl_img.setPixmap(pixmap)
        self.lbl_img.adjustSize()


app = QApplication(sys.argv)
ex = MyApp()
sys.exit(app.exec_())

