from PyQt5 import QtWidgets
from mfg import Ui_MainWindow  # импорт нашего сгенерированного файла
import sys

# -*- coding: utf-8 -*-
#|Основная надпись

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import time
import MiscellaneousHelpers as MH

      #  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
MH.iKompasObject  = kompas_object

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))
MH.iApplication  = application


Documents = application.Documents
#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()
 
 
class mywindow(QtWidgets.QMainWindow):
   def __init__(self):
      super(mywindow, self).__init__()
      self.ui = Ui_MainWindow()
      self.ui.setupUi(self)
      self.ui.pushButton_3.clicked.connect(self.sprav)


 
   def sprav(none):
      iStamp = iDocument2D.GetStamp()
      iStamp.ksOpenStamp()
      iStamp.ksColumnNumber(24)

      iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
      iTextLineParam.Init()
      iTextLineParam.style = 32768
      iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
      iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
      iTextItemParam.Init()
      iTextItemParam.iSNumb = 0
      iTextItemParam.s = self.lineEdit_3.text()
      iTextItemParam.type = 0
      iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
      iTextItemFont.Init()
      iTextItemFont.bitVector = 4096
      iTextItemFont.color = 0
      iTextItemFont.fontName = "GOST type A"
      iTextItemFont.height = 3.5
      iTextItemFont.ksu = 1
      iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
      iTextLineParam.SetTextItemArr(iTextItemArray)

      iStamp.ksTextLine(iTextLineParam)
      
app = QtWidgets.QApplication([])
application = mywindow()
application.show()             
 
sys.exit(app.exec())