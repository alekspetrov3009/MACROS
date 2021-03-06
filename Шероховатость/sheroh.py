# -*- coding: utf-8 -*-
#|Шерох

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH
from tkinter import *

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

iDrawingDocument = kompas_document._oleobj_.QueryInterface(kompas_api7_module.IDrawingDocument.CLSID, pythoncom.IID_IDispatch)
iDrawingDocument = kompas_api7_module.IDrawingDocument(iDrawingDocument)
iSpecRough = iDrawingDocument.SpecRough
iSpecRough.SignType = kompas6_constants.ksNoProcessingType

iSpecRough.Text = "Ra 12,5"
iSpecRough.Distance = 2
iSpecRough.AddSign = True


def функция():
    v = var.get()
    if v==1:
        iSpecRough.Delete()
        print('Удалить')
    if v==2:
        iSpecRough.Text = "Ra 12,5"
        iSpecRough.Distance = 2
        iSpecRough.AddSign = True
        print('Ra 12.5')
    if v==3:
        iSpecRough.Text = "Rz 100"
        iSpecRough.Distance = 2
        iSpecRough.AddSign = True
        print('Rz 100')
    iSpecRough.Update()
win = Tk()
win.wm_attributes('-topmost',1)
var = IntVar()

# Создадим кнопки
none = Radiobutton(win, text='Нет', value=1, variable=var).grid(row=0, column=0)                 #Для каждой кнопки своя value
Ra12_5 = Radiobutton(win, text='✓Ra 12.5', value=2, variable=var).grid(row=0, column=1)
Rz100 = Radiobutton(win, text='✓Rz 100', value=3, variable=var).grid(row=0, column=2)
knopka = Button(win,text='Выполнить',command = функция).grid(row=1, column=0)
win.mainloop()
input('Press ENTER to exit')