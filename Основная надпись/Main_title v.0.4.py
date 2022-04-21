# -*- coding: utf-8 -*-
#|Основная надпись

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import time
import MiscellaneousHelpers as MH
import tkinter as tk

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

class side_format():
    def sprav_nomer():
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
        iTextItemParam.s = "Т2303 ЭТЦПК-6300/10-У3"
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
        iStamp.ksCloseStamp()

    def perv_prim():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(25)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Т2303."
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
        iStamp.ksCloseStamp()

    def type():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(999)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Тип: ЭОДЦН-8200/10-У3"
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
        iStamp.ksCloseStamp()

class bottom_format():    
    def razrab():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(110)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Петров"
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
        iStamp.ksCloseStamp()

    def prover():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(111)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Сорокин"
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
        iStamp.ksCloseStamp()

    def tehnolog():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(112)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Маркин"
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
        iStamp.ksCloseStamp()

    def rukov():  
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(113)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Уфрутов"
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
        iStamp.ksCloseStamp()

    def n_kontr(): 
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(114)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Шкурин"
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
        iStamp.ksCloseStamp()

    def utverdil():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(115)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "Горбунов"
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
        iStamp.ksCloseStamp()

    def date():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(130)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = time.strftime("%d.%m.%y")
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
        iStamp.ksCloseStamp()

class podpisi():
    def podp_razrab():
        iStamp = iDocument2D.GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(120)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 32768
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "TIeTpob"
        iTextItemParam.type = 0
        iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "Staccato222 BT"
        iTextItemFont.height = 3.5
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)

        iStamp.ksTextLine(iTextLineParam)
        iStamp.ksCloseStamp()
# side_format.sprav_nomer()
# side_format.perv_prim()
# side_format.type()
# bottom_format.date()
# bottom_format.n_kontr()
# bottom_format.prover()
# bottom_format.razrab()
# bottom_format.rukov()
# bottom_format.tehnolog()
# bottom_format.utverdil()

# side_format.type()

window = tk.Tk()
#win.wm_attributes('-topmost',1)
#var = IntVar()

label = tk.Label(text = "Разработал").grid(row = 0, column = 0)
entry = tk.Entry(width=20, bg="white", fg="black").grid(row = 0, column = 1)

label = tk.Label(text = "Проверил").grid(row = 1, column = 0)
entry = tk.Entry(width=20, bg="white", fg="black").grid(row = 1, column = 1)

button = tk.Button(command = podpisi.podp_razrab, text = "Вставить").grid(row = 2, column = 1)
 
#entry.insert(0, "Проверил")
 
window.mainloop()
