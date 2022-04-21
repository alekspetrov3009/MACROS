# -*- coding: utf-8 -*-
#|Основная надпись

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import time
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

iStamp = iDocument2D.GetStamp()
iStamp.ksOpenStamp()

def sprav_nomer():
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

# def perv_prim():
#     iStamp.ksOpenStamp()
#     iStamp.ksColumnNumber(25)

#     iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
#     iTextLineParam.Init()
#     iTextLineParam.style = 32768
#     iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
#     iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
#     iTextItemParam.Init()
#     iTextItemParam.iSNumb = 0
#     iTextItemParam.s = "Т2303."
#     iTextItemParam.type = 0
#     iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
#     iTextItemFont.Init()
#     iTextItemFont.bitVector = 4096
#     iTextItemFont.color = 0
#     iTextItemFont.fontName = "GOST type A"
#     iTextItemFont.height = 3.5
#     iTextItemFont.ksu = 1
#     iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
#     iTextLineParam.SetTextItemArr(iTextItemArray)

#     iStamp.ksTextLine(iTextLineParam)

def type():
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



iStamp = iDocument2D.GetStamp()
iStamp.ksOpenStamp()

def razrab():
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
sprav_nomer()

def prover():
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

def tehnolog():
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

def rukov():  
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

def n_kontr(): 
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

def utverdil():
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

def podp_razrab():
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
    # iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
    # iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
    # iTextItemParam.Init()
    # iTextItemParam.iSNumb = 0
    # iTextItemParam.s = ""
    # iTextItemParam.type = 0
    # iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
    # iTextItemFont.Init()
    # iTextItemFont.bitVector = 0
    # iTextItemFont.color = 0
    # iTextItemFont.fontName = "GOST type A"
    # iTextItemFont.height = 3.5
    # iTextItemFont.ksu = 1
    iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
    iTextLineParam.SetTextItemArr(iTextItemArray)

    iStamp.ksTextLine(iTextLineParam)

def date():
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



win = Tk()
win.wm_attributes('-topmost',1)
var = IntVar()