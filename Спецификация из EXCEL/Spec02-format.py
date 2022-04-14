# -*- coding: utf-8 -*-
#|заполнить сп


import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
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
kompas_document_spc = kompas_api7_module.ISpecificationDocument(kompas_document)
iDocumentSpc = kompas_object.SpcActiveDocument()
###
import os
put=os.getcwd()
print (put)

import win32com.client
Excel=win32com.client.Dispatch('Excel.Application')
ex=Excel.Workbooks.Open(put+'\SP.xlsx')

listEx=ex.ActiveSheet
stroka1=0
stroka2=90


while stroka1<stroka2:
    stroka1=stroka1+1
    vivod1 =listEx.Cells(stroka1,1).value # Формат
    vivod4=listEx.Cells(stroka1,4).value     # Обозначение
    vivod5=listEx.Cells(stroka1,5).value    # Наименование
    vivod6 = listEx.Cells(stroka1,6).value  # Количество
    vivod7 =listEx.Cells(stroka1,7).value   # Примечание
    if vivod4==None:
        break
    else:
        print (vivod4, vivod5,vivod6)
        iSpc = iDocumentSpc.GetSpecification()
        iSpc.ksSpcObjectCreate("", 0, 20, 0, 0, 0)
        obj = iSpc.ksSpcObjectEnd()


        iSpcObjParam = kompas6_api5_module.ksSpcObjParam(kompas_object.GetParamStruct(kompas6_constants.ko_SpcObjParam))
        iSpc.ksSpcObjectEdit(obj)

        iDocumentSpc.ksGetObjParam(obj, iSpcObjParam, LDefin2D.ALLPARAM)
        '''
        iSpcObjParam.blockNumber = 2
        iSpcObjParam.draw = 1
        iSpcObjParam.firstOnSheet = 0
        iSpcObjParam.ispoln = 0
        iSpcObjParam.posInc = 0
        iSpcObjParam.posInc = 0
        '''

        iDocumentSpc.ksSetObjParam(obj, iSpcObjParam, LDefin2D.ALLPARAM)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)





        if vivod1==None:
            iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 0
            iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = ""
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 5
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)
        else:
            iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 0
            iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = vivod1
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 5
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)
####################################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(1, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
        iTextItemParam.type = 0
        iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = 5
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
################################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(2, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ''
        iTextItemParam.type = 0
        iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = 5
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
############################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(3, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = vivod4
        iTextItemParam.type = 0
        iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = 5
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
############################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(4, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = vivod5
        iTextItemParam.type = 0
        iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = 5
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
###################################################################################
        if vivod6==None:
            iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
            iSpc.ksSetSpcObjectColumnTextEx(5, 1, 0, iTextLineArray)
            iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

            iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 0
            iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = ""               #"3"     kol-vo########################################################################################
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 5
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)

        else:

            iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
            iSpc.ksSetSpcObjectColumnTextEx(5, 1, 0, iTextLineArray)
            iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

            iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 0
            iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = vivod6               #"3"     kol-vo########################################################################################
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 5
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)

##########################################################################
        if vivod7==None:
            iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
            iSpc.ksSetSpcObjectColumnTextEx(6, 1, 0, iTextLineArray)
            iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

            iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 0
            iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = ""   # ПРИМЕЧАНИЕ пустое
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 5
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)
        else:
            iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
            iSpc.ksSetSpcObjectColumnTextEx(6, 1, 0, iTextLineArray)
            iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

            iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 0
            iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = vivod7    # ПРИМЕЧАНИЕ
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 5
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)
        ##################################################################################################

        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(7, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
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
###########################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(8, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
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
###########################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 13, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
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
############################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 14, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
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
############################################################################################################
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 15, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
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

        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 16, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = ""
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

        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 4, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "eeee"
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

        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 1, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "qqqqq"
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

        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 2, 0, iTextLineArray)
        iTextLineArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_LINE_ARR)

        iTextLineParam = kompas6_api5_module.ksTextLineParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 0
        iTextItemArray = kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = kompas6_api5_module.ksTextItemParam(kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = "wwwww" ############################################################################################################################
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

        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iSpc.ksSetSpcObjectColumnTextEx(10, 3, 0, iTextLineArray)
        obj = iSpc.ksSpcObjectEnd()