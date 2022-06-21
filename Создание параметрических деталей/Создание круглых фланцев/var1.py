# -*- coding: utf-8 -*-


#Импортируем необходимые библиотеки
import os
from pydoc import doc
import pythoncom
from win32com.client import Dispatch, gencache

import LDefin3D
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

#Нахождение пути
path = path = os.path.dirname(os.path.abspath(__file__)) + "/"
os.chdir(path)


for i in range(100):
    i = i + 2
    # Подключение переменных из экселя
    Excel = Dispatch('Excel.Application')
    book = Excel.Workbooks.open(r"D:\МАКРОСЫ\Создание параметрических деталей\БТЛИ.711142 Фланец.xls").ActiveSheet
    listEx=Excel.ActiveSheet
    d = book.Cells(i,1).value
    Dvn = book.Cells(i,2).value
    s = book.Cells(i,3).value
    d1 = book.Cells(i,4).value
    n = book.Cells(i,5).value
    Dmo = book.Cells(i,6).value
    obozn = listEx.Cells(i,7).value
    name = listEx.Cells(i,8).value
    file_name = str(obozn + ' ' + name)

    #Подключаемся к активной модели
    Documents = application.Documents
    #  Получим активный документ
    kompas_document = application.ActiveDocument
    kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
    iDocument3D = kompas_object.ActiveDocument3D()

    #Получаем интерфейс компонента
    iPart = iDocument3D.GetPart(LDefin3D.pTop_Part)

    #Получаем коллекцию внешних переменных
    VariableCollection = iPart.VariableCollection()

    #print(VariableCollection)

    #обновляем коллекцию внешних переменных
    VariableCollection.refresh()

    #Получаем интерфейс переменной по её имени
    var1 = VariableCollection.GetByName('d',True,True)
    var1.value = d
    iPart.RebuildModel()# 'А' - имя
    var2 = VariableCollection.GetByName('Dvn',True,True)
    var2.value = Dvn
    iPart.RebuildModel()
    var3 = VariableCollection.GetByName('s',True,True)
    var3.value = s
    iPart.RebuildModel()
    var4 = VariableCollection.GetByName('d1',True,True)
    var4.value = d1
    iPart.RebuildModel()
    var5 = VariableCollection.GetByName('n',True,True)
    var5.value = n
    iPart.RebuildModel()
    var6 = VariableCollection.GetByName('Dmo',True,True)
    var6.value = Dmo
    iPart.RebuildModel()
    # #Задаём новое значение переменной
    # var1.value = d
    # iPart.RebuildModel()
    # var2.value = Dvn
    # iPart.RebuildModel()
    # var3.value = s
    # iPart.RebuildModel()
    # var4.value = d1
    # iPart.RebuildModel()
    # var5.value = n
    # iPart.RebuildModel()
    # var6.value = D1

    #Перестраиваем модель
    iPart.RebuildModel()
    print(path)

    # Заполнение обозначения и наименования
    iPart.marking = obozn
    iPart.name = name
    iPart.Update()

    kompas_document.SaveAs(path + file_name + '.m3d')
    print(file_name)
kompas_document.Close(False)