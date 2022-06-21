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

# Получим активное приложение
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
iApplication = kompas_api_object.Application
iKompasDocument = iApplication.ActiveDocument

#Нахождение пути
path = path = os.path.dirname(os.path.abspath(__file__)) + "/"
os.chdir(path)


for i in range(10):
    i = i + 2
    # Подключение переменных из экселя
    Excel = Dispatch('Excel.Application')
    book = Excel.Workbooks.open(r"D:\МАКРОСЫ\Ограничительный перечень\ФЛАНЦЫ\Фланцы круглые с резьбой\БТЛИ.711142 Фланец.xls").ActiveSheet
    listEx=Excel.ActiveSheet
    D1 = book.Cells(i,1).value
    Dnar = book.Cells(i,2).value
    S = book.Cells(i,3).value
    Dotv = book.Cells(i,4).value
    N = book.Cells(i,5).value
    D2 = book.Cells(i,6).value
    l2 = book.Cells(i,7).value
    l1 = book.Cells(i,8).value
    d = book.Cells(i,9).value
    shag = book.Cells(i,10).value
    oboz = listEx.Cells(i,11).value
    name = listEx.Cells(i,12).value
    file_name = str(oboz + ' ' + name)

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
    var1 = VariableCollection.GetByName('D1',True,True)
    var1.value = D1
    iPart.RebuildModel()# 'А' - имя
    var2 = VariableCollection.GetByName('Dnar',True,True)
    var2.value = Dnar
    iPart.RebuildModel()
    var3 = VariableCollection.GetByName('S',True,True)
    var3.value = S
    iPart.RebuildModel()
    var4 = VariableCollection.GetByName('Dotv',True,True)
    var4.value = Dotv
    iPart.RebuildModel()
    var5 = VariableCollection.GetByName('N',True,True)
    var5.value = N
    iPart.RebuildModel()
    var6 = VariableCollection.GetByName('D2',True,True)
    var6.value = D2
    iPart.RebuildModel()
    var7 = VariableCollection.GetByName('l2', True, True)
    var7.value = l2
    iPart.RebuildModel()
    var8 = VariableCollection.GetByName('l1', True, True)
    var8.value = l1
    iPart.RebuildModel()
    var9 = VariableCollection.GetByName('d', True, True)
    var9.value = d
    iPart.RebuildModel()
    var10 = VariableCollection.GetByName('shag', True, True)
    var10.value = shag
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
    iPart.marking = oboz
    iPart.name = name
    iPart.Update()
    kompas_document.SaveAs(path + file_name + '.m3d')

    # name = iKompasDocument.Name
    # name = name.split('.m3d')
    # print(name)
    # name1 = name.pop(0)
    # name = name1.split(' ')  # разделитель
    # ob = name.pop(0)
    # name = ' '.join(name)
    # print(ob)
    # print(name)


    print(file_name)
kompas_document.Close(False)