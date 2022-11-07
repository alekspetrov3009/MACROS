import os
import pythoncom
from Kompas6API5 import KompasObject
from win32com.client import Dispatch, gencache
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook
import datetime
import LDefin2D
import MiscellaneousHelpers as MH
from openpyxl.styles.numbers import BUILTIN_FORMATS
import psutil
from tkinter import *
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import ttk



kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Получи API интерфейсов версии 5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                             pythoncom.IID_IDispatch))

#  Получи API интерфейсов версии 7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID,
                                                             pythoncom.IID_IDispatch))
application = kompas_api_object.Application
MH.iApplication = application
Documents = application.Documents


dxfversion = 2010 # Какую версию dxf использовать: 2000, 2004, 2007, 2010, 2013
AutoCAD = {2000:4, 2004:5, 2007:6, 2010:7, 2013:8, 2018:9} # Поддерживаемые версии DXF/DWG

iConverter = iApplication.Converter(KompasObject.ksSystemPath(1) + '\\ImpExp\\dwgdxfExp.rtw') # Конвертер файлов КОМПАС
iConverterParameters = iConverter.ConverterParameters(1) # Получить интерфейс параметров конвертирования (для dxf: command = 1)
currentAcadFileVersion = iConverterParameters.AcadFileVersion # Сохранить текущую версию формата dxf
iConverterParameters.AcadFileVersion = AutoCAD[dxfversion] # Версия AutoCAD, в которую осуществляем запись
iConverter.Convert('', 'c:\\1.dxf', 1, False) # Процесс конвертации (файл или текущий документ, новый файл, номер команды, диалог)
iConverterParameters.AcadFileVersion = currentAcadFileVersion # Вернуть текущую версию формата dxf



cdw_path = os.getcwd()

# Проверка, запущен ли компас
kompas_ex = False

for proc in psutil.process_iter():
    name = proc.name()
    if name == "KOMPAS.Exe":
        print('ОК')
        kompas_ex = True
        break

if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу окно выбора файлов
    filenames = askopenfilenames(title="Выберите чертежи деталей",
                                 filetypes=[('Компас 3D', '*.cdw'), ('Компас 3D', '*.spw')])
    print(filenames)

root.destroy()  # Уничтожаем основное окно
root.mainloop()

def to_DWG():
    for i in filenames:
        #  Открываем документ
        kompas_document = Documents.Open(i, False, False)
        kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
        iDocument2D = kompas_object.ActiveDocument2D()
        # Получаем свойства листов
        layout_sheets = kompas_document.LayoutSheets
        layout_sheet = layout_sheets.ItemByNumber(1)
        sheet_format = layout_sheet.Format
        sheet = kompas_document.LayoutSheets.Count  # счетчик листов
        # Считывание штампа
        iStamp = layout_sheet.Stamp
        iText = iStamp.Text(1)  # номер ячейки для считывания данных наименования
        naim = iText.Str
        iText = iStamp.Text(2)  # номер ячейки для считывания данных обозначения
        oboz = iText.Str
        new_name = cdw_path+ '/' + oboz + ' ' + naim + '.dxf'
        iDocument2D.ksSaveToDXF(new_name)
        print(new_name)

##################################################
if kompas_ex == True:
    print("Запущен")
    to_DWG()
else:
    print("Не запущен")

    # kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    # kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
    #
    # #  Получи API интерфейсов версии 5
    # kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    # kompas_object = kompas6_api5_module.KompasObject(
    #     Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
    #                                                              pythoncom.IID_IDispatch))
    #
    # #  Получи API интерфейсов версии 7
    # kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    # kompas_api_object = kompas_api7_module.IKompasAPIObject(
    #     Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID,
    #                                                              pythoncom.IID_IDispatch))
    # application = kompas_api_object.Application
    # MH.iApplication = application
    # Documents = application.Documents
    MH.iApplication = application
    MH.iApplication.Visible = False  # Видимость компаса

    # Вызов функции
    to_DWG()

    # Закрытие фонового приложения Компаса
    MH.iApplication.Quit()








