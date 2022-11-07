import os, time
import pythoncom
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
iDocument2D = kompas_object.ActiveDocument2D()

iDocument2D.ksSaveDocument("C:/Users/Администратор/Desktop/Новая папка/Чертеж1.pdf")

