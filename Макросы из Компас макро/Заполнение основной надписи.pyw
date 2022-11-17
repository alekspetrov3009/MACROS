# -*- coding: utf-8 -*-

#  PyKompasMacro https://slaviationsoft.blogspot.com
#  КОМПАС-3D (20, 0, 0, 0)
#  PyKompasMacro.exe (1.7.51.104)
#  aleks 14-11-2022 19:26:27

import pythoncom
from win32com.client import Dispatch, gencache, VARIANT

#  Получи константы
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Получи API интерфейсов версии 5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Получи API интерфейсов версии 7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
application = kompas_api_object.Application

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(2)
text.Str = "Обозначение"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(1)
text.Str = "Наименование"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(3)
text.Str = "Масса"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(110)
text.Str = "Разраб"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(111)
text.Str = "Провер"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(112)
text.Str = "Технолог"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(114)
text.Str = "нконтр"
stamp.Update()

#  Получи интерфейс активного документа
kompas_document = application.ActiveDocument

#  Отредактируй ячейки основной надписи на листе 1
layout_sheets = kompas_document.LayoutSheets
layout_sheet = layout_sheets.Item(0)
stamp = layout_sheet.Stamp
text = stamp.Text(115)
text.Str = "утвер"
stamp.Update()
