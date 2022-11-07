# -*- coding: utf-8 -*-
title = "save_date_name"
import pythoncom
import datetime
from win32com.client import Dispatch, gencache
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
application = kompas_api_object.Application
kompas_document = application.ActiveDocument
name_doc = kompas_document.Name
#print name_doc
if  str(datetime.date.today()) == name_doc.split('_')[0]:
    #print 'date'
    kompas_document.SaveAs(kompas_document.Path + "\\" + kompas_document.Name)
    application.MessageBoxEx(u"Cохранили файл !", title, 64)
else:
    #print 'no date'
    kompas_document.SaveAs(kompas_document.Path + "\\" + str(datetime.date.today()) + "_" + kompas_document.Name)
    application.MessageBoxEx(u"Добавили к имени файла год-месяц-число !", title, 64)
