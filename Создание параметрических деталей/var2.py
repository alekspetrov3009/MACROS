# -*- coding: utf-8 -*-
#|var2

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
kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()

iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

obj = iPart.NewEntity(kompas6_constants_3d.o3d_thread)
iDefinition = obj.GetDefinition()
iDefinition.allLength = True
iDefinition.autoDefinDr = False
iDefinition.dr = 12
iDefinition.faceValue = True
iDefinition.length = 40
iDefinition.p = 1.75
iCollection = iPart.EntityCollection(kompas6_constants_3d.o3d_face)
iCollection.SelectByPoint(-28.268722831604, -41.550089564762, 20)
iFace = iCollection.First()
iDefinition.SetBaseObject(iFace)
iDefinition.SetFaceBegin(iFace)
iDefinition.SetFaceEnd(iFace)
obj.name = "Условное изображение резьбы:1"
iColorParam = obj.ColorParam()
iColorParam.color = 6008319
obj.Create()
