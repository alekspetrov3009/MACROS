# -*- coding: cp1251 -*-
#|�������������� ���������� ��������� �������

import pythoncom
from win32com.client import Dispatch, gencache

################################################################
n1 = 1 # ����� ����� ���������� ������� ����� �� ����������
n2 = 2 # ����� ����� ���������� ����������� ������ �� ����������
################################################################

#  ��������� �������� ����������� API7
KAPI7 =  gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
iApplication = KAPI7.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(KAPI7.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch)).Application

iDocument = iApplication.ActiveDocument

if iDocument:

    if iDocument.DocumentType == 1:
        iLayoutSheets = iDocument.LayoutSheets

        for i in range (iLayoutSheets.Count):
            iLayoutSheet = iLayoutSheets.Item(i)

            if i:
                iLayoutSheet.LayoutStyleNumber = n2
            else:
                iLayoutSheet.LayoutStyleNumber = n1
            iLayoutSheet.Update()

        iApplication.MessageBoxEx( "���������� ��������.", "", 64)