# -*- coding:  utf-8 -*-

import pythoncom, os
from win32com.client import Dispatch, gencache
from pdfrw import PdfReader, PdfWriter, IndirectPdfDict, PageMerge


def split(path, output):
    pdf_obj = PdfReader(path)
    total_pages = len(pdf_obj.pages)

    writer = PdfWriter()

    for page in range(2, total_pages):
        if page <= total_pages:
            writer.addpage(pdf_obj.pages[page])

    writer.write(output)


def concatenate(paths, output):
    writer = PdfWriter()

    for path in paths:
        reader = PdfReader(path)
        writer.addpages(reader.pages)

    writer.write(output)


def get4(srcpages):
    scale = 1
    srcpages = PageMerge() + srcpages
    x_increment, y_increment = (scale * i for i in srcpages.xobj_box[2:])
    for i, page in enumerate(srcpages):
        page.scale(scale)
        page.x = x_increment if i & 1 else 0
        # page.y = 0 if i & 2 else y_increment
    return srcpages.render()


def scale_pdf(path, output):
    pages = PdfReader(path).pages
    writer = PdfWriter(output)
    four_pages = get4(pages[0: 2])
    writer.addpage(four_pages)
    writer.write()


if __name__ == '__main__':
    #  Подключим описание интерфейсов API5
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                                 pythoncom.IID_IDispatch))

    #  Подключим описание интерфейсов API7
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = kompas_api7_module.IApplication(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                                 pythoncom.IID_IDispatch))

    iConverter = application.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")

    iDocument = application.ActiveDocument

    if iDocument:

        if iDocument.DocumentType in (1, 3):

            iConverter.Convert(iDocument.PathName, iDocument.Path + "\\" + iDocument.Name[:-4] + ".pdf", 0, False)
            application.MessageBoxEx("Создан файл\n" + iDocument.Name[:-4] + ".pdf", "Сохранение в *.pdf", 64)
        else:
            application.MessageBoxEx("Активный документ не является чертежом или спецификацией!", "Сохранение в *.pdf",
                                     48)
    else:
        application.MessageBoxEx("Нет активного документа!", "Сохранение в *.pdf", 48)

    scale_pdf(iDocument.Path + "\\" + iDocument.Name[:-4] + ".pdf", 'Чертеж2.pdf')
    split(iDocument.Path + "\\" + iDocument.Name[:-4] + ".pdf", 'Чертеж3.pdf')
    paths = ['Чертеж2.pdf', 'Чертеж3.pdf']
    concatenate(paths, iDocument.Path + "\\" + iDocument.Name[:-4] + ".pdf")

    list_pdf = ['Чертеж2.pdf', 'Чертеж3.pdf']

    for j in list_pdf:
        if os.path.isfile(j):
            os.remove(j)
            print("success")
        else:
            print("File doesn't exists!")
