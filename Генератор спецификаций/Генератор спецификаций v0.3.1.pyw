# -*- coding: utf-8 -*-
# Генератор спецификаций для КОМПАС-3D
# http://forum.ascon.ru/index.php/topic,31378.0.html

title = u'Генератор спецификаций v0.3.1'

################  Н А С Т Р О Й К И  ###########################################################
# Данные для заполнения основной надписи (чтобы ячейка не заполнялась, заменить значение на u'')
Razrab = u'ТрындецЪ'
Prov = u'VLaD-Sh'
N_kontr = u'Helicoid'
Utv = u'Вират Лакх'
Organizaciya = u'http://vk.com/k_experts'

# Масса в ячейке "Примечание" для БЧ
mass_prim = True        # Показывать ли массу в ячейке Примечание(Да - True,  Нет - False)
do = u''                # Текст перед значением массы
posle = u' кг'          # Текст после значения массы
tochnost = 2            # кол-во знаков после запятой в значении массы

# Номера разделов СП
razdel_sp = {u"Документация":5,
        u"Комплексы": 10,
        u"Сборочные единицы": 15,
        u"Детали": 20,
        u"Стандартные изделия": 25,
        u"Прочие изделия": 30,
        u"Материалы": 35,
        u"Комплекты": 40}

# Типы чертежей для раздела "Документация"
type_doc = {u"СБ" : u"Сборочный чертеж",
        u"ВО" : u"Чертеж общего вида",
        u"ТЧ" : u"Теоретический чертеж",
        u"ГЧ" : u"Габаритный чертеж",
        u"МЭ" : u"Электромонтажный чертеж",
        u"МЧ" : u"Монтажный чертеж",
        u"УЧ" : u"Упаковочный чертеж"}
#############################################################################

import pythoncom, time, os, re, traceback
from win32com.client import Dispatch, gencache
try:
    import Tkinter as tk
    import tkMessageBox
except:
    import tkinter as tk
    import tkinter.messagebox as tkMessageBox

try:
    def EditSpc():
        '''
        Создаёт спецификацию
        '''
        iDocumentSpc = iKompasObject.SpcDocument()
        iDocumentParam = KAPI.ksDocumentParam(iKompasObject.GetParamStruct(35))     # ko_DocumentParam
        iDocumentParam.Init()
        iDocumentParam.type = 4

        iSheetParam = KAPI.ksSheetPar(iDocumentParam.GetLayoutParam())
        iSheetParam.Init()
        SystemPath = iKompasObject.ksSystemPath(0)                                  # путь к системной папке
        iSheetParam.layoutName = SystemPath + "\graphic.lyt"
        iSheetParam.shtType = 1                                                     #номер стиля создаваемой СП
        iDocumentSpc.ksCreateDocument(iDocumentParam)

        iDocumentSp = iApplication.ActiveDocument
        iSpecificationDescriptions = iDocumentSp.SpecificationDescriptions
        iSpecificationDescription = iSpecificationDescriptions.Active           # текущее описание СП
        iSpecificationBaseObjects = iSpecificationDescription.BaseObjects       # коллекция базовых ОС
        iSpc = iDocumentSpc.GetSpecification()
        iSpcObjParam = KAPI.ksSpcObjParam(iKompasObject.GetParamStruct(95)) # ko_SpcObjParam

        for i in OS_collection:
            iSpc.ksSpcObjectCreate("", 0, razdel_sp[OS_collection[i].razdel], 1, 0, 0) # создали ОС
            obj = iSpc.ksSpcObjectEnd()

            iSpc.ksSpcObjectEdit(obj)
            iDocumentSpc.ksGetObjParam(obj, iSpcObjParam, -1)
            iSpcObjParam.blockNumber = 0
            iSpcObjParam.draw = 1
            iSpcObjParam.firstOnSheet = 0
            iSpcObjParam.ispoln = 0
            iSpcObjParam.posInc = 1
            iSpcObjParam.posNotDraw = 0
            iDocumentSpc.ksSetObjParam(obj, iSpcObjParam, -1)

            try:
                int(i.split('-')[-1])                                           # проверка на исполнение
                isp = True
            except:
                isp = False

            if not OS_collection[i].doc:
                if OS_collection[i].razdel not in (u"Прочие изделия", u"Стандартные изделия", u"Материалы"):
                    iSpc.ksSetSpcObjectColumnText(4, 1, 0, OS_collection[i].oboznachenie )  # Обозначение

                iSpc.ksSetSpcObjectColumnText(5, 1, 0, OS_collection[i].naimenovanie)   # Наименование

                if OS_collection[i].razdel == u"Детали" and not isp:
                    iSpc.ksSetSpcObjectColumnText(1, 1, 0, u'БЧ')   # Формат для БЧ
                    if mass_prim:
                        iSpc.ksSetSpcObjectColumnText(7, 1, 0, '%s%s%s' %(do, str(round(OS_collection[i].massa, tochnost)).replace('.',',').rstrip('0').rstrip(','), posle))   # Масса в примечание
                elif OS_collection[i].razdel == u"Сборочные единицы" and not isp:
                    iSpc.ksSetSpcObjectColumnText(1, 1, 0, u'А4')   # Формат для Сб.ед.

            if OS_collection[i].razdel != u"Документация":
                iSpc.ksSetSpcObjectColumnText(6, 1, 0, OS_collection[i].kolichestvo)    # Кол-во
                iSpc.ksSpcMassa(str(OS_collection[i].massa))

            obj = iSpc.ksSpcObjectEnd()
            iSpecificationObject = iSpecificationBaseObjects.Item(obj)              # получим базовый ОС по Reference

            if OS_collection[i].doc:
                iAttachedDocuments = iSpecificationObject.AttachedDocuments         # интерфейс присоединенных документов
                iAttachedDocuments.Add( OS_collection[i].doc, True)                 # присоединить документ

            if OS_collection[i].razdel  not in (u"Прочие изделия", u"Стандартные изделия", u"Материалы"):

                if isp:
                    iSpecificationObject.Performance = True

            iSpecificationObject.Update()                                           # обновить ОС

            if OS_collection[i].doc:
                ## меняем цвет текста на синий
                for y in ( 1, 4, 5, 7):                                                       # для колонок Формат, Обозначение, Наименование
                    iSpecificationColumns = iSpecificationObject.Columns
                    iSpecificationColumn = iSpecificationColumns.Column(y, 1, 0)
                    iText = iSpecificationColumn.Text
                    TextLines = iText.TextLines

                    if not isinstance(TextLines, tuple):
                        TextLines = tuple(TextLines)

                    for iTextLine in TextLines:
                        TextItems = iTextLine.TextItems

                        if not isinstance(TextItems, tuple):
                            TextItems = tuple(TextItems)

                        for iTextItem in TextItems:
                            iTextFont = KAPI7.ITextFont(iTextItem)
                            iTextFont.Color = 0xff0000
                            iTextItem.Update()

                iSpecificationObject.Update()

        iDoc = iKompasObject.SpcActiveDocument()
        iStamp = iDoc.GetStamp()
        iStamp.ksOpenStamp()

        for key in Dictionary:                                                       # заполнение штампа
            iStamp.ksColumnNumber(key)
            iTextLineParam = KAPI.ksTextLineParam(iKompasObject.GetParamStruct(29)) # ko_TextLineParam
            iTextLineParam.Init()
            iTextItemArray = KAPI.ksDynamicArray(iKompasObject.GetDynamicArray(4))
            iTextItemParam = KAPI.ksTextItemParam(iKompasObject.GetParamStruct(31)) # ko_TextItemParam
            iTextItemParam.Init()
            iTextItemParam.s = Dictionary[key]
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)
            iStamp.ksTextLine(iTextLineParam)

        iStamp.ksCloseStamp()

        path_sp = '%s%s %s.spw' % (directory, Dictionary[2].replace('$|', ''), Dictionary[1])

        if os.path.exists( path_sp ):
            res = iApplication.MessageBoxEx(u'Спецификация с именем "%s %s.spw" уже имеется в папке. Перезаписать файл?' % (Dictionary[2].replace('$|', ''), Dictionary[1]), title, 4)
            if res == 6:
                iDocumentSpc.ksSaveDocument( path_sp ) # сохранение СП
        else:
            iDocumentSpc.ksSaveDocument( path_sp ) # сохранение СП


    class OS(object):

        def __init__(self, oboznachenie = '', naimenovanie = ''):
            self.razdel = ''
            self.oboznachenie = oboznachenie
            self.naimenovanie = naimenovanie
            self.kolichestvo = 0
            self.doc = None
            self.massa = 0


    #############################################################################################################################################
    Dictionary = {130: time.strftime("%d.%m.%y"), 110: Razrab, 111: Prov, 114: N_kontr, 115: Utv, 9: Organizaciya}

    #  Подключим описание интерфейсов API5
    KAPI = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    iKompasObject = KAPI.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(KAPI.KompasObject.CLSID, pythoncom.IID_IDispatch))

    #  Подключим описание интерфейсов API7
    KAPI7 =  gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    iApplication = KAPI7.IKompasAPIObject(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(KAPI7.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch)).Application
    const = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

    iDocument = iApplication.ActiveDocument                                         # получим активный документ
    iDocuments = iApplication.Documents                                             # получим коллекцию документов
    directory = iDocument.Path                                                      # путь к файлу документа

    if iDocument.DocumentType == 5:                                                 # если сборка

        iKompasDocument3D = KAPI7.IKompasDocument3D(iDocument)
        iPart7_SB = iKompasDocument3D.TopPart
        iParts7 = iPart7_SB.Parts                                                   # коллекция компонентов

        if iParts7.Count:                                                           # если есть компоненты
            ## Формирование словаря ОС #############################################
            OS_collection = {}                                                      # коллекция будущих ОС

            iPropertyMng = KAPI7.IPropertyMng(iApplication)
            iProperty_obozn = iPropertyMng.GetProperty(iDocument, 4.0)              # получим интерфейс свойства Обозначение
            iProperty_naimen = iPropertyMng.GetProperty(iDocument, 5.0)             # получим интерфейс свойства Наименование
            iProperty_razdel = iPropertyMng.GetProperty(iDocument, 20.0)            # получим интерфейс свойства Раздел спецификации
            ## сбор информации о главной сборке ##
            iPropertyKeeper = KAPI7.IPropertyKeeper(iPart7_SB)
            value_obozn = KAPI7.IEmbodimentsManager(iPart7_SB).CurrentEmbodiment.GetMarking(-1, True)                              # получим значение свойства Обозначение в настроенных единицах измерения
            value_naimen = iPropertyKeeper.GetPropertyValue( iProperty_naimen, None, True, True )[1]                            # получим значение свойства Наименование в настроенных единицах измерения


            if not value_obozn:
                OS_collection[value_naimen.replace(' ', '')] =  OS('', value_naimen)  # для сборки без обозначения
                OS_collection[value_naimen.replace(' ', '')].razdel = u'Документация'
                Dictionary[2] = ''

            else:
                obozn = '$' + value_obozn.replace('$|', '').replace(' ', '')
                tip_d = u"Сборочный чертеж"                                     # тип документа по умолчанию
                for tip in type_doc:
                    if obozn.endswith(tip):
                        obozn = obozn.rstrip(tip)
                        tip_d = type_doc[tip]
                        break

                OS_collection[obozn] =  OS(value_obozn.rstrip(u'СБ') + u'СБ', tip_d)                  # метим ключ словаря символом '$' для сборки имеющей обозначение
                OS_collection[obozn].razdel = u'Документация'
                Dictionary[2] = obozn[1:]                                       # получаем Обозначение

            if value_naimen:
                Dictionary[1] = naimen = value_naimen.rstrip(' ')

            else:
                Dictionary[1] = ''
            ## обход компонентов ##
            List_not_razdel = []                                                                                                 # детали без раздела СП

            for i in range(iParts7.Count):                                                                                       # перебираем все компоненты
                iPart7 = iParts7.Part (i)

                if KAPI7.IFeature7(iPart7).Excluded:                                                                             # если компонент исключен из расчёта
                    continue                                                                                                     # пропускаем его

                iPropertyKeeper = KAPI7.IPropertyKeeper(iPart7)
                value_razdel = iPropertyKeeper.GetPropertyValue( iProperty_razdel, None, True, True )[1]                         # получим значение свойства Раздел спецификациие в настроенных единицах измерения
                value_obozn = iPropertyKeeper.GetPropertyValue( iProperty_obozn, None, True, True )[1]                            # получим значение свойства Обозначение в настроенных единицах измерения

                if not value_obozn:
                    value_obozn = ''

                value_naimen = iPropertyKeeper.GetPropertyValue( iProperty_naimen, None, True, True )[1]                         # получим значение свойства Наименование в настроенных единицах измерения

                if not value_naimen:
                    value_naimen = ''

                if value_razdel:

                    if value_obozn != '':
                        obozn = '$' + value_obozn.replace('$|', '').replace(' ', '')

                        for tip in type_doc:
                            obozn = obozn.replace(tip, '')

                        OS_collection.setdefault(obozn, OS(value_obozn, value_naimen)).kolichestvo += 1 # заполняем словарь объектами класса, считая кол-во одинаковых объектов

                        if OS_collection[obozn].razdel == '':
                            OS_collection[obozn].razdel = value_razdel
                            OS_collection[obozn].massa = iPart7.Mass/1000                                                                 # записали массу                                                               # записали раздел спецификации
                    else:
                        OS_collection.setdefault(value_naimen.replace(' ', ''), OS(value_obozn, value_naimen)).kolichestvo += 1      # заполняем словарь объектами класса, считая кол-во одинаковых объектов

                        if OS_collection[value_naimen.replace(' ', '')].razdel == '':
                            OS_collection[value_naimen.replace(' ', '')].razdel = value_razdel
                            OS_collection[value_naimen.replace(' ', '')].massa = iPart7.Mass/1000                                        # записали массу                                       # записали раздел спецификации
                else:
                    if (value_obozn, value_naimen) not in List_not_razdel:
                        List_not_razdel.append((value_obozn, value_naimen))


            ## Поиск чертежей и спецификаций #######################################
            list_obozn_det, list_obozn_sbed = [], []                                # список обозначений Деталей и Сборочных единиц для поиска их чертежей и СП

            for i in OS_collection:

                if i[0] == '$':                                                     # если ОС с Обозначением

                    if OS_collection[i].razdel == u'Сборочные единицы':
                        list_obozn_sbed.append(i)                                   # заполняем список обозначений Сборочных единиц
                    elif OS_collection[i].razdel in (u'Детали', u'Документация'):
                        list_obozn_det.append(i)                                    # заполняем список обозначений Деталей
            rassh = []

            if len (list_obozn_sbed) > 0:
                rassh.append('.spw')                                                # если есть Сборочные единицы, то ведём поиск и по спецификациям

            if len (list_obozn_det) > 0:
                rassh.append('.cdw')

            if len(rassh)>0:
                rassh = tuple(rassh)

                for d, dirs, files in os.walk(directory):                               # генератор поддиректорий

                    for f in files:                                                     # для каждого файла в папке

                        if len(list_obozn_sbed) == 0 and len(list_obozn_det) ==0:            # если нашли все документы прерываем поиск файлов
                            ff = 5
                            break

                        if f.endswith(rassh):                                               # если файл чертёж или СП
                            PathName = d.replace('\\', '/')+ '/' + f                         # формируем полный путь к файлу
                            iDoc = iDocuments.Open (PathName, False, True)          # открываем файл False - в невидимом режиме True - только для чтения

                            if iDoc:
                                iLayoutSheets = iDoc.LayoutSheets
                                iLayoutSheet = iLayoutSheets.ItemByNumber (1)
                                iStamp = iLayoutSheet.Stamp
                                Obozn = iStamp.Text( 2 ).Str.replace('$|', '').replace(' ', '') # получаем Обозначение

                                for tip in type_doc:
                                    Obozn = Obozn.replace(tip, '')

                                iDoc.Close(0)

                                if Obozn == '':
                                    Obozn = None
                            else:
                                 Obozn = None

                            if Obozn:                                                   # если Обозначение заполнено

                                if f.endswith('.spw') and len(list_obozn_sbed) > 0:
                                    if '$' + Obozn in list_obozn_sbed:
                                        OS_collection['$' + Obozn].doc = PathName       # добавляем путь к спецификации в свойства ОС
                                        list_obozn_sbed.remove('$' + Obozn)
                                else:
                                    if '$' + Obozn in list_obozn_sbed:
                                        OS_collection['$' + Obozn].doc = PathName       # добавляем путь к спецификации в свойства ОС

                                    if '$' + Obozn in list_obozn_det:
                                        OS_collection['$' + Obozn].doc = PathName       # добавляем путь к чертежу в свойства ОС
                                        list_obozn_det.remove('$' + Obozn)


            EditSpc()                                                               # создаём и заполняем СП

            if len(List_not_razdel) >0:
                text =u'Генерация спецификации завершена.\nВ сборке есть детали, не имеющие раздела спецификаци:'

                for n in List_not_razdel:
                    text += '\n%s %s' %(n[0], n[1])

                iApplication.MessageBoxEx(text, title, 48)

            else:
                iApplication.MessageBoxEx(u'Генерация спецификации завершена', title, 64)

        else:
            iApplication.MessageBoxEx(u'Активная сборка не имеет компонентов!', title, 48)
    else:
        iApplication.MessageBoxEx(u'Активный документ должен быть сборкой!', title, 48)

except:
    root = tk.Tk()
    root.withdraw()
    tkMessageBox.showwarning(title, traceback.format_exc().decode('cp1251'))	# показываем окно с выводом ошибки
    root.destroy()