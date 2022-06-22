import os
import re
import subprocess
import pythoncom
from win32com.client import Dispatch, gencache
from tkinter import Tk
from tkinter.filedialog import askopenfilenames


# Подключение к API7 программы Компас 3D
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const


def is_running():
    proc_list = \
        subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[
            0]
    return True if proc_list else False


# Посчитаем количество листов каждого из формата
def amount_sheet(doc7):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
    return sheets


def parse_design_documents(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    table = []  # Создаём таблицу парметров
    for path in paths:
        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=True,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять

        row = amount_sheet(doc7)  # Посчитаем кол-во листов каждого формат
        row.update(stamp(doc7))  # Читаем основную надпись
        row.update({
            "Filename": doc7.Name,  # Имя файла
            "CountTD": count_demand(doc7, module7),  # Количество пунктов технических требований
            "CountDim": count_dimension(doc7, module7),  # Количество пунктов технических требований
        })
        table.append(row)  # Добавляем строку параметров в таблицу

        doc7.Close(const7.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return table


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу окно выбора файлов

    filenames = askopenfilenames(title="Выберети чертежи деталей", filetypes=[('Компас 3D', '*.cdw'), ])

    root.destroy()  # Уничтожаем основное окно
    root.mainloop()
