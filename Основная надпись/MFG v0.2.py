from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow
import os

Form, Window = uic.loadUiType("Main_frame_GUI.ui")
app = QMainWindow([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

path = os.path.dirname(os.path.abspath(__file__))
os.chdir(path)


def on_click():
    global p_number, p_name
    p_number = form.textEdit.toPlainText()
    p_name = form.textEdit_2.toPlainText()
    print(p_number, p_name)
    
def make_dir():
    os.mkdir(f'{p_number} {p_name}')
    os.path.abspath(path)
    os.chdir(f'{p_number} {p_name}')

def make_folders():
    os.mkdir('!Проект')
    os.mkdir('!Стандартные и покупные изделия')
    os.mkdir(p_number + '.00.00.00 Трансформатор')
    os.mkdir(p_number + '.00.00.00УЧ Упаковка')
    os.makedirs(f'{p_number}.10.00.00 Бак/{p_number}.10.10.00 Бак')
    os.makedirs(f'{p_number}.20.00.00 Активная часть/{p_number}.21.00.00 Крышка')
    os.makedirs(f'{p_number}.40.00.00 Расширитель/{p_number}.41.00.00 Расширитель')
    os.mkdir(p_number + '.50.00.00 Система охлаждения')
    os.mkdir(p_number + '.60.00.00 Вторичные цепи')
    

form.btn.clicked.connect(on_click)
form.btn.clicked.connect(make_dir)
form.btn.clicked.connect(make_folders)

#form.label.setText('asdasdsa')


app.exec_()