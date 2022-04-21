from PyQt5 import uic
from PyQt5.QtWidgets import QApplication
import os

Window = uic.loadUiType("Main_frame_GUI.ui")
app = QApplication([])
window = Window()
window.show()

app.exec_()