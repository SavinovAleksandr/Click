# -*- coding: utf-8 -*-
"""
Click v.5 - Макрос для экспорта графики из RG2 в Word
Точка входа приложения.
"""
from PyQt5.QtWidgets import QApplication
import multiprocessing

import Click_GUI

if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = QApplication([])
    window = Click_GUI.FileOrDirectorySelector()
    window.show()
    app.exec_()
