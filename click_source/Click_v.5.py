# -*- coding: utf-8 -*-
"""
Click v.5 - Макрос для экспорта графики из RG2 в Word
Точка входа приложения.
"""
import logging
import multiprocessing

from PyQt5.QtWidgets import QApplication

import Click_GUI

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = QApplication([])
    window = Click_GUI.FileOrDirectorySelector()
    window.show()
    app.exec_()
