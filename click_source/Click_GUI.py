# -*- coding: utf-8 -*-
"""
Click_GUI - Графический интерфейс макроса Click
Использует PyQt5 для создания окна выбора файлов и параметров.
"""
import logging
import time
from os import cpu_count, path as os_path, listdir
from pathlib import Path

from PyQt5.QtWidgets import QtWidgets as pq
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from concurrent.futures import ProcessPoolExecutor

from Click_Rastr import create_dxf
from Click_word import process_directory, process_directory_com
from config import get_icon_path

logger = logging.getLogger(__name__)


def _create_dxf_safe(fn, path_grf, com_value, path_sch=None):
    """Обёртка create_dxf с обработкой исключений."""
    try:
        if path_sch:
            create_dxf(fn, path_grf, com_value, path_sch)
        else:
            create_dxf(fn, path_grf, com_value)
    except Exception as e:
        logger.exception('Ошибка при обработке %s: %s', fn, e)

try:
    from Click_word_r7 import process_directory_r7
    R7_AVAILABLE = True
except ImportError:
    R7_AVAILABLE = False


class HeavyWorker(QThread):
    """Поток для выполнения тяжёлых операций (экспорт RG2 -> DXF -> PNG/SVG -> Word)."""
    finished_status = pyqtSignal(float)

    def __init__(self, max_w, rg2, path_grf, path_sch, path_wrd, format_value,
                 position_value, text_value, com_value, flag_del, use_r7=False):
        super().__init__()
        self._is_running = True
        self.rg2 = rg2
        self.path_grf = path_grf
        self.path_sch = path_sch
        self.path_wrd = path_wrd
        self.format_value = format_value
        self.position_value = position_value
        self.text_value = text_value
        self.com_value = com_value
        self.flag_del = flag_del
        self.max_w = max_w
        self.use_r7 = use_r7

    def run(self):
        """Основной метод, выполняющийся в отдельном потоке."""
        start_time = time.time()
        del_files = []

        with ProcessPoolExecutor(max_workers=self.max_w) as executor:
            futures = []
            for fn in self.rg2:
                f = executor.submit(_create_dxf_safe, fn, self.path_grf, self.com_value, self.path_sch)
                futures.append(f)
            for f in futures:
                f.result()

        if self.com_value == 'PNG':
            process_directory(self.rg2, self.path_wrd, self.format_value,
                             self.position_value, self.text_value)
        elif self.use_r7 and R7_AVAILABLE:
            process_directory_r7(self.rg2, self.path_wrd, self.format_value,
                                self.position_value, self.text_value,
                                image_format='svg')
        else:
            process_directory_com(self.rg2, self.path_wrd, self.format_value,
                                 self.position_value, self.text_value)

        if self.flag_del:
            for filename in self.rg2:
                directory = os_path.dirname(filename)
                base_path = os_path.splitext(filename)[0]
                base_name = os_path.basename(base_path)
                for suffix in ('.dxf', '.png', '.svg'):
                    item = Path(os_path.join(directory, base_name + suffix))
                    if item.exists():
                        try:
                            item.unlink()
                            del_files.append(str(item))
                        except Exception:
                            pass
            if del_files:
                log_dir = os_path.dirname(self.rg2[0])
                with open(os_path.join(log_dir, 'log.txt'), 'w', encoding='utf-8') as log_file:
                    log_file.write('Удалены следующие файлы:\n' + '\n'.join(del_files))

        self._is_running = False
        end_time = time.time()
        execution_time = (end_time - start_time) / 60
        try:
            self.finished_status.emit(execution_time)
        except Exception as e:
            logger.exception('Ошибка при завершении: %s', e)

    def stop(self):
        """Безопасная остановка потока."""
        self._is_running = False
        self.quit()
        self.wait(5000)


class FileOrDirectorySelector(pq.QWidget):
    """Главное окно выбора файлов и параметров."""

    def __init__(self):
        super().__init__()
        self.flag_del = False
        self.format_value = 'A4'
        self.position_value = 'Книжная'
        self.add_value = False
        self.com_value = 'SVG'
        self.text_value = False
        self.use_r7 = False  # True = Р7-Офис вместо Microsoft Word
        self.f_p = []
        self.path_grf = None
        self.path_sch = None
        self.path_wrd = None
        self.num_cpu = self.num_cpu()
        self.cpu_number = str(self.num_cpu)
        self.worker = None
        self.init_ui()

    def init_ui(self):
        self.main_layout = pq.QVBoxLayout()
        self.create_input_interface()
        self.create_loading_interface()
        self.show_input_interface()
        self.setLayout(self.main_layout)
        self.resize(500, 300)

    def num_cpu(self, max_w=16):
        return min(cpu_count() or 1, max_w)

    def create_input_interface(self):
        self.input_widget = pq.QWidget()
        layout = pq.QVBoxLayout()

        self.del_checkbox = pq.QCheckBox('Удалить промежуточные файлы (*.dxf,*.png)')
        self.del_checkbox.toggled.connect(self.on_checkbox_toggled_del)

        self.add_checkbox = pq.QCheckBox('Изображения приложения')
        self.add_checkbox.stateChanged.connect(self.on_checkbox_toggled_add)

        h_layout1 = pq.QHBoxLayout()
        h_layout1.addWidget(self.del_checkbox)
        h_layout1.addWidget(self.add_checkbox)
        layout.addLayout(h_layout1)

        v_layout_1 = pq.QVBoxLayout()
        label_format = pq.QLabel('Формат листа:')
        self.combo1 = pq.QComboBox()
        self.combo1.addItems(['A4', 'A3'])
        self.combo1.currentTextChanged.connect(self.on_format_changed)
        v_layout_1.addWidget(label_format)
        v_layout_1.addWidget(self.combo1)

        label_position = pq.QLabel('Ориентация листа:')
        self.combo2 = pq.QComboBox()
        self.combo2.addItems(['Книжная', 'Альбомная'])
        self.combo2.currentTextChanged.connect(self.on_position_changed)
        v_layout_1.addWidget(label_position)
        v_layout_1.addWidget(self.combo2)
        layout.addLayout(v_layout_1)

        v_layout_4 = pq.QVBoxLayout()
        self.label_cpu_number = pq.QLabel('Количество процессов:')
        self.path_entry_cpu_number = pq.QLineEdit()
        self.path_entry_cpu_number.setText(str(self.cpu_number))
        self.path_entry_cpu_number.setReadOnly(False)
        self.path_entry_cpu_number.editingFinished.connect(self.update_variable)
        v_layout_4.addWidget(self.label_cpu_number)
        v_layout_4.addWidget(self.path_entry_cpu_number)
        layout.addLayout(v_layout_4)

        v_layout_2 = pq.QVBoxLayout()
        label_com = pq.QLabel('Формат рисунка:')
        self.combo3 = pq.QComboBox()
        self.combo3.addItems(['SVG', 'PNG'])
        self.combo3.currentTextChanged.connect(self.COM_changed)
        v_layout_2.addWidget(label_com)
        v_layout_2.addWidget(self.combo3)
        layout.addLayout(v_layout_2)

        self.r7_checkbox = pq.QCheckBox('Р7-Офис (вместо Microsoft Word)')
        self.r7_checkbox.setEnabled(R7_AVAILABLE)
        if not R7_AVAILABLE:
            self.r7_checkbox.setToolTip('Модуль Click_word_r7 не найден')
        self.r7_checkbox.toggled.connect(self.on_r7_toggled)
        layout.addWidget(self.r7_checkbox)

        label_wrd = pq.QLabel('Шаблон *.docx:')
        self.path_entry_wrd = pq.QLineEdit()
        select_button_wrd = pq.QPushButton('Выбрать_файл *.docx')
        select_button_wrd.clicked.connect(self.on_select_button_clicked_wrd)
        layout.addWidget(label_wrd)
        layout.addWidget(self.path_entry_wrd)
        layout.addWidget(select_button_wrd)

        label_rg2 = pq.QLabel('Файлы *.rg2:')
        self.path_entry = pq.QTextEdit()
        select_button_1 = pq.QPushButton('Выбрать файлы')
        select_button_1.clicked.connect(self.on_select_button_clicked_rg2)
        select_button_2 = pq.QPushButton('Выбрать папку')
        select_button_2.clicked.connect(self.on_select_button_clicked_dir)
        layout.addWidget(label_rg2)
        layout.addWidget(self.path_entry)
        layout.addWidget(select_button_1)
        layout.addWidget(select_button_2)

        label_grf = pq.QLabel('Файл *.grf:')
        self.path_entry_grf = pq.QLineEdit()
        select_button_grf = pq.QPushButton('Выбрать grf')
        select_button_grf.clicked.connect(self.on_select_button_clicked_grf)
        layout.addWidget(label_grf)
        layout.addWidget(self.path_entry_grf)
        layout.addWidget(select_button_grf)

        label_sch = pq.QLabel('Файл *.sch:')
        self.path_entry_sch = pq.QLineEdit()
        select_button_sch = pq.QPushButton('Выбрать sch')
        select_button_sch.clicked.connect(self.on_select_button_clicked_sch)
        layout.addWidget(label_sch)
        layout.addWidget(self.path_entry_sch)
        layout.addWidget(select_button_sch)

        select_button_start = pq.QPushButton('Старт!')
        select_button_start.clicked.connect(self.on_select_button_clicked_start)
        layout.addWidget(select_button_start)

        self.input_widget.setLayout(layout)
        self.setWindowTitle('Click by Guriev')
        icon_path = get_icon_path()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))

    def update_variable(self):
        text = self.path_entry_cpu_number.text()
        try:
            self.cpu_number = int(text)
        except ValueError:
            logger.warning('Некорректное значение количества процессов')
            self.path_entry_cpu_number.setText(str(self.num_cpu))

    def on_select_button_clicked_dir(self):
        dir_path = pq.QFileDialog.getExistingDirectory(self, 'Выберите папку')
        if dir_path:
            file_paths = [os_path.join(dir_path, filename)
                         for filename in listdir(dir_path)
                         if filename.endswith('.rg2')]
            self.path_entry.setPlainText('\n'.join(file_paths))
            self.f_p = file_paths

    def on_select_button_clicked_rg2(self):
        file_paths, _ = pq.QFileDialog.getOpenFileNames(
            self, 'Выберите файлы', '', 'Все файлы (*.rg2)')
        if file_paths:
            self.path_entry.setPlainText('\n'.join(file_paths))
            self.f_p = file_paths

    def on_select_button_clicked_grf(self):
        file_path, _ = pq.QFileDialog.getOpenFileName(
            self, 'Выберите файл', '', 'grf файлы (*.grf)')
        if file_path:
            self.path_entry_grf.setText(file_path)
            self.path_grf = file_path

    def on_select_button_clicked_wrd(self):
        file_path, _ = pq.QFileDialog.getOpenFileName(
            self, 'Выберите файл', '', 'grf файлы (*.docx)')
        if file_path:
            self.path_entry_wrd.setText(file_path)
            self.path_wrd = file_path

    def on_select_button_clicked_sch(self):
        file_path, _ = pq.QFileDialog.getOpenFileName(
            self, 'Выберите файл', '', 'sch файлы (*.sch)')
        if file_path:
            self.path_entry_sch.setText(file_path)
            self.path_sch = file_path

    def on_select_button_clicked_start(self):
        if not self.f_p or not self.path_grf or not self.path_wrd:
            return
        self.show_loading_interface()
        self.worker = HeavyWorker(
            self.cpu_number, self.f_p, self.path_grf, self.path_sch,
            self.path_wrd, self.format_value, self.position_value,
            self.text_value, self.com_value, self.flag_del, self.use_r7)
        self.worker.finished_status.connect(self.done)
        self.worker.start()

    def done(self, execution_time):
        self.anim_timer.stop()
        self.status_label.setText('✅ Готово!')
        self.animation_label.close()
        msg = MSG(execution_time)
        msg.exec_()

    def on_checkbox_toggled_del(self, state):
        self.flag_del = bool(state)

    def on_format_changed(self, text):
        self.format_value = text if text in ('A4', 'A3') else 'A4'

    def on_position_changed(self, text):
        self.position_value = text if text in ('Книжная', 'Альбомная') else 'Книжная'

    def COM_changed(self, text):
        self.com_value = text if text in ('PNG', 'SVG') else 'SVG'

    def on_checkbox_toggled_add(self, state):
        self.text_value = (state == 2)

    def on_r7_toggled(self, state):
        self.use_r7 = bool(state)

    def create_loading_interface(self):
        self.loading_widget = pq.QWidget()
        layout = pq.QVBoxLayout()
        self.animation_label = pq.QLabel('⭕')
        self.animation_label.setAlignment(Qt.AlignCenter)
        self.animation_label.setStyleSheet('font-size: 100px;')
        self.status_label = pq.QLabel('Работаем...')
        layout.addStretch()
        layout.addWidget(self.animation_label)
        layout.addWidget(self.status_label)
        layout.addStretch()
        self.loading_widget.setLayout(layout)
        self.animation_steps = ['⭕', '⏳', '⌛', '🔵', '🔶', '🟢']
        self.anim_index = 0
        self.anim_timer = QTimer(self)
        self.anim_timer.timeout.connect(self.animate)

    def animate(self):
        self.anim_index = (self.anim_index + 1) % len(self.animation_steps)
        self.animation_label.setText(self.animation_steps[self.anim_index])

    def show_input_interface(self):
        self.clear_layout()
        self.main_layout.addWidget(self.input_widget)

    def show_loading_interface(self):
        self.clear_layout()
        self.main_layout.addWidget(self.loading_widget)
        self.anim_timer.start(200)

    def clear_layout(self):
        for i in reversed(range(self.main_layout.count())):
            item = self.main_layout.itemAt(i)
            if item and item.widget():
                item.widget().hide()
            self.main_layout.removeItem(item)


class MSG(pq.QMessageBox):
    """Окно сообщения о завершении работы."""

    def __init__(self, ex_time):
        super().__init__()
        self.execution_time = ex_time
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Click by Guriev')
        icon_path = get_icon_path()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))
        self.setText(f'Работа макроса завершена!\nВремя выполнения: {self.execution_time:.4f} минут')
        self.setIcon(pq.QMessageBox.Information)
        self.setStandardButtons(pq.QMessageBox.Ok)
