import sys

from PyQt6.QtCore import Qt, QCoreApplication
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QCheckBox, QComboBox, QLabel, QHBoxLayout, QLineEdit, QSpinBox


class WindowLauncher(QWidget):
    FF8_VERSION = ['PC 2000', 'Steam 2013', 'Steam Remaster']


    def __init__(self, ifrit_manager, list_lang, list_launch_option):
        QWidget.__init__(self)
        self.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.ifrit_manager = ifrit_manager
        self.setWindowTitle("Ifrit-XLSX Launcher")
        # self.setMinimumSize(300,200)

        self.lang_option = QComboBox()
        self.lang_option.addItems(list_lang)
        self.lang_option_label = QLabel("Language: ")

        self.ff8_version_option = QComboBox()
        self.ff8_version_option.addItems(self.FF8_VERSION)
        self.ff8_version_option_label = QLabel("Version: ")

        self.launch_option = QComboBox()
        self.launch_option.addItems(list_launch_option)
        self.launch_option_label = QLabel("Launch option: ")

        self.launch_button = QPushButton("Launch")
        self.launch_button.clicked.connect(self.launch_click)

        self.limit_option = QSpinBox()
        self.limit_option.setMaximum(200)
        self.limit_option.setMinimum(-1)
        self.limit_option.setValue(-1)
        self.limit_option_label = QLabel("File ID (-1 for all)")

        self.no_pack_option = QCheckBox("Don't pack when xlsx_to_dat")
        self.launch_option.currentTextChanged.connect(self.launch_option_changed)

        self.open_xlsx = QCheckBox("Open xlsx when finish")

        self.delete = QCheckBox("Delete dat files (only keep battle.fs)")
        self.delete.hide()

        self.analyse_ia = QCheckBox("Analyse IA")
        self.analyse_ia.setChecked(True)

        self.copy_option = QLineEdit()
        self.copy_option_label = QLabel("Copy battle.f* to given path (FF8 path).\n Let empty if not needed")
        self.copy_option.hide()
        self.copy_option_label.hide()

        self.general_layout = QVBoxLayout()
        self.ff8_version_layout = QHBoxLayout()
        self.lang_layout = QHBoxLayout()
        self.launch_option_layout = QHBoxLayout()
        self.copy_option_layout = QHBoxLayout()
        self.limit_option_layout = QHBoxLayout()
        self.setup_layout()
        self.show_window()

    def show_window(self):
        self.show()

    def setup_layout(self):
        self.ff8_version_layout.addWidget(self.ff8_version_option_label)
        self.ff8_version_layout.addWidget(self.ff8_version_option)
        self.lang_layout.addWidget(self.lang_option_label)
        self.lang_layout.addWidget(self.lang_option)
        self.launch_option_layout.addWidget(self.launch_option_label)
        self.launch_option_layout.addWidget(self.launch_option)
        self.limit_option_layout.addWidget(self.limit_option_label)
        self.limit_option_layout.addWidget(self.limit_option)
        self.copy_option_layout.addWidget(self.copy_option_label)
        self.copy_option_layout.addWidget(self.copy_option)

        self.general_layout.addLayout(self.ff8_version_layout)
        self.general_layout.addLayout(self.lang_layout)
        self.general_layout.addLayout(self.launch_option_layout)
        self.general_layout.addWidget(self.delete)
        self.general_layout.addWidget(self.no_pack_option)
        self.general_layout.addWidget(self.open_xlsx)
        self.general_layout.addLayout(self.limit_option_layout)
        self.general_layout.addLayout(self.copy_option_layout)
        self.general_layout.addWidget(self.analyse_ia)
        self.general_layout.addWidget(self.launch_button)

        self.setLayout(self.general_layout)

    def launch_click(self):
        if  self.ff8_version_option.currentIndex() == 2:
            remaster = True
        else:
            remaster = False
        lang = str(self.lang_option.currentText())
        launch_option = str(self.launch_option.currentText())
        limit_option = self.limit_option.value()
        no_pack_option = self.no_pack_option.isChecked()
        open_xlsx_option = self.open_xlsx.isChecked()
        delete_option = self.delete.isChecked()
        copy_option = self.copy_option.text()
        analyse_ai_option = self.analyse_ia.isChecked()
        self.ifrit_manager.exec(lang, launch_option, limit_option, no_pack_option, open_xlsx_option, delete_option, copy_option, analyse_ai_option, remaster)
        # QCoreApplication.quit()

    def launch_option_changed(self):
        if str(self.launch_option.currentText()) == 'fs_to_xlsx':
            self.delete.hide()
            self.copy_option.hide()
            self.copy_option_label.hide()
        elif str(self.launch_option.currentText()) == 'xlsx_to_fs':
            self.delete.show()
            self.copy_option.show()
            self.copy_option_label.show()
        else:
            self.delete.show()
            self.copy_option.show()
            self.copy_option_label.show()
