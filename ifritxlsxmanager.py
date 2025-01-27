import argparse
import glob
import os
import pathlib
import re
import shutil
import sys

from PyQt6.QtWidgets import QApplication

from . import xlsxmanager
from .ennemy import Ennemy
from .gamedata import GameData
from .xlsxmanager import XlsxToDat, DatToXlsx


class IfritXlsxManager:
    def __init__(self):
        self._dat_xlsx_manager = DatToXlsx()
        self._xlsx_to_dat_manager = XlsxToDat()

    def create_file(self, xlsx_file):
        self._dat_xlsx_manager.create_file(xlsx_file)

    def load_file(self, xlsx_file):
        self._xlsx_to_dat_manager.load_file(xlsx_file)

    def dat_to_xlsx(self, file_list, analyse_ai=False):
        print("Getting game data")
        game_data = GameData()
        game_data.load_all()

        print("Reading ennemy files")
        for monster_file in file_list:
            file_name = os.path.basename(monster_file)
            file_index = int(re.search(r'\d{3}', file_name).group())
            if file_index == 0 or file_index == 127 or file_index > 143:  # Avoid working on garbage file
                continue
            print("Reading file {}".format(file_name))
            monster = Ennemy(game_data)
            monster.load_file_data(monster_file, game_data)
            monster.analyse_loaded_data(game_data, analyse_ai)

            print("Writing to xlsx file")
            self._dat_xlsx_manager.export_to_xlsx(monster, file_name, game_data, analyse_ai)

        self._dat_xlsx_manager.create_ref_data(game_data)

    def xlsx_to_dat(self, file_list, local_limit, analyse_ai=False):
        game_data = GameData()
        game_data.load_all()
        for sheet in self._xlsx_to_dat_manager.workbook:
            if sheet.title != xlsxmanager.REF_DATA_SHEET_TITLE:
                monster_index = int(re.search(r'\d+', sheet.title).group())
                if local_limit > 0 and local_limit != monster_index:  # Only doing the monster asked
                    continue
                else:
                    current_dat_file = [pathlib.Path(text).resolve().parent for text in file_list if int(pathlib.Path(text).name.replace('c0m','').replace('.dat', '')) == monster_index]
                    if current_dat_file:
                        current_dat_file = current_dat_file[0]
                    else:
                        print(f"Unexpected monster index: {monster_index}")
                        return
                ennemy = self._xlsx_to_dat_manager.import_from_xlsx(sheet, game_data, current_dat_file, local_limit, analyse_ai)
                if ennemy:
                    ennemy.write_data_to_file(game_data, current_dat_file, analyse_ai)