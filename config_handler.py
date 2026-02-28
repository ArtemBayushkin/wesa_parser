import sys
import os
import shutil
import json


class ConfigHandler():
    def __init__(self):
        pass

    def get_relative_path(self, relative_path):
        """Возвращает абсолютный путь к ресурсу, учитывая упаковку PyInstaller.
        Если приложение упаковано (--onefile), использует _MEIPASS; иначе — текущую директорию.
        Args:
            relative_path (str): Относительный путь к ресурсу (например, 'config.json').

        Returns:
            str: Абсолютный путь к ресурсу.
        """
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.getcwd(), relative_path)    
    
    
    def load_config(self):
        config_path = self.get_relative_path('config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        else: 
            raise FileNotFoundError("No config file")   
    def get_projects(self):
        return list(self.load_config().keys())
    
    def edit_config(self):
        config_path = self.get_relative_path('config.json')
        if os.path.exists(config_path):
            shutil.copy(config_path, config_path)
            os.startfile(config_path)
            

if __name__ != '__main__':
    config_handler = ConfigHandler()
    config_data = config_handler.load_config()
    config_projects = config_handler.get_projects()
    