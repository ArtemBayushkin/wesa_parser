import os
import re
import glob
import logging
from excel_parser import ExcelProcessor
from word_parser import WordProcessor
from dwg_parser import AutoCADProcessor
from config_handler import config_data


class FileHandler():
    def __init__(self, input_folder, project, replacement_digit, config_data=config_data, logger=None):
        self.project = project
        self.input_folder = input_folder
        self.output_folder = input_folder + "_processed"
        self.processed_files_counter = 0
        self.files = []
        self.replacement_digit = replacement_digit
        self.config_data = config_data
        self.logger = logger or logging.getLogger(__name__)
        os.makedirs(self.output_folder, exist_ok=True)

    
    def select_files(self):
        '''
        Метод для загрузки файлов различного расширения
        :param input_dir:
        :return:
        '''
        excel_files = glob.glob(os.path.join(self.input_folder, '*.xls*'))
        word_files = glob.glob(os.path.join(self.input_folder, '*.do*'))
        dwg_files = glob.glob(os.path.join(self.input_folder, '*.dwg'))
        sha_files = glob.glob(os.path.join(self.input_folder, '*.sha'))
        self.files = excel_files + word_files + dwg_files + sha_files
        return self.files

    def process_files(self):
        self.logger.log(logging.INFO, "Обработка файлов начата")
        sha_processor = None
        sha_app_started = False
        input_files = self.select_files()
        try:
            for input_path in input_files:
                try:
                    filename = os.path.basename(input_path)
                    name, ext = os.path.splitext(filename)
                    new_name = name
                    file_rename_rules = self.config_data.get(self.project, {}).get("file_rename", {})
                    for rule_name, rule in file_rename_rules.items():
                        try:
                            pattern = eval(rule["pattern"], {"re": re})
                            repl_str = rule["replacement"]
                            repl = eval(repl_str, {"replacement_digit": self.replacement_digit})
                            new_name = pattern.sub(repl, new_name)
                        except Exception as e:
                            pass
                    if new_name == name:
                        new_name = f"{name}"

                    if ext == ".xls":
                        output_path = os.path.join(self.output_folder, new_name + ".xlsm")
                    else:
                        output_path = os.path.join(self.output_folder, new_name + ext)
                    extension = ext.lower()

                    if extension in ('.doc', '.docx', '.dotx', 'dot'):
                        rules = self.config_data.get(self.project, {}).get("word_parser", {})
                        processor = WordProcessor(self.replacement_digit, self.project, rules, logger=self.logger)
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.logger.log(logging.INFO, f"Успешно: {filename}")
                            self.processed_files_counter += 1
                        else:
                            self.logger.log(logging.INFO, f"Ошибка обработки: {filename}")

                    elif extension in ('.xls', '.xlsx', '.xlsm'):
                        rules = self.config_data.get(self.project, {}).get("excel_parser", {})
                        processor = ExcelProcessor(self.replacement_digit, self.project, rules, logger=self.logger)
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.logger.log(logging.INFO, f"Успешно: {filename}")
                            self.processed_files_counter += 1
                        else:
                            self.logger.log(logging.INFO, f"Ошибка обработки: {filename}")

                    elif extension == '.dwg':
                        rules = self.config_data.get(self.project, {}).get("dwg_parser", {})
                        processor = AutoCADProcessor(self.replacement_digit, self.project, rules, logger=self.logger)
                        output_path = os.path.join(self.output_folder, new_name + ext)  # Используем new_name!
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.logger.log(logging.INFO, f"Успешно: {filename}")
                            self.processed_files_counter += 1
                        else:
                            self.logger.log(logging.INFO, f"Ошибка обработки: {filename}")

                    elif extension == '.sha':
                        rules = self.config_data.get(self.project, {}).get("sha_parser", {})
                        if rules:  # Проверяем, есть ли правила
                            if not sha_processor:
                                from sha_parser import ShaProcessorWinAPI
                                sha_processor = ShaProcessorWinAPI(self.replacement_digit, self.project, rules, logger=self.logger)
                            if not sha_app_started:
                                try:
                                    sha_processor.start_app()
                                    sha_app_started = True
                                except Exception as e:
                                    self.logger.log(logging.INFO, f"Ошибка запуска SmartSketch для {filename}: {str(e)}")
                                    continue
                            success = sha_processor.process_file(input_path, output_path)
                            if success:
                                self.logger.log(logging.INFO, f"Успешно: {filename}")
                                self.processed_files_counter += 1
                            else:
                                self.logger.log(logging.INFO, f"Ошибка обработки: {filename}")
                        else:
                            self.logger.log(logging.INFO, f"Пропуск {filename} (нет правил для sha_parser в config)")

                    else:
                        self.logger.log(logging.INFO, f"Пропуск {filename} (неподдерживаемый формат: {extension})")

                except Exception as e:
                    self.logger.log(logging.ERROR, f"Критическая ошибка {filename}: {str(e)}")

        finally:
            if sha_app_started and sha_processor:
                sha_processor.stop_app()

    
   