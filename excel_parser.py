import os
import re
from shutil import rmtree
from tempfile import mkdtemp
from zipfile import ZipFile
from lxml import etree as ET
import logging


try:
    import win32com.client as win32
except ImportError:
    win32 = None

class ExcelProcessor:
    def __init__(self, replacement_digit, project, rules, logger=None):
        self.replacement_digit = str(replacement_digit)
        self.logger = logger or logging.getLogger()
        self.patterns = self._load_patterns(rules)

    def _load_patterns(self, rules):
        patterns = []
        try:
            for rule_name, rule in rules.items():
                try:
                    pattern = eval(rule["pattern"], {"re": re})
                    replacement = eval(rule["replacement"], {"self": self})
                    patterns.append((pattern, replacement))
                    self.logger.log(logging.DEBUG, f"Загружено правило '{rule_name}'")
                except Exception as e:
                    pass
                self.logger.log(logging.DEBUG, f"Ошибка загрузки правила '{rule_name}': {e}")
        except Exception as e:
            pass
            self.logger.log(logging.DEBUG, f"Ошибка обработки rules: {e}")
        if not patterns:
            pass
            self.logger.log(logging.DEBUG, "Предупреждение: Нет patterns для этого парсера")
        return patterns

    def _apply_replacements(self, text):
        if text is None:
            return None
        original_text = text
        for pattern, repl in self.patterns:
            text = pattern.sub(repl, text)
        if text != original_text:
            self.logger.log(logging.DEBUG, f"Замена текста: '{original_text}' → '{text}'")
        return text

    def _process_xml_tree(self, tree):
        modified = False
        nsmap = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for elem in tree.iter():
            if elem.text:
                new_text = self._apply_replacements(elem.text)
                if new_text != elem.text:
                    elem.text = new_text
                    modified = True
            if elem.tail:
                new_tail = self._apply_replacements(elem.tail)
                if new_tail != elem.tail:
                    elem.tail = new_tail
                    modified = True
        return modified

    def process_file(self, input_path, output_path):
        tmp_dir = mkdtemp()
        self.logger.log(logging.DEBUG, f"Открыт файл: {input_path}")
        modified_files = set()
        converted = False
        temp_input = None

        try:
            if input_path.lower().endswith('.xls'):
                self.logger.log(logging.DEBUG, f"Обнаружен .xls файл. Конвертируем в .xlsm...")
                temp_input = os.path.join(tmp_dir, 'converted.xlsm')

                if win32:
                    excel = win32.Dispatch('Excel.Application')
                    excel.Visible = False
                    wb = excel.Workbooks.Open(os.path.abspath(input_path))
                    wb.SaveAs(os.path.abspath(temp_input), FileFormat=52)  # 52 = xlsm
                    wb.Close()
                    excel.Quit()
                    self.logger.log(logging.DEBUG, f"Конвертация завершена: {temp_input}")
                else:
                    raise ImportError(
                        "pywin32 не установлен. Установите 'pip install pywin32' для конвертации на Windows.")

                input_path = temp_input
                converted = True

            with ZipFile(input_path) as zip_in:
                filenames = zip_in.namelist()
                zip_in.extractall(tmp_dir)

            target_files = ['xl/sharedStrings.xml']
            target_files += [f for f in filenames if f.startswith('xl/worksheets/sheet')]

            for fname in target_files:
                full_path = os.path.join(tmp_dir, fname)
                if not os.path.exists(full_path) or os.path.getsize(full_path) == 0:
                    self.logger.log(logging.DEBUG, f"Пропущен файл (отсутствует или пуст): {fname}")
                    continue
                try:
                    parser = ET.XMLParser(remove_blank_text=True)
                    tree = ET.parse(full_path, parser)
                    modified = self._process_xml_tree(tree)
                    if modified:
                        tree.write(full_path, encoding='UTF-8', xml_declaration=True, pretty_print=True)
                        modified_files.add(fname)
                        self.logger.log(logging.DEBUG, f"Файл изменен: {fname}")
                except ET.XMLSyntaxError as e:
                    self.logger.log(logging.DEBUG, f"Ошибка XML в {fname}: {e}")

            with ZipFile(output_path, 'w') as zip_out:
                for fname in filenames:
                    zip_out.write(os.path.join(tmp_dir, fname), fname)

            self.logger.log(logging.DEBUG, f"Файл успешно обработан: {output_path}")
            return True

        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка обработки {input_path}: {str(e)}")
            return False

        finally:
            rmtree(tmp_dir, ignore_errors=True)