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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('ExcelProcessor')


class ExcelProcessor:
    def __init__(self, replacement_digit, log_callback=None, debug=False):
        self.replacement_digit = str(replacement_digit)
        self.debug = debug  # Флаг отладки
        self.log = log_callback or (lambda msg: None)  # Если колбэк не передан — молчит
        self._log(f"Инициализация ExcelProcessor с цифрой: {self.replacement_digit}")
        self.patterns = [
            # ED.D.* — меняем только последнюю цифру
            (re.compile(r'\b(ED\.D\.[A-Z]\d\d\d\.)\d'),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
            # 10UKD — меняем только первую цифру, буквы сохраняем
            (re.compile(r'\b([0-9])(0[A-Z]{3})', flags=re.IGNORECASE),
             lambda m: f"{self.replacement_digit}{m.group(2)}"),
            # C02 -> C01 (если оставляем как раньше)
            (re.compile(r'&R&11C0[2-9]\b'),
             '&R&11C01'),
            # C02 -> C01 (если оставляем как раньше)
            (re.compile(r'&RC0[2-9]\b'),
             '&RC01'),
            # Для нижнего колонтитула - шифра
            (re.compile(r'((?:&[LCR](?:&\d{2})?)?ED\.D\.[A-Z]\d\d\d\.)\d'),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
        ]

    def _log(self, message):
        # Логи, которые всегда записываются
        always_log = (
                message.startswith("Успешно: ") or
                message.startswith("Ошибка обработки: ") or
                message.startswith("Критическая ошибка ") or
                message.startswith("Файлы не найдены.") or
                message.startswith("Пропуск ") or
                message.startswith("Файл успешно обработан: ")
        )
        # Если отладка включена или это обязательный лог, вызываем callback
        if self.debug or always_log:
            self.log(message)

    def _apply_replacements(self, text):
        if text is None:
            return None
        original_text = text
        for pattern, repl in self.patterns:
            text = pattern.sub(repl, text)
        if text != original_text:
            self._log(f"Замена текста: '{original_text}' → '{text}'")
        return text

    def _process_xml_tree(self, tree):
        modified = False
        nsmap = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        # --- Проход по всем узлам ---
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
        self._log(f"Открыт файл: {input_path}")
        modified_files = set()
        converted = False
        temp_input = None

        try:
            # Проверка и конвертация .xls в .xlsm
            if input_path.lower().endswith('.xls'):
                self._log(f"Обнаружен .xls файл. Конвертируем в .xlsm...")
                temp_input = os.path.join(tmp_dir, 'converted.xlsm')

                # Вариант 1: Через pywin32 и Excel (Windows)
                if win32:
                    excel = win32.Dispatch('Excel.Application')
                    excel.Visible = False
                    wb = excel.Workbooks.Open(os.path.abspath(input_path))
                    wb.SaveAs(os.path.abspath(temp_input), FileFormat=52)  # 52 = xlsm
                    wb.Close()
                    excel.Quit()
                    self._log(f"Конвертация завершена: {temp_input}")
                else:
                    raise ImportError(
                        "pywin32 не установлен. Установите 'pip install pywin32' для конвертации на Windows.")



                input_path = temp_input  # Теперь обрабатываем конвертированный файл
                converted = True

            # Основная обработка (как раньше)
            with ZipFile(input_path) as zip_in:
                filenames = zip_in.namelist()
                zip_in.extractall(tmp_dir)

            target_files = ['xl/sharedStrings.xml']
            target_files += [f for f in filenames if f.startswith('xl/worksheets/sheet')]

            for fname in target_files:
                full_path = os.path.join(tmp_dir, fname)
                if not os.path.exists(full_path) or os.path.getsize(full_path) == 0:
                    self._log(f"Пропущен файл (отсутствует или пуст): {fname}")
                    continue
                try:
                    parser = ET.XMLParser(remove_blank_text=True)
                    tree = ET.parse(full_path, parser)
                    modified = self._process_xml_tree(tree)
                    if modified:
                        tree.write(full_path, encoding='UTF-8', xml_declaration=True, pretty_print=True)
                        modified_files.add(fname)
                        self._log(f"Файл изменен: {fname}")
                except ET.XMLSyntaxError as e:
                    self._log(f"Ошибка XML в {fname}: {e}")

            with ZipFile(output_path, 'w') as zip_out:
                for fname in filenames:
                    zip_out.write(os.path.join(tmp_dir, fname), fname)

            self._log(f"Файл успешно обработан: {output_path}")
            return True

        except Exception as e:
            self._log(f"Ошибка обработки {input_path}: {str(e)}")
            return False

        finally:
            rmtree(tmp_dir, ignore_errors=True)