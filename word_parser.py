import os
import re
from shutil import rmtree
from tempfile import mkdtemp
from zipfile import ZipFile
from lxml import etree as ET
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('WordProcessor')

class WordProcessor:
    def __init__(self, replacement_digit, log_callback=None, debug=False):
        self.replacement_digit = str(replacement_digit)
        self.debug = debug  # Флаг отладки
        self.log = log_callback or (lambda msg: None)  # Если колбэк не передан — молчит
        self._log(f"Инициализация WordProcessor с цифрой: {self.replacement_digit}")

        self.patterns = [
            # ED.D.*  — меняем только последнюю цифру
            (re.compile(r'\b(ED\.D\.[A-Z]\d\d\d\.)\d\b'),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),

            # 10UKD  — меняем только первую цифру, буквы сохраняем
            (re.compile(r'\b([0-9])0([A-Z]{3})\b', flags=re.IGNORECASE),
             lambda m: f"{self.replacement_digit}0{m.group(2)}"),

            # C02 -> C01 (если оставляем как раньше)
            (re.compile(r'C0[2-9]\b'),
             'C01'),
            # Замена Блока / Unit
            (re.compile(r'(Unit\s*)\d\b', flags=re.IGNORECASE),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
            (re.compile(r'(блока №\s*)\d\b', flags=re.IGNORECASE),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
        ]

        if self.replacement_digit in ("3", "4"):
            self.patterns.append(
                (re.compile(r'\bED\.B\.P000\.S\b'),
                 "ED.B.P000.W")
            )

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
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # --- 1. Проход по всем узлам ---
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

        # --- 2. Дополнительный проход — для случаев, когда "C0" и цифра разделены ---
        for parent in tree.findall('.//w:p', namespaces=nsmap) + tree.findall('.//w:sdtContent', namespaces=nsmap):
            texts = parent.findall('.//w:t', namespaces=nsmap)
            if len(texts) < 2:
                continue

            # Склеиваем текст всех <w:t> в этом контейнере
            full_text = ''.join(t.text or '' for t in texts)
            new_full_text = self._apply_replacements(full_text)

            if new_full_text != full_text:
                modified = True
                # Распределяем обратно по тем же <w:t>
                idx = 0
                for t in texts:
                    part_len = len(t.text or '')
                    t.text = new_full_text[idx:idx + part_len]
                    idx += part_len

        # --- 3. Новый блок: Очистка текста в столбцах таблицы "Лист регистрации изменений" или "Record of revisions" ---
        for p in tree.findall('.//w:p', namespaces=nsmap):
            para_texts = ''.join(t.text or '' for t in p.findall('.//w:t', namespaces=nsmap)).strip()
            if re.search(r'Лист\s+регистрации\s+изменений|Record\s+of\s+revisions', para_texts, re.IGNORECASE):
                # Находим следующую таблицу после параграфа
                tbl = p.getnext()
                while tbl is not None and tbl.tag != '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl':
                    tbl = tbl.getnext()
                if tbl is not None:
                    self._log("Найдена таблица 'Лист регистрации изменений' или 'Record of revisions'. Очистка данных в столбцах.")
                    rows = tbl.findall('w:tr', namespaces=nsmap)
                    if len(rows) > 1:
                        for row in rows[2:]:  # Обрабатываем только строки данных, пропуская заголовок (первая строка)
                            cells = row.findall('w:tc', namespaces=nsmap)
                            for cell in cells:
                                for t in cell.findall('.//w:t', namespaces=nsmap):
                                    if t.text and t.text.strip():
                                        self._log(f"Очистка текста в ячейке: '{t.text.strip()}' → ''")
                                        t.text = ''
                            modified = True
                    else:
                        self._log("Таблица найдена, но не содержит строк с данными для очистки.")

        return modified

    def process_file(self, input_path, output_path):
        tmp_dir = mkdtemp()
        self._log(f"Открыт файл: {input_path}")
        modified_files = set()

        try:
            with ZipFile(input_path) as zip_in:
                filenames = zip_in.namelist()
                zip_in.extractall(tmp_dir)

            target_files = ['word/document.xml', 'docProps/core.xml']
            target_files += [f for f in filenames if f.startswith('word/header') or f.startswith('word/footer')]

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