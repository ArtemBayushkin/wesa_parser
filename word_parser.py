# word_parser.py
import os
import re
from shutil import rmtree
from tempfile import mkdtemp
from zipfile import ZipFile
from lxml import etree as ET
import logging


class WordProcessor:
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
                    self.logger.log(logging.DEBUG, f"Ошибка загрузки правила '{rule_name}': {e}")
        except Exception as e:
            self.logger.log(logging.DEBUG, f"Ошибка обработки rules: {e}")
        if not patterns:
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
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

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

        for parent in tree.findall('.//w:p', namespaces=nsmap) + tree.findall('.//w:sdtContent', namespaces=nsmap):
            texts = parent.findall('.//w:t', namespaces=nsmap)
            if len(texts) < 2:
                continue

            full_text = ''.join(t.text or '' for t in texts)
            new_full_text = self._apply_replacements(full_text)

            if new_full_text != full_text:
                modified = True
                idx = 0
                for t in texts:
                    if t.text is not None:
                        part_len = len(t.text)
                        t.text = new_full_text[idx:idx + part_len]
                        idx += part_len

        for p in tree.findall('.//w:p', namespaces=nsmap):
            para_texts = ''.join(t.text or '' for t in p.findall('.//w:t', namespaces=nsmap)).strip()
            if re.search(r'Лист\s+регистрации\s+изменений|Record\s+of\s+revisions', para_texts, re.IGNORECASE):
                tbl = p.getnext()
                while tbl is not None and tbl.tag != '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl':
                    tbl = tbl.getnext()
                if tbl is not None:
                    self.logger.log(logging.DEBUG, 
                         "Найдена таблица 'Лист регистрации изменений' или 'Record of revisions'. Очистка данных в столбцах.")
                    rows = tbl.findall('w:tr', namespaces=nsmap)
                    if len(rows) > 1:
                        for row in rows[2:]:
                            cells = row.findall('w:tc', namespaces=nsmap)
                            for cell in cells:
                                for t in cell.findall('.//w:t', namespaces=nsmap):
                                    if t.text and t.text.strip():
                                        self.logger.log(logging.DEBUG, f"Очистка текста в ячейке: '{t.text.strip()}' → ''")
                                        t.text = ''
                            modified = True
                    else:
                        self.logger.log(logging.DEBUG, "Таблица найдена, но не содержит строк с данными для очистки.")

        return modified

    def process_file(self, input_path, output_path):
        tmp_dir = mkdtemp()
        self.logger.log(logging.DEBUG, f"Открыт файл: {input_path}")
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