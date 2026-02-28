import os
import re
import time
import win32com.client
import pythoncom
import psutil  # Для завершения процессов
import logging

class AutoCADProcessor:
    def __init__(self, replacement_digit, project, rules, logger=None):
        pythoncom.CoInitialize()
        self.replacement_digit = str(replacement_digit)
        self.logger = logger or logging.getLogger()
        self.logger.log(logging.DEBUG, f"Инициализация AutoCADProcessor с цифрой: {self.replacement_digit} и проектом: {project}")
        self.patterns = self._load_patterns(rules)
        self.com_app = None
        self.com_doc = None
        self._initialize_autocad()

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
                    self.logger.log(logging.ERROR, f"Ошибка загрузки правила '{rule_name}': {e}")
        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка обработки rules: {e}")
        if not patterns:
            self.logger.log(logging.DEBUG, "Предупреждение: Нет patterns для этого парсера")
        return patterns

    def _initialize_autocad(self):
        retries = 3
        for attempt in range(retries):
            try:
                self._terminate_autocad()  # Очистка перед созданием нового экземпляра
                self.com_app = win32com.client.Dispatch("AutoCAD.Application")
                if self.wait_for_object_ready(self.com_app, timeout=20.0, check_type="app"):
                    self.logger.log(logging.DEBUG, "Экземпляр AutoCAD создан")
                    return
                else:
                    self.logger.log(logging.DEBUG, f"Экземпляр AutoCAD не готов на попытке {attempt + 1}")
            except Exception as e:
                self.logger.log(logging.ERROR, f"Не удалось создать экземпляр AutoCAD на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)  # Увеличенная задержка
                else:
                    raise Exception(f"Не удалось создать экземпляр AutoCAD после {retries} попыток: {e}")
            finally:
                pythoncom.CoUninitialize()
                pythoncom.CoInitialize()

    def wait_for_object_ready(self, obj, timeout=20.0, check_type="app"):
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                pythoncom.PumpWaitingMessages()
                if obj is not None:
                    if check_type == "app":
                        _ = obj.Version
                    else:
                        _ = obj.Name
                    return True
            except Exception as e:
                self.logger.log(logging.ERROR, f"Ошибка проверки готовности объекта ({check_type}): {e}")
            time.sleep(0.2)
        self.logger.log(logging.DEBUG, f"Объект ({check_type}) не готов после {timeout} секунд")
        return False

    def _terminate_autocad(self):
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower().startswith('acad'):
                    proc.kill()
                    self.logger.log(logging.DEBUG, "Процесс AutoCAD завершен")
            time.sleep(1)
        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка при завершении процесса AutoCAD: {e}")
        self.com_app = None
        self.com_doc = None

    def _apply_replacements(self, text):
        if not text:
            return text
        original = text
        new_text = text
        for pattern, repl in self.patterns:
            if pattern.search(new_text):
                new_text = pattern.sub(repl, new_text) if callable(repl) else pattern.sub(repl, new_text)
        if new_text != original:
            self.logger.log(logging.DEBUG, f"Замена: {original} → {new_text}")
        return new_text

    def _process_entity(self, entity, depth=0, location=""):
        retries = 3
        for attempt in range(retries):
            try:
                if not hasattr(entity, 'ObjectName'):
                    self.logger.log(logging.DEBUG, f"Объект в {location} не имеет ObjectName, пропуск")
                    return
                etype = entity.ObjectName
                self.logger.log(logging.DEBUG, f"Обработка объекта {etype} в {location}")
                if etype in ("AcDbText", "AcDbMText"):
                    try:
                        txt = entity.TextString
                        new_txt = self._apply_replacements(txt)
                        if new_txt != txt:
                            entity.TextString = new_txt
                            self.logger.log(logging.DEBUG, f"Замена в {location}: {txt} → {new_txt}")
                    except Exception as e:
                        self.logger.log(logging.DEBUG, f"Ошибка обработки текста в {location}: {e}")
                elif etype == "AcDbMLeader":
                    try:
                        txt = entity.TextString
                        new_txt = self._apply_replacements(txt)
                        if new_txt != txt:
                            entity.TextString = new_txt
                            self.logger.log(logging.DEBUG, f"Замена в {location} (MLeader): {txt} → {new_txt}")
                    except Exception as e:
                        self.logger.log(logging.DEBUG, f"Ошибка обработки MLeader в {location}: {e}")
                elif etype == "AcDbBlockReference" and hasattr(entity, "GetAttributes"):
                    try:
                        attributes = entity.GetAttributes()
                        for attr in attributes:
                            try:
                                txt = attr.TextString
                                new_txt = self._apply_replacements(txt)
                                if new_txt != txt:
                                    attr.TextString = new_txt
                                    self.logger.log(logging.DEBUG, f"Замена в атрибуте блока {location}: {txt} → {new_txt}")
                            except Exception as e:
                                self.logger.log(logging.DEBUG, f"Ошибка обработки атрибута в {location}: {e}")
                                continue
                    except Exception as e:
                        self.logger.log(logging.ERROR, f"Ошибка доступа к атрибутам блока в {location}: {e}")
                return
            except Exception as e:
                self.logger.log(logging.DEBUG, f"Ошибка объекта в {location} на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)
                else:
                    self.logger.log(logging.DEBUG, f"Не удалось обработать объект в {location} после {retries} попыток: {e}")
                    return

    def _process_blocks(self):
        retries = 3
        for attempt in range(retries):
            try:
                if self.com_doc is None:
                    self.logger.log(logging.DEBUG, "Документ не инициализирован, пропуск обработки блоков")
                    return
                block_table = self.com_doc.Blocks
                for block in block_table:
                    if not block.IsLayout and not block.IsXRef:
                        self.logger.log(logging.DEBUG, f"Обработка блока: {block.Name}")
                        try:
                            for entity in block:
                                self._process_entity(entity, depth=1, location=f"block {block.Name}")
                        except Exception as e:
                            self.logger.log(logging.DEBUG, f"Пропуск блока {block.Name} из-за ошибки: {e}")
                            continue
                return
            except Exception as e:
                self.logger.log(logging.DEBUG, f"Ошибка обработки блоков на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)
                    try:
                        if self.com_doc is not None:
                            self.com_doc.Close(False)  # Отклонить изменения
                            self.com_doc = None
                        if self.com_app is not None:
                            self.com_app.Quit()
                            self.com_app = None
                        self._initialize_autocad()
                    except Exception as reinf_err:
                        self.logger.log(logging.DEBUG, f"Не удалось переинициализировать AutoCAD: {reinf_err}")
                else:
                    self.logger.log(logging.DEBUG, f"Не удалось обработать блоки после {retries} попыток: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    return

    def _process_all_entities(self):
        retries = 3
        for attempt in range(retries):
            try:
                if self.com_doc is None:
                    self._log("Документ не инициализирован, пропуск обработки")
                    return False
                self.logger.log(logging.DEBUG, "Обработка ModelSpace...")
                for entity in self.com_doc.ModelSpace:
                    self._process_entity(entity, location="ModelSpace")
                self.logger.log(logging.DEBUG, "Обработка блоков...")
                self._process_blocks()
                self.logger.log(logging.DEBUG, "Обработка листов...")
                for layout in self.com_doc.Layouts:
                    if layout.Name.lower() in ['model', 'модель']:
                        continue
                    self.logger.log(logging.DEBUG, f"Лист: {layout.Name}")
                    try:
                        for entity in layout.Block:
                            self._process_entity(entity, location=f"Layout {layout.Name}")
                    except Exception as e:
                        self.logger.log(logging.DEBUG, f"Пропуск листа {layout.Name} из-за ошибки: {e}")
                        continue
                return True
            except Exception as e:
                self.logger.log(logging.DEBUG, f"Ошибка обработки объектов на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)
                    try:
                        if self.com_doc is not None:
                            self.com_doc.Close(False)  # Отклонить изменения
                            self.com_doc = None
                        if self.com_app is not None:
                            self.com_app.Quit()
                            self.com_app = None
                        self._initialize_autocad()
                    except Exception as reinf_err:
                        self.logger.log(logging.DEBUG, f"Не удалось переинициализировать AutoCAD: {reinf_err}")
                else:
                    self.logger.log(logging.DEBUG, f"Не удалось обработать объекты после {retries} попыток: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    return False

    def process_file(self, input_path, output_path):
        retries = 3
        success = False
        for attempt in range(retries):
            try:
                if not self.wait_for_object_ready(self.com_app, timeout=20.0, check_type="app"):
                    self.logger.log(logging.DEBUG, f"AutoCAD не готов для открытия {input_path} на попытке {attempt + 1}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    continue
                self.com_doc = self.com_app.Documents.Open(os.path.abspath(input_path))
                if self.wait_for_object_ready(self.com_doc, timeout=20.0, check_type="doc"):
                    self.logger.log(logging.DEBUG, f"Открыт: {os.path.basename(input_path)}")
                    try:
                        self.com_app.Visible = False
                    except Exception as e:
                        self.logger.log(logging.DEBUG, f"Не удалось установить Visible = False: {e}")
                    try:
                        self.com_doc.SendCommand("(setvar \"FILEDIA\" 0)\n")
                        self.com_doc.SendCommand("(setvar \"CMDDIA\" 0)\n")
                        self.com_doc.SendCommand("(setvar \"AUTOSAVE\" 0)\n")
                    except Exception as e:
                        self.logger.log(logging.DEBUG, f"Не удалось отключить диалоговые окна или автосохранение: {e}")
                    try:
                        self.com_doc.SendCommand("RECOVER\n")
                        self.logger.log(logging.DEBUG, f"Выполнена команда RECOVER для {input_path}")
                        time.sleep(2)
                    except Exception as e:
                        self.logger.log(logging.DEBUG, f"Ошибка выполнения RECOVER для {input_path}: {e}")
                    if self._process_all_entities():
                        self.com_doc.SaveAs(os.path.abspath(output_path))
                        self.logger.log(logging.DEBUG, f"Сохранено: {output_path}")
                        success = True
                    else:
                        self.logger.log(logging.ERROR, f"Обработка {input_path} не удалась, изменения не сохраняются")
                    return success
                else:
                    self.logger.log(logging.DEBUG, f"Документ не готов на попытке {attempt + 1}")
            except Exception as e:
                self.logger.log(logging.ERROR, f"Критическая ошибка в {input_path} на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)
                    try:
                        if self.com_doc is not None:
                            self.com_doc.Close(False)
                            self.com_doc = None
                        if self.com_app is not None:
                            self.com_app.Quit()
                            self.com_app = None
                        self._initialize_autocad()
                    except Exception as reinf_err:
                        self.logger.log(logging.ERROR, f"Не удалось переинициализировать AutoCAD: {reinf_err}")
                else:
                    self.logger.log(logging.ERROR, f"Не удалось обработать {input_path} после {retries} попыток: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    return False
            finally:
                try:
                    if self.com_doc is not None:
                        self.com_doc.Close(False)
                        self.com_doc = None
                except Exception as e:
                    self.logger.log(logging.ERROR, f"Ошибка закрытия документа: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()

    def process_files(self, input_files, output_dir):
        results = {}
        for input_path in input_files:
            if not os.path.isfile(input_path):
                self.logger.log(logging.ERROR, f"Файл не найден: {input_path}")
                results[input_path] = False
                continue

            filename = os.path.basename(input_path)
            name, ext = os.path.splitext(filename)
            output_path = os.path.join(output_dir, name + ext)  # Will be overwritten
            try:
                results[input_path] = self.process_file(input_path, output_path)
            except Exception as e:
                self.logger.log(logging.ERROR, f"Критическая ошибка обработки {input_path}: {e}")
                results[input_path] = False
                try:
                    if self.com_doc is not None:
                        self.com_doc.Close(False)
                        self.com_doc = None
                    if self.com_app is not None:
                        self.com_app.Quit()
                        self.com_app = None
                    self._initialize_autocad()
                except Exception as reinf_err:
                    self.logger.log(logging.ERROR, f"Не удалось сбросить AutoCAD для следующего файла: {reinf_err}")
                    self._terminate_autocad()
                    self._initialize_autocad()
        return results

    def __del__(self):
        try:
            if self.com_doc is not None:
                self.com_doc.Close(False)
                self.com_doc = None
            if self.com_app is not None:
                self.com_app.Quit()
                self.com_app = None
        except Exception as e:
            self.logger.log(logging.DEBUG, f"Ошибка очистки ресурсов AutoCAD: {e}")
            self._terminate_autocad()
        finally:
            pythoncom.CoUninitialize()