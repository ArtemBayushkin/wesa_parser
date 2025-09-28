import os
import re
import time
import win32com.client
import pythoncom
import psutil  # Для завершения процессов


class AutoCADProcessor:
    def __init__(self, replacement_digit, log_callback=None, debug=False):
        pythoncom.CoInitialize()  # Инициализация COM
        self.debug = debug  # Флаг отладки
        if not re.match(r'^\d$', str(replacement_digit)):
            raise ValueError("Replacement digit must be 0-9")

        self.replacement_digit = str(replacement_digit)
        self.log = log_callback or (lambda msg: print(msg))
        self.com_app = None
        self.com_doc = None
        self._initialize_autocad()

        self.patterns = [
            (re.compile(r'\b(ED\.D\.[A-Z]\d{3}\.)\d\b'),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
            (re.compile(r'\b\d\d[A-Z]{3}\d\d[A-Z]{1,2}\d{3,4}\b'),
             lambda m: self.replacement_digit + m.group(0)[1:]),
            (re.compile(r'\((\d{2}[A-Z]{3,})\)'),
             lambda m: f"({self.replacement_digit}{m.group(1)[1:]})"),
            (re.compile(r'\b\d\d[A-Z]{3}\d\d\b'),
             lambda m: self.replacement_digit + m.group(0)[1:]),
            (re.compile(r'\b([0-9])0([A-Z]{3})\b', flags=re.IGNORECASE),
             lambda m: f"{self.replacement_digit}0{m.group(2)}"),
            (re.compile(r'C0[2-9]\b'),
             'C01'),
            (re.compile(r'(Unit )\d\b', flags=re.IGNORECASE),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
            (re.compile(r'(Блок )\d\b', flags=re.IGNORECASE),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),
        ]

        self.delete_text_patterns = [
            re.compile(r'^C0[0-9]$'),
            re.compile(r'^-+$'),
            re.compile(r'^Repl\.$'),
            re.compile(r'^Зам\.$'),
            re.compile(r'^\d{4,5}-\d{2}$'),
            re.compile(r'^\d{2}\.\d{2,4}$'),
        ]

        # Новый параметр: толерантность для группировки по Y (учитываем float-погрешности)
        self.y_tolerance = 0.1  # Можно настроить под ваши чертежи

        # Словарь для кандидатов на удаление: {approx_y: [(entity, text, exact_y, x)]}
        self.delete_candidates = {}

    def _initialize_autocad(self):
        """Инициализация или переинициализация COM-интерфейса AutoCAD с повторами."""
        retries = 3
        for attempt in range(retries):
            try:
                self._terminate_autocad()  # Очистка перед созданием нового экземпляра
                self.com_app = win32com.client.Dispatch("AutoCAD.Application")
                if self.wait_for_object_ready(self.com_app, timeout=20.0, check_type="app"):
                    self._log("Экземпляр AutoCAD создан")
                    return
                else:
                    self._log(f"Экземпляр AutoCAD не готов на попытке {attempt + 1}")
            except Exception as e:
                self._log(f"Не удалось создать экземпляр AutoCAD на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)  # Увеличенная задержка
                else:
                    raise Exception(f"Не удалось создать экземпляр AutoCAD после {retries} попыток: {e}")
            finally:
                pythoncom.CoUninitialize()
                pythoncom.CoInitialize()

    def wait_for_object_ready(self, obj, timeout=20.0, check_type="app"):
        """Ожидание готовности COM-объекта с улучшенной проверкой."""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                pythoncom.PumpWaitingMessages()
                if obj is not None:
                    if check_type == "app":
                        _ = obj.Version  # Проверка свойства Version для приложения
                    else:
                        _ = obj.Name  # Проверка свойства Name для документа
                    return True
            except Exception as e:
                self._log(f"Ошибка проверки готовности объекта ({check_type}): {e}")
            time.sleep(0.2)
        self._log(f"Объект ({check_type}) не готов после {timeout} секунд")
        return False

    def _terminate_autocad(self):
        """Принудительное завершение процессов AutoCAD при их зависании."""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower().startswith('acad'):
                    proc.kill()
                    self._log("Процесс AutoCAD завершен")
            time.sleep(1)  # Даем время на завершение процесса
        except Exception as e:
            self._log(f"Ошибка при завершении процесса AutoCAD: {e}")
        self.com_app = None
        self.com_doc = None

    def _log(self, message):
        always_log = (
                message.startswith("Успешно: ") or
                message.startswith("Ошибка обработки: ") or
                message.startswith("Критическая ошибка ") or
                message.startswith("Файлы не найдены.") or
                message.startswith("Пропуск ") or
                message.startswith("Saved: ")
        )
        if self.debug or always_log:
            self.log(message)

    def _is_text_to_delete(self, text):
        return any(pattern.match(text) for pattern in self.delete_text_patterns)

    def _apply_replacements(self, text):
        if not text:
            return text
        original = text
        new_text = text
        for pattern, repl in self.patterns:
            if pattern.search(new_text):
                new_text = pattern.sub(repl, new_text) if callable(repl) else pattern.sub(repl, new_text)
        if new_text != original:
            self._log(f"Замена: {original} → {new_text}")
        return new_text

    def _process_entity(self, entity, depth=0, location=""):
        retries = 3
        for attempt in range(retries):
            try:
                if not hasattr(entity, 'ObjectName'):
                    self._log(f"Объект в {location} не имеет ObjectName, пропуск")
                    return
                etype = entity.ObjectName
                self._log(f"Обработка объекта {etype} в {location}")
                if etype in ("AcDbText", "AcDbMText"):
                    try:
                        insertion_point = entity.InsertionPoint
                        x, y = insertion_point[0], insertion_point[1]
                        txt = entity.TextString
                    except Exception as e:
                        self._log(f"Ошибка доступа к свойствам текста в {location}: {e}")
                        return

                    # Вместо проверки области и немедленного удаления:
                    if self._is_text_to_delete(txt):
                        # Приближённый Y для группировки (округляем до толерантности)
                        approx_y = round(y / self.y_tolerance) * self.y_tolerance
                        if approx_y not in self.delete_candidates:
                            self.delete_candidates[approx_y] = []
                        self.delete_candidates[approx_y].append((entity, txt, y, x))
                        self._log(f"Кандидат на удаление: {txt} в ({x}, {y}) {location}")
                        # Не удаляем сразу — это сделаем позже

                    # Продолжаем с заменами (если нужно, но замена и удаление — отдельно)
                    new_txt = self._apply_replacements(txt)
                    if new_txt != txt:
                        try:
                            entity.TextString = new_txt
                            self._log(f"Замена в {location}: {txt} → {new_txt}")
                        except Exception as e:
                            self._log(f"Ошибка установки TextString в {location}: {e}")
                elif etype == "AcDbMLeader":
                    try:
                        txt = entity.TextString
                        new_txt = self._apply_replacements(txt)
                        if new_txt != txt:
                            entity.TextString = new_txt
                            self._log(f"Замена в {location} (MLeader): {txt} → {new_txt}")
                    except Exception as e:
                        self._log(f"Ошибка обработки MLeader в {location}: {e}")
                elif etype == "AcDbBlockReference" and hasattr(entity, "GetAttributes"):
                    try:
                        attributes = entity.GetAttributes()
                        for attr in attributes:
                            try:
                                txt = attr.TextString
                                new_txt = self._apply_replacements(txt)
                                if new_txt != txt:
                                    attr.TextString = new_txt
                                    self._log(f"Замена в атрибуте блока {location}: {txt} → {new_txt}")
                            except Exception as e:
                                self._log(f"Ошибка обработки атрибута в {location}: {e}")
                                continue  # Пропускаем проблемный атрибут
                    except Exception as e:
                        self._log(f"Ошибка доступа к атрибутам блока в {location}: {e}")
                return
            except Exception as e:
                self._log(f"Ошибка объекта в {location} на попытке {attempt + 1}: {e}")
                if attempt < retries - 1:
                    time.sleep(3)
                else:
                    self._log(f"Не удалось обработать объект в {location} после {retries} попыток: {e}")
                    return

    def _process_blocks(self):
        retries = 3
        for attempt in range(retries):
            try:
                if self.com_doc is None:
                    self._log("Документ не инициализирован, пропуск обработки блоков")
                    return
                block_table = self.com_doc.Blocks
                for block in block_table:
                    if not block.IsLayout and not block.IsXRef:
                        self._log(f"Обработка блока: {block.Name}")
                        try:
                            for entity in block:
                                self._process_entity(entity, depth=1, location=f"block {block.Name}")
                        except Exception as e:
                            self._log(f"Пропуск блока {block.Name} из-за ошибки: {e}")
                            continue
                return
            except Exception as e:
                self._log(f"Ошибка обработки блоков на попытке {attempt + 1}: {e}")
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
                        self._log(f"Не удалось переинициализировать AutoCAD: {reinf_err}")
                else:
                    self._log(f"Не удалось обработать блоки после {retries} попыток: {e}")
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
                self._log("Обработка ModelSpace...")
                for entity in self.com_doc.ModelSpace:
                    self._process_entity(entity, location="ModelSpace")
                self._log("Обработка блоков...")
                self._process_blocks()
                self._log("Обработка листов...")
                for layout in self.com_doc.Layouts:
                    if layout.Name.lower() in ['model', 'модель']:
                        continue
                    self._log(f"Лист: {layout.Name}")
                    try:
                        for entity in layout.Block:
                            self._process_entity(entity, location=f"Layout {layout.Name}")
                    except Exception as e:
                        self._log(f"Пропуск листа {layout.Name} из-за ошибки: {e}")
                        continue

                # После обработки всех entities: анализируем и удаляем кандидаты
                self._delete_grouped_candidates()

                return True  # Успешная обработка
            except Exception as e:
                self._log(f"Ошибка обработки объектов на попытке {attempt + 1}: {e}")
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
                        self._log(f"Не удалось переинициализировать AutoCAD: {reinf_err}")
                else:
                    self._log(f"Не удалось обработать объекты после {retries} попыток: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    return False

    def _delete_grouped_candidates(self):
        """Удаление групп кандидатов на одной Y."""
        for approx_y, group in self.delete_candidates.items():
            if len(group) >= 2:  # Удаляем, если в группе >=2 (настройте по вкусу)
                # Опционально: сортируем по X для лога
                group.sort(key=lambda item: item[3])  # По X
                texts = [item[1] for item in group]
                self._log(f"Группа на Y≈{approx_y}: {texts} — удаление")
                for entity, txt, y, x in group:
                    try:
                        entity.Delete()
                        self._log(f"Удален: {txt} в ({x}, {y})")
                    except Exception as e:
                        self._log(f"Ошибка удаления: {txt} — {e}")
            else:
                self._log(f"Одиночный на Y≈{approx_y}: не удаляем")

        # Очищаем candidates после обработки
        self.delete_candidates = {}

    def process_file(self, input_path, output_path):
        self.delete_candidates = {}  # Сброс перед каждым файлом
        retries = 3
        success = False
        for attempt in range(retries):
            try:
                # Проверяем, что AutoCAD готов перед открытием файла
                if not self.wait_for_object_ready(self.com_app, timeout=20.0, check_type="app"):
                    self._log(f"AutoCAD не готов для открытия {input_path} на попытке {attempt + 1}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    continue
                self.com_doc = self.com_app.Documents.Open(os.path.abspath(input_path))
                if self.wait_for_object_ready(self.com_doc, timeout=20.0, check_type="doc"):
                    self._log(f"Открыт: {os.path.basename(input_path)}")
                    # Устанавливаем Visible = False после открытия документа
                    try:
                        self.com_app.Visible = False
                    except Exception as e:
                        self._log(f"Не удалось установить Visible = False: {e}")
                    # Отключаем диалоговые окна и автосохранение
                    try:
                        self.com_doc.SendCommand("(setvar \"FILEDIA\" 0)\n")
                        self.com_doc.SendCommand("(setvar \"CMDDIA\" 0)\n")
                        self.com_doc.SendCommand("(setvar \"AUTOSAVE\" 0)\n")
                    except Exception as e:
                        self._log(f"Не удалось отключить диалоговые окна или автосохранение: {e}")
                    # Выполняем RECOVER для исправления файла
                    try:
                        self.com_doc.SendCommand("RECOVER\n")
                        self._log(f"Выполнена команда RECOVER для {input_path}")
                        time.sleep(2)  # Увеличенная задержка
                    except Exception as e:
                        self._log(f"Ошибка выполнения RECOVER для {input_path}: {e}")
                    if self._process_all_entities():  # Проверяем успешность обработки
                        self.com_doc.SaveAs(os.path.abspath(output_path))
                        self._log(f"Сохранено: {output_path}")
                        success = True
                    else:
                        self._log(f"Обработка {input_path} не удалась, изменения не сохраняются")
                    return success
                else:
                    self._log(f"Документ не готов на попытке {attempt + 1}")
            except Exception as e:
                self._log(f"Критическая ошибка в {input_path} на попытке {attempt + 1}: {e}")
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
                        self._log(f"Не удалось переинициализировать AutoCAD: {reinf_err}")
                else:
                    self._log(f"Не удалось обработать {input_path} после {retries} попыток: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()
                    return False
            finally:
                try:
                    if self.com_doc is not None:
                        self.com_doc.Close(False)  # Отклонить изменения
                        self.com_doc = None
                except Exception as e:
                    self._log(f"Ошибка закрытия документа: {e}")
                    self._terminate_autocad()
                    self._initialize_autocad()

    def process_files(self, input_files, output_dir):
        results = {}
        for input_path in input_files:
            if not os.path.isfile(input_path):
                self._log(f"Файл не найден: {input_path}")
                results[input_path] = False
                continue

            filename = os.path.basename(input_path)
            name, ext = os.path.splitext(filename)
            if name[0].isdigit():
                new_name = f"{self.replacement_digit}{name[1:]}"
            elif name.startswith("ED.D."):
                new_name = re.sub(
                    r'(ED\.D\.[A-Z]\d{3}\.)(\d)',
                    lambda m: f"{m.group(1)}{self.replacement_digit}",
                    name
                )
            else:
                new_name = f"processed_{name}"
            output_path = os.path.join(output_dir, new_name + ext)
            try:
                results[input_path] = self.process_file(input_path, output_path)
            except Exception as e:
                self._log(f"Критическая ошибка обработки {input_path}: {e}")
                results[input_path] = False
                try:
                    if self.com_doc is not None:
                        self.com_doc.Close(False)  # Отклонить изменения
                        self.com_doc = None
                    if self.com_app is not None:
                        self.com_app.Quit()
                        self.com_app = None
                    self._initialize_autocad()
                except Exception as reinf_err:
                    self._log(f"Не удалось сбросить AutoCAD для следующего файла: {reinf_err}")
                    self._terminate_autocad()
                    self._initialize_autocad()
        return results

    def __del__(self):
        try:
            if self.com_doc is not None:
                self.com_doc.Close(False)  # Отклонить изменения
                self.com_doc = None
            if self.com_app is not None:
                self.com_app.Quit()
                self.com_app = None
        except Exception as e:
            self._log(f"Ошибка очистки ресурсов AutoCAD: {e}")
            self._terminate_autocad()
        finally:
            pythoncom.CoUninitialize()