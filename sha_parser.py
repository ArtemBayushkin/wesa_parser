# sha_parser.py
"""Модуль sha_parser.py: Обработка файлов SHA через WinAPI и COM-интерфейс SmartSketch.

Этот модуль предоставляет инструменты для автоматизации замены текста в файлах SHA
с использованием регулярных выражений и COM-объектов. Он ориентирован на Windows
и требует установки pywin32. Основной класс: ShaProcessorWinAPI.

Зависимости: os, re, sys, winreg, win32com.client, pythoncom, pywintypes, time, json.
"""
import os
import re
import winreg
import win32com.client
import pythoncom
import pywintypes
import time
import logging

def get_license_servers_from_registry():
    """
    Читает серверы лицензий из реестра Windows и формирует строку для INGR_LICENSE_PATH.

    Открывает ключ реестра, извлекает значение 'server_names', добавляет порт 27000
    и объединяет в строку с разделителем ';'. Если ключ не найден или ошибка - возвращает
    пустую строку.
    :return: Строка с серверами лицензий в формате '27000@server1;27000@server2' или ''.
    :raise: OSError: Если доступ к реестру запрещён или ключ не существует.
    """
    try:
        path = r"SOFTWARE\WOW6432Node\Intergraph\Pdlice_etc\server_names"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
            value, _ = winreg.QueryValueEx(key, "server_names")
            servers = value.split()
            servers_with_port = [f"27000@{s}" for s in servers]
            return ";".join(servers_with_port)
    except Exception:
        return ""

def wait_for_object_ready(obj, timeout=3.0):
    """
    Ожидает готовности COM-объекта в цикле с обработкой сообщений.

    Использует pythoncom.PumpWaitingMessages() для обработки очереди сообщений COM.
    Проверяет, что объект не None. Если таймаут истёк - возвращает False.
    Необходима для исключения RPC ошибки
    :param obj: COM-объект для проверки (например, документ или приложение).
    :param timeout: Максимальное время ожидания в секундах. По дефолту 3.0 сек.
    :return: True, если объект готов; False, если таймаут истёк.
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            pythoncom.PumpWaitingMessages()
            if obj is not None:
                return True
        except Exception:
            pass
        time.sleep(0.1)
    return False

class ShaProcessorWinAPI:
    """
    Класс для автоматической обработки файлов SmartSketch (Shape2DServer) через WinAPI/COM.

    Основные возможности:
    - Запуск и завершение SmartSketch через COM-интерфейс.
    - Чтение правил замены текста (regex) и применение их ко всем текстовым объектам документа.
    - Поиск и изменение текста в различных элементах: TextBox, надписи, свойства объектов, группы и вложенные группы.
    - Сохранение измененного документа в новый файл, если были найдены и применены изменения.

    Параметры:
        replacement_digit (str|int):
            Цифра или символ, используемый в функциях замены.

        project (str):
            Название или код проекта (не используется напрямую внутри класса, но полезен для логов или внешней логики).

        rules (dict):
            Набор правил замены. Каждый ключ — имя правила, значение — словарь с полями:
                "pattern"      - выражение регулярного поиска (как строка, интерпретируемая через eval),
                "replacement"  - выражение замены (также исполняется через eval).

            Пример:
            {
                "rule1": {
                    "pattern": "re.compile(r\"ABC\")",
                    "replacement": "lambda m: \"XYZ\""
                }
            }

        log_callback (callable, необязательно):
            Функция для вывода логов. Если не задана, логирование отключено.
            Передаваемая функция должна принимать один аргумент: строку сообщения.

        debug (bool):
            Если True - логируются все сообщения. Иначе — только важные.

    Основные методы:
        start_app():
            Запускает приложение SmartSketch и готовит COM-среду.

        stop_app():
            Корректно завершает приложение SmartSketch и освобождает COM-ресурсы.

        process_file(input_path, output_path):
            Открывает файл SmartSketch, выполняет поиск и замену текста по правилам,
            сохраняет результат в указанный путь (если были изменения),
            а затем закрывает документ.

    Примечания:
        - Класс рассчитан на работу только в Windows и требует установленного SmartSketch.
        - Для корректной работы должны быть доступны серверы лицензирования SmartSketch.
        - Правила замены должны быть корректно сформированы, поскольку используются через eval().
    """
    def __init__(self, replacement_digit, project, rules, logger=None):
        """
        Инициализирует экземпляр ShaProcessorWinAPI.

        Конвертирует replacement_digit в строку, устанавливает флаги и логирование,
        загружает паттерны из rules.

        :param replacement_digit: Цифра для замены (будет конвертирована в str).
        :param project: Название проекта для логирования.
        :param rules: Словарь правил с 'pattern' и 'replacement'.
        :param log_callback: Параметр для логирования. По дефолту None.
        :param debug: Включает отладочное логирование.
        """
        self.replacement_digit = str(replacement_digit)    # цифра, которая участвует в заменах.
        
        self.logger = logger or logging.getLogger()
        self.logger.log(logging.DEBUG, f"Инициализация ShaProcessorWinAPI с цифрой: {self.replacement_digit} и проектом: {project}")

        self.patterns = self._load_patterns(rules)         # загружаем и компилируем правила замен

    def _load_patterns(self, rules):
        """
        Загружает паттерны замен из словаря правил.

        Для каждого правила eval'ит pattern (re.compile) и replacement.
        Логирует загрузку или ошибки. Если правил нет - выдаёт предупреждение.

        :param rules: Словарь {rule_name: {'pattern': str, 'replacement': str}}.
        :return: Список кортежей (compiled_pattern, replacement_func_or_str).
        """
        patterns = []
        try:
            for rule_name, rule in rules.items():
                try:
                    pattern = eval(rule["pattern"], {"re": re})
                    replacement = eval(rule["replacement"], {"self": self})
                    patterns.append((pattern, replacement))
                    self._log(f"Загружено правило '{rule_name}'")
                except Exception as e:
                    self._log(f"Ошибка загрузки правила '{rule_name}': {e}")
        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка обработки rules: {e}")
        if not patterns:
            self.logger.log(logging.DEBUG,"Предупреждение: Нет patterns для этого парсера")
        return patterns

    def start_app(self):
        """
        Запускает приложение.
        Инициализирует COM, получает лицензии, устанавливает env-var, диспатчит COM-объект SmartSketch, логирует.
        :return: None
        :raise: RuntimeError: Если приложение не запустилось.
        """
        pythoncom.CoInitialize()

        servers = get_license_servers_from_registry()
        if servers:
            os.environ["INGR_LICENSE_PATH"] = servers
            self.logger.log(logging.DEBUG,f"[ЛИЦЕНЗИИ] Используются сервера: {servers}")
        else:
            self.logger.log(logging.DEBUG,"[ЛИЦЕНЗИИ] Не удалось найти сервера в реестре")

        try:
            self.app = win32com.client.Dispatch("Shape2DServer.Application")
            self.logger.log(logging.DEBUG, "SmartSketch запущен успешно")
        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка запуска SmartSketch: {e}")
            self.stop_app()
            raise

    def stop_app(self):
        """
        Останавливает приложение SmartSketch и деинициализирует COM.
        Вызывает Quit на app, если оно существует, и очищает ресурсы.

        :return: None
        """
        try:
            if self.app:
                self.app.Quit()
                self.logger.log(logging.DEBUG, "SmartSketch закрыт")
        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка при закрытии SmartSketch: {e}")
        finally:
            self.app = None
            pythoncom.CoUninitialize()

    def _replace_text_in_object(self, text_obj, obj_name):
        """
        Заменяет текст в объекте, если присутствует атрибут 'Text'.
        Применяет все паттерны из self.patterns. Логирует изменения или ошибки.

        :param text_obj: COM-объект с возможным атрибутом 'Text'.
        :param obj_name: Имя объекта для логирования.
        :return: True, если текст был изменён; False иначе
        """
        try:
            if hasattr(text_obj, "Text"):
                text = text_obj.Text
                if text and isinstance(text, str):
                    original_text = text
                    for pattern, replacement in self.patterns:
                        text = pattern.sub(replacement, text)
                    if text != original_text:
                        text_obj.Text = text
                        self._log(f"[ИЗМЕНЕНО] {obj_name}: '{original_text}' → '{text}'")
                        return True
        except Exception as e:
            self.logger.log(logging.DEBUG, f"[ОШИБКА] {obj_name}: {e}")
        return False

    def _process_group(self, group, group_name, depth=0):
        """
        Рекурсивно обрабатывает группу объектов.
        Ограничивает глубину 3, чтобы избежать бесконечной рекурсии.
        Для каждого item: заменяет текст generic'ом, рекурсивно обрабатывает подгруппы.
        :param group: COM-группа объектов.
        :param group_name: Имя группы для логирования.
        :param depth: Текущая глубина рекурсии.
        :return: True, если были изменения; False иначе.
        """
        if depth > 3:
            return False
        changes = False

        try:
            if hasattr(group, "Item") and hasattr(group, "Count"):
                for i in range(1, group.Count + 1):
                    try:
                        item = group.Item(i)
                        if self._replace_text_generic(item, f"Item {i} в {group_name}"):
                            changes = True
                        if self._process_group(item, f"Item {i} в {group_name}", depth + 1):
                            changes = True
                    except Exception:
                        continue
        except Exception:
            pass

        return changes

    def _replace_text_generic(self, obj, obj_name):
        """
        Универсально заменяет текст в объекте по списку свойств.
        Проверяет свойства вроде 'Text', 'Caption' и т.д., применяет паттерны, устанавливает новое, логирует.

        :param obj: COM-объект для обработки.
        :param obj_name: Имя объекта для логирования.
        :return: True, если были изменения; False иначе.
        """
        text_properties = ["Text", "TextString", "Caption", "Value", "String",
                           "Content", "Name", "Label", "Description"]
        changed = False
        for prop in text_properties:
            if hasattr(obj, prop):
                try:
                    val = getattr(obj, prop)
                except Exception:
                    continue
                if isinstance(val, str) and val.strip():
                    new_val = val
                    for pattern, repl in self.patterns:
                        new_val = pattern.sub(repl, new_val)
                    if new_val != val:
                        try:
                            setattr(obj, prop, new_val)
                            changed = True
                            self.logger.log(logging.DEBUG, f"[ИЗМЕНЕНО] {obj_name}.{prop}: '{val}' → '{new_val}'")
                        except Exception:
                            pass
        return changed

    def process_file(self, input_path, output_path):
        """
        Обрабатывает файл SHA: открывает, заменяет текст, сохраняет если изменения.
        Обрабатывает TextBoxes и Groups на каждом листе. Закрывает документ в finally.

        :param input_path: Путь к входному файлу.
        :param output_path: Путь для сохранения выходного файла.
        :return: True, если обработка успешна; False при ошибках.
        :raise: RuntimeError: Если приложение не запущено.
        """
        if not self.app:
            raise RuntimeError("SmartSketch не запущен")

        try:
            doc = self.app.Documents.Open(os.path.abspath(input_path))
            wait_for_object_ready(doc)

            changes_made = False

            for sheet_idx, sheet in enumerate(doc.Sheets, start=1):
                self.logger.log(logging.DEBUG, f"--- Лист {sheet_idx}/{doc.Sheets.Count} ---")

                if hasattr(sheet, "TextBoxes") and sheet.TextBoxes is not None:
                    for tb_idx, tb in enumerate(sheet.TextBoxes, start=1):
                        if self._replace_text_in_object(tb, f"TextBox {tb_idx} на Листе {sheet_idx}"):
                            changes_made = True

                if hasattr(sheet, "Groups") and sheet.Groups is not None:
                    for group_idx, group in enumerate(sheet.Groups, start=1):
                        if self._process_group(group, f"Group {group_idx} на Листе {sheet_idx}"):
                            changes_made = True

            if changes_made:
                doc.SaveAs(output_path)
                self.logger.log(logging.DEBUG, f"Документ сохранён: {output_path}")
            else:
                self.logger.log(logging.DEBUG, f"Изменений не найдено, сохранение пропущено")

            return True

        except pywintypes.com_error as e:
            self.logger.log(logging.DEBUG, f"COM ошибка при обработке {input_path}: {e}")
            return False

        except Exception as e:
            self.logger.log(logging.ERROR, f"Ошибка обработки {input_path}: {e}")
            return False

        finally:
            try:
                doc.Close(False)
                wait_for_object_ready(doc)
            except Exception:
                pass
            finally:
                del doc