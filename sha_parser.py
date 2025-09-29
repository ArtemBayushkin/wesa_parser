import os
import re
import winreg
import win32com.client
import pythoncom
import pywintypes
import time

def get_license_servers_from_registry():
    """Читаем серверы лицензий из реестра и формируем строку INGR_LICENSE_PATH"""
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
    """Ожидание готовности COM-объекта (без лишних логов)."""
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
    def __init__(self, replacement_digit, log_callback=None, debug=False):
        self.replacement_digit = str(replacement_digit)
        self.debug = debug  # Флаг отладки
        self.log = log_callback or (lambda msg: None)  # Если колбэк не передан — молчит
        self.app = None
        self._log(f"Инициализация ShaProcessorWinAPI с цифрой: {self.replacement_digit}")

        self.patterns = [
            # ED.D.P000.N → замена N
            (re.compile(r'(ED\.D\.[A-Z]\d{3}\.)(\d)'),
             lambda m: f"{m.group(1)}{self.replacement_digit}"),

            # N0&&&&&BQ2200 → замена N
            (re.compile(r'([1-9])(0&&&&&[A-Z]{2}\d{4})'),
             lambda m: f"{self.replacement_digit}{m.group(2)}"),

            # N0KTC → замена N
            (re.compile(r'([1-9])(0[A-Z]{3})'),
             lambda m: f"{self.replacement_digit}{m.group(2)}"),

            # alt="C0N" → alt="C01"
            (re.compile(re.escape('<?xml version="1.0"?><body><intstgxml stream="Revision" select="/Revision/RevisionRecord[last()-0]/MajorRev_ForRevise" alt="C01"/><intstgxml stream="Revision" select="/Revision/RevisionRecord[last()-0]/MinorRev_ForRevise" alt=""/></body>')),
                        '<?xml version="1.0"?><body><intstgxml stream="Revision" select="/Revision/RevisionRecord[last()-10]/MajorRev_ForRevise" alt="C01"/><intstgxml stream="Revision" select="/Revision/RevisionRecord[last()-10]/MinorRev_ForRevise" alt=""/></body>'),

            # C0N → C01
            (re.compile(r'\bC0[2-9]\b'),
             'C01')
        ]

    def _log(self, message):
        # Логи, которые всегда записываются
        always_log = (
            message.startswith("Успешно: ") or
            message.startswith("Ошибка обработки: ") or
            message.startswith("Критическая ошибка ") or
            message.startswith("Файлы не найдены.") or
            message.startswith("Пропуск ") or
            message.startswith("Документ сохранён: ") or
            message.startswith("SmartSketch запущен успешно") or
            message.startswith("SmartSketch закрыт") or
            message.startswith("COM ошибка при обработке ") or
            message.startswith("Ошибка обработки ")
        )
        # Если отладка включена или это обязательный лог, вызываем callback
        if self.debug or always_log:
            self.log(message)

    def start_app(self):
        """Запуск SmartSketch один раз с установкой лицензии."""
        pythoncom.CoInitialize()

        servers = get_license_servers_from_registry()
        if servers:
            os.environ["INGR_LICENSE_PATH"] = servers
            self._log(f"[ЛИЦЕНЗИИ] Используются сервера: {servers}")
        else:
            self._log("[ЛИЦЕНЗИИ] Не удалось найти сервера в реестре")

        try:
            self.app = win32com.client.Dispatch("Shape2DServer.Application")
            self._log("SmartSketch запущен успешно")
        except Exception as e:
            self._log(f"Ошибка запуска SmartSketch: {e}")
            self.stop_app()
            raise

    def stop_app(self):
        """Закрытие SmartSketch."""
        try:
            if self.app:
                self.app.Quit()
                self._log("SmartSketch закрыт")
        except Exception as e:
            self._log(f"Ошибка при закрытии SmartSketch: {e}")
        finally:
            self.app = None
            pythoncom.CoUninitialize()

    def _replace_text_in_object(self, text_obj, obj_name):
        """Замена текста в объекте."""
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
            self._log(f"[ОШИБКА] {obj_name}: {e}")
        return False

    def _process_group(self, group, group_name, depth=0):
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
        """Универсальная замена текста по набору свойств (для объектов в Group.Item)."""
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
                            self._log(f"[ИЗМЕНЕНО] {obj_name}.{prop}: '{val}' → '{new_val}'")
                        except Exception:
                            pass
        return changed

    def process_file(self, input_path, output_path):
        """Открыть файл, заменить текст и сохранить новый."""
        if not self.app:
            raise RuntimeError("SmartSketch не запущен")

        try:
            doc = self.app.Documents.Open(os.path.abspath(input_path))
            wait_for_object_ready(doc)

            changes_made = False

            for sheet_idx, sheet in enumerate(doc.Sheets, start=1):
                self._log(f"--- Лист {sheet_idx}/{doc.Sheets.Count} ---")

                # Текстовые блоки на листе
                if hasattr(sheet, "TextBoxes") and sheet.TextBoxes is not None:
                    for tb_idx, tb in enumerate(sheet.TextBoxes, start=1):
                        if self._replace_text_in_object(tb, f"TextBox {tb_idx} на Листе {sheet_idx}"):
                            changes_made = True

                # Группы на листе
                if hasattr(sheet, "Groups") and sheet.Groups is not None:
                    for group_idx, group in enumerate(sheet.Groups, start=1):
                        if self._process_group(group, f"Group {group_idx} на Листе {sheet_idx}"):
                            changes_made = True

            # Сохраняем, если есть изменения
            if changes_made:
                doc.SaveAs(output_path)
                self._log(f"Документ сохранён: {output_path}")
            else:
                self._log(f"Изменений не найдено, сохранение пропущено")

            return True

        except pywintypes.com_error as e:
            self._log(f"COM ошибка при обработке {input_path}: {e}")
            return False

        except Exception as e:
            self._log(f"Ошибка обработки {input_path}: {e}")
            return False

        finally:
            try:
                doc.Close(False)
                wait_for_object_ready(doc)
            except Exception:
                pass
            finally:
                del doc