# pdf_parser.py
import os
import re
import fitz
import logging
import platform

class PdfProcessor:
    def __init__(self, replacement_digit, project, rules, log_callback=None, debug=False):
        self.replacement_digit = str(replacement_digit)
        self.debug = debug
        self.log = log_callback or (lambda msg: None)
        self._log(f"Инициализация PdfProcessor с цифрой: {self.replacement_digit} и проектом: {project}")
        self.patterns = self._load_patterns(rules)
        self._font_cache = {}  # кеш для fitz.Font объектов

    def _load_patterns(self, rules):
        patterns = []
        try:
            for rule_name, rule in rules.items():
                try:
                    pattern = eval(rule["pattern"], {"re": re})
                    replacement = eval(rule["replacement"], {"self": self})
                    patterns.append((pattern, replacement, rule_name))
                    self._log(f"Загружено правило '{rule_name}' для pdf_parser")
                except Exception as e:
                    self._log(f"Ошибка загрузки правила '{rule_name}': {e}")
        except Exception as e:
            self._log(f"Ошибка обработки rules: {e}")
        if not patterns:
            self._log("Предупреждение: Нет patterns для этого проекта/парсера")
        return patterns

    def _log(self, message):
        always_log = (
            message.startswith("Успешно: ") or
            message.startswith("Ошибка обработки: ") or
            message.startswith("Критическая ошибка ") or
            message.startswith("Файлы не найдены.") or
            message.startswith("Пропуск ") or
            message.startswith("Файл успешно обработан: ")
        )
        if self.debug or always_log:
            self.log(message)

    def _apply_replacements(self, text):
        if text is None:
            return None
        original_text = text
        for pattern, repl, rule_name in self.patterns:
            text = pattern.sub(repl, text) if callable(repl) else pattern.sub(repl, text)
        if text != original_text:
            self._log(f"Замена текста: '{original_text}' → '{text}'")
        return text

    def _get_style_for_text(self, page, old_str, rect):
        """Извлекает шрифт, размер и цвет из dict-структуры страницы для совпадающего текста."""
        text_dict = page.get_text("dict")
        for block in text_dict.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    span_text = span.get("text", "")
                    if old_str in span_text:
                        font = span.get("font", "helv")  # Fallback на helv
                        size = span.get("size", 8.0)     # Fallback на 8 pt
                        color = span.get("color", 0)     # Может быть int или tuple
                        origin = span.get("origin", (rect.x0, rect.y0))
                        self._log(f"Извлечён стиль для '{old_str}': font={font}, size={size}, color={color}")
                        return font, size, color, origin[1]  # Возвращаем y для baseline
        # Fallback, если не найдено
        return "helv", 8.0, 0, rect.y0

    def _color_to_tuple(self, color):
        """
        Преобразует цвет span (возможно int или tuple) в формат, который принимает insert_text:
        кортеж 3 floats в диапазоне 0..1.
        """
        try:
            # если уже кортеж/список из 3 чисел (0..255 или 0..1), нормализуем
            if isinstance(color, (list, tuple)) and len(color) >= 3:
                r, g, b = color[0], color[1], color[2]
                # detect 0..255 ints
                if any(isinstance(c, int) and c > 1 for c in (r, g, b)):
                    return (r/255.0, g/255.0, b/255.0)
                else:
                    return (float(r), float(g), float(b))
            # если цвет — целое (обычно 0 для чёрного)
            if isinstance(color, int):
                # PyMuPDF обычно хранит цвет как 0xRRGGBB (целое). Разложим.
                val = color
                r = (val >> 16) & 0xFF
                g = (val >> 8) & 0xFF
                b = val & 0xFF
                # если val==0 -> чёрный
                return (r/255.0, g/255.0, b/255.0)
        except Exception:
            pass
        # fallback — чёрный
        return (0.0, 0.0, 0.0)

    def _find_system_font_file(self, font_name):
        """
        Пробует найти файл шрифта по имени на типичных системных путях.
        Возвращает путь к файлу или None.
        """
        if not font_name:
            return None

        # Убираем пробелы и возможные постфиксы (например 'Bold', 'Italic' и т.п.)
        base = font_name.split('+')[-1].split(',')[0].strip()
        base_clean = re.sub(r'[^A-Za-z0-9]', '', base).lower()

        candidates = []
        system = platform.system()
        if system == "Windows":
            win_fonts = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
            # типичные расширения
            for ext in (".ttf", ".TTF", ".otf", ".ttc"):
                candidates.append(os.path.join(win_fonts, base + ext))
                candidates.append(os.path.join(win_fonts, base_clean + ext))
                # часто имя файла может быть arial.ttf для Arial
                candidates.append(os.path.join(win_fonts, base.split()[0] + ext))
        else:
            # Linux / macOS пути
            unix_paths = [
                "/usr/share/fonts",
                "/usr/local/share/fonts",
                os.path.expanduser("~/.fonts"),
                "/Library/Fonts",
                "/System/Library/Fonts"
            ]
            for p in unix_paths:
                for ext in (".ttf", ".otf", ".ttc"):
                    candidates.append(os.path.join(p, base + ext))
                    candidates.append(os.path.join(p, base_clean + ext))
                    candidates.append(os.path.join(p, base.split()[0] + ext))

        # проверить кандидатов
        for c in candidates:
            if c and os.path.exists(c):
                return c

        # дополнительно: попробовать найти любой файл, в имени которого есть base_clean
        for root in (os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts"), "/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"), "/Library/Fonts"):
            if not root or not os.path.isdir(root):
                continue
            try:
                for fname in os.listdir(root):
                    if base_clean and base_clean in fname.lower():
                        full = os.path.join(root, fname)
                        if os.path.isfile(full):
                            return full
            except Exception:
                continue

        return None

    def _ensure_fitz_font(self, font_name):
        """
        Возвращает либо объект fitz.Font (если удалось загрузить файл),
        либо встроенное имя шрифта (str) для использования в insert_text.
        Кеширует объекты.
        """
        if not font_name:
            return "helv"

        # Если шрифт — одно из стандартных встроенных имён, вернуть его имя
        builtin = {"helv", "times", "courier"}
        if font_name.lower() in builtin or any(b in font_name.lower() for b in builtin):
            return font_name.lower()

        # Кеширование по имени
        if font_name in self._font_cache:
            return self._font_cache[font_name]

        # Попробуем найти файл шрифта в системе
        font_path = self._find_system_font_file(font_name)
        if font_path:
            try:
                fobj = fitz.Font(file=font_path)
                self._font_cache[font_name] = fobj
                self._log(f"Загружен системный файл шрифта для '{font_name}': {font_path}")
                return fobj
            except Exception as e:
                self._log(f"Не удалось загрузить файл шрифта '{font_path}' для '{font_name}': {e}")

        # Если не удалось — попробуем простую конвертацию имени (например, Arial -> helv)
        self._log(f"Файл шрифта для '{font_name}' не найден. Использую fallback 'helv'.")
        self._font_cache[font_name] = "helv"
        return "helv"

    def process_file(self, input_path, output_path):
        try:
            doc = fitz.open(input_path)
            self._log(f"Открыт PDF файл: {input_path}")
            changes_made = False

            for page_num in range(len(doc)):
                page = doc[page_num]
                full_text = page.get_text("text")
                new_text = self._apply_replacements(full_text)

                if new_text != full_text:
                    for pattern, repl, rule_name in self.patterns:
                        matches = pattern.finditer(full_text)
                        for match in matches:
                            old_str = match.group(0)
                            new_str = repl(match) if callable(repl) else repl

                            # Находим все вхождения old_str
                            rects = page.search_for(old_str)
                            for rect in rects:
                                # Извлекаем оригинальный стиль
                                font, size, color, baseline_y = self._get_style_for_text(page, old_str, rect)

                                try:
                                    # закрасить прямоугольник белым (чтобы "стереть" старый текст)
                                    page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

                                    # подготовка шрифта и цвета
                                    font_obj_or_name = self._ensure_fitz_font(font)
                                    color_tuple = self._color_to_tuple(color)

                                    # Вставляем текст в тот же rect (insert_textbox сам позаботится о переносах)
                                    if isinstance(font_obj_or_name, fitz.Font):
                                        # если у вас объект fitz.Font, удобнее использовать стандартное имя для insert_textbox,
                                        # но если нужно — можно пробовать font=font_obj_or_name (в зависимости от версии PyMuPDF).
                                        page.insert_textbox(rect, new_str, fontsize=size, fontname="helv",
                                                            color=color_tuple, align=0)
                                    else:
                                        page.insert_textbox(rect, new_str, fontsize=size, fontname=font_obj_or_name,
                                                            color=color_tuple, align=0)

                                    self._log(
                                        f"Замена (textbox) на странице {page_num + 1} по правилу '{rule_name}': '{old_str}' → '{new_str}' (font={font}, size={size})")
                                except Exception as e:
                                    self._log(f"Ошибка вставки (draw_rect+textbox) для '{old_str}': {e}")
                                    # fallback: попытаться вставить с helv и чёрным цветом
                                    try:
                                        page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
                                        page.insert_textbox(rect, new_str, fontsize=size, fontname="helv",
                                                            color=(0, 0, 0), align=0)
                                        self._log(
                                            f"Успешная вставка (fallback helv) на странице {page_num + 1}: '{new_str}'")
                                    except Exception as e2:
                                        self._log(f"Критическая ошибка вставки текста (fallback): {e2}")

                    changes_made = True

            if changes_made:
                doc.save(output_path)
                self._log(f"Файл успешно обработан и сохранен: {output_path}")
            else:
                self._log(f"Изменений не найдено в {input_path}, копируем оригинал")
                doc.save(output_path)

            doc.close()
            return True

        except Exception as e:
            self._log(f"Ошибка обработки PDF {input_path}: {str(e)}")
            return False
