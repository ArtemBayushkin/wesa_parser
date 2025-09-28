import os, sys
import re
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
from excel_parser import ExcelProcessor
from word_parser import WordProcessor
from dwg_parser import AutoCADProcessor
from sha_parser import ShaProcessorWinAPI
from PIL import Image, ImageTk

class FileProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик Excel, Word, DWG и SHA для АЭС \"Эль-Дабаа\"")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        self.replacement_digit = tk.StringVar()
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.debug_logging = tk.BooleanVar(value=False)  # Галочка для отладочных логов
        self.log_file = None
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Выберите Блок:").pack(anchor="w", padx=10, pady=5)

        frame_right = tk.Frame(self.root, width=200, height=120, bg="SystemButtonFace")
        frame_right.place(x=480, y=30)

        lbl_text = tk.Label(frame_right, text="АЭС \"Эль-Дабаа\"", font=("Arial", 14, "italic"))
        lbl_text.pack(expand=True)

        self.replacement_digit.set("1")

        frame_digits = tk.Frame(self.root)
        frame_digits.pack(anchor="w", padx=10)

        digits = [("1", "1"), ("2", "2"), ("3", "3"), ("4", "4")]
        for text, value in digits:
            tk.Radiobutton(frame_digits, text=text, variable=self.replacement_digit,
                           value=value).pack(side="left", padx=5)



        tk.Label(self.root, text="Папка с исходными файлами:").pack(anchor="w", padx=10, pady=5)
        frame_in = tk.Frame(self.root)
        frame_in.pack(anchor="w", padx=10)
        tk.Entry(frame_in, textvariable=self.input_dir, width=50).pack(side="left")
        tk.Button(frame_in, text="Выбрать...", command=self.choose_input_dir).pack(side="left", padx=5)

        tk.Label(self.root, text="Папка для сохранения новых файлов:").pack(anchor="w", padx=10, pady=5)
        frame_out = tk.Frame(self.root)
        frame_out.pack(anchor="w", padx=10)
        tk.Entry(frame_out, textvariable=self.output_dir, width=50).pack(side="left")
        tk.Button(frame_out, text="Выбрать...", command=self.choose_output_dir).pack(side="left", padx=5)

        tk.Checkbutton(self.root, text="Отладочные логи", variable=self.debug_logging).pack(anchor="w", padx=10, pady=5)
        btn_run = tk.Button(frame_right, text="Запустить обработку",
                            command=self.run_processing,
                            bg="green", fg="white", font=("Arial", 11), padx=5, pady=5)
        btn_run.pack(pady=50)

        tk.Label(self.root, text="Процесс обработки:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self.root, width=80, height=5)
        self.log_text.pack(padx=10, pady=5, fill="both", expand=True)
        # Настройка тегов для форматирования текста
        self.log_text.tag_configure("error", foreground="red", font=("Arial", 10, "bold"))
        self.log_text.tag_configure("skip", foreground="red", font=("Arial", 10, "bold"))


        tk.Label(self.root, text="by UKA (Артем Баюшкин)", font=("Arial", 9, "italic")).pack(anchor="e", padx=10, pady=5)
        tk.Button(self.root, text="О программе", command=self.show_about).pack(anchor="e", padx=10, pady=5)

    def choose_input_dir(self):
        folder = filedialog.askdirectory(title="Выберите папку с исходными файлами")
        if folder:
            self.input_dir.set(folder)

    def choose_output_dir(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder:
            self.output_dir.set(folder)

    def log_to_file(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        if self.log_file:
            try:
                self.log_file.write(log_message + "\n")
                self.log_file.flush()
            except Exception as e:
                self.log_to_gui(f"Ошибка записи в лог-файл: {str(e)}")

    def log_to_gui(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        # Отображаем только указанные сообщения в GUI
        if (message.startswith("=== Запуск обработки ===") or
            message.startswith("Успешно: ") or
            message.startswith("Ошибка обработки: ") or
            message.startswith("Обработка завершена. ")):
            # Применяем тег "error" для сообщений об ошибках
            tag = "error" if message.startswith("Ошибка обработки: ") else None
            self.log_text.insert(tk.END, log_message + "\n", tag)
            self.log_text.see(tk.END)
            self.root.update()

    def log(self, message):
        # Логи, которые всегда записываются в файл
        always_log = (
            message.startswith("=== Запуск обработки ===") or
            message.startswith("Успешно: ") or
            message.startswith("Ошибка обработки: ") or
            message.startswith("Обработка завершена. ") or
            message.startswith("Результаты сохранены в: ")
        )
        # Если отладочные логи включены или это обязательный лог, записываем в файл
        if self.debug_logging.get() or always_log:
            self.log_to_file(message)
        # Отображаем в GUI только указанные сообщения
        self.log_to_gui(message)

    def select_files(self, input_dir):
        excel_files = glob.glob(os.path.join(input_dir, '*.xls*'))
        word_files = glob.glob(os.path.join(input_dir, '*.do*'))
        dwg_files = glob.glob(os.path.join(input_dir, '*.dwg'))
        sha_files = glob.glob(os.path.join(input_dir, '*.sha'))
        files = excel_files + word_files + dwg_files + sha_files
        return files

    def process_files(self, input_files, output_dir, replacement_digit):
        os.makedirs(output_dir, exist_ok=True)
        processed = 0

        sha_processor = ShaProcessorWinAPI(replacement_digit, log_callback=self.log, debug=self.debug_logging.get())
        sha_app_started = False
        try:
            for input_path in input_files:
                try:
                    filename = os.path.basename(input_path)
                    name, ext = os.path.splitext(filename)

                    if name[0].isdigit():
                        new_name = f"{replacement_digit}{name[1:]}"
                    elif name.startswith("ED.D."):
                        new_name = re.sub(
                            r'(ED\.D\.[A-Z]\d{3}\.)(\d)',
                            lambda m: f"{m.group(1)}{replacement_digit}",
                            name
                        )
                    else:
                        new_name = f"processed_{name}"

                    if ext == ".xls":
                        output_path = os.path.join(output_dir, new_name + ".xlsm")
                    else:
                        output_path = os.path.join(output_dir, new_name + ext)
                    extension = ext.lower()

                    if extension in ('.doc', '.docx', '.dotx'):
                        processor = WordProcessor(replacement_digit, log_callback=self.log, debug=self.debug_logging.get())
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    elif extension in ('.xls', '.xlsx', '.xlsm'):
                        processor = ExcelProcessor(replacement_digit, log_callback=self.log, debug=self.debug_logging.get())
                        success = processor.process_file(input_path, output_path)
                        if success:
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    elif extension == '.dwg':
                        processor = AutoCADProcessor(replacement_digit, log_callback=self.log, debug=self.debug_logging.get())
                        success_list = processor.process_files([input_path], output_dir)
                        if all(success_list.values()):  # Check the values of the dictionary
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    elif extension == '.sha':
                        if not sha_app_started:
                            sha_processor.start_app()
                            sha_app_started = True
                        success = sha_processor.process_file(input_path, output_path)
                        if success:
                            self.log(f"Успешно: {filename}")
                            processed += 1
                        else:
                            self.log(f"Ошибка обработки: {filename}")

                    else:
                        self.log(f"Пропуск {filename} (неподдерживаемый формат: {extension})")

                except Exception as e:
                    self.log(f"Критическая ошибка {filename}: {str(e)}")

        finally:
            if sha_app_started:
                sha_processor.stop_app()

        return processed

    def run_processing(self):
        repl_digit = self.replacement_digit.get().strip()
        if not repl_digit.isdigit():
            messagebox.showerror("Ошибка", "Введите корректную цифру для замены!")
            return

        input_dir = self.input_dir.get().strip()
        output_dir = self.output_dir.get().strip()

        if not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", "Выберите существующую папку с исходными файлами!")
            return
        if not output_dir:
            messagebox.showerror("Ошибка", "Выберите папку для сохранения файлов!")
            return

        try:
            log_file_path = os.path.join(output_dir, "log.txt")
            self.log_file = open(log_file_path, 'a', encoding='utf-8')
            self.log("=== Запуск обработки ===")
        except Exception as e:
            self.log(f"Ошибка открытия лог-файла: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось открыть лог-файл: {str(e)}")
            return

        try:
            input_files = self.select_files(input_dir)
            if not input_files:
                self.log("Файлы не найдены.")
                return

            processed_count = self.process_files(input_files, output_dir, repl_digit)
            self.log(f"Обработка завершена. Успешно обработано: {processed_count}/{len(input_files)}")
            self.log(f"Результаты сохранены в: {output_dir}")

            messagebox.showinfo(
                "Готово",
                f"Обработка завершена.\nУспешно обработано: {processed_count}/{len(input_files)}"
            )
        finally:
            if self.log_file:
                self.log_file.close()
                self.log_file = None
                self.log("Лог-файл закрыт")

    def show_about(self):
        about_win = tk.Toplevel(self.root)
        about_win.title("О программе")
        about_win.geometry("350x250")
        about_win.resizable(False, False)

        text = (
            "ED_Parser\n"
            "Обработчик Excel, Word, DWG и SHA для АЭС \"Эль-Дабаа\"\n"
            "--------------------------------------\n"
            "Как работает:\n"
            "1. Выберите нужный блок для замены.\n"
            "2. Выберите папку с исходными файлами.\n"
            "3. Укажите папку для сохранения.\n"
            "4. Программа переименует файлы, заменит\n"
            "   цифру Блока в содержимом, а также\n"
            "   ревизию на C01.\n\n"
            "Поддерживаемые форматы:\n"
            ".doc, .docx, .dotx, .xls, .xlsx, .xlsm, .dwg, .sha.\n\n"
            "Автор: Артем Баюшкин (UKA)\n"
            "Версия: 0.95 beta"
        )

        lbl = tk.Label(about_win, text=text, justify="left", padx=5, pady=5)
        lbl.pack(fill="both", expand=True)

def set_icon(root, icon_path):
    img = Image.open(icon_path)
    icon = ImageTk.PhotoImage(img)
    root.iconphoto(False, icon)

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    set_icon(root, resource_path("icon.png"))
    app = FileProcessorGUI(root)
    root.mainloop()