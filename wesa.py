"""
Модуль основной программы: Графический интерфейс для обработки файлов Excel, Word, DWG и SHA.

Это приложение предоставляет GUI на базе Tkinter для выбора проекта, блока, входной/выходной папок
и запуска обработки файлов с заменой текста по правилам из config.json. Поддерживает логирование,
переименование файлов и интеграцию с парсерами для различных форматов.

Зависимости: os, sys, re, glob, tkinter, datetime, PIL, json, shutil, а также импорт парсеров
(excel_parser, word_parser, dwg_parser, sha_parser).

Программа предназначена для автоматизации задач в контексте АЭС (атомных электростанций).
"""
import os
import logging
import tkinter as tk
from Logger import GUILogHandler
from tkinter import filedialog, messagebox, scrolledtext, ttk
from datetime import datetime
from PIL import Image, ImageTk
from config_handler import config_data, config_projects, config_handler
from file_hander import FileHandler


class FileProcessorGUI:
    """Класс для графического интерфейса обработки файлов.

        Создаёт окно Tkinter с элементами для выбора проекта, блока, папок и запуска обработки.
        Загружает конфигурацию из JSON, обрабатывает файлы через специализированные парсеры,
        логирует процесс в файл и GUI.

        Attributes:
            root (tk.Tk): Корневое окно Tkinter.
            replacement_digit (tk.StringVar): Переменная для выбранной цифры замены.
            project (tk.StringVar): Переменная для выбранного проекта.
            input_dir (tk.StringVar): Путь к входной папке.
            output_dir (tk.StringVar): Путь к выходной папке.
            debug_logging (tk.BooleanVar): Флаг отладочного логирования.
            config_data (dict): Загруженная конфигурация из JSON.
            lbl_project (tk.Label): Метка для отображения текущего проекта.
        """
    def __init__(self, root, config_data=config_data):
        """Инициализирует GUI и атрибуты.

            Устанавливает заголовок, размер окна, переменные Tkinter, загружает конфиг
            и вызывает создание виджетов.
            Args:
                root (tk.Tk): Корневое окно приложения.
        """
        self.root = root
        self.root.title("Обработчик Excel, Word, DWG и SHA для АЭС")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        self.replacement_digit = tk.StringVar()
        self.project = tk.StringVar()
        self.input_dir = tk.StringVar()
        self.debug_logging = tk.BooleanVar(value=False)
        self.lbl_project = None
        self.config_data = config_data
        
        self.create_widgets()
        self.setup_logger()
        self.logger.log(logging.INFO, "===Запуск Программы===")
    
    def create_widgets(self):
        tk.Label(self.root, text="Выберите проект:").pack(anchor="w", padx=10, pady=5)
        projects = config_projects
        self.project_combobox = ttk.Combobox(self.root, textvariable=self.project, values=projects, state="readonly")
        self.project_combobox.pack(anchor="w", padx=10, ipadx=50)
        self.project_combobox.bind("<<ComboboxSelected>>", self.update_digits)
        if projects:
            self.project.set(projects[0])

        tk.Label(self.root, text="Выберите Блок:").pack(anchor="w", padx=10, pady=5)

        frame_right = tk.Frame(self.root, width=100, height=120, bg="SystemButtonFace")
        frame_right.place(x=400, y=30)

        self.lbl_project = tk.Label(frame_right, text=self.project.get(), font=("Arial", 14, "italic"))
        self.lbl_project.pack(expand=True)

        self.frame_digits = tk.Frame(self.root)
        self.frame_digits.pack(anchor="w", padx=10)
        self.update_digits()

        tk.Label(self.root, text="Папка с исходными файлами:").pack(anchor="w", padx=10, pady=5)
        frame_in = tk.Frame(self.root)
        frame_in.pack(anchor="w", padx=10)
        tk.Entry(frame_in, textvariable=self.input_dir, width=50).pack(side="left")
        tk.Button(frame_in, text="Выбрать...", command=self.choose_input_dir).pack(side="left", padx=5)

        
        tk.Checkbutton(self.root, text="Отладочные логи", variable=self.debug_logging).pack(anchor="w", padx=10, pady=5)
        btn_run = tk.Button(frame_right, text="Запустить обработку",
                            command=self.run_processing,
                            bg="green", fg="white", font=("Arial", 11), padx=50, pady=5)
        btn_run.pack(pady=50)

        tk.Label(self.root, text="Процесс обработки:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self.root, width=80, height=5)
        self.log_text.pack(padx=10, pady=5, fill="both", expand=True)
        self.log_text.tag_configure("error", foreground="red", font=("Arial", 10, "bold"))
        self.log_text.tag_configure("skip", foreground="red", font=("Arial", 10, "bold"))

        tk.Label(self.root, text="by Артем Баюшкин", font=("Arial", 9, "italic")).pack(anchor="e", padx=10, pady=5)
        tk.Button(self.root, text="О программе", command=self.show_about).pack(anchor="e", padx=10, pady=5)
        tk.Button(self.root, text="Редактировать config", command=config_handler.edit_config).pack(anchor="e", padx=10, pady=5)

    def setup_logger(self):
        self.logger = logging.getLogger(__name__)
        
        self.logger.setLevel(logging.DEBUG)
               
        # Создаем обработчик для текстового поля
        text_handler = GUILogHandler(self.log_text)
        text_handler.setLevel(logging.INFO)
        
        # Форматирование
        formatter = logging.Formatter('%(asctime)s -%(levelname)s- %(message)s', datefmt="%X")
        text_handler.setFormatter(formatter)
        log_to_file_handler = logging.FileHandler(filename='l.txt', encoding='utf-8', mode='w')
        log_to_file_handler.setFormatter(formatter)
        self.logger.addHandler(text_handler)
        self.logger.addHandler(log_to_file_handler)

    def update_digits(self, event=None):
        project = self.project.get()
        if self.lbl_project:
            self.lbl_project.config(text=project)
        digits = self.config_data.get(project, {}).get("digits", [])
        for widget in self.frame_digits.winfo_children():
            widget.destroy()
        if digits:
            self.replacement_digit.set(digits[0][1])
            for text, value in digits:
                tk.Radiobutton(self.frame_digits, text=text, variable=self.replacement_digit,
                               value=value).pack(side="left", padx=5)

    def choose_input_dir(self):
        folder = filedialog.askdirectory(title="Выберите папку с исходными файлами")
        if folder:
            self.input_dir.set(folder)

    def run_processing(self):
        
        project = self.project.get()
        input_dir = self.input_dir.get().strip()
        repl_digit = self.replacement_digit.get().strip()
        
        if not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", "Выберите существующую папку с исходными файлами!")
            return
        
        file_handler = FileHandler(input_dir, project, repl_digit, logger=self.logger)
        file_handler.process_files()   

        messagebox.showinfo(
            "Готово",
            f"Обработка завершена.\nУспешно обработано: {file_handler.processed_files_counter}/{len(file_handler.files)}"
        )

    def show_about(self):
        about_win = tk.Toplevel(self.root)
        about_win.title("О программе")
        about_win.geometry("400x400")
        about_win.resizable(False, False)

        text = (
            "WESA_Parser\n"
            "Обработчик Excel, Word, DWG и SHA\n\n"
            "--------------------------------------\n\n"
            "Принцип работы программы:\n\n"
            "1. Выберите проект и блок, НА который\n"
            "необходимо произвести замену.\n\n"
            "2. Выберите папку с исходными файлами.\n\n"
            "3. Укажите папку для сохранения.\n\n"
            "4. Программа переименует файлы, заменит\n"
            "   цифру Блока в содержимом, а также\n"
            "   ревизию на C01.\n\n"
            "Поддерживаемые форматы:\n"
            ".doc, .docx, .dotx, .xls, .xlsx, .xlsm, .dwg, .sha.\n\n"
            "Автор: Артем Баюшкин\n"
            "Версия: 1.02 "
            "   MIT License"
        )

        lbl = tk.Label(about_win, text=text, justify="left", padx=5, pady=5)
        lbl.pack(fill="both", expand=True)

def set_icon(root, icon_path):
    img = Image.open(icon_path)
    icon = ImageTk.PhotoImage(img)
    root.iconphoto(False, icon)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileProcessorGUI(root)
    set_icon(root, config_handler.get_relative_path("icon.ico"))
    root.mainloop()