# **WESA_Parser**

File handler (Word, Excel, SmartSketch SHA and AutoCAD DWG).

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)

[![Code Style: Black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

## **Description**

WESA_Parser is a graphical Python application with the Tkinter interface designed for automated file processing. The program replaces the digit (or any other value specified by the user in the code) in the file names and their contents, updates the revision to "C01", clears the change tables (in Word) and performs other specific substitutions. Supported formats:

* Word: .doc, .docx, .dotx
* Excel: .xls, .xlsx, .xlsm (with conversion .xls to .xlsm)
* AutoCAD: .dwg
* SmartSketch: .sha

The application only works on Windows, as it uses COM interfaces (win32com) to interact with AutoCAD, Excel, Word and SmartSketch.

Author: Artem Bayushkin 

Version: 0.95 beta

## **Functions**

Replacing the block number (for example, "1" with "2") in the text of files and their names.
Updating the revision to "C01" (for example, "C02" → "C01").
Clearing the "Change Registration Sheet" or "Record of revisions" tables in Word.
Processing of specific patterns such as "ED.D.P000.N", "10UKD", "Unit N", etc.
Logging the process (to a file log.txt and GUI).
Debugging mode for detailed logs.
Automatic reinitialization of AutoCAD in case of errors.

## **Requirements**

* OS: Windows (64-bit recommended).
* Python: 3.8+.

### **Installed programs:**

* AutoCAD (for processing .dwg).
* Microsoft Office (Excel and Word for conversion and processing).
* SmartSketch (for .sha, with license settings in the registry).


### **Python Libraries:**

* tkinter (built-in).
* win32com (install via pip install pywin32).
* psutil (for managing AutoCAD processes).
* lxml (for parsing XML in Word/Excel).
* PIL (Pillow for GUI icons).
* pythoncom, pywintypes (from pywin32).
* Others: os, re, glob, shutil, zipfile, datetime.

**Install dependencies**:

`pip install pywin32 psutil lxml pillow`

## **Installation**

Clone the repository

`git clone https://github.com/ArtemBayushkin/wesa_parser`

`cd yourrepo`

Install the dependencies (see above).

Make sure that AutoCAD, Office, and SmartSketch are installed and licensed.

(Optional) Compile to EXE using PyInstaller:

`pyinstaller --onefile --icon=icon.ico main.py`

## **Using**

1. Launch the app:

`python main.py`

2. In the interface:
* Select the block number (1-4).
* Specify the folder with the source files.
* Specify the output folder.
* Enable "Debugging Logs" if necessary.
* Click "Start processing".

3. Results:
* The processed files are saved in the specified folder with new names.
* The log is recorded in log.txt in the output folder.
* The GUI displays key messages (success/errors).

Processing example:

Source file: 10UKD.docx → New: 20UKD.docx with replacements inside.

**Attention**: .sha requires a running SmartSketch with a license. The program automatically starts/closes it.

## **Contributing**

If you want to make changes:

1. Fork the repository.
2. Create a branch: `git checkout -b feature/new-feature'.
3. Commit the changes: `git commit -m "Added new feature"'.
4. Launch: `git push origin feature/new-feature'.
5. Create A Pull Request.

## **License**

This project is distributed under the MIT license.

## **Contacts**
* Author: Artem Bayushkin 
* Email: artem.bayushkin@mail.ru
* GitHub: ArtemBayushkin

If you have any problems, create an Issue in the repository.

# ======================================
# **WESA_Parser**

Обработчик файлов (Word, Excel, SmartSketch SHA и AutoCAD DWG).

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)

[![Code Style: Black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

## **Описание**

WESA_Parser — это графическое приложение на Python с интерфейсом Tkinter, предназначенное для автоматизированной обработки файлов. Программа заменяет цифру (или любое другое значение, заданное пользователем в коде) в именах файлов и их содержимом, обновляет ревизию на "C01", очищает таблицы изменений (в Word) и выполняет другие специфические замены. Поддерживаются форматы:

* Word: .doc, .docx, .dotx
* Excel: .xls, .xlsx, .xlsm (с конвертацией .xls в .xlsm)
* AutoCAD: .dwg
* SmartSketch: .sha

Приложение работает только на Windows, так как использует COM-интерфейсы (win32com) для взаимодействия с AutoCAD, Excel, Word и SmartSketch.

Автор: Артем Баюшкин 

Версия: 0.95 beta

## **Функции**

Замена номера блока (например, "1" на "2") в тексте файлов и их именах.
Обновление ревизии на "C01" (например, "C02" → "C01").
Очистка таблиц "Лист регистрации изменений" или "Record of revisions" в Word.
Обработка специфических паттернов, таких как "ED.D.P000.N", "10UKD", "Unit N" и т.д.
Логирование процесса (в файл log.txt и GUI).
Отладочный режим для детальных логов.
Автоматическая переинициализация AutoCAD при ошибках.

## **Требования**

* ОС: Windows (64-bit рекомендуется).
* Python: 3.8+.

### **Установленные программы:**

* AutoCAD (для обработки .dwg).
* Microsoft Office (Excel и Word для конвертации и обработки).
* SmartSketch (для .sha, с настройкой лицензии в реестре).


### **Библиотеки Python:**

* tkinter (встроенная).
* win32com (установите через pip install pywin32).
* psutil (для управления процессами AutoCAD).
* lxml (для парсинга XML в Word/Excel).
* PIL (Pillow для иконок GUI).
* pythoncom, pywintypes (из pywin32).
* Другие: os, re, glob, shutil, zipfile, datetime.

**Установите зависимости**:

`pip install pywin32 psutil lxml pillow`

## **Установка**

Клонируйте репозиторий

`git clone https://github.com/ArtemBayushkin/wesa_parser`

`cd yourrepo`

Установите зависимости (см. выше).

Убедитесь, что AutoCAD, Office и SmartSketch установлены и лицензированы.

(Опционально) Скомпилируйте в EXE с помощью PyInstaller:

`pyinstaller --onefile --icon=icon.ico main.py`

## **Использование**

1. Запустите приложение:

`python main.py`

2. В интерфейсе:
* Выберите номер блока (1-4).
* Укажите папку с исходными файлами.
* Укажите папку для вывода.
* Включите "Отладочные логи" при необходимости.
* Нажмите "Запустить обработку".

3. Результаты:
* Обработанные файлы сохраняются в указанной папке с новыми именами.
* Лог записывается в log.txt в папке вывода.
* В GUI отображаются ключевые сообщения (успех/ошибки).

Пример обработки:

Исходный файл: 10UKD.docx → Новый: 20UKD.docx с заменами внутри.

**Внимание**: Для .sha требуется запущенный SmartSketch с лицензией. Программа автоматически запускает/закрывает его.

## **Контрибьютинг**

Если хотите внести изменения:

1. Форкните репозиторий.
2. Создайте ветку: `git checkout -b feature/new-feature`.
3. Зафиксируйте изменения: `git commit -m "Добавлена новая функция"`.
4. Запушьте: `git push origin feature/new-feature`.
5. Создайте Pull Request.

## **Лицензия**

Этот проект распространяется под лицензией MIT.

## **Контакты**
* Автор: Артем Баюшкин 
* Email: artem.bayushkin@mail.ru
* GitHub: ArtemBayushkin

Если возникли проблемы, создайте Issue в репозитории.