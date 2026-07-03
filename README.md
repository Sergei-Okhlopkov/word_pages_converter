# Pages Converter (Windows)

Набор инструментов для работы с документами: конвертация Word в DOCX с изображениями страниц и создание гиперссылок на задачи в tracker.rddm.team.

## 📑 Table of Contents

- [English Instructions](#english-instructions)
- [Инструкция на русском](#инструкция-на-русском)

---

## English Instructions

### Overview

**Pages Converter** is a desktop app with a sidebar of tools:

1. **Word to «PDF in Word» conversion** — takes an input Word document and creates a new `.docx` where each original page is inserted as an image.
2. **Project hyperlinks** — reads a CSV or Excel file and turns issue numbers in the first column (from row 2 onward) into clickable hyperlinks to `https://tracker.rddm.team/issues/{number}`. The result is always saved as Excel (`.xlsx`).

The app supports light and dark themes (toggle at the bottom of the sidebar). The selected theme is remembered between sessions.

### Requirements

- Windows
- Installed Microsoft Word (required for the Word conversion tool)
- Python 3.10+

### Quick Start

    py -3 -m venv .venv
    .\.venv\Scripts\Activate.ps1
    pip install -r requirements.txt
    python app.py

### Build EXE

    build.bat

After build, the executable will be available at:

- `dist/WordPagesToImages.exe`

### Tool 1: Word to «PDF in Word»

1. Select **«Конвертация Word в «PDF в Word'е»»** in the sidebar
2. Click **«Выбрать файл»** and pick the input file (`.doc`, `.docx`, `.rtf`, `.odt`, `.txt`, `.html`, etc.)
3. Click **«Конвертировать»**

The output file is created automatically next to the input file as `Converted_<name>.docx` (or `Converted_<name>1.docx` if the name already exists).

### Tool 2: Project hyperlinks

1. Select **«Гиперссылки для проектов»** in the sidebar
2. Click **«Выбрать файл»** and pick a CSV or Excel file (`.csv`, `.xlsx`, `.xlsm`)
3. Click **«Создать гиперссылки»**

The first column (`#`) is processed from row 2: each issue number becomes a hyperlink. The output is saved next to the input file as `Hyper_<name>.xlsx` (or `Hyper_<name>1.xlsx` if the name already exists).

### Important

- During Word conversion, the document must not be open in another application
- Image quality depends on DPI (currently set to 220 in code)
- For hyperlinks, only the first column from row 2 is processed; the header row is left unchanged

---

## Инструкция на русском

### Описание

**Pages Converter** — настольное приложение с боковым меню инструментов:

1. **Конвертация Word в «PDF в Word'е»** — берёт входной Word-документ и создаёт новый `.docx`, где каждая исходная страница вставлена как изображение.
2. **Гиперссылки для проектов** — загружает CSV или Excel: номера в первом столбце (со 2-й строки) преобразуются в кликабельные гиперссылки на `https://tracker.rddm.team/issues/{номер}`. Результат всегда сохраняется в Excel (`.xlsx`).

Приложение поддерживает светлую и тёмную тему (переключатель внизу бокового меню). Выбранная тема сохраняется между запусками.

### Требования

- Windows
- Установленный Microsoft Word (обязательно для инструмента конвертации Word)
- Python 3.10+

### Быстрый запуск

    py -3 -m venv .venv
    .\.venv\Scripts\Activate.ps1
    pip install -r requirements.txt
    python app.py

### Сборка EXE

    build.bat

После сборки файл будет доступен:

- `dist/WordPagesToImages.exe`

### Инструмент 1: Конвертация Word в «PDF в Word'е»

1. В боковом меню выберите **«Конвертация Word в «PDF в Word'е»»**
2. Нажмите **«Выбрать файл»** и укажите входной файл (`.doc`, `.docx`, `.rtf`, `.odt`, `.txt`, `.html` и др.)
3. Нажмите **«Конвертировать»**

Выходной файл создаётся автоматически рядом с исходным: `Converted_<имя>.docx` (или `Converted_<имя>1.docx`, если имя уже занято).

### Инструмент 2: Гиперссылки для проектов

1. В боковом меню выберите **«Гиперссылки для проектов»**
2. Нажмите **«Выбрать файл»** и выберите CSV или Excel (`.csv`, `.xlsx`, `.xlsm`)
3. Нажмите **«Создать гиперссылки»**

Обрабатывается первый столбец (`#`) со 2-й строки: каждый номер задачи становится гиперссылкой. Результат сохраняется рядом с исходным файлом как `Hyper_<имя>.xlsx` (или `Hyper_<имя>1.xlsx`, если имя уже занято).

### Важно

- Во время конвертации Word-документ не должен быть открыт в другом приложении
- Качество изображений зависит от DPI (в коде сейчас 220)
- Для гиперссылок обрабатывается только первый столбец со 2-й строки; строка заголовка не изменяется
