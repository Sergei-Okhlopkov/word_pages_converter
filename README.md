# Word Pages to Images (Windows)

## 📑 Table of Contents

- [English Instructions](#english-instructions)
- [Инструкция на русском](#инструкция-на-русском)

---

## English Instructions

### Overview

The application takes an input Word document and creates a new `.docx` file where each original page is inserted as an image.

### Requirements

- Windows
- Installed Microsoft Word (required)
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

### Usage

1. Click `Browse...` and select the input file (`.doc`, `.docx`, `.rtf`, `.odt`, `.txt`, `.html`, etc.)
2. Click `Save as...` and choose the output `.docx` path
3. Click `Convert`

### Important

- The document must not be open in another application during conversion
- Image quality depends on DPI (currently set to 220 in code)

---

## Инструкция на русском

### Описание

Приложение берёт входной Word-документ и создаёт новый `.docx`, где каждая исходная страница вставлена как изображение.

### Требования

- Windows
- Установленный Microsoft Word (обязательно)
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

### Как использовать

1. Нажмите `Обзор...` и выберите входной файл (`.doc`, `.docx`, `.rtf`, `.odt`, `.txt`, `.html` и др.)
2. Нажмите `Сохранить как...` и укажите путь для `.docx`
3. Нажмите `Конвертировать`

### Важно

- Во время конвертации документ не должен быть открыт в другом приложении
- Качество изображений зависит от DPI (в коде сейчас 220)
