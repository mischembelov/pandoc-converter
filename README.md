# Pandoc Converter

Инструмент для пакетной конвертации документов с использованием [Pandoc](https://pandoc.org/) и простым графическим интерфейсом для Windows.

Подходит как для разработчиков (CLI / скрипты), так и для обычных пользователей (GUI-приложение).

---

## 📋 Возможности

Форматы и режимы:

- **DOCX → Markdown** — извлечение текста из Word-документов в чистый Markdown (без картинок)
- **Markdown → DOCX** — генерация Word-документов из файлов Markdown
- **PDF → DOCX** — конвертация текстовых PDF в редактируемые Word-файлы
- **GUI-приложение**:
  - Drag & Drop файлов
  - Выбор режима конвертации
  - Прогресс-бар и лог
  - Открытие папки с результатом
  - Кнопка «Сохранить файл как (в Загрузки)» — копирует последний результат в системную папку загрузок пользователя

---

## 🚀 Быстрый старт

### 0. Системные требования

- **Windows 10/11** (GUI ориентирован на Windows)
- **Python 3.8+** ([скачать](https://www.python.org/downloads/))
- **Git** ([скачать](https://git-scm.com/downloads))
- Подключение к интернету (для первой установки Pandoc и зависимостей)

---

## 📥 Шаг 1. Загрузка проекта из репозитория

Открой PowerShell и выполни:

```bash
git clone https://github.com/mischembelov/pandoc-converter.git
cd pandoc-converter
```

Проверка содержимого:

```text
pandoc-converter/
├── app.py               # GUI-приложение
├── convert.py           # CLI-скрипт (docx↔md, pdf→docx)
├── setup.py             # Скрипт первичной настройки проекта
├── requirements.txt     # Python-зависимости
├── .gitignore
└── README.md
```

---

## ⚙️ Шаг 2. Настройка локального окружения

### 2.1. Создание и активация виртуального окружения

**Windows (PowerShell):**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

После активации в начале строки должно появиться `(.venv)`.

### 2.2. Установка Python-зависимостей

```powershell
pip install -r requirements.txt
```

Устанавливаются:

- `pdf2docx` — конвертация PDF → DOCX
- `pypandoc` — обёртка вокруг Pandoc
- `customtkinter`, `tkinterdnd2` — GUI и Drag & Drop

### 2.3. Автоматическая настройка проекта и установка Pandoc

```powershell
python setup.py
```

Скрипт:

1. Создаёт рабочие папки:
   - `input/` — DOCX для `to_md`
   - `input_md/` — MD для `to_docx`
   - `input_pdf/` — PDF для `pdf_to_docx`
   - `output/` — результат `.md`
   - `output_docx/` — результат `.docx`
2. Устанавливает зависимости из `requirements.txt`
3. Проверяет наличие Pandoc и при необходимости скачивает его через `pypandoc`

> Если автоматическая установка Pandoc не удалась — скачай установщик вручную с [pandoc.org/installing.html](https://pandoc.org/installing.html) и переустанови.

---

## 🪟 Готовый `.exe` для Windows

Для пользователей Windows доступен готовый исполняемый файл, собранный из этого репозитория.

### Как скачать и запустить

1. Открой страницу проекта на GitHub:  
   `https://github.com/mischembelov/pandoc-converter`
2. Перейди в раздел **Releases** (если есть) или в **Code → dist/**.
3. Скачай файл `Pandoc Converter.exe`.
4. Перемести его в удобную папку (например, `C:\Programs\PandocConverter\`).
5. Запусти двойным кликом по `Pandoc Converter.exe`.

> При первом запуске Windows может показать предупреждение SmartScreen.  
> Нажми «Подробнее» → «Выполнить в любом случае», если доверяешь источнику.

---

## 🖥 Шаг 3. Использование GUI-приложения (рекомендуется для пользователей)

### Запуск GUI из исходников (Windows)

Из корня проекта при активном виртуальном окружении:

```powershell
python app.py
```

Откроется окно **Pandoc Converter**.

### Элементы интерфейса

- **Зона Drag & Drop** — можно перетащить файлы мышкой
- **Кнопка «✕»** — очистка списка добавленных файлов
- **Кнопка «Выбрать файлы»** — стандартный диалог выбора файлов
- **Режим конвертации**:
  - `DOCX → Markdown`
  - `Markdown → DOCX`
  - `PDF → DOCX`
- **Кнопка «Конвертировать»** — запускает обработку всех добавленных файлов
- **Прогресс-бар** — показывает ход выполнения
- **Лог** — показывает список файлов, статусы ✅/❌ и ошибки
- **Кнопка «Открыть папку с результатом»** — открывает `output/` или `output_docx/` в проводнике
- **Кнопка «Сохранить файл как (в Загрузки)»** — копирует *последний сконвертированный файл* в папку `Загрузки` пользователя

### Сценарий использования GUI

1. Запусти `python app.py` или готовый `Pandoc Converter.exe`.
2. Перетащи файлы в зону Drag & Drop или нажми «Выбрать файлы».
3. Выбери режим конвертации.
4. Нажми «Конвертировать».
5. После завершения:
   - Посмотри лог (успехи/ошибки).
   - Нажми «Открыть папку с результатом», если нужно перейти в рабочую папку.
   - Нажми «Сохранить файл как (в Загрузки)», чтобы положить последнюю конвертацию в `Downloads`.

> Ограничение: «Сохранить файл как (в Загрузки)» всегда берёт **последний успешно сконвертированный файл** в текущем запуске.

---

## 🍏 Запуск на macOS

Готового `.app` / `.dmg` в репозитории пока нет, но приложение можно запустить из исходников.

### Вариант 1. Запуск через Python (проще всего)

1. Установить **Homebrew** (если ещё не установлен):  
   [https://brew.sh](https://brew.sh)
2. Установить Python 3:

   ```bash
   brew install python
   ```

3. Клонировать репозиторий:

   ```bash
   git clone https://github.com/mischembelov/pandoc-converter.git
   cd pandoc-converter
   ```

4. Создать и активировать виртуальное окружение:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

5. Установить зависимости:

   ```bash
   pip install -r requirements.txt
   ```

6. Установить Pandoc:

   ```bash
   brew install pandoc
   ```

7. Запустить GUI:

   ```bash
   python3 app.py
   ```

Откроется то же самое окно **Pandoc Converter**, что и на Windows.

### Вариант 2. Сборка `.app` (для двойного клика)

Если вы на macOS и хотите полноценное приложение:

1. Установите PyInstaller:

   ```bash
   pip install pyinstaller
   ```

2. В корне проекта выполните:

   ```bash
   pyinstaller --windowed --name "Pandoc Converter" app.py
   ```

3. В папке `dist/` появится:

   ```text
   dist/
   └── Pandoc Converter.app
   ```

Его можно запускать двойным кликом и перенести в `/Applications`.

> macOS может блокировать запуск приложений не из App Store.  
> В этом случае откройте «System Settings → Privacy & Security» и разрешите запуск приложения.

---

## 🔄 Шаг 4. Использование CLI-скрипта `convert.py` (для разработчиков / автоматизации)

Скрипт поддерживает три режима:

```bash
python convert.py to_md          # DOCX → MD
python convert.py to_docx        # MD → DOCX
python convert.py pdf_to_docx    # PDF → DOCX
```

### DOCX → Markdown

1. Положи `.docx` в `input/`.
2. Выполни:

   ```powershell
   python convert.py to_md
   ```

3. Забери `.md` из папки `output/`.

Особенности:

- Сохраняются заголовки, списки, таблицы, базовое форматирование
- Изображения не извлекаются (упор на текст)

### Markdown → DOCX

1. Положи `.md` в `input_md/`.
2. Выполни:

   ```powershell
   python convert.py to_docx
   ```

3. Результат появится в `output_docx/`.

### PDF → DOCX

1. Положи текстовые `.pdf` в `input_pdf/`.
2. Выполни:

   ```powershell
   python convert.py pdf_to_docx
   ```

3. Готовые `.docx` будут в `output_docx/`.

> PDF должны содержать текстовый слой (текст выделяется мышкой). Для сканов требуется отдельный OCR-пайплайн.

---

## 📂 Структура проекта после настройки

```text
pandoc-converter/
├── .venv/                     # виртуальное окружение (не коммитится)
├── app.py                     # GUI (CustomTkinter + Drag & Drop)
├── convert.py                 # CLI-конвертер
├── setup.py                   # автоматическая настройка проекта
├── requirements.txt
├── README.md
├── .gitignore
├── input/                     # .docx для режима to_md
├── input_md/                  # .md для режима to_docx
├── input_pdf/                 # .pdf для режима pdf_to_docx
├── output/                    # результаты .md
└── output_docx/               # результаты .docx
```

> Папки `input*` и `output*` исключены из Git — они используются только локально.

---

## 🛠 Типичные проблемы

### Pandoc не найден

Сообщение вида:

```text
pandoc: command not found
```

Решение:

1. Проверь:

   ```powershell
   pandoc --version
   ```

2. Если не найден — установи вручную:

   - Windows: скачай `.msi` с [pandoc.org/installing.html](https://pandoc.org/installing.html) и установи.
   - macOS: `brew install pandoc`.

3. После установки перезапусти терминал / IDE.

---

### Ошибка `ModuleNotFoundError` для библиотек

Если видишь:

```text
ModuleNotFoundError: No module named 'pdf2docx'
```

Проверь, что:

1. Виртуальное окружение активно (`(.venv)` в начале строки).
2. Выполнена установка:

   ```powershell
   pip install -r requirements.txt
   ```

---

### PDF конвертируется «криво» или пустой DOCX

Чаще всего это сканированный PDF (картинки страниц без текста). Решение:

- Пропусти файл через OCR (например, OCRmyPDF + Tesseract), чтобы добавить текстовый слой.
- После этого повтори конвертацию PDF → DOCX.

## 📄 Лицензия

MIT License — свободное использование и модификация.



