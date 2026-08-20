# Парсер сметы — себестоимость из прайса

Скрипт берёт `.xls` файл со сметой (из Базис-Мебельщик), подтягивает цены из базы поставщика и сохраняет результат в `.xlsx`.

## Быстрый запуск

### 1. Установить Python 3.9+

**Windows:** https://python.org — скачать установщик, при установке поставить галочку **"Add Python to PATH"**.

Проверить после установки (в терминале/командной строке):

```
python --version
```

### 2. Скачать скрипт

Скопировать папку `estimate_parser` целиком на компьютер.

### 3. Установить зависимости

Открыть терминал (командную строку) в папке `estimate_parser`:

```
pip install -r requirements.txt
```

### 4. Положить смету

Положить `.xls` файл в папку `xls/` (если папки нет — создать рядом со скриптом).

### 5. Запустить

```
python estimate_parser.py
```

Готовый файл `{имя}_с_ценами.xlsx` появится рядом со скриптом.

## Переменные окружения (обязательно)

Скрипт берёт параметры подключения к БД из переменных окружения. Задайте их перед запуском:

**Linux/macOS:**
```
export KITCHEN_DB_HOST=your_db_host
export KITCHEN_DB_NAME=your_db_name
export KITCHEN_DB_USER=your_db_user
export KITCHEN_DB_PASSWORD=your_db_password
export KITCHEN_DB_PORT=5432
python estimate_parser.py
```

**Windows (cmd):**
```
set KITCHEN_DB_HOST=your_db_host
set KITCHEN_DB_NAME=your_db_name
set KITCHEN_DB_USER=your_db_user
set KITCHEN_DB_PASSWORD=your_db_password
set KITCHEN_DB_PORT=5432
python estimate_parser.py
```

**Windows (PowerShell):**
```
$env:KITCHEN_DB_HOST="your_db_host"
$env:KITCHEN_DB_NAME="your_db_name"
$env:KITCHEN_DB_USER="your_db_user"
$env:KITCHEN_DB_PASSWORD="your_db_password"
$env:KITCHEN_DB_PORT="5432"
python estimate_parser.py
```

Либо скопировать `.env.example` в `.env` — скрипт подхватит автоматически (установка `python-dotenv` уже в `requirements.txt`).

## Формат входного файла

Ожидается стандартный экспорт сметы из Базис-Мебельщик. В первых 20 строках должен быть заголовок с колонками:

- **Наименование материала**
- **Количество в изделии**
- **Ед. изм.**

## Формат выходного файла

Колонки: Наименование, Количество, Ед. изм., Цена за ед., Сумма.  
Внизу строка **Итого** с общей стоимостью.

Для позиций ЛДСП, ДСП и ХДФ цена за лист делится на **5.8** (пересчёт на м²).
