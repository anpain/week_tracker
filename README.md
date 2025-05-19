# Генератор отчетов WEEEK с выгрузкой в Google Sheets

## Предварительные требования

1.  **Python**
2.  **Аккаунт WEEEK** с доступом к API.
3.  **Аккаунт Google** и доступ к Google Cloud Platform.
4.  **pip** для установки зависимостей Python.

## Настройка

### 1. Настройка WEEEK API

-   Получите ваш API токен из настроек профиля WEEEK (Настройки -> Интеграции -> API).

### 2. Настройка Google Cloud Platform и Google Sheets API

    1.  **Создайте проект** в [Google Cloud Console](https://console.cloud.google.com/).
    2. Перейдите в **APIs & Services** -> **Library** -> **Google Sheets API**
    2.  **Включите "Google Sheets API"**
    3.  **Создайте Сервисный аккаунт:**
        *   Перейдите в "IAM & Admin" -> "Service accounts".
        *   Создайте новый сервисный аккаунт.
        *   Предоставьте ему роль **"Редактор" (Editor)**.
        *   Перейдите в него и нажмите *Keys*
        *   Создайте для него JSON-ключ и скачайте его. Переименуйте этот файл, например, в `credentials.json`.
    4.  **Поделитесь вашей Google Таблицей** (в которую будут записываться данные) с email-адресом созданного сервисного аккаунта, предоставив ему права **"Редактора"**.
    5.  **Скопируйте ID вашей Google Таблицы**. Его можно найти в URL таблицы: `https://docs.google.com/spreadsheets/d/ВАШ_ID_ТАБЛИЦЫ/edit`.

### 3. Настройка проекта локально/на сервере

1.  **Клонируйте репозиторий (если он есть) или скопируйте файлы скрипта.**
2.  **Создайте и активируйте виртуальное окружение (рекомендуется):**
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    # venv\Scripts\activate
    ```
3.  **Установите зависимости:**
    ```bash
    pip install -r requirements.txt
    ```
    Если файла `requirements.txt` или вручную:
    ```bash
    pip install requests python-dateutil python-dotenv gspread google-api-python-client google-auth-httplib2 google-auth-oauthlib
    ```
4.  **Создайте файл `.env`** в корневой директории проекта (рядом со скриптом `weeek_report_generator.py`) и заполните его следующими данными:
    ```env
    WEEEK_API_TOKEN="ВАШ_WEEEK_API_ТОКЕН"
    GOOGLE_SHEET_ID="ID_ВАШЕЙ_GOOGLE_ТАБЛИЦЫ"
    GOOGLE_CREDENTIALS_FILENAME="credentials.json" # Или полный путь к вашему JSON-ключу
    ```
    Замените значения на ваши реальные данные.
5.  **Поместите файл JSON-ключа** (например, `credentials.json`) в директорию проекта или по пути, указанному в `GOOGLE_CREDENTIALS_FILENAME`. Добавьте его и `credentials` в `.gitignore`.

## Использование

Скрипт можно запустить из командной строки:

```bash
python weeek_report_generator.py