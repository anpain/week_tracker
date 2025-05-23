# Генератор отчетов WEEEK -> Google Sheets (Apps Script)

## Предварительные требования

1.  **API WEEEK**
2.  **Аккаунт Google**

## Настройка Скрипта

1.  **Откройте Google Таблицу**, в которую будет выгружаться отчет.
2.  Перейдите в **Расширения > Apps Script**.
3.  **Скопируйте код:**
    *   Удалите содержимое и вставьте в `Code.gs` предоставленный код скрипта.
4.  **Сохраните скрипт** (значок дискеты).
6.  Нажмите **Выполнить** сверху слева и предоставьте разрешения
    *   **Дополнительные настройки** -> **Перейти на страницу "ВАШ ПРОЕКТ" (небезопасно)** -> **Выбрать все** -> **Продолжить**
5.  **Настройте ID Таблицы в коде:**
    *   Перейдите в настройки проекта в **Apps Script**.
    *   Внизу станицы будут **свойства скрипта**, перейдите в них.
    *   Свойство: `GOOGLE_SHEET_ID`, значение: `ID ТАБЛИЦЫ`
    *   Свойство: `WEEEK_API_TOKEN`, значение: `API TOKEN`
    *   Сохраните свойства

## Использование

*   **Ручной запуск:** Выберите в меню **Отчет WEEEK -> Сгенерировать Отчет**.
*   **Автоматический запуск:**
    1.  В редакторе Apps Script перейдите в раздел "Триггеры" (значок будильника слева).
    2.  Нажмите "+ Добавить триггер".
    3.  Выберите функцию: `runReportGeneration`
    4.  Выберите источник события: "На основе времени".
    5.  Настройте расписание (например, "Таймер по дням", "Каждый день в 9:00").
    6.  Сохраните триггер.
