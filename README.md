# PDFParser

## Описание логики работы

Приложение PDFParser выполняет следующие основные шаги:

1. **Создание нового файла Excel**:
    - Создается новый рабочий файл Excel с заголовками для колонок: "Номер заказа", "Номер заказа и КТ", "Дата", "Сумма", "Базовая станция".
    - Определяется стиль для ячеек с красным фоном (`red_fill`).

2. **Выбор и обработка файлов**:
    - Пользователь выбирает файл Excel (файл 7.15.2) с помощью графического интерфейса.
    - Функция `process_files` загружает выбранный Excel-файл и читает данные из первой страницы.
    - Программа получает список всех PDF-файлов из указанной папки (`./pdf`).

3. **Извлечение и обработка данных из PDF**:
    - Для каждого PDF-файла:
        - Открывается PDF-файл и читается количество страниц.
        - Извлекается текст с первой страницы для получения номера заказа и даты.
        - Ищутся совпадения с помощью регулярных выражений для извлечения базовых станций (БС) и сумм на каждой странице.
        - Извлекаются и сохраняются данные о заказе, дате, сумме и базовых станциях.

4. **Сравнение данных с данными в Excel**:
    - Для каждого найденного совпадения:
        - Сравниваются суммы и базовые станции с данными из выбранного Excel-файла.
        - Если данные совпадают, соответствующий номер заказа добавляется в новый Excel-файл.
        - Если данных нет или они не совпадают, ячейки заполняются красным фоном и добавляется значение "NULL".

5. **Сохранение результатов**:
    - Результаты обработки сохраняются в новый Excel-файл `Результат.xlsx`.

6. **Отправка отчета**:
    - Функция `send_report` отправляет HTTP-запрос с информацией о процессе обработки.

Программа предоставляет графический интерфейс для выбора файлов и запуска процесса обработки, используя библиотеку `tkinter`.

# Pyinstaller
pyinstaller --onefile --hidden-import PyPDF2 PDFHandler.py
