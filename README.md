# Access ↔ Excel GUI

Небольшое desktop-приложение на Python (Tkinter), которое автоматизирует обмен данными между базой Microsoft Access и Excel-книгой.

## Возможности

- Выбор файла Access через проводник.
- Выбор Excel-файла через проводник.
- Выгрузка таблицы из Access в лист Excel (`Access → Excel`).
- Загрузка листа Excel в таблицу Access (`Excel → Access`) с полной очисткой таблицы перед загрузкой.

## Установка

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

> Важно: в системе должен быть установлен ODBC-драйвер Microsoft Access (`Microsoft Access Driver (*.mdb, *.accdb)`).

## Запуск

```bash
python access_excel_gui.py
```

## Сценарий работы

1. Выберите файл Access (`.accdb`/`.mdb`).
2. Выберите Excel-файл (`.xlsx`).
3. Укажите имя таблицы Access и листа Excel.
4. Выберите операцию:
   - `Access → Excel` — выгрузка данных из таблицы в Excel.
   - `Excel → Access` — загрузка данных из Excel с полной перезаписью таблицы.
5. Нажмите **Запустить**.

## Примечания

- Для режима `Excel → Access` названия столбцов в Excel должны совпадать с полями таблицы Access.
- Пустые значения Excel записываются как `NULL`.
