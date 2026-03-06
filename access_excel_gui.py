import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    import pyodbc
except ImportError:  # pragma: no cover
    pyodbc = None


class AccessExcelSyncApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Access ↔ Excel Sync")
        self.root.geometry("760x360")

        self.access_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.table_name = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Sheet1")
        self.operation = tk.StringVar(value="access_to_excel")

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="Файл Access (.accdb / .mdb):").grid(row=0, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.access_path, width=75).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(main, text="Выбрать", command=self._choose_access).grid(row=1, column=1, sticky="ew")

        ttk.Label(main, text="Файл Excel (.xlsx):").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(main, textvariable=self.excel_path, width=75).grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(main, text="Выбрать", command=self._choose_excel).grid(row=3, column=1, sticky="ew")

        ttk.Label(main, text="Таблица Access:").grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(main, textvariable=self.table_name, width=30).grid(row=5, column=0, sticky="w")

        ttk.Label(main, text="Лист Excel:").grid(row=4, column=1, sticky="w", pady=(12, 0))
        ttk.Entry(main, textvariable=self.sheet_name, width=20).grid(row=5, column=1, sticky="ew")

        ops_frame = ttk.LabelFrame(main, text="Операция", padding=10)
        ops_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(16, 0))

        ttk.Radiobutton(
            ops_frame,
            text="Access → Excel (выгрузка)",
            value="access_to_excel",
            variable=self.operation,
        ).pack(anchor="w")

        ttk.Radiobutton(
            ops_frame,
            text="Excel → Access (полная перезапись таблицы)",
            value="excel_to_access",
            variable=self.operation,
        ).pack(anchor="w", pady=(4, 0))

        self.run_btn = ttk.Button(main, text="Запустить", command=self._run)
        self.run_btn.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(16, 0))

        note = (
            "Важно: для работы нужен установленный драйвер ODBC для Microsoft Access.\n"
            "Перед загрузкой в Access таблица будет очищена (DELETE FROM)."
        )
        ttk.Label(main, text=note, foreground="#555").grid(row=8, column=0, columnspan=2, sticky="w", pady=(14, 0))

        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)

    def _choose_access(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите файл Access",
            filetypes=[("Access DB", "*.accdb *.mdb"), ("All files", "*.*")],
        )
        if path:
            self.access_path.set(path)

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    def _validate(self) -> bool:
        if pyodbc is None:
            messagebox.showerror("Ошибка", "Не установлен pyodbc. Установите зависимости из requirements.txt")
            return False

        if not self.access_path.get() or not os.path.isfile(self.access_path.get()):
            messagebox.showerror("Ошибка", "Выберите корректный файл Access")
            return False

        if not self.excel_path.get() or (
            self.operation.get() == "excel_to_access" and not os.path.isfile(self.excel_path.get())
        ):
            messagebox.showerror("Ошибка", "Выберите корректный файл Excel")
            return False

        if not self.table_name.get().strip():
            messagebox.showerror("Ошибка", "Укажите таблицу Access")
            return False

        if not self.sheet_name.get().strip():
            messagebox.showerror("Ошибка", "Укажите лист Excel")
            return False

        return True

    def _connect(self):
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.access_path.get()};"
        )
        return pyodbc.connect(conn_str)

    def _run(self) -> None:
        if not self._validate():
            return

        self.run_btn.configure(state="disabled")
        try:
            if self.operation.get() == "access_to_excel":
                self._access_to_excel()
            else:
                self._excel_to_access()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))
        finally:
            self.run_btn.configure(state="normal")

    def _access_to_excel(self) -> None:
        table = self.table_name.get().strip()
        sheet = self.sheet_name.get().strip()

        with self._connect() as conn:
            df = pd.read_sql(f"SELECT * FROM [{table}]", conn)

        mode = "a" if os.path.exists(self.excel_path.get()) else "w"
        with pd.ExcelWriter(self.excel_path.get(), engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

        messagebox.showinfo("Готово", f"Успешно выгружено {len(df)} строк в Excel лист '{sheet}'.")

    def _excel_to_access(self) -> None:
        table = self.table_name.get().strip()
        sheet = self.sheet_name.get().strip()

        df = pd.read_excel(self.excel_path.get(), sheet_name=sheet)
        if df.empty:
            raise ValueError("Лист Excel пустой. Нечего загружать.")

        columns = [str(col).strip() for col in df.columns]
        if any(not col for col in columns):
            raise ValueError("Обнаружены пустые названия колонок в Excel.")

        placeholders = ",".join(["?"] * len(columns))
        quoted_cols = ",".join([f"[{c}]" for c in columns])

        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM [{table}]")

            insert_sql = f"INSERT INTO [{table}] ({quoted_cols}) VALUES ({placeholders})"
            values = [tuple(None if pd.isna(val) else val for val in row) for row in df.itertuples(index=False, name=None)]
            cursor.fast_executemany = True
            cursor.executemany(insert_sql, values)
            conn.commit()

        messagebox.showinfo("Готово", f"Успешно загружено {len(df)} строк в таблицу Access '{table}'.")


if __name__ == "__main__":
    root = tk.Tk()
    app = AccessExcelSyncApp(root)
    root.mainloop()
