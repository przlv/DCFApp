import tkinter as tk
from tkinter import ttk, filedialog
import openpyxl

class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Денежн. поток на буд. ст-ть")
        self.geometry("620x300")
        self.resizable(False, False)

        # Установка стиля
        style = ttk.Style()
        style.configure("My.TFrame", background="gray")
        style.configure("My.TButton", background="black", foreground="white")

        # Контейнер для виджетов "От" и "До"
        input_frame = ttk.Frame(self, style="My.TFrame")
        input_frame.pack(pady=10)

        # Виджеты "От" и "До"
        self.from_label = ttk.Label(input_frame, text="От:", foreground="white", background="gray")
        self.from_label.pack(side="left", padx=5)

        self.from_entry = ttk.Entry(input_frame)
        self.from_entry.pack(side="left", padx=5)

        self.to_label = ttk.Label(input_frame, text="До:", foreground="white", background="gray")
        self.to_label.pack(side="left", padx=5)

        self.to_entry = ttk.Entry(input_frame)
        self.to_entry.pack(side="left", padx=5)

        # Кнопка выбора файла
        self.file_button = ttk.Button(self, text="Выбрать", command=self.select_file, style="My.TButton")
        self.file_button.pack(pady=10)

        # Таблица с основными показателями проекта
        table_frame = ttk.Frame(self, style="My.TFrame")
        table_frame.pack(pady=10)

        table_label = ttk.Label(table_frame, text="Основные показатели проекта:", foreground="white", background="gray")
        table_label.pack()

        columns = ("Показатель", "Ед. изм.", "Значение")
        self.data = [
            ("NPV", "тыс. руб", ""),
            ("DPI", "коэф.", ""),
            ("IRR", "%", ""),
            ("PP", "лет", ""),
            ("DPP", "лет", "")
        ]

        self.table = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.table.pack()

        for column in columns:
            self.table.heading(column, text=column)

        for item in self.data:
            self.table.insert("", tk.END, values=item)

        self.center_columns()

        self.mainloop()

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            print("Выбранный файл:", file_path)

            # Открываем выбранный файл и считываем данные
            try:
                workbook = openpyxl.load_workbook(file_path)
                worksheet = workbook.active

                for i, row in enumerate(worksheet.iter_rows(values_only=True, min_row=2, max_row=6, min_col=3, max_col=3)):
                    value = row[0]
                    self.data[i] = (*self.data[i][:2], value)
                    self.table.item(i, values=self.data[i])
                    

            except Exception as e:
                print("Ошибка при чтении файла:", e)

    def center_columns(self):
        for column in self.table["columns"]:
            self.table.column(column, anchor="center")

app = MyApp()
