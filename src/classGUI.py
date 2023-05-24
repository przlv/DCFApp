import tkinter as tk
from tkinter import ttk, filedialog
from src.classDataProcessing import *
from openpyxl import load_workbook


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
        self.file_button = ttk.Button(self, text="Выбрать файл", command=self.select_file,)
        self.file_button.pack(pady=10)

        # # Таблица с основными показателями проекта
        # self.table_frame = ttk.Frame(self, style="My.TFrame")
        # self.table_frame.pack(pady=10)

        # table_label = ttk.Label(self.table_frame, text="Основные показатели проекта:", foreground="white", background="gray")
        # table_label.pack()

        # columns = ("Показатель", "Ед. изм.", "Значение")
        # self.data = [
        #     ("NPV", "тыс. руб", ""),
        #     ("DPI", "коэф.", ""),
        #     ("IRR", "%", ""),
        #     ("PP", "лет", ""),
        #     ("DPP", "лет", "")
        # ]

        # self.table = ttk.Treeview(self.table_frame, columns=columns, show="headings")
        # self.table.pack()

        # for column in columns:
        #     self.table.heading(column, text=column)

        # for item in self.data:
        #     self.table.insert("", tk.END, values=item)

        # self.center_columns()

        self.mainloop()

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("Excel Files", "*.xlsm")])
        if file_path:
            print("Выбранный файл:", file_path)
            self.filepath = file_path
        
        if '.xlsm' in file_path:
            # Указать путь к исходному файлу XLSM
            xlsm_file_path = file_path

            # Указать путь для сохранения файла XLSX
            xlsx_file_path = file_path.replace('.xlsm', '.xlsx')

            # Загрузить файл XLSM
            wb = load_workbook(xlsm_file_path, keep_vba=True)

            # Сохранить файл в формате XLSX
            wb.save(xlsx_file_path)
            
            self.filepath = xlsx_file_path
        
        self.run_indic()

    def center_columns(self):
        for column in self.table["columns"]:
            self.table.column(column, anchor="center")
    
    def run_indic(self):
        appdataproc = DataProccessing(self.filepath, 1, 1)
        self.indicators = appdataproc.run()
        
        self.table_frame = ttk.Frame(self, style="My.TFrame")
        self.table_frame.pack(pady=10)
        
        table_label = ttk.Label(self.table_frame, text="Основные показатели проекта:", foreground="white", background="gray")
        table_label.pack()
        
        columns = ("Показатель", "Ед. изм.", "Значение")
        self.data = [
            ("NPV", "тыс. руб", str(self.indicators['NPV'])),
            ("DPI", "коэф.", str(self.indicators['DPI'])),
            ("IRR", "%", str(self.indicators['IRR'])),
            ("PP", "лет", str(self.indicators['PP'])),
            ("DPP", "лет", str(self.indicators['DPP']))
        ]

        self.table = ttk.Treeview(self.table_frame, columns=columns, show="headings")
        self.table.pack()

        for column in columns:
            self.table.heading(column, text=column)

        for item in self.data:
            self.table.insert("", tk.END, values=item)

        self.center_columns()

app = MyApp()
