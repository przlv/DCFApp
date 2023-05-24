import tkinter as tk

class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Мое приложение")
        self.geometry("1280x720")
        self.resizable(False, False)
        # Добавьте здесь свои виджеты и компоненты интерфейса
        
        self.mainloop()