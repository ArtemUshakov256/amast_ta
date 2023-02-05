import tkinter as tk

from ..utils import *
from .description_window import DescriptionWindow


class MainWindow:
    def __init__(
        self,
        width=350,
        height=200,
        title="Основное окно",
        resizable=(False, False),
        bg="#DFD8FF",
        icon=None
    ):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry(f"{width}x{height}+200+200")
        self.root.resizable(resizable[0], resizable[1])
        self.root.config(bg=bg)
        if icon:
            self.root.iconbitmap(icon)
        self.entry = tk.Entry(self.root,
            width=30,
            justify="left",
            relief="sunken",
            bd=3)

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        tk.Label(self.root, text="Путь к файлу", bg="#DFD8FF")\
        .grid(row=0, column=0, stick="e")
        
        self.entry.grid(row=0, column=1)

        tk.Button(self.root, text="Обзор", command=self.callback)\
        .grid(row=0, column=2)

    def callback(self):
        file_path = make_path() 
        self.entry.insert("insert", file_path)

    def open_description_window(self,
        width=400,
        height=200,
        title="Описание программы",
        resizable=(False, False),
        bg="#DFD8FF",
        icon=None
        ):
        DescriptionWindow(self.root,
            width,
            height,
            title,
            resizable,
            bg,
            icon)


if __name__ == "__main__":
    main_page = MainWindow(icon="logo.ico")
    main_page.open_description_window(icon="logo.ico")
    main_page.run()