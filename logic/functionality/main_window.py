import tkinter as tk

from ..utils import *
from .description_window import DescriptionWindow


class MainWindow:
    def __init__(
        self,
        width=350,
        height=250,
        title="Основное окно",
        resizable=(False, False),
        bg="#FFFFFF",
        bg_pic=None,
        icon=None
    ):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry(f"{width}x{height}+200+200")
        self.root.resizable(resizable[0], resizable[1])
        self.root.config(bg=bg)
        if icon:
            self.root.iconbitmap(icon)

        self.canvas = tk.Canvas(self.root, height=height-100, width=width)
        self.image_file = tk.PhotoImage(file=bg_pic)
        self.image = self.canvas.create_image(0,0, anchor='nw', image=self.image_file)

        self.entry = tk.Entry(
            self.root,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.path_label = tk.Label(self.root, text="Путь к файлу", bg="#FFFFFF")
        self.browse_button = tk.Button(self.root, text="Обзор", command=self.callback)
        self.convert_button = tk.Button(
            self.root, text="Конвертировать",
            command=self.convert_txt_to_xlsx
        )
        # self.convertation_result = tk.Label(self.root, text=f"{self.result}", bg="#FFFFFF")

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        self.canvas.pack(side='top')

        # self.path_label.grid(row=0, column=0, stick="e")
        # self.entry.grid(row=0, column=1)
        # self.browse_button.grid(row=0, column=2)
        # self.convert_button.grid(row=1, column=1)

        self.path_label.place(x=0, y=160)
        self.entry.place(x=80, y=160)
        self.browse_button.place(x=270, y=160)
        self.convert_button.place(x=120, y=190)
        # self.convertation_result.place(x=120, y=230)

    def callback(self):
        self.file_path = make_path() 
        self.entry.insert("insert", self.file_path)

    def convert_txt_to_xlsx(self):
        self.result = extract_txt_data(path=self.file_path)
        return self.result

    # def open_description_window(self,
    #     width=400,
    #     height=200,
    #     title="Описание программы",
    #     resizable=(False, False),
    #     bg="#DFD8FF",
    #     icon=None
    #     ):
    #     DescriptionWindow(self.root,
    #         width,
    #         height,
    #         title,
    #         resizable,
    #         bg,
    #         icon)


if __name__ == "__main__":
    main_page = MainWindow(icon="logo.ico")
    main_page.open_description_window(icon="logo.ico")
    main_page.run()