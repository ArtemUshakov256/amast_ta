import tkinter as tk


from PIL import ImageTk
from tkinter import messagebox as mb


from core.db.db_connector import Database
from core.utils import (
    # tempFile_back,
    # tempFile_open,
    # tempFile_save,
    make_path_xlsx,
    make_path_png
)
from core.exceptions import AddPlsPolePathException
from core.functionality.kozu.utils import (
    make_tkr,
    make_pz
)


class KozuSchema(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("КОЗУ")
        self.geometry("480x550+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        self.module_bg = tk.Frame(
            self,
            width=460,
            height=540,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            # image=self.back_icon,
            text="Назад",
            command=self.back_to_main_window
        )
        
        self.article_label = tk.Label(
            self,
            text='Данные по КОЗУ:',
            width=18,
            anchor="e",
            font=("standard", 10, "bold")
        )
        
        self.general_info_label = tk.Label(
            self,
            text='Информация по объекту',
            width=28,
            anchor="e"
        )
        self.general_info_entry = tk.Entry(
            self,
            width=40,
            relief="sunken",
            bd=2
        )