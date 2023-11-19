import tkinter as tk


from PIL import ImageTk
from core.db.db_connector import Database
from tkinter import messagebox as mb


from core.utils import (
    tempFile_back,
    make_path_png,
    make_path_xlsx,
    AssemblyAPI
)
from core.constants import typical_ground_dict, coef_nadej_dict
from core.functionality.foundation.utils import *
from core.functionality.foundation.utils import (
    make_rpzf
)


class RpzfGeneration(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Расчет фундамента")
        self.geometry("781x140+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        self.back_icon = ImageTk.PhotoImage(
            file=tempFile_back
        )

        self.module_bg = tk.Frame(
            self,
            width=761,
            height=135,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            image=self.back_icon,
            command=self.back_to_main_window
        )

        self.ige_name_label = tk.Label(
            self,
            text="Наименование отчета ИГЭ",
            width=22,
            anchor="e"
        )
        self.ige_name_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.building_adress_label = tk.Label(
            self,
            text="Адрес строительства",
            width=22,
            anchor="e"
        )
        self.building_adress_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.razrez_skvajin_label = tk.Label(
            self,
            text="Номер скважины",
            width=22,
            anchor="e"
        )
        self.razrez_skvajin_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.media_label = tk.Label(
            self,
            text="РПЗФ:",
            width=5,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.picture1_label = tk.Label(
            self,
            text="Параметры выбранной скважины",
            width=28,
            anchor="e"
        )
        self.picture1_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.picture2_label = tk.Label(
            self,
            text="Рекомедуемые параметры грунтов",
            width=28,
            anchor="e"
        )
        self.picture2_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.browse_for_pic1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pic1
        )
        self.browse_for_pic2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pic2
        )

        self.xlsx_svai_label = tk.Label(
            self,
            text="Эксель посчитанной сваи",
            width=28,
            anchor="e"
        )
        self.xlsx_svai_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )
        self.browse_for_xlsx_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_xlsx
        )

        self.rpzf_button = tk.Button(
            self,
            text="Создать РПЗФ",
            command=self.call_make_rpzf
        )

    def run(self):
        self.draw_widgets()
        self.get_rpzf_data_from_db()
        self.mainloop()

    def get_rpzf_data_from_db(self):
        try:
            self.ige_name_entry.delete(0, tk.END)
            self.ige_name_entry.insert(0, self.parent.ige_name)
            self.building_adress_entry.delete(0, tk.END)
            self.building_adress_entry.insert(0, self.parent.building_adress)
            self.razrez_skvajin_entry.delete(0, tk.END)
            self.razrez_skvajin_entry.insert(0, self.parent.razrez_skvajin)
            self.picture1_entry.delete(0, tk.END)
            self.picture1_entry.insert(0, self.parent.picture1)
            self.picture2_entry.delete(0, tk.END)
            self.picture2_entry.insert(0, self.parent.picture2)
            self.xlsx_svai_entry.delete(0, tk.END)
            self.xlsx_svai_entry.insert(0, self.parent.xlsx_svai)
        except Exception as e:
            print("!INFO!: Сохраненные данные отсутствуют.")

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=3)
        self.rpzf_button.place(x=335, y=104)
        self.ige_name_label.place(x=12, y=29)
        self.ige_name_entry.place(x=173, y=29)
        self.building_adress_label.place(x=12, y=53)
        self.building_adress_entry.place(x=173, y=53)
        self.razrez_skvajin_label.place(x=12, y=76)
        self.razrez_skvajin_entry.place(x=173, y=76)
        self.picture1_label.place(x=340, y=29)
        self.picture1_entry.place(x=545, y=29)
        self.browse_for_pic1_button.place(x=702, y=24)
        self.picture2_label.place(x=340, y=53)
        self.picture2_entry.place(x=545, y=53)
        self.browse_for_pic2_button.place(x=702, y=50)
        self.xlsx_svai_label.place(x=340, y=76)
        self.xlsx_svai_entry.place(x=545, y=76)
        self.browse_for_xlsx_button.place(x=702, y=75)

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()

    def browse_for_pic1(self):
        self.file_path = make_path_png()
        self.picture1_entry.delete("0", "end") 
        self.picture1_entry.insert("insert", self.file_path)

    def browse_for_pic2(self):
        self.file_path = make_path_png()
        self.picture2_entry.delete("0", "end") 
        self.picture2_entry.insert("insert", self.file_path)

    def browse_for_xlsx(self):
        self.file_path = make_path_xlsx()
        self.xlsx_svai_entry.delete("0", "end") 
        self.xlsx_svai_entry.insert("insert", self.file_path)

    def call_make_rpzf(self):
        pic1 = self.picture1_entry.get().split("Удаленка")[1]
        pic2 = self.picture2_entry.get().split("Удаленка")[1]
        xlsx = self.xlsx_svai_entry.get().split("Удаленка")[1]
        self.db.add_rpzf_data(
            initial_data_id=self.parent.initial_data_id,
            ige_name=self.ige_name_entry.get(),
            building_adress=self.building_adress_entry.get(),
            razrez_skvajin=self.razrez_skvajin_entry.get(),
            picture1=pic1,
            picture2=pic2,
            xlsx_svai=xlsx
        )
        svai_data = self.db.get_svai_data(self.parent.initial_data_id)
        make_rpzf(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            pole_code=self.parent.pole_code,
            developer=self.parent.developer,
            diam_svai=self.parent.pls_pole_data["bot_diam"],
            deepness_svai=svai_data["deepness_svai"],
            height_svai=svai_data["height_svai"],
            moment1=self.parent.pls_pole_data["moment"],
            vert_force1=self.parent.pls_pole_data["vert_force"],
            shear_force1=self.parent.pls_pole_data["shear_force"],
            ige_name=self.ige_name_entry.get(),
            building_adress=self.building_adress_entry.get(),
            razrez_skvajin=self.razrez_skvajin_entry.get(),
            picture1=self.picture1_entry.get(),
            picture2=self.picture2_entry.get(),
            xlsx_svai=self.xlsx_svai_entry.get()
        )