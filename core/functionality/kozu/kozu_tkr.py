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


class KozuTkr(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("КОЗУ")
        self.geometry("480x550+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        # self.back_icon = ImageTk.PhotoImage(
        #     file=tempFile_back
        # )

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

        self.klimat_label = tk.Label(
            self,
            text='Климат',
            width=28,
            anchor="e"
        )
        self.klimat_entry = tk.Entry(
            self,
            width=40,
            relief="sunken",
            bd=2
        )

        self.relief_label = tk.Label(
            self,
            text='Рельеф',
            width=28,
            anchor="e"
        )
        self.relief_entry = tk.Entry(
            self,
            width=40,
            relief="sunken",
            bd=2
        )

        self.geologia_label = tk.Label(
            self,
            text='Геологические условия',
            width=28,
            anchor="e"
        )
        self.geologia_entry = tk.Entry(
            self,
            width=40,
            relief="sunken",
            bd=2
        )

        self.flora_label = tk.Label(
            self,
            text='Растительность',
            width=28,
            anchor="e"
        )
        self.flora_entry = tk.Entry(
            self,
            width=40,
            relief="sunken",
            bd=2
        )

        self.gidrologia_label = tk.Label(
            self,
            text='Гидрологические условия',
            width=28,
            anchor="e"
        )
        self.gidrologia_entry = tk.Entry(
            self,
            width=40,
            relief="sunken",
            bd=2
        )

        self.diam_osn_label = tk.Label(
            self,
            text='Диаметр основания объекта, мм',
            width=28,
            anchor="e"
        )
        self.diam_osn_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.diam_verha_label = tk.Label(
            self,
            text='Диаметр верха объекта, мм',
            width=28,
            anchor="e"
        )
        self.diam_verha_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.h_label = tk.Label(
            self,
            text='Высота объекта, мм',
            width=28,
            anchor="e"
        )
        self.h_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.teor_massa_metalla_label = tk.Label(
            self,
            text='Общая теоретическая масса, т',
            width=28,
            anchor="e"
        )
        self.teor_massa_metalla_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ploschad_uchastka_label = tk.Label(
            self,
            text='Площадь участка под объект, м2',
            width=28,
            anchor="e"
        )
        self.ploschad_uchastka_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.territoria_raspoloj_label = tk.Label(
            self,
            text='Кому принадлежит объект',
            width=28,
            anchor="e"
        )
        self.territoria_raspoloj_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.god_vvoda_v_ekspl_label = tk.Label(
            self,
            text='Год ввода в эксплуатацию',
            width=28,
            anchor="e"
        )
        self.god_vvoda_v_ekspl_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.seism_rayon_label = tk.Label(
            self,
            text='Сейсмичность участка',
            width=28,
            anchor="e"
        )
        self.seism_rayon_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.pz_button = tk.Button(
            self,
            text="Создать том ПЗ",
            command=self.call_make_pz
        )

        self.tkr_button = tk.Button(
            self,
            text="Создать том ТКР",
            command=self.call_make_tkr
        )

    def run(self):
        self.draw_widgets()
        # self.get_rpzf_data_from_db()
        self.mainloop()

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=2)
        self.article_label.place(x=100, y=19)
        self.general_info_label.place(x=15, y=42)
        self.general_info_entry.place(x=220, y=42, height=44)
        self.klimat_label.place(x=15, y=88)
        self.klimat_entry.place(x=220, y=88, height=44)
        self.relief_label.place(x=15, y=134)
        self.relief_entry.place(x=220, y=134, height=44)
        self.geologia_label.place(x=15, y=180)
        self.geologia_entry.place(x=220, y=180, height=44)
        self.flora_label.place(x=15, y=226)
        self.flora_entry.place(x=220, y=226, height=44)
        self.gidrologia_label.place(x=15, y=272)
        self.gidrologia_entry.place(x=220, y=272, height=44)
        self.diam_osn_label.place(x=15, y=318)
        self.diam_osn_entry.place(x=220, y=318)
        self.diam_verha_label.place(x=15, y=341)
        self.diam_verha_entry.place(x=220, y=341)
        self.h_label.place(x=15, y=364)
        self.h_entry.place(x=220, y=364)
        self.teor_massa_metalla_label.place(x=15, y=387)
        self.teor_massa_metalla_entry.place(x=220, y=387)
        self.ploschad_uchastka_label.place(x=15, y=410)
        self.ploschad_uchastka_entry.place(x=220, y=410)
        self.territoria_raspoloj_label.place(x=15, y=433)
        self.territoria_raspoloj_entry.place(x=220, y=433)
        self.god_vvoda_v_ekspl_label.place(x=15, y=456)
        self.god_vvoda_v_ekspl_entry.place(x=220, y=456)
        self.seism_rayon_label.place(x=15, y=479)
        self.seism_rayon_entry.place(x=220, y=479)
        self.tkr_button.place(x=110, y=505)
        self.pz_button.place(x=220, y=505)

    def call_make_tkr(self):
        make_tkr(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            developer=self.parent.developer,
            general_info=self.general_info_entry.get(),
            klimat=self.klimat_entry.get(),
            relief=self.relief_entry.get(),
            geologia=self.geologia_entry.get(),
            flora=self.flora_entry.get(),
            gidrologia=self.gidrologia_entry.get(),
            diam_osn=self.diam_osn_entry.get(),
            diam_verha=self.diam_verha_entry.get(),
            h=self.h_entry.get(),
            teor_massa_metala=self.teor_massa_metalla_entry.get(),
            ploschad_uchastka=self.ploschad_uchastka_entry.get(),
            territoria_raspoloj=self.territoria_raspoloj_entry.get(),
            god_vvoda_v_ekspl=self.god_vvoda_v_ekspl_entry.get(),
            wind_region=self.parent.wind_region,
            wind_pressure=self.parent.wind_pressure,
            area=self.parent.area,
            ice_region=self.parent.ice_region,
            ice_thickness=self.parent.ice_thickness,
            ice_wind_pressure=self.parent.ice_wind_pressure,
            year_average_temp=self.parent.year_average_temp,
            min_temp=self.parent.min_temp,
            wind_temp=self.parent.wind_temp,
            ice_temp=self.parent.ice_temp,
            max_temp=self.parent.max_temp,
            wind_reg_coef=self.parent.wind_reg_coef,
            ice_reg_coef=self.parent.ice_reg_coef,
            seism_rayon=self.seism_rayon_entry.get()
        )

    def call_make_pz(self):
        make_pz(
            project_code=self.parent.project_code,
            developer=self.parent.developer
        )

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()