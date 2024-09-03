import tkinter as tk


from PIL import ImageTk
from tkinter import messagebox as mb
from tkinter.ttk import Combobox


from core.constants import (
    sp_wind_reg_dict,
    sp_snow_reg_dict
)
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
    # make_pzg,
    # make_pz
)


class KozuTkr(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("КОЗУ")
        self.geometry("380x444+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        # self.back_icon = ImageTk.PhotoImage(
        #     file=tempFile_back
        # )

        self.module_bg = tk.Frame(
            self,
            width=360,
            height=434,
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

        self.sp_wind_reg_label = tk.Label(
            self,
            text='Ветровой район по СП',
            width=28,
            anchor="e"
        )
        self.sp_wind_reg_combobox = Combobox(
            self,
            values=("Ia", "I", "II", "III", "IV", "V", "VI", "VII"),
            width=12,
            validate="key"
        )
        self.sp_wind_reg_combobox.bind("<<ComboboxSelected>>", self.paste_wind)

        self.wind_nagr_label = tk.Label(
            self,
            text='Норм. ветровое давление, кПа',
            width=28,
            anchor="e"
        )
        self.wind_nagr_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.sp_sneg_reg_label = tk.Label(
            self,
            text='Снеговой район по СП',
            width=28,
            anchor="e"
        )
        self.sp_sneg_reg_combobox = Combobox(
            self,
            values=("I", "II", "III", "IV", "V", "VI", "VII", "VIII"),
            width=12,
            validate="key"
        )
        self.sp_sneg_reg_combobox.bind("<<ComboboxSelected>>", self.paste_sneg)

        self.snow_nagr_label = tk.Label(
            self,
            text='Нормативный вес снега, кН/м2',
            width=28,
            anchor="e"
        )
        self.snow_nagr_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.golol_rayon_label = tk.Label(
            self,
            text='Гололедный район по СП',
            width=28,
            anchor="e"
        )
        self.golol_rayon_combobox = Combobox(
            self,
            values=("I", "II", "III", "IV", "V"),
            width=12,
        )

        self.rvs_label = tk.Label(
            self,
            text='Марка РВС',
            width=28,
            anchor="e"
        )
        self.rvs_entry = tk.Entry(
            self,
            width=15,
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

        self.speca_label = tk.Label(
            self,
            text='Специф. материалов (ТКР).png',
            width=28,
            anchor="e"
        )
        self.speca_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.browse_for_speca_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_speca
        )

        self.speca_pz_label = tk.Label(
            self,
            text='Специф. материалов (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.speca_pz_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.browse_for_speca_pz_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_speca_pz
        )

        self.tkr_button = tk.Button(
            self,
            text="Создать документацию",
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
        self.sp_wind_reg_label.place(x=15, y=42)
        self.sp_wind_reg_combobox.place(x=220, y=42)
        self.wind_nagr_label.place(x=15, y=65)
        self.wind_nagr_entry.place(x=220, y=65)
        self.sp_sneg_reg_label.place(x=15, y=88)
        self.sp_sneg_reg_combobox.place(x=220, y=88)
        self.snow_nagr_label.place(x=15, y=111)
        self.snow_nagr_entry.place(x=220, y=111)
        self.golol_rayon_label.place(x=15, y=134)
        self.golol_rayon_combobox.place(x=220, y=134)
        self.rvs_label.place(x=15, y=157)
        self.rvs_entry.place(x=220, y=157)
        self.diam_osn_label.place(x=15, y=180)
        self.diam_osn_entry.place(x=220, y=180)
        self.diam_verha_label.place(x=15, y=203)
        self.diam_verha_entry.place(x=220, y=203)
        self.h_label.place(x=15, y=226)
        self.h_entry.place(x=220, y=226)
        self.teor_massa_metalla_label.place(x=15, y=249)
        self.teor_massa_metalla_entry.place(x=220, y=249)
        self.ploschad_uchastka_label.place(x=15, y=272)
        self.ploschad_uchastka_entry.place(x=220, y=272)
        self.territoria_raspoloj_label.place(x=15, y=295)
        self.territoria_raspoloj_entry.place(x=220, y=295)
        self.god_vvoda_v_ekspl_label.place(x=15, y=318)
        self.god_vvoda_v_ekspl_entry.place(x=220, y=318)
        self.speca_label.place(x=15, y=341)
        self.speca_entry.place(x=220, y=341)
        self.browse_for_speca_button.place(x=317, y=339)
        self.speca_pz_label.place(x=15, y=364)
        self.speca_pz_entry.place(x=220, y=364)
        self.browse_for_speca_pz_button.place(x=317, y=362)
        self.tkr_button.place(x=120, y=393)

    def call_make_tkr(self):
        make_tkr(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            developer=self.parent.developer,
            sp_wind_region=self.sp_wind_reg_combobox.get(),
            wind_nagr=self.wind_nagr_entry.get(),
            sp_ice_region=self.sp_sneg_reg_combobox.get(),
            snow_nagr=self.snow_nagr_entry.get(),
            golol_rayon=self.golol_rayon_combobox.get(),
            rvs=self.rvs_entry.get(),
            diam_osn=self.diam_osn_entry.get(),
            diam_verha=self.diam_verha_entry.get(),
            h=self.h_entry.get(),
            teor_massa_metala=self.teor_massa_metalla_entry.get(),
            ploschad_uchastka=self.ploschad_uchastka_entry.get(),
            territoria_raspoloj=self.territoria_raspoloj_entry.get(),
            god_vvoda_v_ekspl=self.god_vvoda_v_ekspl_entry.get(),
            min_temp=self.parent.min_temp,
            max_temp=self.parent.max_temp,
            speca=self.speca_entry.get(),
            speca_pz=self.speca_pz_entry.get()
        )

    def paste_wind(self, event):
        wind_key = self.sp_wind_reg_combobox.get()
        self.wind_nagr_entry.delete(0, tk.END)
        self.wind_nagr_entry.insert(0, sp_wind_reg_dict[wind_key])

    def paste_sneg(self, event):
        sneg_key = self.sp_sneg_reg_combobox.get()
        self.snow_nagr_entry.delete(0, tk.END)
        self.snow_nagr_entry.insert(0, sp_snow_reg_dict[sneg_key])

    def browse_for_speca(self):
        self.file_path = make_path_png()
        self.speca_entry.delete("0", "end") 
        self.speca_entry.insert("insert", self.file_path)

    def browse_for_speca_pz(self):
        self.file_path = make_path_png()
        self.speca_pz_entry.delete("0", "end") 
        self.speca_pz_entry.insert("insert", self.file_path)

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()