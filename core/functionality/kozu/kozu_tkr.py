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
    make_path_png,
    make_path_pdf
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
        self.geometry("620x777+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        # self.back_icon = ImageTk.PhotoImage(
        #     file=tempFile_back
        # )

        self.module_bg = tk.Frame(
            self,
            width=600,
            height=767,
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

        self.quantity_of_rvs_label = tk.Label(
            self,
            text='Количество марок РВС, шт',
            width=28,
            anchor="e"
        )
        self.quantity_of_rvs_combobox = Combobox(
            self,
            width=12,
            values=(
                "1",
                "2",
                "3",
                "4"
            )
        )
        self.quantity_of_rvs_combobox.bind("<<ComboboxSelected>>", self.activate_rvs)

        self.rvs_1_label = tk.Label(
            self,
            text='РВС 1',
            width=8,
            anchor="e"
        )
        self.rvs_2_label = tk.Label(
            self,
            text='РВС 2',
            width=8,
            anchor="e"
        )
        self.rvs_3_label = tk.Label(
            self,
            text='РВС 3',
            width=8,
            anchor="e"
        )
        self.rvs_4_label = tk.Label(
            self,
            text='РВС 4',
            width=8,
            anchor="e"
        )
        
        self.ob_rvs_label = tk.Label(
            self,
            text='Объем РВС, м3',
            width=28,
            anchor="e"
        )
        self.diam_osn_label = tk.Label(
            self,
            text='Диаметр осн. КОЗ-У, мм',
            width=28,
            anchor="e"
        )
        self.diam_verha_label = tk.Label(
            self,
            text='Диаметр верха КОЗ-У, мм',
            width=28,
            anchor="e"
        )
        self.h_label = tk.Label(
            self,
            text='Высота КОЗ-У, мм',
            width=28,
            anchor="e"
        )
        self.massa_rvs_label = tk.Label(
            self,
            text='Масса 1-го КОЗ-У, т',
            width=28,
            anchor="e"
        )
        self.obsch_massa_rvs_label = tk.Label(
            self,
            text='Общ. масса КОЗ-У, т',
            width=28,
            anchor="e"
        )

        self.rvs1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.rvs2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.rvs3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.rvs4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.zasch_obj_label = tk.Label(
            self,
            text='Защищаемый объект (ПЗ)',
            width=28,
            anchor="e"
        )
        self.zasch_obj_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.diam_osn1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.diam_osn2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.diam_osn3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.diam_osn4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.diam_verha1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.diam_verha2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.diam_verha3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.diam_verha4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.h1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.h2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.h3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.h4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.massa_rvs1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.massa_rvs2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.massa_rvs3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.massa_rvs4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.obsch_massa_rvs1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.obsch_massa_rvs2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.obsch_massa_rvs3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.obsch_massa_rvs4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
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

        self.rayon_str_label = tk.Label(
            self,
            text='Район строительства',
            width=28,
            anchor="e"
        )
        self.rayon_str_entry = tk.Entry(
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
            text='Спецификация (ТКР).png',
            width=28,
            anchor="e"
        )
        self.speca_entry = tk.Entry(
            self,
            width=45,
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
            text='Спецификация 1 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.speca_pz_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_speca_pz_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_speca_pz
        )

        self.speca_pz2_label = tk.Label(
            self,
            text='Спецификация 2 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.speca_pz2_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_speca_pz2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_speca_pz2
        )

        self.speca_pz3_label = tk.Label(
            self,
            text='Спецификация 3 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.speca_pz3_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_speca_pz3_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_speca_pz3
        )

        self.speca_pz4_label = tk.Label(
            self,
            text='Спецификация 4 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.speca_pz4_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_speca_pz4_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_speca_pz4
        )

        self.vid_kozu_label = tk.Label(
            self,
            text='Вид КОЗ-У 1 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.vid_kozu_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_vid_kozu_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_vid_kozu
        )

        self.vid_kozu2_label = tk.Label(
            self,
            text='Вид КОЗ-У 2 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.vid_kozu2_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_vid_kozu2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_vid_kozu2
        )

        self.vid_kozu3_label = tk.Label(
            self,
            text='Вид КОЗ-У 3 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.vid_kozu3_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_vid_kozu3_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_vid_kozu3
        )

        self.vid_kozu4_label = tk.Label(
            self,
            text='Вид КОЗ-У 4 тип РВС (ПЗ).png',
            width=28,
            anchor="e"
        )
        self.vid_kozu4_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_vid_kozu4_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_vid_kozu4
        )

        self.eskiz_kozu_label = tk.Label(
            self,
            text='Эскиз КОЗ-У.pdf',
            width=28,
            anchor="e"
        )
        self.eskiz_kozu_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_eskiz_kozu_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_eskiz_kozu
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
        self.quantity_of_rvs_label.place(x=15, y=157)
        self.quantity_of_rvs_combobox.place(x=220, y=157)
        self.rvs_1_label.place(x=220, y=180)
        self.rvs_2_label.place(x=314, y=180)
        self.rvs_3_label.place(x=408, y=180)
        self.rvs_4_label.place(x=502, y=180)
        self.ob_rvs_label.place(x=15, y=203)
        self.diam_osn_label.place(x=15, y=226)
        self.diam_verha_label.place(x=15, y=249)
        self.h_label.place(x=15, y=272)
        self.massa_rvs_label.place(x=15, y=295)
        self.obsch_massa_rvs_label.place(x=15, y=318)
        self.rvs1_entry.place(x=220, y=203)
        self.rvs2_entry.place(x=314, y=203)
        self.rvs3_entry.place(x=408, y=203)
        self.rvs4_entry.place(x=502, y=203)
        self.diam_osn1_entry.place(x=220, y=226)
        self.diam_osn2_entry.place(x=314, y=226)
        self.diam_osn3_entry.place(x=408, y=226)
        self.diam_osn4_entry.place(x=502, y=226) 
        self.diam_verha1_entry.place(x=220, y=249)
        self.diam_verha2_entry.place(x=314, y=249)
        self.diam_verha3_entry.place(x=408, y=249)
        self.diam_verha4_entry.place(x=502, y=249)
        self.h1_entry.place(x=220, y=272)
        self.h2_entry.place(x=314, y=272)
        self.h3_entry.place(x=408, y=272)
        self.h4_entry.place(x=502, y=272)
        self.massa_rvs1_entry.place(x=220, y=295)
        self.massa_rvs2_entry.place(x=314, y=295)
        self.massa_rvs3_entry.place(x=408, y=295)
        self.massa_rvs4_entry.place(x=502, y=295)
        self.obsch_massa_rvs1_entry.place(x=220, y=318)
        self.obsch_massa_rvs2_entry.place(x=314, y=318)
        self.obsch_massa_rvs3_entry.place(x=408, y=318)
        self.obsch_massa_rvs4_entry.place(x=502, y=318)
        self.zasch_obj_label.place(x=15, y=341)
        self.zasch_obj_entry.place(x=220, y=341)
        self.ploschad_uchastka_label.place(x=15, y=364)
        self.ploschad_uchastka_entry.place(x=220, y=364)
        self.territoria_raspoloj_label.place(x=15, y=387)
        self.territoria_raspoloj_entry.place(x=220, y=387)
        self.rayon_str_label.place(x=15, y=410)
        self.rayon_str_entry.place(x=220, y=410)
        self.god_vvoda_v_ekspl_label.place(x=15, y=433)
        self.god_vvoda_v_ekspl_entry.place(x=220, y=433)
        self.vid_kozu_label.place(x=15, y=461)
        self.vid_kozu_entry.place(x=220, y=461)
        self.browse_for_vid_kozu_button.place(x=497, y=459)
        self.vid_kozu2_label.place(x=15, y=489)
        self.vid_kozu2_entry.place(x=220, y=489)
        self.browse_for_vid_kozu2_button.place(x=497, y=487)
        self.vid_kozu3_label.place(x=15, y=516)
        self.vid_kozu3_entry.place(x=220, y=516)
        self.browse_for_vid_kozu3_button.place(x=497, y=514)
        self.vid_kozu4_label.place(x=15, y=543)
        self.vid_kozu4_entry.place(x=220, y=543)
        self.browse_for_vid_kozu4_button.place(x=497, y=541)
        self.speca_label.place(x=15, y=570)
        self.speca_entry.place(x=220, y=570)
        self.browse_for_speca_button.place(x=497, y=568)
        self.speca_pz_label.place(x=15, y=597)
        self.speca_pz_entry.place(x=220, y=597)
        self.browse_for_speca_pz_button.place(x=497, y=595)
        self.speca_pz2_label.place(x=15, y=624)
        self.speca_pz2_entry.place(x=220, y=624)
        self.browse_for_speca_pz2_button.place(x=497, y=622)
        self.speca_pz3_label.place(x=15, y=651)
        self.speca_pz3_entry.place(x=220, y=651)
        self.browse_for_speca_pz3_button.place(x=497, y=649)
        self.speca_pz4_label.place(x=15, y=678)
        self.speca_pz4_entry.place(x=220, y=678)
        self.browse_for_speca_pz4_button.place(x=497, y=676)
        self.eskiz_kozu_label.place(x=15, y=705)
        self.eskiz_kozu_entry.place(x=220, y=705)
        self.browse_for_eskiz_kozu_button.place(x=497, y=703)
        self.tkr_button.place(x=250, y=734)

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
            rvs1=self.rvs1_entry.get(),
            rvs2=self.rvs2_entry.get(),
            rvs3=self.rvs3_entry.get(),
            rvs4=self.rvs4_entry.get(),
            diam_osn1=self.diam_osn1_entry.get(),
            diam_osn2=self.diam_osn2_entry.get(),
            diam_osn3=self.diam_osn3_entry.get(),
            diam_osn4=self.diam_osn4_entry.get(),
            diam_verha1=self.diam_verha1_entry.get(),
            diam_verha2=self.diam_verha2_entry.get(),
            diam_verha3=self.diam_verha3_entry.get(),
            diam_verha4=self.diam_verha4_entry.get(),
            h1=self.h1_entry.get(),
            h2=self.h2_entry.get(),
            h3=self.h3_entry.get(),
            h4=self.h4_entry.get(),
            massa_rvs1=self.massa_rvs1_entry.get(),
            massa_rvs2=self.massa_rvs2_entry.get(),
            massa_rvs3=self.massa_rvs3_entry.get(),
            massa_rvs4=self.massa_rvs4_entry.get(),
            obsch_massa_rvs1=self.obsch_massa_rvs1_entry.get(),
            obsch_massa_rvs2=self.obsch_massa_rvs2_entry.get(),
            obsch_massa_rvs3=self.obsch_massa_rvs3_entry.get(),
            obsch_massa_rvs4=self.obsch_massa_rvs4_entry.get(),
            zasch_obj=self.zasch_obj_entry.get(),
            ploschad_uchastka=self.ploschad_uchastka_entry.get(),
            territoria_raspoloj=self.territoria_raspoloj_entry.get(),
            god_vvoda_v_ekspl=self.god_vvoda_v_ekspl_entry.get(),
            min_temp=self.parent.min_temp,
            max_temp=self.parent.max_temp,
            speca=self.speca_entry.get(),
            speca_pz1=self.speca_pz_entry.get(),
            speca_pz2=self.speca_pz2_entry.get(),
            speca_pz3=self.speca_pz3_entry.get(),
            speca_pz4=self.speca_pz4_entry.get(),
            vid_kozu1=self.vid_kozu_entry.get(),
            vid_kozu2=self.vid_kozu2_entry.get(),
            vid_kozu3=self.vid_kozu3_entry.get(),
            vid_kozu4=self.vid_kozu4_entry.get(),
            rayon_str=self.rayon_str_entry.get(),
            eskiz_kozu=self.eskiz_kozu_entry.get(),
            quantity_of_rvs=self.quantity_of_rvs_combobox.get()
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

    def browse_for_speca_pz2(self):
        self.file_path = make_path_png()
        self.speca_pz2_entry.delete("0", "end") 
        self.speca_pz2_entry.insert("insert", self.file_path)

    def browse_for_speca_pz3(self):
        self.file_path = make_path_png()
        self.speca_pz3_entry.delete("0", "end") 
        self.speca_pz3_entry.insert("insert", self.file_path)

    def browse_for_speca_pz4(self):
        self.file_path = make_path_png()
        self.speca_pz4_entry.delete("0", "end") 
        self.speca_pz4_entry.insert("insert", self.file_path)

    def browse_for_vid_kozu(self):
        self.file_path = make_path_png()
        self.vid_kozu_entry.delete("0", "end") 
        self.vid_kozu_entry.insert("insert", self.file_path)

    def browse_for_vid_kozu2(self):
        self.file_path = make_path_png()
        self.vid_kozu2_entry.delete("0", "end") 
        self.vid_kozu2_entry.insert("insert", self.file_path)

    def browse_for_vid_kozu3(self):
        self.file_path = make_path_png()
        self.vid_kozu3_entry.delete("0", "end") 
        self.vid_kozu3_entry.insert("insert", self.file_path)

    def browse_for_vid_kozu4(self):
        self.file_path = make_path_png()
        self.vid_kozu4_entry.delete("0", "end") 
        self.vid_kozu4_entry.insert("insert", self.file_path)

    def browse_for_eskiz_kozu(self):
        self.file_path = make_path_pdf()
        self.eskiz_kozu_entry.delete("0", "end") 
        self.eskiz_kozu_entry.insert("insert", self.file_path)

    def activate_rvs(self, event):
        if self.quantity_of_rvs_combobox.get() == "1":
            self.rvs2_entry.delete(0, tk.END)
            self.rvs3_entry.delete(0, tk.END)
            self.rvs4_entry.delete(0, tk.END)
            self.rvs2_entry.config(state="disabled")
            self.rvs3_entry.config(state="disabled")
            self.rvs4_entry.config(state="disabled")
            self.rvs1_entry.config(state="normal")
            self.diam_osn2_entry.delete(0, tk.END)
            self.diam_osn3_entry.delete(0, tk.END)
            self.diam_osn4_entry.delete(0, tk.END)
            self.diam_osn2_entry.config(state="disabled")
            self.diam_osn3_entry.config(state="disabled")
            self.diam_osn4_entry.config(state="disabled")
            self.diam_osn1_entry.config(state="normal")
            self.diam_verha2_entry.delete(0, tk.END)
            self.diam_verha3_entry.delete(0, tk.END)
            self.diam_verha4_entry.delete(0, tk.END)
            self.diam_verha2_entry.config(state="disabled")
            self.diam_verha3_entry.config(state="disabled")
            self.diam_verha4_entry.config(state="disabled")
            self.diam_verha1_entry.config(state="normal")
            self.h2_entry.delete(0, tk.END)
            self.h3_entry.delete(0, tk.END)
            self.h4_entry.delete(0, tk.END)
            self.h2_entry.config(state="disabled")
            self.h3_entry.config(state="disabled")
            self.h4_entry.config(state="disabled")
            self.h1_entry.config(state="normal")
            self.massa_rvs2_entry.delete(0, tk.END)
            self.massa_rvs3_entry.delete(0, tk.END)
            self.massa_rvs4_entry.delete(0, tk.END)
            self.massa_rvs2_entry.config(state="disabled")
            self.massa_rvs3_entry.config(state="disabled")
            self.massa_rvs4_entry.config(state="disabled")
            self.massa_rvs1_entry.config(state="normal")
            self.obsch_massa_rvs2_entry.delete(0, tk.END)
            self.obsch_massa_rvs3_entry.delete(0, tk.END)
            self.obsch_massa_rvs4_entry.delete(0, tk.END)
            self.obsch_massa_rvs2_entry.config(state="disabled")
            self.obsch_massa_rvs3_entry.config(state="disabled")
            self.obsch_massa_rvs4_entry.config(state="disabled")
            self.obsch_massa_rvs1_entry.config(state="normal")
        elif self.quantity_of_rvs_combobox.get() == "2":
            self.rvs3_entry.delete(0, tk.END)
            self.rvs4_entry.delete(0, tk.END)
            self.rvs3_entry.config(state="disabled")
            self.rvs4_entry.config(state="disabled")
            self.rvs1_entry.config(state="normal")
            self.rvs2_entry.config(state="normal")
            self.diam_osn3_entry.delete(0, tk.END)
            self.diam_osn4_entry.delete(0, tk.END)
            self.diam_osn3_entry.config(state="disabled")
            self.diam_osn4_entry.config(state="disabled")
            self.diam_osn1_entry.config(state="normal")
            self.diam_osn2_entry.config(state="normal")
            self.diam_verha3_entry.delete(0, tk.END)
            self.diam_verha4_entry.delete(0, tk.END)
            self.diam_verha3_entry.config(state="disabled")
            self.diam_verha4_entry.config(state="disabled")
            self.diam_verha1_entry.config(state="normal")
            self.diam_verha2_entry.config(state="normal")
            self.h3_entry.delete(0, tk.END)
            self.h4_entry.delete(0, tk.END)
            self.h3_entry.config(state="disabled")
            self.h4_entry.config(state="disabled")
            self.h1_entry.config(state="normal")
            self.h2_entry.config(state="normal")
            self.massa_rvs3_entry.delete(0, tk.END)
            self.massa_rvs4_entry.delete(0, tk.END)
            self.massa_rvs3_entry.config(state="disabled")
            self.massa_rvs4_entry.config(state="disabled")
            self.massa_rvs1_entry.config(state="normal")
            self.massa_rvs2_entry.config(state="normal")
            self.obsch_massa_rvs3_entry.delete(0, tk.END)
            self.obsch_massa_rvs4_entry.delete(0, tk.END)
            self.obsch_massa_rvs3_entry.config(state="disabled")
            self.obsch_massa_rvs4_entry.config(state="disabled")
            self.obsch_massa_rvs1_entry.config(state="normal")
            self.obsch_massa_rvs2_entry.config(state="normal")
        elif self.quantity_of_rvs_combobox.get() == "3":
            self.rvs4_entry.delete(0, tk.END)
            self.rvs4_entry.config(state="disabled")
            self.rvs1_entry.config(state="normal")
            self.rvs2_entry.config(state="normal")
            self.rvs3_entry.config(state="normal")
            self.diam_osn4_entry.delete(0, tk.END)
            self.diam_osn4_entry.config(state="disabled")
            self.diam_osn1_entry.config(state="normal")
            self.diam_osn2_entry.config(state="normal")
            self.diam_osn3_entry.config(state="normal")
            self.diam_verha4_entry.delete(0, tk.END)
            self.diam_verha4_entry.config(state="disabled")
            self.diam_verha1_entry.config(state="normal")
            self.diam_verha2_entry.config(state="normal")
            self.diam_verha3_entry.config(state="normal")
            self.h4_entry.delete(0, tk.END)
            self.h4_entry.config(state="disabled")
            self.h1_entry.config(state="normal")
            self.h2_entry.config(state="normal")
            self.h3_entry.config(state="normal")
            self.massa_rvs4_entry.delete(0, tk.END)
            self.massa_rvs4_entry.config(state="disabled")
            self.massa_rvs1_entry.config(state="normal")
            self.massa_rvs2_entry.config(state="normal")
            self.massa_rvs3_entry.config(state="normal")
            self.obsch_massa_rvs4_entry.delete(0, tk.END)
            self.obsch_massa_rvs4_entry.config(state="disabled")
            self.obsch_massa_rvs1_entry.config(state="normal")
            self.obsch_massa_rvs2_entry.config(state="normal")
            self.obsch_massa_rvs3_entry.config(state="normal")
        else:
            self.rvs1_entry.config(state="normal")
            self.rvs2_entry.config(state="normal")
            self.rvs3_entry.config(state="normal")
            self.rvs4_entry.config(state="normal")
            self.diam_osn1_entry.config(state="normal")
            self.diam_osn2_entry.config(state="normal")
            self.diam_osn3_entry.config(state="normal")
            self.diam_osn4_entry.config(state="normal")
            self.diam_verha1_entry.config(state="normal")
            self.diam_verha2_entry.config(state="normal")
            self.diam_verha3_entry.config(state="normal")
            self.diam_verha4_entry.config(state="normal")
            self.h1_entry.config(state="normal")
            self.h2_entry.config(state="normal")
            self.h3_entry.config(state="normal")
            self.h4_entry.config(state="normal")
            self.massa_rvs1_entry.config(state="normal")
            self.massa_rvs2_entry.config(state="normal")
            self.massa_rvs3_entry.config(state="normal")
            self.massa_rvs4_entry.config(state="normal")
            self.obsch_massa_rvs1_entry.config(state="normal")
            self.obsch_massa_rvs2_entry.config(state="normal")
            self.obsch_massa_rvs3_entry.config(state="normal")
            self.obsch_massa_rvs4_entry.config(state="normal")

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()