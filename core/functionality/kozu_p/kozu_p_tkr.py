import tkinter as tk

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
    make_path_pdf,
    make_multiple_path
)
from core.exceptions import AddPlsPolePathException
from core.functionality.kozu.utils import (
    make_tkr,
    make_vor
    # make_pzg,
    # make_pz
)


class KozuPTkr(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("КОЗ-У-П")
        self.geometry("840x833+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        self.module_bg = tk.Frame(
            self,
            width=820,
            height=823,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            text="Назад",
            command=self.back_to_main_window
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

        self.v_m_bpla_label = tk.Label(
            self,
            text='Скор. и масса БПЛА',
            width=28,
            anchor="e"
        )
        self.v_m_bpla_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.zasch_obj_label = tk.Label(
            self,
            text='Защищаемый объект',
            width=28,
            anchor="e"
        )
        self.zasch_obj_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.steel_label = tk.Label(
            self,
            text='Используемые марки стали',
            width=28,
            anchor="e"
        )
        self.steel_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.anal_steel_label = tk.Label(
            self,
            text='Аналог стали',
            width=28,
            anchor="e"
        )
        self.anal_steel_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.konst_label = tk.Label(
            self,
            text='Конструкция КОЗУ: сети',
            width=28,
            anchor="e"
        )
        self.konst_entry = tk.Entry(
            self,
            width=50,
            relief="sunken",
            bd=2
        )

        self.kp_bolt_label = tk.Label(
            self,
            text='Классы прочности болтов',
            width=28,
            anchor="e"
        )
        self.kp_bolt_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.fundament_label = tk.Label(
            self,
            text='Фундамент',
            width=28,
            anchor="e"
        )
        self.fundament_combobox = Combobox(
            self,
            values=(
                "лежневый фундамент",
                "плита"
            ),
            width=47,
        )

        self.fund_osn_label = tk.Label(
            self,
            text='Фундаментная подготовка',
            width=28,
            anchor="e"
        )
        self.fund_osn_combobox = Combobox(
            self,
            values=(
                "песчаный слой, геотекстиль и щебеночный слой",
                "котлован с полной выемкой и заменой грунта",
                "котлован с частичной выемкой и заменой грунта",
                "насыпь"
            ),
            width=47,
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

        self.kol_obj_label = tk.Label(
            self,
            text="ТКР: Количество КОЗ-У-П, шт",
            width=28,
            anchor="e"
        )
        self.kol_obj_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
        )

        self.massa_obsch_label = tk.Label(
            self,
            text="ТКР: Общая металлоемкость, т",
            width=28,
            anchor="e"
        )
        self.massa_obsch_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
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

        self.mont_vremya_label = tk.Label(
            self,
            text='Время на монтаж',
            width=28,
            anchor="e"
        )
        self.mont_vremya_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.expl_god_label = tk.Label(
            self,
            text='Год ввода в экспл.',
            width=28,
            anchor="e"
        )
        self.expl_god_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.shag_yach_label = tk.Label(
            self,
            text='ТКР: Шаг ячейки м/п сетки, м',
            width=28,
            anchor="e"
        )
        self.shag_yach_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.quantity_of_obj_label = tk.Label(
            self,
            text='Количество марок КОЗ-У-П, шт',
            width=28,
            anchor="e"
        )
        self.quantity_of_obj_combobox = Combobox(
            self,
            width=12,
            values=(
                "1",
                "2",
                "3",
                "4"
            )
        )
        self.quantity_of_obj_combobox.bind("<<ComboboxSelected>>", self.activate_kozu_p)

        self.kozup_1_label = tk.Label(
            self,
            text='КОЗ-У-П 1',
            width=8,
            anchor="e"
        )
        self.kozup_2_label = tk.Label(
            self,
            text='КОЗ-У-П 2',
            width=8,
            anchor="e"
        )
        self.kozup_3_label = tk.Label(
            self,
            text='КОЗ-У-П 3',
            width=8,
            anchor="e"
        )
        self.kozup_4_label = tk.Label(
            self,
            text='КОЗ-У-П 4',
            width=8,
            anchor="e"
        )
        
        self.zaschichaemyi_obj_label = tk.Label(
            self,
            text='Защищаемый объект',
            width=28,
            anchor="e"
        )
        self.length_kozup_label = tk.Label(
            self,
            text='Длина в осях, мм',
            width=28,
            anchor="e"
        )
        self.width_kozup_label = tk.Label(
            self,
            text='Ширина в осях, мм',
            width=28,
            anchor="e"
        )
        self.h_label = tk.Label(
            self,
            text='Высота, мм',
            width=28,
            anchor="e"
        )
        self.massa_kozup_label = tk.Label(
            self,
            text='Масса 1-го КОЗ-У-П, т',
            width=28,
            anchor="e"
        )

        self.zaschichaemyi_obj1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.zaschichaemyi_obj2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.zaschichaemyi_obj3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.zaschichaemyi_obj4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.length_kozup1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.length_kozup2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.length_kozup3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.length_kozup4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.width_kozup1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.width_kozup2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.width_kozup3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.width_kozup4_entry = tk.Entry(
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

        self.massa_kozup1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.massa_kozup2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.massa_kozup3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.massa_kozup4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.post_nagr_label = tk.Label(
            self,
            text='ПЗ:6.2. Пост. нагрузки',
            width=28,
            anchor="e"
        )
        self.post_nagr_entry = tk.Entry(
            self,
            width=50,
            relief="sunken",
            bd=2
        )

        self.strela_label = tk.Label(
            self,
            text='ПЗ:6.3 Стрела провиса, м',
            width=28,
            anchor="e"
        )
        self.strela_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.natyajenie_label = tk.Label(
            self,
            text='ПЗ:6.3 Натяжение нити, кг/м',
            width=28,
            anchor="e"
        )
        self.natyajenie_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.klass_betona_label = tk.Label(
            self,
            text='ПЗФ: Класс бетона',
            width=28,
            anchor="e"
        )
        self.klass_betona_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.morozostoikost_label = tk.Label(
            self,
            text='ПЗФ: Морозостойкость бетона',
            width=28,
            anchor="e"
        )
        self.morozostoikost_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.vodonepronicaemost_label = tk.Label(
            self,
            text='ПЗФ: Водонепрониц. бетона',
            width=28,
            anchor="e"
        )
        self.vodonepronicaemost_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.dlina_elem_label = tk.Label(
            self,
            text='ПЗФ: Длина лежня, дм',
            width=28,
            anchor="e"
        )
        self.dlina_elem_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.speca_label = tk.Label(
            self,
            text='ТКР: Спецификация.png',
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

        self.vid_kozu_p_label = tk.Label(
            self,
            text='ПЗ: Виды КОЗ-У-П.png',
            width=28,
            anchor="e"
        )
        self.vid_kozu_p_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_vid_kozu_p_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_vid_kozup
        )

        self.table3_label = tk.Label(
            self,
            text='ПЗ: Таблица 3.png',
            width=28,
            anchor="e"
        )
        self.table3_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_table3_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_table3
        )

        self.table4_label = tk.Label(
            self,
            text='ПЗ: Таблица 4.png',
            width=28,
            anchor="e"
        )
        self.table4_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_table4_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_table4
        )

        self.table5_label = tk.Label(
            self,
            text='ПЗ: Таблица 5.png',
            width=28,
            anchor="e"
        )
        self.table5_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_table5_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_table5
        )

        self.table6_label = tk.Label(
            self,
            text='ПЗ: Таблица 6.png',
            width=28,
            anchor="e"
        )
        self.table6_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_table6_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_table6
        )

        self.raschet_model_label = tk.Label(
            self,
            text='Прил.Б, Расчет модель.png',
            width=28,
            anchor="e"
        )
        self.raschet_model_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_raschet_model_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_raschet_model
        )

        self.usiliya1_label = tk.Label(
            self,
            text='Прил.B, Усилия N.png',
            width=28,
            anchor="e"
        )
        self.usiliya1_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usiliya1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usiliya1
        )

        self.usiliya2_label = tk.Label(
            self,
            text='Прил.B, Усилия My.png',
            width=28,
            anchor="e"
        )
        self.usiliya2_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usiliya2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usiliya2
        )

        self.usiliya3_label = tk.Label(
            self,
            text='Прил.B, Усилия Mz.png',
            width=28,
            anchor="e"
        )
        self.usiliya3_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usiliya3_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usiliya3
        )

        self.usiliya4_label = tk.Label(
            self,
            text='Прил.B, Усилия Qz.png',
            width=28,
            anchor="e"
        )
        self.usiliya4_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usiliya4_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usiliya4
        )

        self.generate_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usiliya4
        )

    def run(self):
        self.draw_widgets()
        # self.paste_kozu_project_data()
        self.mainloop()

    # def paste_kozu_project_data(self):
    #     try:
            # self.sp_wind_reg_combobox.set(self.parent.sp_wind_reg)
            # self.wind_nagr_entry.insert(0, self.parent.wind_nagr)
            # self.sp_sneg_reg_combobox.set(self.parent.sp_sneg_reg)
            # self.snow_nagr_entry.insert(0, self.parent.snow_nagr)
            # self.golol_rayon_combobox.set(self.parent.golol_rayon)
            # self.zasch_obj_entry.insert(0, self.parent.zasch_obj)
            # self.territoria_raspoloj_entry.insert(0, self.parent.territoria_raspoloj)
            # self.rayon_str_entry.insert(0, self.parent.rayon_str)
            # self.quantity_of_rvs_combobox.set(self.parent.quantity_of_rvs)
            # self.activate_rvs(None)
            # self.rvs1_entry.insert(0, self.parent.rvs1)
            # self.rvs2_entry.insert(0, self.parent.rvs2)
            # self.rvs3_entry.insert(0, self.parent.rvs3)
            # self.rvs4_entry.insert(0, self.parent.rvs4)
            # self.diam_osn1_entry.insert(0, self.parent.diam_osn1)
            # self.diam_osn2_entry.insert(0, self.parent.diam_osn2)
            # self.diam_osn3_entry.insert(0, self.parent.diam_osn3)
            # self.diam_osn4_entry.insert(0, self.parent.diam_osn4)
            # self.diam_verha1_entry.insert(0, self.parent.diam_verha1)
            # self.diam_verha2_entry.insert(0, self.parent.diam_verha2)
            # self.diam_verha3_entry.insert(0, self.parent.diam_verha3)
            # self.diam_verha4_entry.insert(0, self.parent.diam_verha4)
            # self.h1_entry.insert(0, self.parent.h1)
            # self.h2_entry.insert(0, self.parent.h2)
            # self.h3_entry.insert(0, self.parent.h3)
            # self.h4_entry.insert(0, self.parent.h4)
            # self.massa_rvs1_entry.insert(0, self.parent.massa_rvs1)
            # self.massa_rvs2_entry.insert(0, self.parent.massa_rvs2)
            # self.massa_rvs3_entry.insert(0, self.parent.massa_rvs3)
            # self.massa_rvs4_entry.insert(0, self.parent.massa_rvs4)
            # self.kol_rvs1_entry.insert(0, self.parent.kol_rvs1)
            # self.kol_rvs2_entry.insert(0, self.parent.kol_rvs2)
            # self.kol_rvs3_entry.insert(0, self.parent.kol_rvs3)
            # self.kol_rvs4_entry.insert(0, self.parent.kol_rvs4)
            # self.ploschad_uchastka1_entry.insert(0, self.parent.ploschad_uchastka1)
            # self.ploschad_uchastka2_entry.insert(0, self.parent.ploschad_uchastka2)
            # self.ploschad_uchastka3_entry.insert(0, self.parent.ploschad_uchastka3)
            # self.ploschad_uchastka4_entry.insert(0, self.parent.ploschad_uchastka4)
            # self.list_sogl_entry.insert(0, self.parent.list_sogl)
            # self.kont_zazel_entry.insert(0, self.parent.kont_zazel)
            # self.mont_schema_entry.insert(0, self.parent.mont_schema)
            # self.vid_kozu_entry.insert(0, self.parent.vid_kozu)
            # self.vid_kozu2_entry.insert(0, self.parent.vid_kozu2)
            # self.vid_kozu3_entry.insert(0, self.parent.vid_kozu3)
            # self.vid_kozu4_entry.insert(0, self.parent.vid_kozu4)
            # self.speca_entry.insert(0, self.parent.speca)
            # self.speca_pz_entry.insert(0, self.parent.speca_pz)
            # self.speca_pz2_entry.insert(0, self.parent.speca_pz2)
            # self.speca_pz3_entry.insert(0, self.parent.speca_pz3)
            # self.speca_pz4_entry.insert(0, self.parent.speca_pz4)
            # self.eskiz_kozu_entry.insert(0, self.parent.eskiz_kozu)
        # except Exception as e:
        #     # print("!INFO!: Сохраненные данные отсутствуют.")
        #     print(e)

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=2)
        self.sp_wind_reg_label.place(x=15, y=42)
        self.sp_wind_reg_combobox.place(x=220, y=42)
        self.wind_nagr_label.place(x=15, y=65)
        self.wind_nagr_entry.place(x=220, y=65)
        self.golol_rayon_label.place(x=15, y=88)
        self.golol_rayon_combobox.place(x=220, y=88)
        self.v_m_bpla_label.place(x=15, y=111)
        self.v_m_bpla_entry.place(x=220, y=111)
        self.zasch_obj_label.place(x=15, y=134)
        self.zasch_obj_entry.place(x=220, y=134)
        self.steel_label.place(x=15, y=157)
        self.steel_entry.place(x=220, y=157)
        self.anal_steel_label.place(x=15, y=180)
        self.anal_steel_entry.place(x=220, y=180)
        self.konst_label.place(x=315, y=157)
        self.konst_entry.place(x=520, y=157)
        self.kp_bolt_label.place(x=15, y=226)
        self.kp_bolt_entry.place(x=220, y=226)
        self.fundament_label.place(x=315, y=42)
        self.fundament_combobox.place(x=520, y=42)
        self.fund_osn_label.place(x=315, y=65)
        self.fund_osn_combobox.place(x=520, y=65)
        self.rayon_str_label.place(x=315, y=88)
        self.rayon_str_entry.place(x=520, y=88)
        self.kol_obj_label.place(x=315, y=111)
        self.kol_obj_entry.place(x=520, y=111)
        self.massa_obsch_label.place(x=315, y=134)
        self.massa_obsch_entry.place(x=520, y=134)
        self.territoria_raspoloj_label.place(x=15, y=203)
        self.territoria_raspoloj_entry.place(x=220, y=203)
        self.mont_vremya_label.place(x=15, y=249)
        self.mont_vremya_entry.place(x=220, y=249)
        self.expl_god_label.place(x=15, y=272)
        self.expl_god_entry.place(x=220, y=272)
        self.shag_yach_label.place(x=15, y=295)
        self.shag_yach_entry.place(x=220, y=295)
        self.post_nagr_label.place(x=315, y=180)
        self.post_nagr_entry.place(x=520, y=180)
        self.strela_label.place(x=315, y=203)
        self.strela_entry.place(x=520, y=203)
        self.natyajenie_label.place(x=315, y=226)
        self.natyajenie_entry.place(x=520, y=226)
        self.klass_betona_label.place(x=315, y=249)
        self.klass_betona_entry.place(x=520, y=249)
        self.morozostoikost_label.place(x=315, y=272)
        self.morozostoikost_entry.place(x=520, y=272)
        self.vodonepronicaemost_label.place(x=315, y=295)
        self.vodonepronicaemost_entry.place(x=520, y=295)
        self.quantity_of_obj_label.place(x=15, y=318)
        self.quantity_of_obj_combobox.place(x=220, y=318)
        self.dlina_elem_label.place(x=315, y=318)
        self.dlina_elem_entry.place(x=520, y=318)
        self.kozup_1_label.place(x=240, y=341)
        self.kozup_2_label.place(x=334, y=341)
        self.kozup_3_label.place(x=428, y=341)
        self.kozup_4_label.place(x=522, y=341)
        self.zaschichaemyi_obj_label.place(x=15, y=364)
        self.length_kozup_label.place(x=15, y=387)
        self.width_kozup_label.place(x=15, y=410)
        self.h_label.place(x=15, y=433)
        self.massa_kozup_label.place(x=15, y=456)
        self.zaschichaemyi_obj1_entry.place(x=220, y=364)
        self.zaschichaemyi_obj2_entry.place(x=314, y=364)
        self.zaschichaemyi_obj3_entry.place(x=408, y=364)
        self.zaschichaemyi_obj4_entry.place(x=502, y=364)
        self.length_kozup1_entry.place(x=220, y=387)
        self.length_kozup2_entry.place(x=314, y=387)
        self.length_kozup3_entry.place(x=408, y=387)
        self.length_kozup4_entry.place(x=502, y=387) 
        self.width_kozup1_entry.place(x=220, y=410)
        self.width_kozup2_entry.place(x=314, y=410)
        self.width_kozup3_entry.place(x=408, y=410)
        self.width_kozup4_entry.place(x=502, y=410)
        self.h1_entry.place(x=220, y=433)
        self.h2_entry.place(x=314, y=433)
        self.h3_entry.place(x=408, y=433)
        self.h4_entry.place(x=502, y=433)
        self.massa_kozup1_entry.place(x=220, y=456)
        self.massa_kozup2_entry.place(x=314, y=456)
        self.massa_kozup3_entry.place(x=408, y=456)
        self.massa_kozup4_entry.place(x=502, y=456)
        self.speca_label.place(x=15, y=484)
        self.speca_entry.place(x=220, y=484)
        self.browse_for_speca_button.place(x=497, y=482)
        self.vid_kozu_p_label.place(x=15, y=512)
        self.vid_kozu_p_entry.place(x=220, y=512)
        self.browse_for_vid_kozu_p_button.place(x=497, y=510)
        self.table3_label.place(x=15, y=540)
        self.table3_entry.place(x=220, y=540)
        self.browse_for_table3_button.place(x=497, y=538)
        self.table4_label.place(x=15, y=568)
        self.table4_entry.place(x=220, y=568)
        self.browse_for_table4_button.place(x=497, y=566)
        self.table5_label.place(x=15, y=596)
        self.table5_entry.place(x=220, y=596)
        self.browse_for_table5_button.place(x=497, y=594)
        self.table6_label.place(x=15, y=624)
        self.table6_entry.place(x=220, y=624)
        self.browse_for_table6_button.place(x=497, y=622)
        self.raschet_model_label.place(x=15, y=652)
        self.raschet_model_entry.place(x=220, y=652)
        self.browse_for_raschet_model_button.place(x=497, y=650)
        self.usiliya1_label.place(x=15, y=680)
        self.usiliya1_entry.place(x=220, y=680)
        self.browse_for_usiliya1_button.place(x=497, y=678)
        self.usiliya2_label.place(x=15, y=708)
        self.usiliya2_entry.place(x=220, y=708)
        self.browse_for_usiliya2_button.place(x=497, y=706)
        self.usiliya3_label.place(x=15, y=736)
        self.usiliya3_entry.place(x=220, y=736)
        self.browse_for_usiliya3_button.place(x=497, y=734)
        self.usiliya4_label.place(x=15, y=764)
        self.usiliya4_entry.place(x=220, y=764)
        self.browse_for_usiliya4_button.place(x=497, y=762)
        
        self.tkr_button.place(x=250, y=790)
        # self.vor_button.place(x=395, y=740)

    def generate(self):
        make_tkr(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            developer=self.parent.developer,
            sp_wind_region=self.sp_wind_reg_combobox.get(),
            wind_nagr=self.wind_nagr_entry.get(),
            golol_rayon=self.golol_rayon_combobox.get(),
            v_m_bpla=self.v_m_bpla_entry.get(),
            zasch_obj=self.zasch_obj_entry.get(),
            steel=self.steel_entry.get(),
            anal_steel=self.anal_steel_entry.get(),
            konst=self.konst_entry.get(),
            kp_bolt=self.kp_bolt_entry.get(),
            fundament=self.fundament_combobox.get(),
            fund_osn=self.fund_osn_combobox.get(),
            rayon_str=self.rayon_str_entry.get(),
            kol_obj=self.kol_obj_entry.get(),
            massa_obsch=self.massa_obsch_entry.get(),
            territoria_raspoloj=self.territoria_raspoloj_entry.get(),
            mont_vremya=self.mont_vremya_entry.get(),
            expl_god=self.expl_god_entry.get(),
            shag_yach=self.shag_yach_entry.get(),
            post_nagr=self.post_nagr_entry.get(),
            strela=self.strela_entry.get(),
            natyajenie=self.natyajenie_entry.get(),
            klass_betona=self.klass_betona_entry.get(),
            morozostoikost=self.morozostoikost_entry.get(),
            vodonepronicaemost=self.vodonepronicaemost_entry.get(),
            quantity_of_obj=self.quantity_of_obj_combobox.get(),
            dlina_elem=self.dlina_elem_entry.get(),
            zaschichaemyi_obj1=self.zaschichaemyi_obj1_entry.get(),
            zaschichaemyi_obj2=self.zaschichaemyi_obj2_entry.get(),
            zaschichaemyi_obj3=self.zaschichaemyi_obj3_entry.get(),
            zaschichaemyi_obj4=self.zaschichaemyi_obj4_entry.get(),
            length_kozup1=self.length_kozup1_entry.get(),
            length_kozup2=self.length_kozup2_entry.get(),
            length_kozup3=self.length_kozup3_entry.get(),
            length_kozup4=self.length_kozup4_entry.get(),
            width_kozup1=self.width_kozup1_entry.get(),
            width_kozup2=self.width_kozup2_entry.get(),
            width_kozup3=self.width_kozup3_entry.get(),
            width_kozup4=self.width_kozup4_entry.get(),
            h1=self.h1_entry.get(),
            h2=self.h2_entry.get(),
            h3=self.h3_entry.get(),
            h4=self.h4_entry.get(),
            massa_kozup1=self.massa_kozup1_entry.get(),
            massa_kozup2=self.massa_kozup2_entry.get(),
            massa_kozup3=self.massa_kozup3_entry.get(),
            massa_kozup4=self.massa_kozup4_entry.get(),
            speca=self.speca_entry.get(),
            vid_kozu_p=self.vid_kozu_p_entry.get(),
            table3=self.table3_entry.get(),
            table4=self.table4_entry.get(),
            table5=self.table5_entry.get(),
            table6=self.table6_entry.get(),
            raschet_model=self.raschet_model_entry.get(),
            usiliya1=self.usiliya1_entry.get(),
            usiliya2=self.usiliya2_entry.get(),
            usiliya3=self.usiliya3_entry.get(),
            usiliya4=self.usiliya4_entry.get(),
        )
        # list_sogl=self.list_sogl_entry.get().split("КОЗУ (инженерная)")[1]
        # kont_zazel=self.kont_zazel_entry.get().split("КОЗУ (инженерная)")[1]
        # mont_schema=self.mont_schema_entry.get().split("КОЗУ (инженерная)")[1]
        # vid_kozu=self.vid_kozu_entry.get().split("КОЗУ (инженерная)")[1]
        # if self.vid_kozu2_entry.get():
        #     vid_kozu2=self.vid_kozu2_entry.get().split("КОЗУ (инженерная)")[1]
        # else:
        #     vid_kozu2=""
        # if self.vid_kozu3_entry.get():
        #     vid_kozu3=self.vid_kozu3_entry.get().split("КОЗУ (инженерная)")[1]
        # else:
        #     vid_kozu3=""
        # if self.vid_kozu4_entry.get():
        #     vid_kozu4=self.vid_kozu4_entry.get().split("КОЗУ (инженерная)")[1]
        # else:
        #     vid_kozu4=""
        # speca=self.speca_entry.get().split("КОЗУ (инженерная)")[1]
        # speca_pz=self.speca_pz_entry.get().split("КОЗУ (инженерная)")[1]
        # if self.speca_pz2_entry.get():
        #     speca_pz2=self.speca_pz2_entry.get().split("КОЗУ (инженерная)")[1]
        # else:
        #     speca_pz2=""
        # if self.speca_pz3_entry.get():
        #     speca_pz3=self.speca_pz3_entry.get().split("КОЗУ (инженерная)")[1]
        # else:
        #     speca_pz3=""
        # if self.speca_pz4_entry.get():
        #     speca_pz4=self.speca_pz4_entry.get().split("КОЗУ (инженерная)")[1]
        # else:
        #     speca_pz4=""
        # eskiz_kozu=self.eskiz_kozu_entry.get().split("КОЗУ (инженерная)")[1]
        # self.db.add_kozu_project_data(
        #     initial_data_id=self.parent.initial_data_id,
        #     is_gabion=self.is_gabion_var.get(),
        #     sp_wind_reg=self.sp_wind_reg_combobox.get(),
        #     wind_nagr=self.wind_nagr_entry.get(),
        #     sp_sneg_reg=self.sp_sneg_reg_combobox.get(),
        #     snow_nagr=self.snow_nagr_entry.get(),
        #     golol_rayon=self.golol_rayon_combobox.get(),
        #     zasch_obj=self.zasch_obj_entry.get(),
        #     territoria_raspoloj=self.territoria_raspoloj_entry.get(),
        #     rayon_str=self.rayon_str_entry.get(),
        #     quantity_of_rvs=self.quantity_of_rvs_combobox.get(),
        #     rvs1=self.rvs1_entry.get(),
        #     rvs2=self.rvs2_entry.get(),
        #     rvs3=self.rvs3_entry.get(),
        #     rvs4=self.rvs4_entry.get(),
        #     diam_osn1=self.diam_osn1_entry.get(),
        #     diam_osn2=self.diam_osn2_entry.get(),
        #     diam_osn3=self.diam_osn3_entry.get(),
        #     diam_osn4=self.diam_osn4_entry.get(),
        #     diam_verha1=self.diam_verha1_entry.get(),
        #     diam_verha2=self.diam_verha2_entry.get(),
        #     diam_verha3=self.diam_verha3_entry.get(),
        #     diam_verha4=self.diam_verha4_entry.get(),
        #     h1=self.h1_entry.get(),
        #     h2=self.h2_entry.get(),
        #     h3=self.h3_entry.get(),
        #     h4=self.h4_entry.get(),
        #     massa_rvs1=self.massa_rvs1_entry.get(),
        #     massa_rvs2=self.massa_rvs2_entry.get(),
        #     massa_rvs3=self.massa_rvs3_entry.get(),
        #     massa_rvs4=self.massa_rvs4_entry.get(),
        #     kol_rvs1=self.kol_rvs1_entry.get(),
        #     kol_rvs2=self.kol_rvs2_entry.get(),
        #     kol_rvs3=self.kol_rvs3_entry.get(),
        #     kol_rvs4=self.kol_rvs4_entry.get(),
        #     ploschad_uchastka1=self.ploschad_uchastka1_entry.get(),
        #     ploschad_uchastka2=self.ploschad_uchastka2_entry.get(),
        #     ploschad_uchastka3=self.ploschad_uchastka3_entry.get(),
        #     ploschad_uchastka4=self.ploschad_uchastka4_entry.get(),
        #     list_sogl=list_sogl,
        #     kont_zazel=kont_zazel,
        #     mont_schema=mont_schema,
        #     vid_kozu=vid_kozu,
        #     vid_kozu2=vid_kozu2,
        #     vid_kozu3=vid_kozu3,
        #     vid_kozu4=vid_kozu4,
        #     speca=speca,
        #     speca_pz=speca_pz,
        #     speca_pz2=speca_pz2,
        #     speca_pz3=speca_pz3,
        #     speca_pz4=speca_pz4,
        #     eskiz_kozu=eskiz_kozu
        # )

    # def call_make_vor(self):
    #     if self.quantity_of_rvs_combobox.get() == "1":
    #         make_vor(
    #             n=self.kol_rvs1_entry.get(),
    #             h=self.h1_entry.get(),
    #             d_nijn=self.diam_osn1_entry.get(),
    #             d_verh=self.diam_verha1_entry.get(),
    #             m=self.massa_rvs1_entry.get(),
    #             is_gabion=self.is_gabion_var.get(),
    #             rvs=self.rvs1_entry.get()
    #         )
    #     else:
    #         mb.showinfo("ERROR", "Количество марок РВС должно быть 1!")

    def paste_wind(self, event):
        wind_key = self.sp_wind_reg_combobox.get()
        self.wind_nagr_entry.delete(0, tk.END)
        self.wind_nagr_entry.insert(0, sp_wind_reg_dict[wind_key])

    def browse_for_speca(self):
        self.file_path = make_path_png()
        self.speca_entry.delete("0", "end") 
        self.speca_entry.insert("insert", self.file_path)

    def browse_for_table3(self):
        self.file_path = make_path_png()
        self.table3_entry.delete("0", "end") 
        self.table3_entry.insert("insert", self.file_path)

    def browse_for_table4(self):
        self.file_path = make_path_png()
        self.table4_entry.delete("0", "end") 
        self.table4_entry.insert("insert", self.file_path)

    def browse_for_table5(self):
        self.file_path = make_path_png()
        self.table5_entry.delete("0", "end") 
        self.table5_entry.insert("insert", self.file_path)

    def browse_for_table6(self):
        self.file_path = make_path_png()
        self.table6_entry.delete("0", "end") 
        self.table6_entry.insert("insert", self.file_path)

    def browse_for_raschet_model(self):
        self.file_path = make_path_png()
        self.raschet_model_entry.delete("0", "end") 
        self.raschet_model_entry.insert("insert", self.file_path)

    def browse_for_usiliya1(self):
        self.file_path = make_path_png()
        self.usiliya1_entry.delete("0", "end") 
        self.usiliya1_entry.insert("insert", self.file_path)

    def browse_for_usiliya2(self):
        self.file_path = make_path_png()
        self.usiliya2_entry.delete("0", "end") 
        self.usiliya2_entry.insert("insert", self.file_path)
    
    def browse_for_usiliya3(self):
        self.file_path = make_path_png()
        self.usiliya3_entry.delete("0", "end") 
        self.usiliya3_entry.insert("insert", self.file_path)

    def browse_for_usiliya4(self):
        self.file_path = make_path_png()
        self.usiliya4_entry.delete("0", "end") 
        self.usiliya4_entry.insert("insert", self.file_path)

    def browse_for_vid_kozup(self):
        self.file_path = make_multiple_path()
        self.usiliya4_entry.delete("0", "end") 
        self.usiliya4_entry.insert("insert", self.file_path)

    def activate_kozu_p(self, event):
        if self.quantity_of_obj_combobox.get() == "1":
            self.zaschichaemyi_obj2_entry.delete(0, tk.END)
            self.zaschichaemyi_obj3_entry.delete(0, tk.END)
            self.zaschichaemyi_obj4_entry.delete(0, tk.END)
            self.zaschichaemyi_obj2_entry.config(state="disabled")
            self.zaschichaemyi_obj3_entry.config(state="disabled")
            self.zaschichaemyi_obj4_entry.config(state="disabled")
            self.zaschichaemyi_obj1_entry.config(state="normal")
            self.length_kozup2_entry.delete(0, tk.END)
            self.length_kozup3_entry.delete(0, tk.END)
            self.length_kozup4_entry.delete(0, tk.END)
            self.length_kozup2_entry.config(state="disabled")
            self.length_kozup3_entry.config(state="disabled")
            self.length_kozup4_entry.config(state="disabled")
            self.length_kozup1_entry.config(state="normal")
            self.width_kozup2_entry.delete(0, tk.END)
            self.width_kozup3_entry.delete(0, tk.END)
            self.width_kozup4_entry.delete(0, tk.END)
            self.width_kozup2_entry.config(state="disabled")
            self.width_kozup3_entry.config(state="disabled")
            self.width_kozup4_entry.config(state="disabled")
            self.width_kozup1_entry.config(state="normal")
            self.h2_entry.delete(0, tk.END)
            self.h3_entry.delete(0, tk.END)
            self.h4_entry.delete(0, tk.END)
            self.h2_entry.config(state="disabled")
            self.h3_entry.config(state="disabled")
            self.h4_entry.config(state="disabled")
            self.h1_entry.config(state="normal")
            self.massa_kozup2_entry.delete(0, tk.END)
            self.massa_kozup3_entry.delete(0, tk.END)
            self.massa_kozup4_entry.delete(0, tk.END)
            self.massa_kozup2_entry.config(state="disabled")
            self.massa_kozup3_entry.config(state="disabled")
            self.massa_kozup4_entry.config(state="disabled")
            self.massa_kozup1_entry.config(state="normal")
        elif self.quantity_of_obj_combobox.get() == "2":
            self.zaschichaemyi_obj3_entry.delete(0, tk.END)
            self.zaschichaemyi_obj4_entry.delete(0, tk.END)
            self.zaschichaemyi_obj3_entry.config(state="disabled")
            self.zaschichaemyi_obj4_entry.config(state="disabled")
            self.zaschichaemyi_obj1_entry.config(state="normal")
            self.zaschichaemyi_obj2_entry.config(state="normal")
            self.length_kozup3_entry.delete(0, tk.END)
            self.length_kozup4_entry.delete(0, tk.END)
            self.length_kozup3_entry.config(state="disabled")
            self.length_kozup4_entry.config(state="disabled")
            self.length_kozup1_entry.config(state="normal")
            self.length_kozup2_entry.config(state="normal")
            self.width_kozup3_entry.delete(0, tk.END)
            self.width_kozup4_entry.delete(0, tk.END)
            self.width_kozup3_entry.config(state="disabled")
            self.width_kozup4_entry.config(state="disabled")
            self.width_kozup1_entry.config(state="normal")
            self.width_kozup2_entry.config(state="normal")
            self.h3_entry.delete(0, tk.END)
            self.h4_entry.delete(0, tk.END)
            self.h3_entry.config(state="disabled")
            self.h4_entry.config(state="disabled")
            self.h1_entry.config(state="normal")
            self.h2_entry.config(state="normal")
            self.massa_kozup3_entry.delete(0, tk.END)
            self.massa_kozup4_entry.delete(0, tk.END)
            self.massa_kozup3_entry.config(state="disabled")
            self.massa_kozup4_entry.config(state="disabled")
            self.massa_kozup1_entry.config(state="normal")
            self.massa_kozup2_entry.config(state="normal")
        elif self.quantity_of_obj_combobox.get() == "3":
            self.zaschichaemyi_obj4_entry.delete(0, tk.END)
            self.zaschichaemyi_obj4_entry.config(state="disabled")
            self.zaschichaemyi_obj1_entry.config(state="normal")
            self.zaschichaemyi_obj2_entry.config(state="normal")
            self.zaschichaemyi_obj3_entry.config(state="normal")
            self.length_kozup4_entry.delete(0, tk.END)
            self.length_kozup4_entry.config(state="disabled")
            self.length_kozup1_entry.config(state="normal")
            self.length_kozup2_entry.config(state="normal")
            self.length_kozup3_entry.config(state="normal")
            self.width_kozup4_entry.delete(0, tk.END)
            self.width_kozup4_entry.config(state="disabled")
            self.width_kozup1_entry.config(state="normal")
            self.width_kozup2_entry.config(state="normal")
            self.width_kozup3_entry.config(state="normal")
            self.h4_entry.delete(0, tk.END)
            self.h4_entry.config(state="disabled")
            self.h1_entry.config(state="normal")
            self.h2_entry.config(state="normal")
            self.h3_entry.config(state="normal")
            self.massa_kozup4_entry.delete(0, tk.END)
            self.massa_kozup4_entry.config(state="disabled")
            self.massa_kozup1_entry.config(state="normal")
            self.massa_kozup2_entry.config(state="normal")
            self.massa_kozup3_entry.config(state="normal")
        else:
            self.zaschichaemyi_obj1_entry.config(state="normal")
            self.zaschichaemyi_obj2_entry.config(state="normal")
            self.zaschichaemyi_obj3_entry.config(state="normal")
            self.zaschichaemyi_obj4_entry.config(state="normal")
            self.length_kozup1_entry.config(state="normal")
            self.length_kozup2_entry.config(state="normal")
            self.length_kozup3_entry.config(state="normal")
            self.length_kozup4_entry.config(state="normal")
            self.width_kozup1_entry.config(state="normal")
            self.width_kozup2_entry.config(state="normal")
            self.width_kozup3_entry.config(state="normal")
            self.width_kozup4_entry.config(state="normal")
            self.h1_entry.config(state="normal")
            self.h2_entry.config(state="normal")
            self.h3_entry.config(state="normal")
            self.h4_entry.config(state="normal")
            self.massa_kozup1_entry.config(state="normal")
            self.massa_kozup2_entry.config(state="normal")
            self.massa_kozup3_entry.config(state="normal")
            self.massa_kozup4_entry.config(state="normal")
    
    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()