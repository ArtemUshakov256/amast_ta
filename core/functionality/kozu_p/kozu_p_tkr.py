import tkinter as tk

from tkinter import messagebox as mb
from tkinter.ttk import Combobox


from core.constants import (
    sp_wind_reg_dict,
    STR_KLIM_ZONA,
    VID_KLIM,
    SBROS,
    SET
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
from core.functionality.kozu_p.utils import (
    make_tkr,
    # make_pzg,
    # make_pz
)


class KozuPTkr(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("КОЗ-У-П")
        self.geometry("1140x966+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        self.module_bg = tk.Frame(
            self,
            width=1120,
            height=956,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            text="Назад",
            command=self.back_to_main_window
        )

        self.klimat_label = tk.Label(
            self,
            text='КЛИМАТ',
            width=28,
            anchor="e"
        )

        self.project_info_label = tk.Label(
            self,
            text='ИНФОРМАЦИЯ О ПРОЕКТЕ',
            width=28,
            anchor="e"
        )

        self.konstr_resh_label = tk.Label(
            self,
            text='КОНСТРУКТИВНЫЕ РЕШЕНИЯ',
            width=28,
            anchor="e"
        )

        self.zazemlenie_label = tk.Label(
            self,
            text='ЗАЗЕМЛЕНИЕ',
            width=28,
            anchor="e"
        )

        self.dop_info_label = tk.Label(
            self,
            text='ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ',
            width=30,
            anchor="e"
        )

        self.is_svaya_var = tk.IntVar()
        self.is_svaya_checkbutton = tk.Checkbutton(
            self,
            text="Есть свая",
            variable=self.is_svaya_var,
        )

        self.is_ferma_usil_var = tk.IntVar()
        self.is_ferma_usil_checkbutton = tk.Checkbutton(
            self,
            text="Есть ферма усиления",
            variable=self.is_ferma_usil_var,
        )

        self.is_tros_ferma_var = tk.IntVar()
        self.is_tros_ferma_checkbutton = tk.Checkbutton(
            self,
            text="Есть тросовая ферма",
            variable=self.is_tros_ferma_var,
        )

        self.is_existing_ground_var = tk.IntVar()
        self.is_existing_ground_checkbutton = tk.Checkbutton(
            self,
            text="Подключение к сущ. конт. заземления",
            variable=self.is_existing_ground_var,
        )

        self.bartal_code_label = tk.Label(
            self,
            text='Шифр Бартала',
            width=28,
            anchor="e"
        )
        self.bartal_code_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.object_titul_label = tk.Label(
            self,
            text='Название объекта',
            width=28,
            anchor="e"
        )
        self.object_titul_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.set_label = tk.Label(
            self,
            text='Устройство сети',
            width=28,
            anchor="e"
        )
        self.set_combobox = Combobox(
            self,
            values=SET,
            width=67,
            validate="key"
        )
        
        self.str_klim_zona_label = tk.Label(
            self,
            text='Строительно-климатическая зона',
            width=28,
            anchor="e"
        )
        self.str_klim_zona_combobox = Combobox(
            self,
            values=STR_KLIM_ZONA,
            width=12,
            validate="key"
        )

        self.vid_klim_label = tk.Label(
            self,
            text='Вид климата',
            width=28,
            anchor="e"
        )
        self.vid_klim_combobox = Combobox(
            self,
            values=VID_KLIM,
            width=12,
            validate="key"
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
        self.golol_rayon_combobox.bind("<<ComboboxSelected>>", self.paste_golol)

        self.golol_thick_label = tk.Label(
            self,
            text='Толщина стенки гололеда, мм',
            width=28,
            anchor="e"
        )
        self.golol_thick_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.seism_label = tk.Label(
            self,
            text='Сейсмичность районов, баллов',
            width=28,
            anchor="e"
        )
        self.seism_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.sbros_label = tk.Label(
            self,
            text='Демонтаж трансформатора: сброс',
            width=28,
            anchor="e"
        )
        self.sbros_combobox = Combobox(
            self,
            values=(
                "в пролете, наиболее удаленном от зоны тр-ра",
                "в пролете, параллельно дороге",
                "в пролете, без проводов"
            ),
            width=67,
        )

        self.grounding_initial_data_label = tk.Label(
            self,
            text='Исх. данные для заземления',
            width=28,
            anchor="e"
        )
        self.grounding_initial_data_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.r1_label = tk.Label(
            self,
            text='Уд. сопротивление ρ1, Ом*м',
            width=28,
            anchor="e"
        )
        self.r1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.r2_label = tk.Label(
            self,
            text='Уд. сопротивление ρ2, Ом*м',
            width=28,
            anchor="e"
        )
        self.r2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.mosch_gr_label = tk.Label(
            self,
            text='Мощность грунта, м',
            width=28,
            anchor="e"
        )
        self.h_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.isol_rast_label = tk.Label(
            self,
            text='Изоляц. расстояние, мм',
            width=28,
            anchor="e"
        )
        self.isol_rast_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.raspr_nagr_label = tk.Label(
            self,
            text='Норм. распр. нагрузка, кг/м.п.',
            width=28,
            anchor="e"
        )
        self.raspr_nagr_entry = tk.Entry(
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
                "Лежневый фундамент",
                "Плита",
                "Свая"
            ),
            width=67,
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
        self.dlina_rigelya1_label = tk.Label(
            self,
            text='Длина рагеля 1, м',
            width=28,
            anchor="e"
        )
        self.dlina_rigelya2_label = tk.Label(
            self,
            text='Длина ригеля 2, м',
            width=28,
            anchor="e"
        )
        self.dlina_stoiki_label = tk.Label(
            self,
            text='Длина стойки, м',
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

        self.dlina_rigelya1_1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya1_2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya1_3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya1_4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya2_1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya2_2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya2_3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_rigelya2_4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_stoiki1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_stoiki2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_stoiki3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.dlina_stoiki4_entry = tk.Entry(
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

        self.ish_schema_label = tk.Label(
            self,
            text='Прил.Г: SCAD-модель.png',
            width=28,
            anchor="e"
        )
        self.ish_schema_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_ish_schema_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_ish_schema
        )

        self.rasch_model_sverhu_label = tk.Label(
            self,
            text='Прил.Г: Модель сверху.png',
            width=28,
            anchor="e"
        )
        self.rasch_model_sverhu_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_rasch_model_sverhu_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_rasch_model_sverhu
        )

        self.rasch_model1_label = tk.Label(
            self,
            text='Прил.Г: Модель. Вид 1.png',
            width=28,
            anchor="e"
        )
        self.rasch_model1_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_rasch_model1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_rasch_model1
        )

        self.rasch_model2_label = tk.Label(
            self,
            text='Прил.Г: Модель. Вид 2.png',
            width=28,
            anchor="e"
        )
        self.rasch_model2_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_rasch_model2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_rasch_model2
        )

        self.coef_isp_label = tk.Label(
            self,
            text='Прил.Д: Коэф. использования.png',
            width=28,
            anchor="e"
        )
        self.coef_isp_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_coef_isp_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_coef_isp
        )

        self.perem_x_label = tk.Label(
            self,
            text='Прил.Е: Перемещения X.png',
            width=28,
            anchor="e"
        )
        self.perem_x_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_perem_x_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_perem_x
        )

        self.perem_y_label = tk.Label(
            self,
            text='Прил.Е: Перемещения Y.png',
            width=28,
            anchor="e"
        )
        self.perem_y_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_perem_y_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_perem_y
        )

        self.perem_z_label = tk.Label(
            self,
            text='Прил.Е: Перемещения Z.png',
            width=28,
            anchor="e"
        )
        self.perem_z_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_perem_z_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_perem_z
        )

        self.prodolnoe_usil_label = tk.Label(
            self,
            text='Прил.Ж: Продольное усилие.png',
            width=28,
            anchor="e"
        )
        self.prodolnoe_usil_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_prodolnoe_usil_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_prodolnoe_usil
        )

        self.m_y_label = tk.Label(
            self,
            text='Прил.Ж: Момент My.png',
            width=28,
            anchor="e"
        )
        self.m_y_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_m_y_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_m_y
        )

        self.m_z_label = tk.Label(
            self,
            text='Прил.Ж: Момент Mz.png',
            width=28,
            anchor="e"
        )
        self.m_z_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_m_z_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_m_z
        )

        self.q_z_label = tk.Label(
            self,
            text='Прил.Ж: Поперечная сила Qz.png',
            width=35,
            anchor="e"
        )
        self.q_z_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_q_z_button = tk.Button(
            self, 
            text="Обзор",
            command=self.browse_for_q_z
        )

        self.q_y_label = tk.Label(
            self,
            text='Прил.Ж: Поперечная сила Qy.png',
            width=35,
            anchor="e"
        )
        self.q_y_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_q_y_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_q_y
        )

        self.nagr_v_rigel_label = tk.Label(
            self,
            text='Прил.З: Удар. нагр. верт. в ригель.png',
            width=35,
            anchor="e"
        )
        self.nagr_v_rigel_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_nagr_v_rigel_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_nagr_v_rigel
        )

        self.sum_peremesch_v_rigel_label = tk.Label(
            self,
            text='Прил.З: Сумм. перемещения (ригель).png',
            width=35,
            anchor="e"
        )
        self.sum_peremesch_v_rigel_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_sum_peremesch_v_rigel_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_sum_peremesch_v_rigel
        )

        self.nagr_v_uzel_label = tk.Label(
            self,
            text='Прил.З: Удар. нагр. верт. в узел.png',
            width=35,
            anchor="e"
        )
        self.nagr_v_uzel_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_nagr_v_uzel_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_nagr_v_uzel
        )

        self.sum_peremesch_uzel_label = tk.Label(
            self,
            text='Прил.З: Сумм. перемещения (узел).png',
            width=35,
            anchor="e"
        )
        self.sum_peremesch_uzel_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_sum_peremesch_uzel_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_sum_peremesch_uzel
        )

        self.nagr_g_rigel_label = tk.Label(
            self,
            text='Прил.З: Удар. нагр. гор. в ригель.png',
            width=35,
            anchor="e"
        )
        self.nagr_g_rigel_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_nagr_g_rigel_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_nagr_g_rigel
        )

        self.sum_peremesch_g_rigel_label = tk.Label(
            self,
            text='Прил.З: Сумм. перемещения (ригель).png',
            width=35,
            anchor="e"
        )
        self.sum_peremesch_g_rigel_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_sum_peremesch_g_rigel_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_sum_peremesch_g_rigel
        )

        self.usil_osn_label = tk.Label(
            self,
            text='Прил.И: Усилия в основании.png',
            width=35,
            anchor="e"
        )
        self.usil_osn_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usil_osn_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usil_osn
        )

        self.usil_n_label = tk.Label(
            self,
            text='Прил.И: Усилия N.png',
            width=35,
            anchor="e"
        )
        self.usil_n_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usil_n_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usil_n
        )

        self.usil_m_label = tk.Label(
            self,
            text='Прил.И: Усилия M.png',
            width=35,
            anchor="e"
        )
        self.usil_m_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usil_m_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usil_m
        )

        self.usil_q_label = tk.Label(
            self,
            text='Прил.И: Усилия Q.png',
            width=35,
            anchor="e"
        )
        self.usil_q_entry = tk.Entry(
            self,
            width=45,
            relief="sunken",
            bd=2
        )
        self.browse_for_usil_q_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_usil_q
        )

        self.generate_button = tk.Button(
            self,
            text="Создать документацию",
            command=self.generate
        )

    def run(self):
        self.draw_widgets()
        self.paste_kozup_project_data()
        self.mainloop()

    def paste_kozup_project_data(self):
        try:
            self.sp_wind_reg_combobox.set(self.parent.sp_wind_reg)
            self.wind_nagr_entry.insert(0, self.parent.wind_nagr)
            self.golol_rayon_combobox.set(self.parent.golol_rayon)
            self.zasch_obj_entry.insert(0, self.parent.zasch_obj)
            self.territoria_raspoloj_entry.insert(0, self.parent.territoria_raspoloj)
            self.rayon_str_entry.insert(0, self.parent.rayon_str)
            self.expl_god_entry.insert(0, self.parent.expl_god)
            self.v_m_bpla_entry.insert(0, self.parent.v_m_bpla)
            self.steel_entry.insert(0, self.parent.steel)
            self.anal_steel_entry.insert(0, self.parent.anal_steel)
            self.konst_entry.insert(0, self.parent.konst)
            self.kp_bolt_entry.insert(0, self.parent.kp_bolt)
            self.fundament_combobox.set(self.parent.fundament)
            self.fund_osn_combobox.set(self.parent.fund_osn)
            self.kol_obj_entry.insert(0, self.parent.kol_obj)
            self.massa_obsch_entry.insert(0, self.parent.massa_obsch)
            self.mont_vremya_entry.insert(0, self.parent.mont_vremya)
            self.shag_yach_entry.insert(0, self.parent.shag_yach)
            self.post_nagr_entry.insert(0, self.parent.post_nagr)
            self.strela_entry.insert(0, self.parent.strela)
            self.natyajenie_entry.insert(0, self.parent.natyajenie)
            self.klass_betona_entry.insert(0, self.parent.klass_betona)
            self.morozostoikost_entry.insert(0, self.parent.morozostoikost)
            self.vodonepronicaemost_entry.insert(0, self.parent.vodonepronicaemost)
            self.quantity_of_obj_combobox.set(self.parent.quantity_of_obj)
            self.activate_kozu_p(None)
            self.dlina_elem_entry.insert(0, self.parent.dlina_elem)
            self.zaschichaemyi_obj1_entry.insert(0, self.parent.zaschichaemyi_obj1)
            self.zaschichaemyi_obj2_entry.insert(0, self.parent.zaschichaemyi_obj2)
            self.zaschichaemyi_obj3_entry.insert(0, self.parent.zaschichaemyi_obj3)
            self.zaschichaemyi_obj4_entry.insert(0, self.parent.zaschichaemyi_obj4)
            self.length_kozup1_entry.insert(0, self.parent.length_kozup1)
            self.length_kozup2_entry.insert(0, self.parent.length_kozup2)
            self.length_kozup3_entry.insert(0, self.parent.length_kozup3)
            self.length_kozup4_entry.insert(0, self.parent.length_kozup4)
            self.width_kozup1_entry.insert(0, self.parent.width_kozup1)
            self.width_kozup2_entry.insert(0, self.parent.width_kozup2)
            self.width_kozup3_entry.insert(0, self.parent.width_kozup3)
            self.width_kozup4_entry.insert(0, self.parent.width_kozup4)
            self.h1_entry.insert(0, self.parent.h1)
            self.h2_entry.insert(0, self.parent.h2)
            self.h3_entry.insert(0, self.parent.h3)
            self.h4_entry.insert(0, self.parent.h4)
            self.massa_kozup1_entry.insert(0, self.parent.massa_kozup1)
            self.massa_kozup2_entry.insert(0, self.parent.massa_kozup2)
            self.massa_kozup3_entry.insert(0, self.parent.massa_kozup3)
            self.massa_kozup4_entry.insert(0, self.parent.massa_kozup4)
            self.speca_entry.insert(0, self.parent.speca)
            vid_list = [
                self.parent.vid_kozu_p1,
                self.parent.vid_kozu_p2,
                self.parent.vid_kozu_p3,
                self.parent.vid_kozu_p4,
            ]
            final_vid_list = ["{" + vid + "}" for vid in vid_list if vid]
            self.vid_kozu_p_entry.insert(0, " ".join(final_vid_list))
            self.table3_entry.insert(0, self.parent.table3)
            self.table4_entry.insert(0, self.parent.table4)
            self.table5_entry.insert(0, self.parent.table5)
            self.table6_entry.insert(0, self.parent.table6)
            self.raschet_model_entry.insert(0, self.parent.raschet_model)
            self.usiliya1_entry.insert(0, self.parent.usiliya1)
            self.usiliya2_entry.insert(0, self.parent.usiliya2)
            self.usiliya3_entry.insert(0, self.parent.usiliya3)
            self.usiliya4_entry.insert(0, self.parent.usiliya4)
            self.usiliya5_entry.insert(0, self.parent.usiliya5)
            
        except Exception as e:
            # print("!INFO!: Сохраненные данные отсутствуют.")
            print(e)

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=2)
        self.is_svaya_checkbutton.place(x=30, y=42)
        self.is_tros_ferma_checkbutton.place(x=145, y=42)
        self.is_ferma_usil_checkbutton.place(x=335, y=42)
        self.is_existing_ground_checkbutton.place(x=515, y=42)
        self.klimat_label.place(x=15, y=157)
        self.sp_wind_reg_label.place(x=15, y=180)
        self.sp_wind_reg_combobox.place(x=220, y=180)
        self.wind_nagr_label.place(x=15, y=203)
        self.wind_nagr_entry.place(x=220, y=203)
        self.golol_rayon_label.place(x=15, y=226)
        self.golol_rayon_combobox.place(x=220, y=226)
        self.golol_thick_label.place(x=15, y=249)
        self.golol_thick_entry.place(x=220, y=249)
        self.str_klim_zona_label.place(x=15, y=272)
        self.str_klim_zona_combobox.place(x=220, y=272)
        self.vid_klim_label.place(x=15, y=295)
        self.vid_klim_combobox.place(x=220, y=295)
        self.seism_label.place(x=15, y=318)
        self.seism_entry.place(x=220, y=318)
        self.project_info_label.place(x=35, y=65)
        self.bartal_code_label.place(x=15, y=88)
        self.bartal_code_entry.place(x=220, y=88)
        self.object_titul_label.place(x=15, y=111)
        self.object_titul_entry.place(x=220, y=111)
        self.rayon_str_label.place(x=15, y=134)
        self.rayon_str_entry.place(x=220, y=134)
        self.konstr_resh_label.place(x=350, y=65)
        self.set_label.place(x=315, y=88)
        self.set_combobox.place(x=520, y=88)
        self.fundament_label.place(x=315, y=111)
        self.fundament_combobox.place(x=520, y=111)
        self.sbros_label.place(x=315, y=134)
        self.sbros_combobox.place(x=520, y=134)
        self.zazemlenie_label.place(x=315, y=157)
        self.grounding_initial_data_label.place(x=315, y=180)
        self.grounding_initial_data_entry.place(x=520, y=180)
        self.r1_label.place(x=315, y=203)
        self.r1_entry.place(x=520, y=203)
        self.r2_label.place(x=315, y=226)
        self.r2_entry.place(x=520, y=226)
        self.mosch_gr_label.place(x=315, y=249)
        self.h_entry.place(x=520, y=249)
        self.dop_info_label.place(x=380, y=272)
        self.isol_rast_label.place(x=315, y=295)
        self.isol_rast_entry.place(x=520, y=295)
        self.raspr_nagr_label.place(x=315, y=318)
        self.raspr_nagr_entry.place(x=520, y=318)
        self.quantity_of_obj_label.place(x=15, y=341)
        self.quantity_of_obj_combobox.place(x=220, y=341)
        self.kozup_1_label.place(x=240, y=364)
        self.kozup_2_label.place(x=334, y=364)
        self.kozup_3_label.place(x=428, y=364)
        self.kozup_4_label.place(x=522, y=364)
        self.zaschichaemyi_obj_label.place(x=15, y=387)
        self.length_kozup_label.place(x=15, y=410)
        self.width_kozup_label.place(x=15, y=433)
        self.h_label.place(x=15, y=456)
        self.massa_kozup_label.place(x=15, y=479)
        self.dlina_rigelya1_label.place(x=15, y=502)
        self.dlina_rigelya2_label.place(x=15, y=525)
        self.dlina_stoiki_label.place(x=15, y=548)
        self.zaschichaemyi_obj1_entry.place(x=220, y=387)
        self.zaschichaemyi_obj2_entry.place(x=314, y=387)
        self.zaschichaemyi_obj3_entry.place(x=408, y=387)
        self.zaschichaemyi_obj4_entry.place(x=502, y=387)
        self.length_kozup1_entry.place(x=220, y=410)
        self.length_kozup2_entry.place(x=314, y=410)
        self.length_kozup3_entry.place(x=408, y=410)
        self.length_kozup4_entry.place(x=502, y=410) 
        self.width_kozup1_entry.place(x=220, y=433)
        self.width_kozup2_entry.place(x=314, y=433)
        self.width_kozup3_entry.place(x=408, y=433)
        self.width_kozup4_entry.place(x=502, y=433)
        self.h1_entry.place(x=220, y=456)
        self.h2_entry.place(x=314, y=456)
        self.h3_entry.place(x=408, y=456)
        self.h4_entry.place(x=502, y=456)
        self.massa_kozup1_entry.place(x=220, y=479)
        self.massa_kozup2_entry.place(x=314, y=479)
        self.massa_kozup3_entry.place(x=408, y=479)
        self.massa_kozup4_entry.place(x=502, y=479)
        self.dlina_rigelya1_1_entry.place(x=220, y=502)
        self.dlina_rigelya1_2_entry.place(x=314, y=502)
        self.dlina_rigelya1_3_entry.place(x=408, y=502)
        self.dlina_rigelya1_4_entry.place(x=502, y=502)
        self.dlina_rigelya2_1_entry.place(x=220, y=525)
        self.dlina_rigelya2_2_entry.place(x=314, y=525)
        self.dlina_rigelya2_3_entry.place(x=408, y=525)
        self.dlina_rigelya2_4_entry.place(x=502, y=525)
        self.dlina_stoiki1_entry.place(x=220, y=548)
        self.dlina_stoiki2_entry.place(x=314, y=548)
        self.dlina_stoiki3_entry.place(x=408, y=548)
        self.dlina_stoiki4_entry.place(x=502, y=548)
        self.vid_kozu_p_label.place(x=15, y=576)
        self.vid_kozu_p_entry.place(x=220, y=576)
        self.browse_for_vid_kozu_p_button.place(x=497, y=574)
        self.ish_schema_label.place(x=15, y=604)
        self.ish_schema_entry.place(x=220, y=604)
        self.browse_for_ish_schema_button.place(x=497, y=602)
        self.rasch_model_sverhu_label.place(x=15, y=632)
        self.rasch_model_sverhu_entry.place(x=220, y=632)
        self.browse_for_rasch_model_sverhu_button.place(x=497, y=630)
        self.rasch_model1_label.place(x=15, y=660)
        self.rasch_model1_entry.place(x=220, y=660)
        self.browse_for_rasch_model1_button.place(x=497, y=658)
        self.rasch_model2_label.place(x=15, y=688)
        self.rasch_model2_entry.place(x=220, y=688)
        self.browse_for_rasch_model2_button.place(x=497, y=686)
        self.coef_isp_label.place(x=15, y=716)
        self.coef_isp_entry.place(x=220, y=716)
        self.browse_for_coef_isp_button.place(x=497, y=714)
        self.perem_x_label.place(x=15, y=744)
        self.perem_x_entry.place(x=220, y=744)
        self.browse_for_perem_x_button.place(x=497, y=742)
        self.perem_y_label.place(x=15, y=772)
        self.perem_y_entry.place(x=220, y=772)
        self.browse_for_perem_y_button.place(x=497, y=770)
        self.perem_z_label.place(x=15, y=800)
        self.perem_z_entry.place(x=220, y=800)
        self.browse_for_perem_z_button.place(x=497, y=798)
        self.prodolnoe_usil_label.place(x=15, y=828)
        self.prodolnoe_usil_entry.place(x=220, y=828)
        self.browse_for_prodolnoe_usil_button.place(x= 497, y=826)
        self.m_y_label.place(x=15, y=856)
        self.m_y_entry.place(x=220, y=856)
        self.browse_for_m_y_button.place(x=497, y=854)
        self.m_z_label.place(x=15, y=884)
        self.m_z_entry.place(x=220, y=884)
        self.browse_for_m_z_button.place(x=497, y=882)
        self.q_z_label.place(x=545, y=576)
        self.q_z_entry.place(x=800, y=576)
        self.browse_for_q_z_button.place(x=1077, y=574)
        self.q_y_label.place(x=545, y=604)
        self.q_y_entry.place(x=800, y=604)
        self.browse_for_q_y_button.place(x=1077, y=602)
        self.nagr_v_rigel_label.place(x=545, y=632)
        self.nagr_v_rigel_entry.place(x=800, y=632)
        self.browse_for_nagr_v_rigel_button.place(x=1077, y=630)
        self.sum_peremesch_v_rigel_label.place(x=545, y=660)
        self.sum_peremesch_v_rigel_entry.place(x=800, y=660)
        self.browse_for_sum_peremesch_v_rigel_button.place(x=1077, y=658)
        self.nagr_v_uzel_label.place(x=545, y=688)
        self.nagr_v_uzel_entry.place(x=800, y=688)
        self.browse_for_nagr_v_uzel_button.place(x=1077, y=686)
        self.sum_peremesch_uzel_label.place(x=545, y=716)
        self.sum_peremesch_uzel_entry.place(x=800, y=716)
        self.browse_for_sum_peremesch_uzel_button.place(x=1077, y=714)
        self.nagr_g_rigel_label.place(x=545, y=744)
        self.nagr_g_rigel_entry.place(x=800, y=744)
        self.browse_for_nagr_g_rigel_button.place(x=1077, y=742)
        self.sum_peremesch_g_rigel_label.place(x=545, y=772)
        self.sum_peremesch_g_rigel_entry.place(x=800, y=772)
        self.browse_for_sum_peremesch_g_rigel_button.place(x=1077, y=770)
        self.usil_osn_label.place(x=545, y=800)
        self.usil_osn_entry.place(x=800, y=800)
        self.browse_for_usil_osn_button.place(x=1077, y=798)
        self.usil_n_label.place(x=545, y=828)
        self.usil_n_entry.place(x=800, y=828)
        self.browse_for_usil_n_button.place(x=1077, y=826)
        self.usil_m_label.place(x=545, y=856)
        self.usil_m_entry.place(x=800, y=856)
        self.browse_for_usil_m_button.place(x=1077, y=854)
        self.usil_q_label.place(x=545, y=884)
        self.usil_q_entry.place(x=800, y=884)
        self.browse_for_usil_q_button.place(x=1077, y=882)
        self.generate_button.place(x=500, y=924)

    def generate(self):
        make_tkr(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            voltage=self.parent.voltage,
            developer=self.parent.developer,
            is_svaya=self.is_svaya_var.get(),
            is_tros=self.is_tros_ferma_var.get(),
            is_ferma=self.is_ferma_usil_var.get(),
            is_existing_ground=self.is_existing_ground_var.get(),
            sp_wind_region=self.sp_wind_reg_combobox.get(),
            wind_nagr=self.wind_nagr_entry.get(),
            golol_rayon=self.golol_rayon_combobox.get(),
            golol_thick=self.golol_thick_entry.get(),
            str_klim=self.str_klim_zona_combobox.get(),
            vid_klim=self.vid_klim_combobox.get(),
            seism=self.seism_entry.get(),
            bartal_code=self.bartal_code_entry.get(),
            object_titul=self.object_titul_entry.get(),
            seti=self.set_combobox.get(),
            fundament=self.fundament_combobox.get(),
            rayon_str=self.rayon_str_entry.get(),
            sbros=self.sbros_combobox.get(),
            grounding_initial_data=self.grounding_initial_data_entry.get(),
            r1=self.r1_entry.get(),
            r2=self.r2_entry.get(),
            h=self.h_entry.get(),
            isol_rast=self.isol_rast_entry.get(),
            raspr_nagr=self.raspr_nagr_entry.get(),
            quantity_of_obj=self.quantity_of_obj_combobox.get(),
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
            dlina_rigelya1_1=self.dlina_rigelya1_1_entry.get(),
            dlina_rigelya1_2=self.dlina_rigelya1_2_entry.get(),
            dlina_rigelya1_3=self.dlina_rigelya1_3_entry.get(),
            dlina_rigelya1_4=self.dlina_rigelya1_4_entry.get(),
            dlina_rigelya2_1=self.dlina_rigelya2_1_entry.get(),
            dlina_rigelya2_2=self.dlina_rigelya2_2_entry.get(),
            dlina_rigelya2_3=self.dlina_rigelya2_3_entry.get(),
            dlina_rigelya2_4=self.dlina_rigelya2_4_entry.get(),
            dlina_stoiki1=self.dlina_stoiki1_entry.get(),
            dlina_stoiki2=self.dlina_stoiki2_entry.get(),
            dlina_stoiki3=self.dlina_stoiki3_entry.get(),
            dlina_stoiki4=self.dlina_stoiki4_entry.get(),
            vid_kozu_p=self.vid_kozu_p_entry.get(),
            ish_schema=self.ish_schema_entry.get(),
            rasch_model_sverhu=self.rasch_model_sverhu_entry.get(),
            rasch_model1=self.rasch_model1_entry.get(),
            rasch_model2=self.rasch_model2_entry.get(),
            coef_isp=self.coef_isp_entry.get(),
            perem_x=self.perem_x_entry.get(),
            perem_y=self.perem_y_entry.get(),
            perem_z=self.perem_z_entry.get(),
            prodolnoe_usil=self.prodolnoe_usil_entry.get(),
            m_y=self.m_y_entry.get(),
            m_z=self.m_z_entry.get(),
            q_z=self.q_z_entry.get(),
            q_y=self.q_y_entry.get(),
            nagr_v_rigel=self.nagr_v_rigel_entry.get(),
            sum_peremesch_v_rigel=self.sum_peremesch_v_rigel_entry.get(),
            nagr_v_uzel=self.nagr_v_uzel_entry.get(),
            sum_peremesch_uzel=self.sum_peremesch_uzel_entry.get(),
            nagr_g_rigel=self.nagr_g_rigel_entry.get(),
            sum_peremesch_g_rigel=self.sum_peremesch_g_rigel_entry.get(),
            usil_osn=self.usil_osn_entry.get(),
            usil_n=self.usil_n_entry.get(),
            usil_m=self.usil_m_entry.get(),
            usil_q=self.usil_q_entry.get(),
        )
        # speca=self.speca_entry.get().split("КОЗУ (инженерная)")[1]
        # vid_data = self.vid_kozu_p_entry.get()
        # if vid_data:
        #     vids = [pic_dir.strip("}{") for pic_dir in vid_data.split("} {")]
        #     for i in range(4):
        #         if len(vids) <4:
        #             vids.append("")
        #         if vids[i]:
        #             vids[i] = vids[i].split("КОЗУ (инженерная)")[1]
        # table3=self.table3_entry.get().split("КОЗУ (инженерная)")[1]
        # table4=self.table4_entry.get().split("КОЗУ (инженерная)")[1]
        # table5=self.table5_entry.get().split("КОЗУ (инженерная)")[1]
        # table6=self.table6_entry.get().split("КОЗУ (инженерная)")[1]
        # raschet_model=self.raschet_model_entry.get().split("КОЗУ (инженерная)")[1]
        # usiliya1=self.usiliya1_entry.get().split("КОЗУ (инженерная)")[1]
        # usiliya2=self.usiliya2_entry.get().split("КОЗУ (инженерная)")[1]
        # usiliya3=self.usiliya3_entry.get().split("КОЗУ (инженерная)")[1]
        # usiliya4=self.usiliya4_entry.get().split("КОЗУ (инженерная)")[1]
        # usiliya5=self.usiliya5_entry.get().split("КОЗУ (инженерная)")[1]
        
        # self.db.add_kozup_project_data(
        #     initial_data_id=self.parent.initial_data_id,
        #     sp_wind_reg=self.sp_wind_reg_combobox.get(),
        #     wind_nagr=self.wind_nagr_entry.get(),
        #     golol_rayon=self.golol_rayon_combobox.get(),
        #     v_m_bpla=self.v_m_bpla_entry.get(),
        #     zasch_obj=self.zasch_obj_entry.get(),
        #     steel=self.steel_entry.get(),
        #     anal_steel=self.anal_steel_entry.get(),
        #     konst=self.konst_entry.get(),
        #     kp_bolt=self.kp_bolt_entry.get(),
        #     fundament=self.fundament_combobox.get(),
        #     fund_osn=self.fund_osn_combobox.get(),
        #     rayon_str=self.rayon_str_entry.get(),
        #     kol_obj=self.kol_obj_entry.get(),
        #     massa_obsch=self.massa_obsch_entry.get(),
        #     territoria_raspoloj=self.territoria_raspoloj_entry.get(),
        #     mont_vremya=self.mont_vremya_entry.get(),
        #     expl_god=self.expl_god_entry.get(),
        #     shag_yach=self.shag_yach_entry.get(),
        #     post_nagr=self.post_nagr_entry.get(),
        #     strela=self.strela_entry.get(),
        #     natyajenie=self.natyajenie_entry.get(),
        #     klass_betona=self.klass_betona_entry.get(),
        #     morozostoikost=self.morozostoikost_entry.get(),
        #     vodonepronicaemost=self.vodonepronicaemost_entry.get(),
        #     quantity_of_obj=self.quantity_of_obj_combobox.get(),
        #     dlina_elem=self.dlina_elem_entry.get(),
        #     zaschichaemyi_obj1=self.zaschichaemyi_obj1_entry.get(),
        #     zaschichaemyi_obj2=self.zaschichaemyi_obj2_entry.get(),
        #     zaschichaemyi_obj3=self.zaschichaemyi_obj3_entry.get(),
        #     zaschichaemyi_obj4=self.zaschichaemyi_obj4_entry.get(),
        #     length_kozup1=self.length_kozup1_entry.get(),
        #     length_kozup2=self.length_kozup2_entry.get(),
        #     length_kozup3=self.length_kozup3_entry.get(),
        #     length_kozup4=self.length_kozup4_entry.get(),
        #     width_kozup1=self.width_kozup1_entry.get(),
        #     width_kozup2=self.width_kozup2_entry.get(),
        #     width_kozup3=self.width_kozup3_entry.get(),
        #     width_kozup4=self.width_kozup4_entry.get(),
        #     h1=self.h1_entry.get(),
        #     h2=self.h2_entry.get(),
        #     h3=self.h3_entry.get(),
        #     h4=self.h4_entry.get(),
        #     massa_kozup1=self.massa_kozup1_entry.get(),
        #     massa_kozup2=self.massa_kozup2_entry.get(),
        #     massa_kozup3=self.massa_kozup3_entry.get(),
        #     massa_kozup4=self.massa_kozup4_entry.get(),
        #     speca=speca,
        #     vid_kozu_p1=vids[0],
        #     vid_kozu_p2=vids[1],
        #     vid_kozu_p3=vids[2],
        #     vid_kozu_p4=vids[3],
        #     table3=table3,
        #     table4=table4,
        #     table5=table5,
        #     table6=table6,
        #     raschet_model=raschet_model,
        #     usiliya1=usiliya1,
        #     usiliya2=usiliya2,
        #     usiliya3=usiliya3,
        #     usiliya4=usiliya4,
        #     usiliya5=usiliya5,
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

    def paste_golol(self, event):
        golol_key = self.golol_rayon_combobox.get()
        self.golol_thick_entry.delete(0, tk.END)
        self.golol_thick_entry.insert(0, sp_wind_reg_dict[golol_key])

    def browse_for_ish_schema(self):
        self.file_path = make_path_png()
        self.ish_schema_entry.delete("0", "end") 
        self.ish_schema_entry.insert("insert", self.file_path)

    def browse_for_rasch_model_sverhu(self):
        self.file_path = make_path_png()
        self.rasch_model_sverhu_entry.delete("0", "end") 
        self.rasch_model_sverhu_entry.insert("insert", self.file_path)

    def browse_for_rasch_model1(self):
        self.file_path = make_path_png()
        self.rasch_model1_entry.delete("0", "end") 
        self.rasch_model1_entry.insert("insert", self.file_path)

    def browse_for_rasch_model2(self):
        self.file_path = make_path_png()
        self.rasch_model2_entry.delete("0", "end") 
        self.rasch_model2_entry.insert("insert", self.file_path)

    def browse_for_coef_isp(self):
        self.file_path = make_path_png()
        self.coef_isp_entry.delete("0", "end") 
        self.coef_isp_entry.insert("insert", self.file_path)

    def browse_for_perem_x(self):
        self.file_path = make_path_png()
        self.perem_x_entry.delete("0", "end") 
        self.perem_x_entry.insert("insert", self.file_path)

    def browse_for_perem_y(self):
        self.file_path = make_path_png()
        self.perem_y_entry.delete("0", "end") 
        self.perem_y_entry.insert("insert", self.file_path)

    def browse_for_perem_z(self):
        self.file_path = make_path_png()
        self.perem_z_entry.delete("0", "end") 
        self.perem_z_entry.insert("insert", self.file_path)

    def browse_for_prodolnoe_usil(self):
        self.file_path = make_path_png()
        self.prodolnoe_usil_entry.delete("0", "end") 
        self.prodolnoe_usil_entry.insert("insert", self.file_path)

    def browse_for_m_y(self):
        self.file_path = make_path_png()
        self.m_y_entry.delete("0", "end") 
        self.m_y_entry.insert("insert", self.file_path)

    def browse_for_m_z(self):
        self.file_path = make_path_png()
        self.m_z_entry.delete("0", "end") 
        self.m_z_entry.insert("insert", self.file_path)

    def browse_for_q_z(self):
        self.file_path = make_path_png()
        self.q_z_entry.delete("0", "end") 
        self.q_z_entry.insert("insert", self.file_path)

    def browse_for_q_y(self):
        self.file_path = make_path_png()
        self.q_y_entry.delete("0", "end") 
        self.q_y_entry.insert("insert", self.file_path)

    def browse_for_nagr_v_rigel(self):
        self.file_path = make_path_png()
        self.nagr_v_rigel_entry.delete("0", "end") 
        self.nagr_v_rigel_entry.insert("insert", self.file_path)
    
    def browse_for_sum_peremesch_v_rigel(self):
        self.file_path = make_path_png()
        self.sum_peremesch_v_rigel_entry.delete("0", "end") 
        self.sum_peremesch_v_rigel_entry.insert("insert", self.file_path)

    def browse_for_nagr_v_uzel(self):
        self.file_path = make_path_png()
        self.nagr_v_uzel_entry.delete("0", "end") 
        self.nagr_v_uzel_entry.insert("insert", self.file_path)

    def browse_for_sum_peremesch_uzel(self):
        self.file_path = make_path_png()
        self.sum_peremesch_uzel_entry.delete("0", "end") 
        self.sum_peremesch_uzel_entry.insert("insert", self.file_path)

    def browse_for_nagr_g_rigel(self):
        self.file_path = make_path_png()
        self.nagr_g_rigel_entry.delete("0", "end") 
        self.nagr_g_rigel_entry.insert("insert", self.file_path)

    def browse_for_sum_peremesch_g_rigel(self):
        self.file_path = make_path_png()
        self.sum_peremesch_g_rigel_entry.delete("0", "end") 
        self.sum_peremesch_g_rigel_entry.insert("insert", self.file_path)

    def browse_for_usil_osn(self):
        self.file_path = make_path_png()
        self.usil_osn_entry.delete("0", "end") 
        self.usil_osn_entry.insert("insert", self.file_path)

    def browse_for_usil_n(self):
        self.file_path = make_path_png()
        self.usil_n_entry.delete("0", "end") 
        self.usil_n_entry.insert("insert", self.file_path)

    def browse_for_usil_m(self):
        self.file_path = make_path_png()
        self.usil_m_entry.delete("0", "end") 
        self.usil_m_entry.insert("insert", self.file_path)

    def browse_for_usil_q(self):
        self.file_path = make_path_png()
        self.usil_q_entry.delete("0", "end") 
        self.usil_q_entry.insert("insert", self.file_path)

    def browse_for_vid_kozup(self):
        self.file_path = make_multiple_path()
        self.vid_kozu_p_entry.delete("0", "end")
        self.vid_kozu_p_entry.insert("insert", self.file_path)

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
            self.dlina_rigelya1_2_entry.delete(0, tk.END)
            self.dlina_rigelya1_3_entry.delete(0, tk.END)
            self.dlina_rigelya1_4_entry.delete(0, tk.END)
            self.dlina_rigelya1_2_entry.config(state="disabled")
            self.dlina_rigelya1_3_entry.config(state="disabled")
            self.dlina_rigelya1_4_entry.config(state="disabled")
            self.dlina_rigelya1_1_entry.config(state="normal")
            self.dlina_rigelya2_2_entry.delete(0, tk.END)
            self.dlina_rigelya2_3_entry.delete(0, tk.END)
            self.dlina_rigelya2_4_entry.delete(0, tk.END)
            self.dlina_rigelya2_2_entry.config(state="disabled")
            self.dlina_rigelya2_3_entry.config(state="disabled")
            self.dlina_rigelya2_4_entry.config(state="disabled")
            self.dlina_rigelya2_1_entry.config(state="normal")
            self.dlina_stoiki2_entry.delete(0, tk.END)
            self.dlina_stoiki3_entry.delete(0, tk.END)
            self.dlina_stoiki4_entry.delete(0, tk.END)
            self.dlina_stoiki2_entry.config(state="disabled")
            self.dlina_stoiki3_entry.config(state="disabled")
            self.dlina_stoiki4_entry.config(state="disabled")
            self.dlina_stoiki1_entry.config(state="normal")
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
            self.dlina_rigelya1_3_entry.delete(0, tk.END)
            self.dlina_rigelya1_4_entry.delete(0, tk.END)
            self.dlina_rigelya1_3_entry.config(state="disabled")
            self.dlina_rigelya1_4_entry.config(state="disabled")
            self.dlina_rigelya1_1_entry.config(state="normal")
            self.dlina_rigelya1_2_entry.config(state="normal")
            self.dlina_rigelya2_3_entry.delete(0, tk.END)
            self.dlina_rigelya2_4_entry.delete(0, tk.END)
            self.dlina_rigelya2_3_entry.config(state="disabled")
            self.dlina_rigelya2_4_entry.config(state="disabled")
            self.dlina_rigelya2_1_entry.config(state="normal")
            self.dlina_rigelya2_2_entry.config(state="normal")
            self.dlina_stoiki3_entry.delete(0, tk.END)
            self.dlina_stoiki4_entry.delete(0, tk.END)
            self.dlina_stoiki3_entry.config(state="disabled")
            self.dlina_stoiki4_entry.config(state="disabled")
            self.dlina_stoiki1_entry.config(state="normal")
            self.dlina_stoiki2_entry.config(state="normal")
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
            self.dlina_rigelya1_4_entry.delete(0, tk.END)
            self.dlina_rigelya1_4_entry.config(state="disabled")
            self.dlina_rigelya1_1_entry.config(state="normal")
            self.dlina_rigelya1_2_entry.config(state="normal")
            self.dlina_rigelya1_3_entry.config(state="normal")
            self.dlina_rigelya2_4_entry.delete(0, tk.END)
            self.dlina_rigelya2_4_entry.config(state="disabled")
            self.dlina_rigelya2_1_entry.config(state="normal")
            self.dlina_rigelya2_2_entry.config(state="normal")
            self.dlina_rigelya2_3_entry.config(state="normal")
            self.dlina_stoiki4_entry.delete(0, tk.END)
            self.dlina_stoiki4_entry.config(state="disabled")
            self.dlina_stoiki1_entry.config(state="normal")
            self.dlina_stoiki2_entry.config(state="normal")
            self.dlina_stoiki3_entry.config(state="normal")
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
            self.dlina_rigelya1_1_entry.config(state="normal")
            self.dlina_rigelya1_2_entry.config(state="normal")
            self.dlina_rigelya1_3_entry.config(state="normal")
            self.dlina_rigelya1_4_entry.config(state="normal")
            self.dlina_rigelya2_1_entry.config(state="normal")
            self.dlina_rigelya2_2_entry.config(state="normal")
            self.dlina_rigelya2_3_entry.config(state="normal")
            self.dlina_rigelya2_4_entry.config(state="normal")
            self.dlina_stoiki1_entry.config(state="normal")
            self.dlina_stoiki2_entry.config(state="normal")
            self.dlina_stoiki3_entry.config(state="normal")
            self.dlina_stoiki4_entry.config(state="normal")
    
    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()