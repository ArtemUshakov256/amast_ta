import tkinter as tk


from PIL import ImageTk
from core.exceptions import CheckCalculationData
from tkinter.ttk import Combobox
from tkinter import messagebox as mb


from core.constants import bolt_dict
from core.utils import (
    tempFile_back,
    tempFile_open,
    tempFile_save,
    make_path_xlsx,
    make_path_png
)
from core.exceptions import AddPlsPolePathException
from core.functionality.ankernie_zakladnie.utils import (
    calculate_bolt,
    save_xlsx,
    make_rpzaz,
    make_pasport
)


class AnkernieZakladnie(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Расчет анкерных закладных")
        self.geometry("650x325+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")

        self.back_icon = ImageTk.PhotoImage(
            file=tempFile_back
        )

        self.module_bg_1 = tk.Frame(
            self,
            width=310,
            height=315,
            borderwidth=2,
            relief="sunken"
        )

        self.module_bg_2 = tk.Frame(
            self,
            width=310,
            height=315,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            image=self.back_icon,
            command=self.back_to_main_window
        )
        
        self.article_label = tk.Label(
            self,
            text='Расчетные данные:',
            width=28,
            anchor="e",
            font=("standard", 10, "bold")
        )
        
        self.pole_diam_label = tk.Label(
            self,
            text='Диаметр опоры, мм',
            width=28,
            anchor="e"
        )
        self.pole_diam_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.bend_moment_label = tk.Label(
            self,
            text='Результирующий момент, кНм',
            width=28,
            anchor="e"
        )
        self.bend_moment_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.vert_force_label = tk.Label(
            self,
            text='Вертикальная сила, кН',
            width=28,
            anchor="e"
        )
        self.vert_force_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.shear_force_label = tk.Label(
            self,
            text='Горизонтальная сила, кН',
            width=28,
            anchor="e"
        )
        self.shear_force_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.bolt_label = tk.Label(
            self,
            text='Марка болта',
            width=28,
            anchor="e"
        )
        self.bolt_variable = tk.StringVar()
        self.bolt_combobox = Combobox(
            self,
            width=12,
            values=(
                "М20",
                "М24",
                "М30",
                "М36",
                "М42",
                "М48",
                "М52",
                "М56",
            ),
            textvariable=self.bolt_variable,
            validate="key"
        )
        self.bolt_combobox.bind("<<ComboboxSelected>>", self.paste_bolt_data)
        self.bolt_combobox["validatecommand"] = (
            self.bolt_combobox.register(self.validate_bolt),
            "%P"
        )

        self.kol_boltov_label = tk.Label(
            self,
            text='Количество болтов',
            width=28,
            anchor="e"
        )
        self.kol_boltov_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.bolt_class_label = tk.Label(
            self,
            text='Класс прочности болтов',
            width=28,
            anchor="e"
        )
        self.bolt_class_combobox = Combobox(
            self,
            width=12,
            values=("5.6", "8.8", "10.9"),
            validate="key"
        )
        self.bolt_class_combobox["validatecommand"] = (
            self.bolt_class_combobox.register(self.validate_bolt_class),
            "%P"
        )

        self.hole_diam_label = tk.Label(
            self,
            text='Диаметр отверстия, мм',
            width=28,
            anchor="e"
        )
        self.hole_diam_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.m_label = tk.Label(
            self,
            text='Расстояние до стенки М, мм',
            width=28,
            anchor="e"
        )
        self.m_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.rez_ras_label = tk.Label(
            self,
            text='Результаты расчета:',
            width=25,
            anchor="e",
            font=("standard", 10, "bold")
        )

        self.xlsx_bolt_label = tk.Label(
            self,
            text="Эксель болтов",
            width=14,
            anchor="e"
        )
        self.xlsx_bolt_entry = tk.Entry(
            self,
            width=23,
            relief="sunken",
            bd=2
        )
        self.browse_for_xlsx_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_xlsx
        )

        self.picture1_label = tk.Label(
            self,
            text="Чертеж",
            width=14,
            anchor="e"
        )
        self.picture1_entry = tk.Entry(
            self,
            width=23,
            relief="sunken",
            bd=2
        )
        self.browse_for_picture1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pic1
        )

        self.calculate_button = tk.Button(
            self,
            text="Расчет",
            command=self.call_calculate_bolt
        )
        self.save_xlsx_button = tk.Button(
            self,
            text="Сохранить",
            command=save_xlsx
        )
        self.rpzaz_button = tk.Button(
            self,
            text="Создать РПЗАБ",
            command=self.call_make_rpzaz
        )
        self.pasport_button = tk.Button(
            self,
            text="Создать паспорт закладной",
            command=self.call_make_pasport
        )

    def run(self):
        self.draw_widgets()
        self.mainloop()

    def draw_widgets(self):
        self.module_bg_1.place(x=10, y=0)
        self.module_bg_2.place(x=330, y=0)
        self.back_to_main_window_button.place(x=15, y=2)
        self.article_label.place(x=15, y=29)
        self.pole_diam_label.place(x=15, y=52)
        self.pole_diam_entry.place(x=220, y=52)
        self.bend_moment_label.place(x=15, y=75)
        self.bend_moment_entry.place(x=220, y=75)
        self.vert_force_label.place(x=15, y=98)
        self.vert_force_entry.place(x=220, y=98)
        self.shear_force_label.place(x=15, y=121)
        self.shear_force_entry.place(x=220, y=121)
        self.bolt_label.place(x=15, y=144)
        self.bolt_combobox.place(x=220, y=144)
        self.kol_boltov_label.place(x=15, y=167)
        self.kol_boltov_entry.place(x=220, y=167)
        self.bolt_class_label.place(x=15, y=190)
        self.bolt_class_combobox.place(x=220, y=190)
        self.hole_diam_label.place(x=15, y=213)
        self.hole_diam_entry.place(x=220, y=213)
        self.m_label.place(x=15, y=236)
        self.m_entry.place(x=220, y=236)
        self.calculate_button.place(x=190, y=269)
        self.save_xlsx_button.place(x=240, y=269)
        self.rez_ras_label.place(x=350, y=5)
        self.rpzaz_button.place(x=353, y=284)
        self.pasport_button.place(x=450, y=284)
        self.xlsx_bolt_label.place(x=335, y=236)
        self.xlsx_bolt_entry.place(x=440, y=236)
        self.browse_for_xlsx_button.place(x=585, y=233)
        self.picture1_label.place(x=335, y=258)
        self.picture1_entry.place(x=440, y=258)
        self.browse_for_picture1_button.place(x=585, y=259)

        try:
            self.pole_diam_entry.delete(0, tk.END)
            self.pole_diam_entry.insert(0, self.parent.pls_pole_data["bot_diam"])
            self.bend_moment_entry.delete(0, tk.END)
            self.bend_moment_entry.insert(0, self.parent.pls_pole_data["moment"])
            self.vert_force_entry.delete(0, tk.END)
            self.vert_force_entry.insert(0, self.parent.pls_pole_data["vert_force"])
            self.shear_force_entry.delete(0, tk.END)
            self.shear_force_entry.insert(0, self.parent.pls_pole_data["shear_force"])
        except Exception as e:
            # mb.showerror("Ошибка","Добавь ссылки к отчетам POLE.")
            # self.destroy()
            print(str(e))

    def open_data(self):
        pass

    def save_data(self):
        pass

    def paste_bolt_data(self, event):
        bolt_key = self.bolt_combobox.get()
        self.hole_diam_entry.delete(0, tk.END)
        self.hole_diam_entry.insert(0, bolt_dict[bolt_key][0])
        self.m_entry.delete(0, tk.END)
        self.m_entry.insert(0, bolt_dict[bolt_key][1])

    def validate_bolt(self, value):
        if value in ["М20", "М24", "М30", "М36", "М42", "М48", "М52", "М56"]:
            return True
        return False
    
    def validate_bolt_class(self, value):
        if value in ["5.6", "8.8", "10.9"]:
            return True
        return False
    
    def browse_for_xlsx(self):
        self.file_path = make_path_xlsx()
        self.xlsx_bolt_entry.delete("0", "end") 
        self.xlsx_bolt_entry.insert("insert", self.file_path)

    def browse_for_pic1(self):
        self.file_path = make_path_png()
        self.picture1_entry.delete("0", "end") 
        self.picture1_entry.insert("insert", self.file_path)
    
    def insert_result(self):
        self.diam_okr_bolt_entry.delete(0, tk.END)
        self.diam_okr_bolt_entry.insert(0, self.result["diam_okr_bolt"])
        self.diam_flanca_entry.delete(0, tk.END)
        self.diam_flanca_entry.insert(0, self.result["diam_flanca"])
        self.bolt_usage_m24_entry.delete(0, tk.END)
        self.bolt_usage_m24_entry.insert(0, self.result["coef_isp_m24"])
        self.bolt_usage_m30_entry.delete(0, tk.END)
        self.bolt_usage_m30_entry.insert(0, self.result["coef_isp_m30"])
        self.bolt_usage_m36_entry.delete(0, tk.END)
        self.bolt_usage_m36_entry.insert(0, self.result["coef_isp_m36"])
        self.bolt_usage_m42_entry.delete(0, tk.END)
        self.bolt_usage_m42_entry.insert(0, self.result["coef_isp_m42"])
        self.bolt_usage_m48_entry.delete(0, tk.END)
        self.bolt_usage_m48_entry.insert(0, self.result["coef_isp_m48"])
        self.bolt_usage_m52_entry.delete(0, tk.END)
        self.bolt_usage_m52_entry.insert(0, self.result["coef_isp_m52"])
        self.bolt_usage_m56_entry.delete(0, tk.END)
        self.bolt_usage_m56_entry.insert(0, self.result["coef_isp_m56"])
    
    def call_calculate_bolt(self):
        self.result = calculate_bolt(
            bend_moment=self.bend_moment_entry.get(),
            vert_force=self.vert_force_entry.get(),
            pole_diam=self.pole_diam_entry.get(),
            bolt=self.bolt_combobox.get(),
            kol_boltov=self.kol_boltov_entry.get(),
            bolt_class=self.bolt_class_combobox.get()
        )
        self.diam_okr_bolt_label = tk.Label(
            self,
            text='Диаметр окружности болтов, мм',
            width=28,
            anchor="e"
        )
        self.diam_okr_bolt_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.diam_flanca_label = tk.Label(
            self,
            text='Диаметр фланца под ключ, мм',
            width=28,
            anchor="e"
        )
        self.diam_flanca_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.bolt_usage_m24_label = tk.Label(
            self,
            text='Загрузка болта М24',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m24_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m24_bg"]
        )

        self.bolt_usage_m30_label = tk.Label(
            self,
            text='Загрузка болта М30',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m30_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m30_bg"]
        )

        self.bolt_usage_m36_label = tk.Label(
            self,
            text='Загрузка болта М36',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m36_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m36_bg"]
        )

        self.bolt_usage_m42_label = tk.Label(
            self,
            text='Загрузка болта М42',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m42_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m42_bg"]
        )

        self.bolt_usage_m48_label = tk.Label(
            self,
            text='Загрузка болта М48',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m48_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m48_bg"]
        )

        self.bolt_usage_m52_label = tk.Label(
            self,
            text='Загрузка болта М52',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m52_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m52_bg"]
        )

        self.bolt_usage_m56_label = tk.Label(
            self,
            text='Загрузка болта М56',
            width=28,
            anchor="e"
        )
        self.bolt_usage_m56_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_m56_bg"]
        )
        self.diam_okr_bolt_label.place(x=335, y=28)
        self.diam_okr_bolt_entry.place(x=540, y=28)
        self.diam_flanca_label.place(x=335, y=51)
        self.diam_flanca_entry.place(x=540, y=51)
        self.bolt_usage_m24_label.place(x=335, y=74)
        self.bolt_usage_m24_entry.place(x=540, y=74)
        self.bolt_usage_m30_label.place(x=335, y=97)
        self.bolt_usage_m30_entry.place(x=540, y=97)
        self.bolt_usage_m36_label.place(x=335, y=120)
        self.bolt_usage_m36_entry.place(x=540, y=120)
        self.bolt_usage_m42_label.place(x=335, y=143)
        self.bolt_usage_m42_entry.place(x=540, y=143)
        self.bolt_usage_m48_label.place(x=335, y=166)
        self.bolt_usage_m48_entry.place(x=540, y=166)
        self.bolt_usage_m52_label.place(x=335, y=189)
        self.bolt_usage_m52_entry.place(x=540, y=189)
        self.bolt_usage_m56_label.place(x=335, y=212)
        self.bolt_usage_m56_entry.place(x=540, y=212)
        self.insert_result()
    
    def call_make_rpzaz(self):
        make_rpzaz(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            developer=self.parent.developer,
            moment=self.bend_moment_entry.get(),
            vert_force=self.vert_force_entry.get(),
            shear_force=self.shear_force_entry.get(),
            bolt_xlsx_path=self.xlsx_bolt_entry.get()
        )

    def call_make_pasport(self):
        make_pasport(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            pole_code=self.parent.pole_code,
            bolt_xlsx_path=self.xlsx_bolt_entry.get(),
            picture1_path=self.picture1_entry.get()
        )

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()