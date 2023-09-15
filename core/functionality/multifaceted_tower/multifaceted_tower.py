import os
import re
import tkinter as tk


from PIL import Image, ImageTk
from tkinter.ttk import Combobox


from core.utils import (
    tempFile_back,
    tempFile_open,
    tempFile_save,
    make_path_txt,
    make_path_png,
    make_multiple_path
)
from core.functionality.multifaceted_tower.utils import *


class MultifacetedTower(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("РПЗО многогранных опор")
        self.geometry(f"730x247+400+10")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")

        self.back_icon = ImageTk.PhotoImage(
            file=tempFile_back
        )
        self.open_icon = ImageTk.PhotoImage(
            file=tempFile_open
        )
        self.save_icon = ImageTk.PhotoImage(
            file=tempFile_save
        )

        self.calculation_clarification_bg1 = tk.Frame(
            self,
            width=350,
            height=100,
            borderwidth=2,
            relief="sunken"
        )

        self.calculation_clarification_bg2 = tk.Frame(
            self,
            width=350,
            height=100,
            borderwidth=2,
            relief="sunken"
        )

        self.media_bg = tk.Frame(
            self,
            width=350,
            height=109,
            borderwidth=2,
            relief="sunken"
        )

        self.generation_bg = tk.Frame(
            self,
            width=350,
            height=109,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            image=self.back_icon,
            command=self.back_to_main_window
        )

        self.open_button = tk.Button(
            self,
            image=self.open_icon,
            command=self.open_data
        )
        
        self.save_button = tk.Button(
            self,
            image=self.save_icon,
            command=self.save_data
        )

        self.is_ground_wire_davit = tk.Label(
            self,
            text="Есть тросовая(ые) траверса(ы)",
            anchor="e", 
            width=30
        )
        self.is_ground_wire_davit_combobox = Combobox(
            self,
            values=("Да", "Нет"),
            width=16,
            validate="key"
        )
        self.is_ground_wire_davit_combobox["validatecommand"] = (
            self.is_ground_wire_davit_combobox.register(self.validate_yes_no),
            "%P"
        )
        self.is_ground_wire_davit_combobox.bind("<KeyRelease>", lambda x: self.check_entries())

        self.deflection = tk.Label(
            self,
            text="Отклонение (если выше нормы), мм",
            anchor="e", 
            width=30
        )
        self.deflection_entry = tk.Entry(
            self,
            width=19,
            relief="sunken",
            bd=2
        )

        self.wire_pos = tk.Label(self, text="Расположение проводов", anchor="e", width=30)
        self.wire_pos_combobox = Combobox(
            self,
            values=("Горизонтальное", "Вертикальное"),
            width=16,
            validate="key"
        )
        self.wire_pos_combobox["validatecommand"] = (
            self.wire_pos_combobox.register(self.validate_wire_pos),
            "%P"
        )

        self.ground_wire_attachment = tk.Label(self, text="Крепление троса", anchor="e", width=30)
        self.ground_wire_attachment_combobox = Combobox(
            self,
            values=("Ниже верха опоры", "К верху опоры"),
            width=16,
            validate="key"
        )
        self.ground_wire_attachment_combobox["validatecommand"] = (
            self.ground_wire_attachment_combobox.register(self.validate_ground_wire_attach),
            "%P"
        )

        self.quantity_of_ground_wire = tk.Label(self, text="Количество тросов", anchor="e", width=30)
        self.quantity_of_ground_wire_combobox = Combobox(
            self,
            values=("1", "2"),
            width=16,
            validate="key"
        )
        self.quantity_of_ground_wire_combobox["validatecommand"] = (
            self.quantity_of_ground_wire_combobox.register(self.validate_quantity_of_ground_wire),
            "%P"
        )

        self.media = tk.Label(
            self,
            text="Ссылки на медиа-файлы:",
            anchor="e",
            width=20,
            bg="#FFFFFF"
        )

        self.pole = tk.Label(
            self,
            text="Общий вид опоры",
            anchor="e",
            width=15
        )
        self.pole_entry = tk.Entry(
            self,
            width=30,
            relief="sunken",
            bd=2
        )
        self.browse_for_pole_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pole
        )

        self.pole_defl = tk.Label(
            self,
            text="Отклонение опоры",
            anchor="e",
            width=15
        )
        self.pole_defl_entry = tk.Entry(
            self,
            width=30,
            relief="sunken",
            bd=2
        )
        self.browse_for_pole_defl_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pole_defl
        )

        self.loads = tk.Label(
            self,
            text="Нагрузки",
            anchor="e",
            width=15
        )
        self.loads_entry = tk.Entry(
            self,
            width=30,
            relief="sunken",
            bd=2
        )
        self.browse_for_loads_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_loads
        )

        self.is_mont_schema = tk.Label(
            self,
            text="Чертеж",
            anchor="e", 
            width=15
        )
        self.is_mont_schema_entry = tk.Entry(
            self,
            width=30,
            relief="sunken",
            bd=2
        )
        self.browse_for_mont_schema_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_mont_schema
        )

        self.generation = tk.Label(
            self,
            text="Сгенерировать и сохранить отчет:",
            anchor="e",
            width=32,
            bg="#FFFFFF"
        )

        self.path_to_txt_1_label = tk.Label(
            self,
            text="Путь к 1 отчету",
            anchor="e",
            width=13
        )
        self.path_to_txt_1_entry = tk.Entry(
            self,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.browse_txt_1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_txt_1
        )

        self.path_to_txt_2_label = tk.Label(
            self,
            text="Путь ко 2 отчету",
            anchor="e",
            width=13
        )
        self.path_to_txt_2_entry = tk.Entry(
            self,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.browse_txt_2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_txt_2
        )
        
        self.generate_and_save_button = tk.Button(
            self, text="Сгенерировать отчет",
            command=self.generate_output,
        )

    def run(self):
        self.draw_widgets()
        self.mainloop()

    def draw_widgets(self):
        self.back_to_main_window_button.place(x=15, y=2)
        self.open_button.place(x=41, y=2)
        self.save_button.place(x=67, y=2)
        self.calculation_clarification_bg1.place(x=10, y=0)
        self.calculation_clarification_bg2.place(x=370, y=0)
        self.media_bg.place(x=10, y=126)
        self.generation_bg.place(x=370, y=126)
        self.wire_pos.place(x=14, y=27)
        self.wire_pos_combobox.place(x=231, y=27)
        self.ground_wire_attachment.place(x=14, y=50)
        self.ground_wire_attachment_combobox.place(x=231, y=50)
        self.quantity_of_ground_wire.place(x=14, y=73)
        self.quantity_of_ground_wire_combobox.place(x=231, y=73)
        self.is_ground_wire_davit.place(x=375, y=27)
        self.is_ground_wire_davit_combobox.place(x=595, y=27)
        self.deflection.place(x=375, y=50)
        self.deflection_entry.place(x=595, y=50)
        self.media.place(x=115,y=102)
        self.pole.place(x=15,y=132)
        self.pole_entry.place(x=126,y=132)
        self.browse_for_pole_button.place(x=313,y=129)
        self.pole_defl.place(x=15,y=155)
        self.pole_defl_entry.place(x=126,y=155)
        self.browse_for_pole_defl_button.place(x=313,y=152)
        self.loads.place(x=15,y=178)
        self.loads_entry.place(x=126,y=178)
        self.browse_for_loads_button.place(x=313,y=175)
        self.is_mont_schema.place(x=15, y=201)
        self.is_mont_schema_entry.place(x=126, y=201)
        self.browse_for_mont_schema_button.place(x=313,y=201)
        self.generation.place(x=430,y=102)
        self.path_to_txt_1_label.place(x=373,y=132)
        self.path_to_txt_1_entry.place(x=470,y=132)
        self.browse_txt_1_button.place(x=657,y=129)
        self.path_to_txt_2_label.place(x=373,y=155)
        self.path_to_txt_2_entry.place(x=470,y=155)
        self.browse_txt_2_button.place(x=657,y=152)
        self.generate_and_save_button.place(x=485,y=191)

    def browse_for_pole(self):
        self.file_path = make_path_png()
        self.pole_entry.delete("0", "end")
        self.pole_entry.insert("insert", self.file_path)

    def browse_for_pole_defl(self):
        self.file_path = make_path_png()
        self.pole_defl_entry.delete("0", "end")
        self.pole_defl_entry.insert("insert", self.file_path)

    def browse_for_loads(self):
        self.file_path = make_multiple_path()
        self.loads_entry.delete("0", "end") 
        self.loads_entry.insert("insert", self.file_path)

    def browse_for_mont_schema(self):
        self.file_path = make_path_png()
        self.is_mont_schema_entry.delete("0", "end") 
        self.is_mont_schema_entry.insert("insert", self.file_path)
        
    def browse_for_txt_1(self):
        self.file_path = make_path_txt()
        self.path_to_txt_1_entry.delete("0", "end") 
        self.path_to_txt_1_entry.insert("insert", self.file_path)

    def browse_for_txt_2(self):
        self.file_path = make_path_txt()
        self.path_to_txt_2_entry.delete("0", "end") 
        self.path_to_txt_2_entry.insert("insert", self.file_path)

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()

    def save_data(self):
        filename = fd.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
        )
        if filename:
            with open(filename, "w") as file:
                file.writelines(
                    [
                    self.is_ground_wire_davit_combobox.get() + "\n",
                    self.deflection_entry.get() + "\n",
                    self.wire_pos_combobox.get() + "\n",
                    self.ground_wire_attachment_combobox.get() + "\n",
                    self.quantity_of_ground_wire_combobox.get() + "\n",
                    self.pole_entry.get() + "\n",
                    self.pole_defl_entry.get() + "\n",
                    self.loads_entry.get() + "\n",
                    self.is_mont_schema_entry.get() + "\n",
                    self.path_to_txt_1_entry.get() + "\n",
                    self.path_to_txt_2_entry.get() + "\n"
                    ]
                )
            
    def open_data(self):
        filename = fd.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filename:
            with open(filename, "r") as file:
                self.is_ground_wire_davit_combobox.set(file.readline().rstrip("\n")),
                self.deflection_entry.delete(0, "end"),
                self.deflection_entry.insert(0, file.readline().rstrip("\n")),
                self.wire_pos_combobox.set(file.readline().rstrip("\n")),
                self.ground_wire_attachment_combobox.set(file.readline().rstrip("\n")),
                self.quantity_of_ground_wire_combobox.set(file.readline().rstrip("\n")),
                self.pole_entry.delete(0, "end"),
                self.pole_entry.insert(0, file.readline().rstrip("\n")),
                self.pole_defl_entry.delete(0, "end"),
                self.pole_defl_entry.insert(0, file.readline().rstrip("\n")),
                self.loads_entry.delete(0, "end"),
                self.loads_entry.insert(0, file.readline().rstrip("\n")),
                self.is_mont_schema_entry.delete(0, "end"),
                self.is_mont_schema_entry.insert(0, file.readline().rstrip("\n")),
                self.path_to_txt_1_entry.delete(0, "end"),
                self.path_to_txt_1_entry.insert(0, file.readline().rstrip("\n")),
                self.path_to_txt_2_entry.delete(0, "end"),
                self.path_to_txt_2_entry.insert(0, file.readline().rstrip("\n"))

    def generate_output(self):
        if re.match(r"\w+\s\d+/\d+", self.parent.wire):
            self.result = put_data(
                project_name=self.parent.project_name,
                project_code=self.parent.project_code,
                pole_code=self.parent.pole_code,
                pole_type=self.parent.pole_type,
                developer=self.parent.developer,
                voltage=self.parent.voltage,
                area=self.parent.area,
                branches=self.parent.branches,
                wind_region=self.parent.wind_region,
                wind_pressure=self.parent.wind_pressure,
                ice_region=self.parent.ice_region,
                ice_thickness=self.parent.ice_thickness,
                ice_wind_pressure=self.parent.ice_wind_pressure,
                year_average_temp=self.parent.year_average_temp,
                min_temp=self.parent.min_temp,
                max_temp=self.parent.max_temp,
                ice_temp=self.parent.ice_temp,
                wind_temp=self.parent.wind_temp,
                wind_reg_coef=self.parent.wind_reg_coef,
                ice_reg_coef=self.parent.ice_reg_coef,
                wire_hesitation=self.parent.wire_hesitation,
                wire=self.parent.wire,
                wire_tencion=self.parent.wire_tencion,
                ground_wire=self.parent.ground_wire,
                oksn=self.parent.oksn,
                wind_span=self.parent.wind_span,
                weight_span=self.parent.weight_span,
                is_stand=self.parent.is_stand,
                is_plate=self.parent.is_plate,
                is_ground_wire_davit=self.is_ground_wire_davit_combobox.get(),
                deflection=self.deflection_entry.get(),
                wire_pos=self.wire_pos_combobox.get(),
                ground_wire_attachment=self.ground_wire_attachment_combobox.get(),
                quantity_of_ground_wire=self.quantity_of_ground_wire_combobox.get(),
                pole=self.pole_entry.get(),
                pole_defl_pic=self.pole_defl_entry.get(),
                loads_str=self.loads_entry.get(),
                mont_schema=self.is_mont_schema_entry.get(),
                path_to_txt_1=self.path_to_txt_1_entry.get(),
                path_to_txt_2=self.path_to_txt_2_entry.get()
            )
        else:
            print("!!!ОШИБКА!!!: Проверь правильность введенного провода!")

    def validate_pole_type(self, value):
        if value in ["Анкерно-угловая", "Концевая", "Отпаечная", "Промежуточная"]:
            return True
        return False
    
    def validate_voltage(self, value):
        return value.isdigit()
    
    def validate_branches(self, value):
        if value in ["1", "2"]:
            return True
        return False

    def validate_float(self, value):
        return re.match(r"^\d*\.?\d*$", value) is not None
    
    def validate_wire_pos(self, value):
        if value in ["Горизонтальное", "Вертикальное", ""]:
            return True
        return False
    
    def validate_ground_wire_attach(self, value):
        if value in ["Ниже верха опоры", "К верху опоры", ""]:
            return True
        return False
    
    def validate_yes_no(self, value):
        if value in ["Да", "Нет"]:
            return True
        return False
    
    def validate_quantity_of_ground_wire(self, value):
        if value in ["1", "2", ""]:
            return True
        return False
