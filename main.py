import tkinter as tk
from tkinter import filedialog as fd

from logic.functionality.main_window import MainWindow
from logic.functionality.description_window import DescriptionWindow

def callback():
    name= fd.askopenfilename() 
    print(name)

main_window = MainWindow(icon="logo.ico")
main_window.run()
