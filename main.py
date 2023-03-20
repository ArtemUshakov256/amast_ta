import base64
import os
import tkinter as tk

from logic.functionality.main_window import MainWindow
from logic.functionality.description_window import DescriptionWindow
from tkinter import filedialog as fd

icondata= base64.b64decode(os.getenv("icon"))
tempFile= "logo.ico"
iconfile= open(tempFile,"wb")
iconfile.write(icondata)
iconfile.close()

main_window = MainWindow(icon=tempFile)
main_window.run()
