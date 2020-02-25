 #! python3
 # -*- coding: utf-8 -*-
from tkinter import *
import tkinter.ttk as ttk
from ttkthemes import ThemedStyle
from PIL import ImageTk, Image
from scripts.functions import open_file, prepare, filtering, get_logo_path

WINDOW_WIDTH = 600
WINDOW_HEIGHT = 220

#basic parameters for window
root = Tk()
root.title("BOM Filtering Utility")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+300+200")
root.resizable(False, False)

#style
style=ThemedStyle(root)
style.set_theme("plastik")

#open file
ttk.Label(text="Excel File:").grid(row=0, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
chosen_file = ttk.Button(text="Openfile")
chosen_file.grid(row=0, column=1, columnspan=3, sticky=W+E, padx=10)

#choose mode
ttk.Label(text="Select Mode:").grid(row=1, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
mode = ttk.Spinbox(width=7, values=("Guides", "Models"))
mode.grid(row=1, column=1, columnspan=3, sticky=W+E, padx=10)

#prepare filtering
ttk.Label(text="Prepare Filtering:").grid(row=2, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
prepare_filtering = ttk.Button(text="Prepare")
prepare_filtering.grid(row=2, column=1, columnspan=3, sticky=W+E, padx=10)

#Filter
ttk.Label(text="Filter Excel File:").grid(row=3, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
filter_file = ttk.Button(text="Prepare")
filter_file.grid(row=3, column=1, columnspan=3, sticky=W+E, padx=10)

#Chekc Slicer name
ttk.Label(text="Name of Slicer:").grid(row=4, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
slicer_name = ttk.Entry()
slicer_name.insert(END, "Slicer_BOM")
slicer_name.grid(row=4, column=1, columnspan=3, sticky=W+E, padx=10)

root.mainloop()