 #! python3
 # -*- coding: utf-8 -*-
from tkinter import *
import tkinter.ttk as ttk
from ttkthemes import ThemedStyle
from PIL import ImageTk, Image
from scripts.functions import open_file, prepare, filtering, get_logo_path

WINDOW_WIDTH = 600
WINDOW_HEIGHT = 300
  
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
chosen_file = ttk.Button(text="Openfile", command=open_file)
chosen_file.grid(row=0, column=1, columnspan=3, sticky=W+E, padx=10)

#Check Slicer name for guides
ttk.Label(text="Name of Slicer for guides:").grid(row=1, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
slicer_name_guides = ttk.Entry()
slicer_name_guides.insert(END, "Slicer_BOM")
slicer_name_guides.grid(row=1, column=1, columnspan=3, sticky=W+E, padx=10)

#Check Slicer name for models
ttk.Label(text="Name of Slicer for models:").grid(row=2, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
slicer_name_models = ttk.Entry()
slicer_name_models.insert(END, "Slicer_BOM1")
slicer_name_models.grid(row=2, column=1, columnspan=3, sticky=W+E, padx=10)

#prepare filtering
ttk.Label(text="Prepare Filtering:").grid(row=3, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
prepare_filtering = ttk.Button(text="Prepare", command=lambda: prepare(slicer_name_guides.get()))
prepare_filtering.grid(row=3, column=1, columnspan=3, sticky=W+E, padx=10)

#Filter
ttk.Label(text="Filter Excel File(Guides):").grid(row=4, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
filter_file = ttk.Button(text="Filter", command=lambda: filtering(slicer_name_guides.get(), 'GUIDES'))
filter_file.grid(row=4, column=1, columnspan=3, sticky=W+E, padx=10)

#Filter
ttk.Label(text="Filter Excel File(Models):").grid(row=5, column=0, sticky=W, pady=10, padx=WINDOW_WIDTH/5)
filter_file = ttk.Button(text="Filter", command=lambda: filtering(slicer_name_models.get(), 'MODELS'))
filter_file.grid(row=5, column=1, columnspan=3, sticky=W+E, padx=10)

#set logo
path = get_logo_path()
image = Image.open(path)
image = image.resize((150, 50), Image.ANTIALIAS)
img = ImageTk.PhotoImage(image)
panel = ttk.Label(root, image = img).grid(row=6, column=0, padx=10)


root.mainloop()