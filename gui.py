import tkinter as tk
import os
from tkinter import filedialog as fd
from scripts.prepare import prepare


FILE_NAME = ''
SLICER_NAME = 'Slicer_BOM'


def open_file():
    file_name = fd.askopenfile()
    global FILE_NAME
    FILE_NAME = file_name.name

    #os.startfile(FILE_NAME)


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()


    def click_open_button(self):
        open_file()


    def click_prepare_button(self):
        prepare(FILE_NAME, SLICER_NAME)
    

    def init_main(self):
        self.toolbar = tk.Frame(bg="#dfd8e0", bd=2)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        self.buttonOpen = tk.Button(self, text="Openfile", command=self.click_open_button)
        self.buttonOpen.pack()
        self.buttonPrepare = tk.Button(self, text="Prepare Filtering", command=self.click_prepare_button)
        self.buttonPrepare.pack()


if __name__ == "__main__":
    root = tk.Tk()
    app = Main(root)
    app.pack()
    root.title("Excel filtering")
    root.geometry("650x450+300+200")
    root.resizable(False, False)
    root.mainloop()