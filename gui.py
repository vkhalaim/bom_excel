 #! python3
 # -*- coding: utf-8 -*-
import tkinter as tk
from PIL import ImageTk, Image
from scripts.functions import open_file, prepare, filtering, get_logo_path


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()

    def init_main(self):
        #creating frames
        self.bottom_frame = tk.Frame(root, width=650, height=150)
        self.bottom_frame.pack(side=tk.BOTTOM)

        self.left_frame = tk.Frame(root, width=325, height=300)
        self.left_frame.pack(side=tk.LEFT)

        self.right_frame = tk.Frame(root, width=325, height=300)
        self.right_frame.pack(side=tk.RIGHT)

        #prepare logo
        path = get_logo_path()
        self.img = ImageTk.PhotoImage(Image.open(path))
        self.panel = tk.Label(self.bottom_frame, image=self.img, padx=-400)
        self.panel.pack(side=tk.LEFT)

        #open file for both models and guides
        self.button_open = tk.Button(text="Openfile", command=open_file)
        self.button_open.pack()

        # prepare guide design panel
        self.button_prepare = tk.Button(self.left_frame, text="Prepare Filtering",
                                       command=prepare)
        self.button_prepare.pack()
        self.button_filter = tk.Button(self.left_frame, text="Filter", command=filtering)
        self.button_filter.pack()

if __name__ == "__main__":
    root = tk.Tk()
    app = Main(root)
    app.pack()
    root.title("Excel filtering")
    root.geometry("650x450+300+200")
    root.resizable(False, False)
    root.mainloop()    
