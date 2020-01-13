import tkinter as tk
from scripts.functions import open_file, prepare, filtering


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()


    def init_main(self):
        self.toolbar = tk.Frame(bg="#dfd8e0", bd=2)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        self.buttonOpen = tk.Button(text="Openfile", command=open_file)
        self.buttonOpen.pack()
        self.buttonPrepare = tk.Button(text="Prepare Filtering", command=prepare)
        self.buttonPrepare.pack()
        self.buttonFilter = tk.Button(text="Filter", command=filtering)
        self.buttonFilter.pack()


if __name__ == "__main__":
    root = tk.Tk()
    app = Main(root)
    app.pack()
    root.title("Excel filtering")
    root.geometry("650x450+300+200")
    root.resizable(False, False)
    root.mainloop()