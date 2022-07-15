""" 
Version 1.0.1
almolina@tenaris.com
07/2022
"""
from tkinter import *
from PIL import ImageTk, Image
import search_critical_data as scd
import compare_pdfs as cpdf
import os, sys

WINDOW_WIDTH = 450
WINDOW_HEIGHT = 150

class Window():
    def run(self):
        # Init window
        self.master = Tk()
        self.master.title("Document analysis tool")
        self.master['background'] = "#0088ff"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.resizable(0,0)
        # Get image and resize
        img = Image.open(Window.resource_path("Tenaris_Logo.png"))
        img = img.resize((WINDOW_WIDTH, WINDOW_HEIGHT))
        # Menu bar
        self.menu()
        # Window content
        img = ImageTk.PhotoImage(img)
        #lb = Label(self.master, image=img)
        #lb.pack()
        Label(image=img).grid(row=0, column=0)

        # Common functionalities
        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)

        # Show window
        self.master.mainloop()        

    # Add menu bar to window
    def menu(self):
        # Critical Data
        self.menu_bar = Menu(self.master)
        self.master.config(menu=self.menu_bar)
        search_menu = Menu(self.menu_bar, tearoff=False)
        search_menu.add_command(label="Search Critical Data", 
                                command=lambda:self.search_critical_data())
        self.menu_bar.add_cascade(label="Critical Data", menu=search_menu)
        # Compare pdfs
        compare_menu = Menu(self.menu_bar, tearoff=False)
        compare_menu.add_command(label="Compare 2 PDF versions",
                                command=lambda:self.compare_pdfs())
        self.menu_bar.add_cascade(label="Compare", menu=compare_menu)
        # PO digitalization
        digit_menu = Menu(self.menu_bar, tearoff=False)
        digit_menu.add_command(label="PO digitalization")
        self.menu_bar.add_cascade(label="PDF digitalization", menu=digit_menu)
        # About
        help_menu = Menu(self.menu_bar, tearoff=False)
        help_menu.add_command(label="About", command=lambda:sys.exit())
        help_menu.add_separator()
        help_menu.add_command(label="Exit", command=lambda:sys.exit())
        self.menu_bar.add_cascade(label="Help", menu=help_menu)
        self.master.config(menu=self.menu_bar)

    def search_critical_data(self):
        # Init window
        crit = scd.Window()
        crit.waiting()
        crit.wait.destroy()
        crit.run()

    def compare_pdfs(self):
        crit = scd.Window()
        crit.waiting()
        # Init window
        cp = cpdf.Window()
        crit.wait.destroy()
        cp.run()

    # Center window
    def center(self, w_width, w_height):
        screen_width = self.master.winfo_screenwidth() # 1920
        screen_height = self.master.winfo_screenheight() # 1080
        x_cordinate = int((screen_width/2) - (w_width/2))
        y_cordinate = int((screen_height/2) - (w_height/2))

        self.master.geometry("{}x{}+{}+{}".format(w_width, w_height, x_cordinate, y_cordinate))

    # Enable cut, copy & paste
    def _onKeyRelease(self, event):
        ctrl  = (event.state & 0x4) != 0
        if event.keycode==88 and  ctrl and event.keysym.lower() != "x": 
            event.widget.event_generate("<<Cut>>")

        if event.keycode==86 and  ctrl and event.keysym.lower() != "v": 
            event.widget.event_generate("<<Paste>>")

        if event.keycode==67 and  ctrl and event.keysym.lower() != "c":
            event.widget.event_generate("<<Copy>>")

    # Resolve images path for local and exe
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath("./imgs")
            
        return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    doc_app = Window()
    doc_app.run()