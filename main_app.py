""" 
Version 1.0.0
almolina@tenaris.com
06/2022
"""
from tkinter import *
from PIL import ImageTk, Image
import os, sys

WINDOW_WIDTH = 450
WINDOW_HEIGHT = 150

class Main_Window():
    def run(self):
        # Init window
        self.master = Tk()
        self.set_window()
        # Get image and resize
        img = Image.open(Main_Window.resource_path("Tenaris_Logo.png"))
        img = img.resize((WINDOW_WIDTH, WINDOW_HEIGHT))
        # Menu bar
        self.menu()
        # Window content
        img = ImageTk.PhotoImage(img)
        lb = Label(self.master, image=img)
        lb.pack()

        # Common functionalities
        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", self.kill_children)

        # Show window
        self.master.mainloop()

    # Init window parameters
    def set_window(self):
        self.master.title("Document analysis tool")
        self.master['background'] = "#0088ff"
        self.master.iconbitmap(default=Main_Window.resource_path("Tenaris_Logo.ico"))
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.resizable(0,0)

    # Add menu bar to window
    def menu(self):
        # Critical Data
        self.menu_bar = Menu(self.master)
        self.master.config(menu=self.menu_bar)
        search_menu = Menu(self.menu_bar, tearoff=False)
        search_menu.add_command(label="Search Critical Data", 
                                command=lambda:self.new_thread("apps\\search_critical_data.exe"))
        self.menu_bar.add_cascade(label="Critical Data", menu=search_menu)
        # Compare pdfs
        compare_menu = Menu(self.menu_bar, tearoff=False)
        compare_menu.add_command(label="Compare 2 PDF versions",
                                command=lambda:self.new_thread("apps\\compare_pdfs.exe"))
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

    # def show_path(self):
    #     self.master.destroy()
    #     self.master = Tk()
        
    #     lb = Label(self.master, text=sys.executable.replace("main_app.exe", "") + "apps\\search_critical_data.exe")
    #     lb.pack()

    #     self.master.mainloop()
   
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

    # Resolve thread argument
    def new_thread(self, process):
        return os.system("start \"\" \"" + sys.executable.replace("main_app.exe", "") + process + "\"")

    def kill_children(self):
        try:
            os.system('taskkill /IM "compare_pdfs.exe" /F')
        except:
            print("Compare app was not running")
        try:
            os.system('taskkill /IM "search_critical_data.exe" /F')
        except:
            print("Critical data app was not running")
        sys.exit()



if __name__ == "__main__":
    window_app = Main_Window()
    window_app.run()
    window_app.show_path()