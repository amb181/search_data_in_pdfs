""" 
Version 1.0.0
almolina@tenaris.com
06/2022
"""
from tkinter import *
from tkinter import filedialog
import glob, os, re, time, fitz, difflib, sys, datetime

WINDOW_WIDTH = 450
WINDOW_HEIGHT = 150

class Window():
    def __init__(self) -> None:
        self.path = ""
        self.folder = []
        self.running = False
        self.html_name = ""
        self.time = 0.0

    def run(self):
        self.search_init()

    def search_init(self):
        self.running = False
        try:
            self.master.destroy()
            self.running = True
        except:
            pass
        self.master = Tk()
    
        self.master.title("Compare pdfs")
        self.master['background'] = "#0088ff"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))

        self.l1 = Label(self.master, text="Folder location:", bg='#0088ff', fg="#fff", font=("Arial", 14))
        self.l1.grid(row=0, column=0)

        self.e1 = Entry(self.master, font=("Arial", 14))
        self.e1.grid(row=0, column=1)

        b = Button(self.master, text="Select", font=("Arial", 10), command=self.get_path)
        b.grid(row=0, column=2)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.columnconfigure(2, weight=1)
        self.master.rowconfigure(0, weight=1)

        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())

        self.master.mainloop()

    def done(self):
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#0088ff"

        l1 = Label(self.master, text='Search done!', bg='#0088ff', fg="#fff", font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        l2 = Label(self.master, text='Elapsed time: {:.2f} min'.format(self.time), 
                    bg='#0088ff', fg="#fff", font=("Arial", 14))
        l2.grid(row=1, columnspan=2)

        b1 = Button(self.master, text="Open report", font=("Arial", 10), command=self.open_report)
        b1.grid(row=2, column=0)
        b2 = Button(self.master, text="Close", font=("Arial", 10), command=lambda:sys.exit())
        b2.grid(row=2, column=1)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        self.master.rowconfigure(2, weight=1)
        
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())

        self.master.mainloop()

    def open_report(self):
        try:
            os.system("start CHROME.EXE " + self.html_name + ".html")
            sys.exit()
        except Exception as e:
            print(e.args)

    def get_path(self):
        self.path = str(self.e1.get())
        if self.path == "":
            self.path = filedialog.askdirectory()
        if self.path == "":
            self.master.destroy()
            return self.search_init()
            
        self.folder = glob.glob(self.path + "/*.pdf")

        if self.running:
            return self.get_params()

        self.master.destroy()

        if self.path == "":
            self.search_init()

    def ask_filename(self):
        self.master = Tk()
        self.master.withdraw()

        today = datetime.datetime.today().strftime('%Y-%m-%d').replace("-", "_")
        csv_placeholder = "Comparison_" + today

        self.html_name = filedialog.asksaveasfilename(
            defaultextension = '.html', 
            filetypes = [('HTML files', '*.html')],
            initialdir = self.path,
            initialfile = csv_placeholder
        )

        if ".html" in self.html_name:
            self.html_name = self.html_name.replace(".html", "")

        self.master.destroy()

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

class Compare():
    def __init__(self) -> None:
        self.Ordinals = {
            'FIRST'     :  1,
            'SECOND'    :  2,
            'THIRD'     :  3,
            'FOURTH'    :  4,
            'FIFTH'     :  5,
            'SIXTH'     :  6,
            'SEVENTH'   :  7,
            'EIGHTH'    :  8,
            'NINTH'     :  9,
            'TENTH'     : 10,
            'ELEVENTH'  : 11 
        }
        self.path = ""
        self.diff = ""
        self.new_file = ""

    def pdf_2_txt(self, pdf):
        name = os.path.basename(pdf).split('/')[:1][0].replace(".pdf", "")
        
        doc = fitz.open(pdf)
        num_pag = doc.page_count
        for i in range(num_pag):
            page = doc.load_page(i)
            text = page.get_text()

            with open(self.path + "/" + name + ".txt", "a", encoding="utf-8") as f:
                f.write(text)
                f.write("- * - * - * - E N D  O F  P A G E  {} - * - * - * -\n".format(i + 1))
            
    def compare_files(self):
        files = {}
        folder_txt = glob.glob(self.path + "/*.txt")

        for file in folder_txt:
            with open(file, "r", encoding="utf-8") as f:
                fi = f.readlines()
            os.remove(file)

            # Find doc version
            version = ""
            for row in fi:
                try:
                    version = re.compile("^([A-Z]+) EDITION").findall(row)[0]
                    break
                except:
                    pass

            version = self.Ordinals[version]
            files[int(version)] = fi

        files = list(sorted(files.items()))

        self.diff = difflib.HtmlDiff(wrapcolumn=70).make_file(files[0][1], files[1][1])
        with open(self.html_name + ".html", "w", encoding="utf-8") as f:
            f.write(self.diff)

        self.new_file = files[1][1]

    def validate_search(self):
        blocks = []
        pattern = re.compile('class="diff_sub">(.*?)</span>')
        
        try:
            total_sub = pattern.findall(self.diff)
            for sub in total_sub:
                text = sub.replace("&nbsp;", " ")
                blocks.append(text)
        except:
            pass
        
        for block in blocks:
            if block in self.new_file:
                print("Already in older file!\n{}\n".format(block))
        
# # # #  M A I N  F U N C T I O N  # # # #

if __name__ == '__main__':
    # Init window
    app = Window()
    comp = Compare()
    app.run()
    # Ask name for csv
    app.ask_filename()
    
    comp.path = app.path
    comp.html_name = app.html_name

    start_time = time.time()

    for pdf in app.folder:
        comp.pdf_2_txt(pdf)

    comp.compare_files()
    comp.validate_search()

    end_time = time.time() - start_time
    app.time = end_time/60
    
    app.done()
    print("Elapsed time: {:.2f} min".format(app.time))