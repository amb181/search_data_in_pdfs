""" 
Version 1.0.1
almolina@tenaris.com
07/2022
"""
from tkinter import *
from tkinter import filedialog
import os, time, difflib, sys, datetime, re
import xml.etree.ElementTree as ET
import search_critical_data as scd

WINDOW_WIDTH = 450
WINDOW_HEIGHT = 150
DEFAULT_PATH = "C:/Users/" + os.getlogin() + "/Documents/"

class Window():
    def __init__(self) -> None:
        self.file1 = ""
        self.file2 = ""
        self.folder = []
        self.html_name = ""
        self.time = 0.0

    def run(self):
        self.search_init()

    def search_init(self):
        self.master = Tk()
    
        self.master.title("Compare pdfs")
        self.master['background'] = "#0066CC"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))

        self.l1 = Label(self.master, text="Choose files:", bg='#0066CC', fg="#fff", font=("Arial", 14))
        self.l1.grid(row=0, column=0)

        # self.e1 = Entry(self.master, font=("Arial", 14))
        # self.e1.grid(row=0, column=1)

        b = Button(self.master, text="Select files", font=("Arial", 10), command=self.get_files)
        b.grid(row=1, column=0) # row=0, column=2

        self.master.columnconfigure(0, weight=1)
        # self.master.columnconfigure(1, weight=1)
        # self.master.columnconfigure(2, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)   # Add

        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:self.master.destroy())

        self.master.mainloop()

    def done(self):
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#0066CC"

        l1 = Label(self.master, text='Search done!', bg='#0066CC', fg="#fff", font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        l2 = Label(self.master, text='Elapsed time: {:.2f} min'.format(self.time), 
                    bg='#0066CC', fg="#fff", font=("Arial", 14))
        l2.grid(row=1, columnspan=2)

        b1 = Button(self.master, text="Open report", font=("Arial", 10), command=self.open_report)
        b1.grid(row=2, column=0)
        b2 = Button(self.master, text="Close", font=("Arial", 10), command=lambda:self.master.destroy())
        b2.grid(row=2, column=1)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        self.master.rowconfigure(2, weight=1)
        
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:self.master.destroy())

        self.master.mainloop()

    def open_report(self):
        try:
            os.system("start CHROME.EXE " + self.html_name + ".html")
            self.master.destroy()
        except Exception as e:
            print(e.args)

    def get_files(self):
        global DEFAULT_PATH
        # self.path = str(self.e1.get())
        # if self.path == "":
        #     self.path = filedialog.askdirectory()
        # if self.path == "":
        #     self.master.destroy()
        #     return self.search_init()
        # Get 1st file
        self.file1 = filedialog.askopenfilename(
            filetypes = [('PDF files', '*.pdf')],
            initialdir = DEFAULT_PATH
        )
        DEFAULT_PATH = os.path.abspath(self.file1).replace(os.path.basename(self.file1), "")
        # Get 2nd file
        self.file2 = filedialog.askopenfilename(
            filetypes = [('PDF files', '*.pdf')],
            initialdir = DEFAULT_PATH
        )

        self.folder = [self.file1, self.file2]

        self.master.destroy()

        if self.file1 == "" or self.file2 == "":
            self.search_init()
        else:
            self.ask_filename()

    def ask_filename(self):
        global DEFAULT_PATH
        self.master = Tk()
        self.master.withdraw()

        today = datetime.datetime.today().strftime('%Y-%m-%d').replace("-", "_")
        csv_placeholder = "Comparison_" + today

        self.html_name = filedialog.asksaveasfilename(
            defaultextension = '.html', 
            filetypes = [('HTML files', '*.html')],
            initialdir = DEFAULT_PATH,
            initialfile = csv_placeholder
        )

        if ".html" in self.html_name:
            self.html_name = self.html_name.replace(".html", "")

        self.master.destroy()
        comp = Compare()
        comp.html_name = self.html_name
        search_critical = scd.Window()
        start_time = time.time()
        search_critical.waiting()
        for pdf in self.folder:
            comp.pdf_2_txt(pdf)

        comp.compare_files()
        #comp.validate_search()

        end_time = time.time() - start_time
        self.time = end_time/60
        search_critical.wait.destroy()
        self.done()
        print("Elapsed time: {:.2f} min".format(self.time))
        
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
        self.diff = ""
        self.txt_files = []

    def pdf_2_txt(self, pdf):
        txt_name = pdf.replace(".pdf", ".txt")
        extract_pdf = "python -m fitz gettext \"{}\" -output \"{}\"".format(pdf, txt_name) #-noformfeed -extra-spaces
        os.system(extract_pdf)
        text = ""
        with open(txt_name, "r", encoding="utf8") as f:
            lines = f.readlines()
        for line in lines:
            if line.strip():
                line = re.sub(r"\s+", " ", line)
                text = text + "\n" + line
            else:
                continue
        os.remove(txt_name)
        self.txt_files.append(text.strip())
            
    def compare_files(self):
        self.diff = difflib.HtmlDiff(wrapcolumn=80).make_file(self.txt_files[0].split("\n"), self.txt_files[1].split("\n"))
        with open(self.html_name + ".html", "w", encoding="utf-8") as f:
            f.write(self.diff)

    def validate_search(self):
        blocks = []
        self.diff = self.diff.replace("&nbsp;", " ")
        tree = ET.fromstring(self.diff)
        counter = 1
        for node in tree.iter('tbody'):
            for elem in node.iter('span'):
                if not elem.tag==node.tag:
                    #print("{}: {}".format(elem.attrib.get("class"), elem.text))
                    if elem.attrib.get("class") == "diff_sub":
                        blocks.append(elem.text)
                        counter = 0
                    else:
                        counter += 1
                        if counter > 1:
                            blocks.append("")
        
        new_blocks = []
        new_item = ""
        for item in blocks:
            if item == "" and new_item != "":
                new_blocks.append(new_item.strip())
                new_item = ""
            elif item != "":
                new_item = new_item + " " + item
        #print(new_blocks)
        # with open(self.html_name + ".txt", "w") as f:
        #     f.write(self.txt_files[1])
        
        for item in new_blocks:
            # Insert fuzzy wuzzy logic in block below #
            if item in self.txt_files[1]:
                print("Section duplicated: {}".format(item))
        
# # # # #  M A I N  F U N C T I O N  # # # #

# if __name__ == '__main__':
#     # Init window
#     app = Window()
#     comp = Compare()
#     app.run()
#     # Ask name for csv
#     app.ask_filename()
#     comp.html_name = app.html_name

#     start_time = time.time()

#     for pdf in app.folder:
#         comp.pdf_2_txt(pdf)

#     comp.compare_files()
#     comp.validate_search()

#     end_time = time.time() - start_time
#     app.time = end_time/60
    
#     app.done()
#     print("Elapsed time: {:.2f} min".format(app.time))