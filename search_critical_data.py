""" 
Version 1.0.0
almolina@tenaris.com
06/2022
"""
# Import libraries
from tkinter import *
from tkinter import filedialog
import pandas as pd
import os, glob, fitz, csv, time, itertools, shutil, sys, datetime

WINDOW_WIDTH = 450
WINDOW_HEIGHT = 150

class Window():
    def __init__(self) -> None:
        self.path = ""
        self.folder = []
        self.running = False
        self.csv_name = ""
        self.words_lookup = []
        self.time = 0.0
        self.data = []

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
    
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))

        self.l1 = Label(self.master, text="Folder location:", bg='#499e03', font=("Arial", 14))
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
    
    def params_window(self):
        self.master = Tk()

        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))

        # Menu bar
        menu_bar = Menu(self.master)
        self.master.config(menu=menu_bar)
        file_menu = Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Change folder", command=self.search_init)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=lambda:sys.exit())
        menu_bar.add_cascade(label="Config", menu=file_menu)

        # Content
        l1 = Label(self.master, text='Parameters (separate each with a ";") :', bg='#499e03', font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        
        self.e1 = Entry(self.master, width=30, font=("Arial", 14))
        self.e1.grid(row=1, column=0)
        
        b = Button(self.master, text="Search", font=("Arial", 10), command=self.get_params)
        b.grid(row=1, column=1)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)

        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())

        self.master.mainloop()

    def ask_filename(self):
        today = datetime.datetime.today().strftime('%Y-%m-%d').replace("-", "_")
        csv_placeholder = "Search_report_" + today

        csv_name = filedialog.asksaveasfilename(
            defaultextension = '.csv', 
            filetypes = [('CSV files', '*.csv')],
            initialdir = self.path,
            initialfile = csv_placeholder
        )

        if csv_name != "":
            self.master.destroy()

            if ".csv" in csv_name:
                self.csv_name = csv_name.replace(".csv", "")

        

    def waiting(self):
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        gif = PhotoImage(file=Window.resource_path("loading.gif"))
        Label(image=gif).grid(row=0, column=0)

        self.center(340, 340)
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())
        
        self.master.update()
        
    def done(self):
        self.master.destroy()
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"

        l1 = Label(self.master, text='Search done!', bg='#499e03', font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        l2 = Label(self.master, text='Elapsed time: {:.2f} min'.format(self.time), bg='#499e03', font=("Arial", 14))
        l2.grid(row=1, columnspan=2)

        b1 = Button(self.master, text="Open report", font=("Arial", 10), command=self.open_report)
        b1.grid(row=2, column=0)
        b2 = Button(self.master, text="Close", font=("Arial", 10), command=self.continue_search)
        b2.grid(row=2, column=1)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        self.master.rowconfigure(2, weight=1)
        
        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())

        self.master.mainloop()

    def continue_search(self):
        self.master.destroy()
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"

        l1 = Label(self.master, text='Search again?', bg='#499e03', font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        
        b1 = Button(self.master, text="Search!", font=("Arial", 12), command=self.master.destroy)
        b1.grid(row=1, column=0)
        b2 = Button(self.master, text="Exit", font=("Arial", 12), command=lambda:sys.exit())
        b2.grid(row=1, column=1)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)

        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", lambda:sys.exit())

        self.master.mainloop()

    def get_path(self):
        self.path = str(self.e1.get())
        if self.path == "":
            self.path = filedialog.askdirectory()
        if self.path != "":
            self.master.destroy()

        self.folder = glob.glob(self.path + "/*.pdf")

    def get_params(self):
        user_input = str(self.e1.get())
        if user_input != "":
            self.words_lookup= user_input.split(";")
        else:
            param_file = filedialog.askopenfilename(
                filetypes = [('Excel files', '*.xlsx'), ('CSV files', '*.csv')],
                initialdir = self.path
            )

            if param_file != "":

                if ".xlsx" in param_file:
                    df = pd.DataFrame(pd.read_excel(param_file))
                    self.words_lookup = df['Parametro'].tolist()
                elif ".csv" in param_file:
                    df = pd.DataFrame(pd.read_csv(param_file))
                    self.words_lookup = df['Parametro'].tolist()
                
                self.ask_filename()
            
        for i in range(len(self.words_lookup)):
            self.words_lookup[i] = self.words_lookup[i].strip().upper()

    def open_report(self):
        try:
            os.system("start EXCEL.EXE " + self.csv_name + ".csv")
        except Exception as e:
            print(e.args)
        self.continue_search()

    def center(self, w_width, w_height):
        screen_width = self.master.winfo_screenwidth() # 1920
        screen_height = self.master.winfo_screenheight() # 1080
        x_cordinate = int((screen_width/2) - (w_width/2))
        y_cordinate = int((screen_height/2) - (w_height/2))

        self.master.geometry("{}x{}+{}+{}".format(w_width, w_height, x_cordinate, y_cordinate))

    def _onKeyRelease(self, event):
        ctrl  = (event.state & 0x4) != 0
        if event.keycode==88 and  ctrl and event.keysym.lower() != "x": 
            event.widget.event_generate("<<Cut>>")

        if event.keycode==86 and  ctrl and event.keysym.lower() != "v": 
            event.widget.event_generate("<<Paste>>")

        if event.keycode==67 and  ctrl and event.keysym.lower() != "c":
            event.widget.event_generate("<<Copy>>")

    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath("./imgs")

        return os.path.join(base_path, relative_path)


class Critical_Data():
    def __init__(self) -> None:
        self.path = ""
        self.csv_name = ""
        self.words_lookup = []
        self.data = []

    # Lookup critical data
    def search_data(self, pdf):
        # Create a copy of original file
        os.makedirs(self.path + "\\highlighted_pdf\\", exist_ok=True)
        pdf_name = os.path.basename(pdf).replace(".pdf", "")
        searched_pdf = self.path + "\\highlighted_pdf\\" + pdf_name + "_" \
                        + os.path.basename(self.csv_name) + ".pdf"
        shutil.copyfile(pdf, searched_pdf)
        # Extract text and search parameter
        doc = fitz.open(searched_pdf)
        num_pag = doc.page_count
        for i in range(num_pag):
            page = doc.load_page(i)
            for param in self.words_lookup:
                # Highlight text if found
                quads = page.search_for(param, quads=True)
                # Add to csv where the text was found
                if quads:
                    self.data.append(['=HYPERLINK("' + searched_pdf + '","' + os.path.basename(searched_pdf) + '")', 
                                    param, 
                                    "Page: " + str(i + 1)])

                page.add_highlight_annot(quads)
        # Save highlighted pdf
        doc.save(searched_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

        return

    # Open excel file and append data
    def WriteCSV(self):
        with open(self.csv_name + ".csv", 'a+', encoding="utf-8") as csv_doc:
            csv_doc = csv.writer(csv_doc)
            csv_doc.writerows(self.data)

            return


# # # #  M A I N  F U N C T I O N  # # # #

if __name__ == "__main__":
    # Init classes
    app = Window()
    crit = Critical_Data()
    app.run()

    while True:
        # Start csv file
        headers = ['Document', 'Parameter', 'Page']
        crit.data.append(headers)

        start_time = time.time()

        # Ask for parameters to search
        app.params_window()
        # Ask name for csv
        # app.ask_filename()
        # Waiting window
        app.waiting()

        # Pass attribute
        crit.path = app.path
        crit.csv_name = app.csv_name
        crit.words_lookup = app.words_lookup
        
        for file_ in app.folder:
            crit.search_data(file_)
        
        end_time = time.time() - start_time
        app.time = end_time/60

        # Delete previous report if exists, without this data will be appended
        try:
            os.remove(app.csv_name + ".csv")
        except:
            pass
        
        # Drop duplicates
        crit.data = list(crit.data for crit.data, _ in itertools.groupby(crit.data))
        crit.WriteCSV()

        # Remove empty rows in excel file
        df = pd.read_csv(app.csv_name + ".csv", encoding="utf-8")
        # Droping the empty rows
        modifiedDF = df.dropna()
        # Saving it to the csv file 
        modifiedDF.to_csv(app.csv_name + ".csv", index=False)

        print("\nSearch finished in {} files.\nElapsed time: {:.2f} min".format(len(app.folder), app.time))
        # Show done window
        app.done()