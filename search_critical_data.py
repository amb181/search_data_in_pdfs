""" 
Version 1.0.0
almolina@tenaris.com
06/2022
"""
print("Starting app\nThis may take a few seconds...")
# Import libraries
from tkinter import *
from tkinter import filedialog
import pandas as pd
import os, glob, fitz, csv, time, itertools, shutil, sys, datetime
import ctypes

# Hide console after letting the user know the app is loading
kernel32 = ctypes.WinDLL('kernel32')
user32 = ctypes.WinDLL('user32')
SW_HIDE = 0
hWnd = kernel32.GetConsoleWindow()
user32.ShowWindow(hWnd, SW_HIDE)


WINDOW_WIDTH = 450
WINDOW_HEIGHT = 150

class Window():
    def run(self):
        self.search_init()

    def search_init(self):
        global running
        running = False
        try:
            self.master.destroy()
            running = True
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
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

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
        file_menu.add_command(label="Exit", command=self.on_closing)
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
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.master.mainloop()

    def ask_filename(self):
        self.master = Tk()
        self.master.withdraw()

        today = datetime.datetime.today().strftime('%Y-%m-%d').replace("-", "_")
        global csv_name
        csv_placeholder = "Search_report_" + today

        csv_name = filedialog.asksaveasfilename(
            defaultextension='.csv', 
            filetypes=[('CSV files', '*.csv')],
            initialdir=path,
            initialfile = csv_placeholder
        )

        if ".csv" in csv_name:
            csv_name = csv_name.replace(".csv", "")

        self.master.destroy()

    def waiting(self):
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        gif = PhotoImage(file=Window.resource_path("loading.gif"))
        Label(image=gif).grid(row=0, column=0)

        self.center(340, 340)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.master.update()
        
    def done(self, time):
        self.master.destroy()
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"

        l1 = Label(self.master, text='Search done!', bg='#499e03', font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        l2 = Label(self.master, text='Elapsed time: {:.2f} min'.format(time), bg='#499e03', font=("Arial", 14))
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
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.master.mainloop()

    def continue_search(self):
        self.master.destroy()
        self.master = Tk()
        self.master.attributes("-topmost", True)
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"

        global continue_flag
        continue_flag = True

        l1 = Label(self.master, text='Search again?', bg='#499e03', font=("Arial", 14))
        l1.grid(row=0, columnspan=2)
        
        b1 = Button(self.master, text="Search!", font=("Arial", 12), command=self.master.destroy)
        b1.grid(row=1, column=0)
        b2 = Button(self.master, text="Exit", font=("Arial", 12), command=self.on_closing)
        b2.grid(row=1, column=1)

        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)

        self.center(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.master.mainloop()

    def get_path(self):
        global path, folder, running
        path = str(self.e1.get())
        if path == "":
            path = filedialog.askdirectory()
        if path == "":
            self.master.destroy()
            return self.search_init()

        folder = glob.glob(path + "/*.pdf")

        if running:
            return self.get_params()

        self.master.destroy()

        if path == "":
            self.search_init()

    def get_params(self):
        global words_lookup
        words_lookup = []
        user_input = str(self.e1.get())
        if user_input != "":
            words_lookup= user_input.split(";")
        else:
            param_file = filedialog.askopenfilename(
                filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')],
                initialdir=path
            )

            if param_file == "":
                self.master.destroy()
                return self.params_window()

            if ".xlsx" in param_file:
                df = pd.DataFrame(pd.read_excel(param_file))
                words_lookup = df['Parametro'].tolist()
            elif ".csv" in param_file:
                df = pd.DataFrame(pd.read_csv(param_file))
                words_lookup = df['Parametro'].tolist()
            
        for i in range(len(words_lookup)):
            words_lookup[i] = words_lookup[i].strip().upper()
            
        self.master.destroy()

    def open_report(self):
        try:
            os.system("start EXCEL.EXE " + csv_name + ".csv")
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

    def on_closing(self):
        self.master.destroy()
        sys.exit()


class Critical_Data():
    # Lookup critical data
    def search_data(self, pdf):
        # Create a copy of original file
        os.makedirs(path + "\\highlighted_pdf\\", exist_ok=True)
        pdf_name = os.path.basename(pdf).replace(".pdf", "")
        searched_pdf = path + "\\highlighted_pdf\\" + pdf_name + "_" + os.path.basename(csv_name) + ".pdf"
        shutil.copyfile(pdf, searched_pdf)
        # Extract text and search parameter
        doc = fitz.open(searched_pdf)
        num_pag = doc.page_count
        for i in range(num_pag):
            page = doc.load_page(i)
            for param in words_lookup:
                # Highlight text if found
                quads = page.search_for(param, quads=True)
                # Add to csv where the text was found
                if quads:
                    data.append(['=HYPERLINK("' + searched_pdf + '","' + os.path.basename(searched_pdf) + '")', param, "Page: " + str(i + 1)])

                page.add_highlight_annot(quads)
        # Save highlighted pdf
        doc.save(searched_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

        return

    # Open excel file and append data
    def WriteCSV(self, data):
        with open(csv_name + ".csv", 'a+', encoding="utf-8") as csv_doc:
            csv_doc = csv.writer(csv_doc)
            csv_doc.writerows(data)

            return


# # # #  M A I N  F U N C T I O N  # # # #

if __name__ == "__main__":
    # Init classes
    app = Window()
    crit = Critical_Data()
    app.run()

    while True:
        global data
        data = []
        
        # Start csv file
        headers = ['Document', 'Parameter', 'Page']
        data.append(headers)

        start_time = time.time()

        # Ask for parameters to search
        app.params_window()
        # Ask name for csv
        app.ask_filename()
        # Waiting window
        app.waiting()
        
        for file_ in folder:
            crit.search_data(file_)

        end_time = time.time() - start_time
        end_time = end_time/60

        # Delete previous report if exists, without this data will be appended
        try:
            os.remove(csv_name + ".csv")
        except:
            pass

        # Drop duplicates
        data = list(data for data, _ in itertools.groupby(data))
        crit.WriteCSV(data)

        # Remove empty rows in excel file
        df = pd.read_csv(csv_name + ".csv", encoding="utf-8")
        # Droping the empty rows
        modifiedDF = df.dropna()
        # Saving it to the csv file 
        modifiedDF.to_csv(csv_name + ".csv", index=False)

        # Show done window
        app.done(end_time)

        print("\nSearch finished in {} files.\nElapsed time: {:.2f} min".format(len(folder), end_time))