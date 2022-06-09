""" 
Version 1.0.0
almolina@tenaris.com
09/2022
"""
# Import libraries
from threading import Thread
from tkinter import *
import pandas as pd
import os, glob, fitz, csv, time, itertools, shutil, sys

class Window():
    def run(self):
        self.master = Tk()
    
        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))
        Label(self.master, text="Files Location", bg='#499e03').grid(row=0, sticky=W)
        
        self.e1 = Entry(self.master)
        self.e1.grid(row=0, column=1, sticky=EW)
        
        b = Button(self.master, text="Select", command=self.get_path)
        b.grid(row=0, column=2,columnspan=2, rowspan=2, padx=5, pady=5)
        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        Window.center(self)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        mainloop()

    def get_path(self):
        global path, folder
        path = str(self.e1.get()) + "\\"
        folder = glob.glob(path + "*.pdf")
        self.master.destroy()
        self.params_window()
        
    def _onKeyRelease(self, event):
        ctrl  = (event.state & 0x4) != 0
        if event.keycode==88 and  ctrl and event.keysym.lower() != "x": 
            event.widget.event_generate("<<Cut>>")

        if event.keycode==86 and  ctrl and event.keysym.lower() != "v": 
            event.widget.event_generate("<<Paste>>")

        if event.keycode==67 and  ctrl and event.keysym.lower() != "c":
            event.widget.event_generate("<<Copy>>")

    def params_window(self):
        self.master = Tk()

        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        self.master.iconbitmap(default=Window.resource_path("Tenaris_Logo.ico"))
        Label(self.master, text='Parameters (separate each with a ";") :', bg='#499e03').grid(row=0, sticky=W)
        
        self.e1 = Entry(self.master)
        self.e1.grid(row=0, column=1, sticky=EW)
        
        b = Button(self.master, text="Search", command=self.get_params)
        b.grid(row=0, column=2,columnspan=2, rowspan=2, padx=5, pady=5)
        self.master.bind_all("<Key>", self._onKeyRelease, "+")
        Window.center(self)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        mainloop()

    def get_params(self):
        global words_lookup
        user_input = str(self.e1.get())
        words_lookup= user_input.split(";")
        for i in range(len(words_lookup)):
            words_lookup[i] = words_lookup[i].strip().upper()
        
        self.master.destroy()

    def waiting(self):
        self.master = Tk()

        self.master.title("Critical Data")
        self.master['background'] = "#499e03"
        gif = PhotoImage(file=Window.resource_path("loading.gif"))
        Label(image=gif).grid(row=0, column=0)
        Window.center(self)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.master.update()
        
    def done(self, time):
        self.master.destroy()
        self.master = Tk()

        self.master.title("Critical Data")
        self.master['background'] = "#499e03"

        Label(self.master, text='Search done!', bg='#499e03').grid(row=0, sticky=EW)
        Label(self.master, text='Elapsed time: {:.2f} min'.format(time), bg='#499e03').grid(row=1, sticky=EW)
        
        b1 = Button(self.master, text="Open report", command=self.open_report)
        b1.grid(row=2, column=0, padx=5, pady=5)
        b2 = Button(self.master, text="Close", command=self.master.destroy)
        b2.grid(row=2, column=1, padx=5, pady=5)
        Window.center(self)
        self.master.protocol("WM_DELETE_WINDOW", self.master.destroy)

        mainloop()

    def open_report(self):
        global report
        report = True
        self.master.destroy()

    def center(self):
        w = self.master.winfo_reqwidth()
        h = self.master.winfo_reqheight()
        ws = self.master.winfo_screenwidth()
        hs = self.master.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.master.geometry('+%d+%d' % (x, y))

    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def on_closing(self):
        self.master.destroy()
        exit()


class Critical_Data():
    # Lookup critical data
    def search_data(self, pdf):
        # Create a copy of original file
        os.makedirs(path + "highlighted_pdf\\", exist_ok=True)
        shutil.copyfile(pdf, path + "highlighted_pdf\\" + os.path.basename(pdf))
        searched_pdf = path + "highlighted_pdf\\" + os.path.basename(pdf)
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
        with open(path + "Search_report.csv", 'a+', encoding="utf-8") as csv_doc:
            csv_doc = csv.writer(csv_doc)
            csv_doc.writerows(data)

            return


# # # #  M A I N  F U N C T I O N  # # # #

if __name__ == "__main__":
    # Init classes
    app = Window()
    crit = Critical_Data()

    app.run()
    global data
    data = []
    report = False
    # Start csv file
    headers = ['Document', 'Parameter', 'Page']
    data.append(headers)

    t = []
    a = 0

    start_time = time.time()

    # Waiting window
    app.waiting()

    for file_ in folder:
        t.append(Thread(target=crit.search_data, args=(file_,)))
        t[a].start()
        a += 1       

    for th in t:
        th.join()

    end_time = time.time() - start_time
    end_time = end_time/60

    # Delete previous report if exists
    try:
        os.remove(path + "Search_report.csv")
    except:
        pass

    # Show done window
    app.done(end_time)

    # Open excel report
    if report:
        try:
            os.system("start EXCEL.EXE " + path + "Search_report.csv")
        except Exception as e:
            print(e.args)

    # Drop duplicates
    data = list(data for data, _ in itertools.groupby(data))
    crit.WriteCSV(data)

    # Remove empty rows in excel file
    df = pd.read_csv(path + "Search_report.csv", encoding="utf-8")
    # Droping the empty rows
    modifiedDF = df.dropna()
    # Saving it to the csv file 
    modifiedDF.to_csv(path + "Search_report.csv", index=False)

    print("\nSearch finished in {} files.\nElapsed time: {:.2f} min".format(len(folder), end_time))