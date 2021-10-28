# PDF yhdistäjä J.M
try:
    from tkinter import *
    from PyPDF2 import PdfFileMerger
    from tkinter import filedialog
    from tkinter import messagebox
    from tkinter import ttk
    from tkinter import Menu
    import pandas as pd
    import datetime
    from random import randrange
    from PIL import Image
    import os
    from imghdr import what
    import webbrowser
except ImportError as I:
    exit()

# print(I)

# Global variables

global pdfs
pdfs = []
global pdfs2
pdfs2 = []
global excelit
excelit = []
global aika

aika = datetime.datetime.now().strftime("%I%M%S%p%B%d%Y")

try:
    def main():

        root = Tk()

        root.geometry("930x360")
        b = Mergergui(master=root)
        root.iconbitmap(default="fileicopdfmerg.ico")  # Adding icon to the GUI window
        '''Commented codeblocks of the code under progress.'''
        root.mainloop()

    class Mergergui(Frame, Menu):

        def __init__(self, master=None):

            global lista  # Listbox

            super().__init__(master)

            self.master = master

            self.master.title("PDFMergeri")

            self.pack(fill=BOTH, expand=1)
            self.menubar = Menu(self)
            # Had to use lambda, because couldn't call a method normally
            self.menubar.add_command(
                label="Help/Usage", command=lambda: self.help())
            self.menubar.add_command(label="Quit", command=lambda: self.quit())

            self.master.config(menu=self.menubar)
            self.scrollbar = Scrollbar(self, orient="vertical")
            self.scrollbar2 = Scrollbar(self, orient="horizontal")

            self.lista = Listbox(self, width=130, height=20, relief=SUNKEN, borderwidth=5,
                                 yscrollcommand=self.scrollbar.set, xscrollcommand=self.scrollbar2.set)
            self.lista.pack(side="left", fill="both")

            self.AddButton1 = ttk.Button(
                self, text="Clear filelist", command=self.clear_list)
            self.AddButton1.place(x=790, y=160)

            # self.AddButton4 = ttk.Button(self, text="Add Excels", command=self.excelmuuntaja)
            # self.AddButton4.place(x=790, y=240)

            self.labeli2 = Label(
                self, text="This version\n only supports \n jpg, png and \n most pdf \nfiles.")
            self.labeli2.place(x=850, y=90, anchor="center")
            # self.labeli2.bind(
            # '<Button-1>', lambda x: webbrowser.open_new(""))

            self.AddButton2 = ttk.Button(
                self, text="Add PDFs", command=self.files)
            self.AddButton2.place(x=790, y=270)

            self.AddButton3 = ttk.Button(
                self, text="Convert images to PDF", command=self.avaakuvat)
            self.AddButton3.place(x=790, y=200)

            self.mergeButton = ttk.Button(
                self, text="Merge", command=self.mergeri)
            self.mergeButton.place(x=790, y=300)

            self.quitButton = ttk.Button(self, text="Exit", command=self.quit)
            self.quitButton.place(x=790, y=333)
            '''There is some unnecessary use of code.
            It is there just in case I wanna do something different, or use in a another program'''

        def files(self):
            try:
                self.lista.delete(0, 'end')

                self.file_path = filedialog.askopenfilenames()

                for self.file in self.file_path:
                    pdfs.append(self.file)
                for self.c in pdfs:

                    self.lista.insert(END, self.c)
            except Exception:
                messagebox.showinfo("Error", "Something went wrong! Try Again")
                self.lista.delete(0, 'end')
                return 0

        def mergeri(self, *args):
            '''try:
            self.lista.delete(0, 'end')
            _a_1 = pdfs
            _b_2 = pdfs2
            _c_3 = self.listaksi
            listakokoelma = _a_1, _b_2, _c_3
            if all(listakokoelma):
            del pdfs[:]
            del pdfs2[:]
            del self.listaksi[:]
            self.lista.insert(END, "----------List empty----------" +
            "\n" + "PDFMergeri1.0")
            return 0'''

            try:
                if pdfs:
                    self.merger = PdfFileMerger()

                    for self.pdf in pdfs:
                        self.merger.append(open(self.pdf, 'rb'))

                    with open('result' + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf", 'wb') as fout:
                        self.merger.write(fout)

                    messagebox.showinfo(
                        "Info", "All PDF files have been merged successfully, and have been saved to current folder with the name 'result + current time.pdf!'")
                    self.lista.delete(0, 'end')
                    del pdfs[:]
                    # del self.listaksi[:]
                    del pdfs2[:]
                    return 0
                else:
                    messagebox.showinfo("Error", "You must choose files first")
                    return 0
            except Exception:
                messagebox.showinfo("Error", "Wrong filetype or no files chosen!")
                self.lista.delete(0, 'end')
                del pdfs[:]
                del pdfs2[:]
                # del self.listaksi[:]
                return 0

                '''if self.listaksi:

                    self.excels = [pd.ExcelFile(self.name)
                                                for self.name in self.listaksi]
                    self.frames1 = [x.parse(x.sheet_names[0], header=None, index_col=None)
                                    for x in self.excels]
                    self.frames1[1:] = [self.ekseli[1:]
                        for self.ekseli in self.frames1[1:]]
                    self.combined = pd.concat(self.frames1)

                    self.combined.to_excel(
                        'result' + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".xlsx", header=False, index=False)
                    self.lista.delete(0, 'end')

                    self.excelit2 = ()

                    messagebox.showinfo(
                        "Info", "All spreadsheet files have been merged successfully, and have been saved to current folder with the name 'result + current time!'")
                    del self.listaksi[:]
                    del pdfs[:]
                    del pdfs2[:]
                    self.lista.delete(0, 'end')
                    return 0'''

            '''except Exception:
            messagebox.showinfo("Error", "Wrong filetype or no files chosen!")
            self.lista.delete(0, 'end')
            del pdfs[:]
            del pdfs2[:]
            del self.listaksi[:]
            return 0'''

        def mergeri_2(self):

            try:
                if pdfs2:

                    self.merger = PdfFileMerger()

                    for self.pdf in pdfs2:
                        self.merger.append(open(self.pdf, 'rb'))

                    with open('result' + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf", 'wb') as fout:
                        self.merger.write(fout)

                    messagebox.showinfo(
                        "Info", "All image files converted to PDF files! A merged result have been made with the name 'result + current time.pdf' and saved to your current directory.")

                    self.lista.delete(0, 'end')
                    del pdfs2[:]
                    del pdfs[:]
                    # del self.listaksi[:]
                    return 0
                else:
                    messagebox.showinfo("Error", "No files chosen")
                    return 0

            except Exception:
                messagebox.showinfo("Error", "Wrong filetype or no files chosen!")
                self.lista.delete(0, 'end')
                del pdfs[:]
                del pdfs2[:]
                # del self.listaksi[:]

        def avaakuvat(self):
            try:
                global kuvat2
                if not pdfs or pdfs2:
                    self.kuvat = filedialog.askopenfilenames()

                    self.kuvat2 = self.tk.splitlist(self.kuvat)
                    for self.c in self.kuvat2:

                        self.lista.insert(END, self.c)
                else:
                    del pdfs[:]
                    del pdfs2[:]
                    self.lista.delete(0, 'end')

                    return self.avaakuvat()

                laskuri2 = 0

                for self.formaatti in self.kuvat2:

                    laskuri1 = 0

                    self.muoto = what(self.formaatti)

                    if self.muoto == "jpeg":

                        laskuri2 = laskuri2 + 1
                        self.im = Image.open(self.formaatti)

                        if self.im.mode == "RBGA":
                            self.im = self.im.convert("RBG")
                        self.newfile = ((self.formaatti) + str(laskuri2) + ".pdf")
                        self.cwd = os.getcwd()
                        self.im.save(self.newfile, "PDF", resolution=100.0)

                        self.tiedostonimi = (str(self.cwd) + "\\" + str(laskuri2) + ".pdf")

                        if os.path.exists(self.tiedostonimi):
                            self.uusi = os.replace(self.newfile, str(laskuri2) + "_" +
                                                   datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf")
                            max_mtime = 0

                            for dirname, subdirs, files in os.walk("."):
                                for fname in files:
                                    full_path = os.path.join(dirname, fname)
                                    mtime = os.stat(full_path).st_mtime
                                    if mtime > max_mtime:
                                        max_mtime = mtime
                                        max_dir = dirname
                                        max_file = fname

                            self.tiedosto = max_file

                            pdfs2.append(self.tiedosto)

                        else:
                            self.uusi = os.rename(self.newfile, str(laskuri2) + "_" +
                                                  datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf")
                            max_mtime = 0

                            for dirname, subdirs, files in os.walk("."):
                                for fname in files:
                                    full_path = os.path.join(dirname, fname)
                                    mtime = os.stat(full_path).st_mtime
                                    if mtime > max_mtime:
                                        max_mtime = mtime
                                        max_dir = dirname
                                        max_file = fname

                            self.tiedosto_2 = max_file
                            pdfs2.append(self.tiedosto_2)

                        laskuri1 = laskuri1 + 1

                    elif self.muoto == "png":

                        self.cwd = os.getcwd()
                        self.tiedostonimi = (str(self.cwd) + "\\" + str(laskuri2) + ".pdf")

                        laskuri2 = laskuri2 + 1
                        self.im = Image.open(self.formaatti)

                        self.ima3 = self.im.convert('RGB').save('kuva' + ".jpg")
                        self.PNG_FILE = ((self.formaatti) + str(laskuri2) + ".pdf")
                        self.ima4 = Image.open('kuva.jpg')

                        self.ima4.save(self.PNG_FILE, "PDF", resolution=100.0)

                        if os.path.exists(self.tiedostonimi):

                            self.PNG_FILE = ((self.formaatti) + str(laskuri2) + ".pdf")
                            self.uusi_2 = os.replace(self.PNG_FILE, str(
                                laskuri2) + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf")
                            max_mtime = 0

                            for dirname, subdirs, files in os.walk("."):
                                for fname in files:
                                    full_path = os.path.join(dirname, fname)
                                    mtime = os.stat(full_path).st_mtime
                                    if mtime > max_mtime:
                                        max_mtime = mtime
                                        max_dir = dirname
                                        max_file = fname

                            self.tiedosto_3 = max_file
                            pdfs2.append(self.tiedosto_3)

                        else:

                            self.PNG_FILE = ((self.formaatti) + str(laskuri2) + ".pdf")
                            self.uusi_2 = os.replace(self.PNG_FILE, str(
                                laskuri2) + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf")
                            max_mtime = 0

                            for dirname, subdirs, files in os.walk("."):
                                for fname in files:
                                    full_path = os.path.join(dirname, fname)
                                    mtime = os.stat(full_path).st_mtime
                                    if mtime > max_mtime:
                                        max_mtime = mtime
                                        max_dir = dirname
                                        max_file = fname

                            self.tiedosto_3 = max_file
                            pdfs2.append(self.tiedosto_3)

                    else:
                        messagebox.showinfo("Warning", "No image format recognized")
                        self.kuvat2 = ("")
                        self.lista.delete(0, 'end')
                        return 0

                self.mergeri_2()

                self.lista.delete(0, 'end')
            except Exception:
                messagebox.showinfo("Error", "Something went wrong! Try Again")
                return 0

        '''def excelmuuntaja(self):

            try:
                if not pdfs or pdfs2:

                    # self.lista.delete(0, 'end')

                    self.excelit2 = filedialog.askopenfilenames()
                    # self.excelit2 = self.tk.splitlist(self.file_path)

                    self.listaksi = list(self.excelit2)

                    for self.file in self.listaksi:
                        self.ekseli = pd.read_excel(self.file)
                        # self.all_data = excelit.append(self.ekseli)

                    for self.c in self.excelit2:
                        self.lista.insert(END, self.c)

                else:
                    del pdfs[:]
                    del pdfs2[:]
                    # self.lista.delete(0, 'end')
                    mergeri(self.listaksi)

                    return self.listaksi

            except Exception:
                messagebox.showinfo("Error", "Something went wrong! Try Again")
                return 0

        def excelmuuntaja2(self):
            try:
                self.excels = [pd.ExcelFile(self.name) for self.name in self.listaksi]
                self.frames1 = [x.parse(x.sheet_names[0], header=None, index_col=None)
                                for x in self.excels]
                self.frames1[1:] = [self.ekseli[1:] for self.ekseli in self.frames1[1:]]
                self.combined = pd.concat(self.frames1)

                self.combined.to_excel(
                    'result' + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".xlsx", header=False, index=False)
                self.lista.delete(0, 'end')

                self.excelit2 = ()
            except Exception:
                messagebox.showinfo("Error", "Something went wrong")
                return 0'''

        def clear_list(self):

            try:
                self.lista.delete(0, 'end')
                a_1 = pdfs
                b_2 = pdfs2
                # c_3 = self.listaksi
                listakokoelma = a_1, b_2
                if any(listakokoelma):
                    del pdfs[:]
                    del pdfs2[:]
                    # del self.listaksi[:]
                    self.lista.insert(END, "----------List empty----------" +
                                      "\n" + "PDFMergeri1.0")
                    # print("Test")

                    # print("Test")
                    '''If all and any fails somehow continue -->'''

                elif pdfs:
                    del pdfs[:]
                    self.lista.insert(END, "----------List empty----------" +
                                      "\n" + "PDFMergeri1.0")
                elif pdfs2:
                    del pdfs2[:]
                    self.lista.insert(END, "----------List empty----------" +
                                      "\n" + "PDFMergeri1.0")
                else:
                    messagebox.showinfo("Info", "List already empty!")
                    self.lista.insert(END, "----------List empty----------" +
                                      "\n" + "PDFMergeri1.0")
                '''elif self.listaksi:
                    # del self.listaksi[:]
                    self.lista.insert(END, "----------List empty----------" +
                                      "\n" + "PDFMergeri1.0")

                else:
                    messagebox.showinfo("Info", "List already empty!")
                    self.lista.insert(END, "----------List empty----------" +
                                      "\n" + "PDFMergeri1.0")'''
            except Exception:
                messagebox.showinfo("Error", "Something went wrong")
                return 0

        def help(self):
            messagebox.showinfo("Help", "Adding PDF files for merging: You can select multiple files by pressing down CTRL and selecting the files, or if you forget to select a file you can still press Add PDFs to add them. After adding the files press the Merge button. \n When converting images to PDF, you can choose multiple files by pressing down CTRL and choose the images. After choosing and opening the images mergeri will automatically convert them to separate PDF files. Merged files will be saved in the folder where the application is located. ")
            return 0

        def quit(self):

            self.master.destroy()

except Exception:
    messagebox.showinfo("Something went wrong! Exiting..")
    Mergergui.master.destroy()

if __name__ == "__main__":
    main()
