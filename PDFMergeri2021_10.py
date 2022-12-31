# PDF yhdistäjä J.M
try:
    import tkinter as t
    from tkinter import messagebox  # Had to import explicitly
    from tkinter import filedialog  # Had to import explicitly
    from PyPDF2 import PdfFileMerger
    # import pandas as pd  # For later development
    import datetime
    from PIL import Image
    import os
    from imghdr import what
except ImportError as I:
    exit()

# print(I)

# Global variables

global PDFS
PDFS = []
global PDFS2
PDFS2 = []
global EXCELIT
EXCELIT = []
global AIKA

AIKA = datetime.datetime.now().strftime("%I%M%S%p%B%d%Y")


def main():

    windo = t.Tk()

    windo.geometry("930x360")
    b = Mergergui(master=windo)
    # Adding icon to the GUI window
    windo.iconbitmap(default="fileicopdfmerg.ico")
    '''Commented codeblocks of the code under progress.'''
    windo.mainloop()


class Mergergui(t.Frame, t.Menu):

    def __init__(self, master=None):

        global lista  # Listbox

        super().__init__(master)

        self.master = master

        self.master.title("PDFMergeri")

        self.pack(fill="both", expand=1)
        self.menubar = t.Menu(self)
        # Had to use lambda, because couldn't call a method normally
        self.menubar.add_command(
            label="Help/Usage", command=lambda: self.help())
        self.menubar.add_command(label="Quit", command=lambda: self.quit())

        self.master.config(menu=self.menubar)
        self.scrollbar = t.Scrollbar(self, orient="vertical")
        self.scrollbar2 = t.Scrollbar(self, orient="horizontal")

        self.lista = t.Listbox(self, width=130, height=20, relief="sunken", borderwidth=5,
                               yscrollcommand=self.scrollbar.set, xscrollcommand=self.scrollbar2.set)
        self.lista.pack(side="left", fill="both")

        self.AddButton1 = t.Button(
            self, text="Clear filelist", command=self.clear_list)
        self.AddButton1.place(x=790, y=160)

        # self.AddButton4 = t.Button(self, text="Add Excels", command=self.excelmuuntaja)
        # self.AddButton4.place(x=790, y=240)

        self.labeli2 = t.Label(
            self, text="This version\n only supports \n jpg, png and \n most pdf \nfiles.")
        self.labeli2.place(x=850, y=90, anchor="center")
        # self.labeli2.bind(
        # '<Button-1>', lambda x: webbrowser.open_new(""))

        self.AddButton2 = t.Button(
            self, text="Add PDFs", command=self.files)
        self.AddButton2.place(x=790, y=270)

        self.AddButton3 = t.Button(
            self, text="Convert images to PDF", command=self.avaakuvat)
        self.AddButton3.place(x=790, y=200)

        self.mergeButton = t.Button(
            self, text="Merge", command=self.mergeri)
        self.mergeButton.place(x=790, y=300)

        self.quitButton = t.Button(self, text="Exit", command=self.quit)
        self.quitButton.place(x=790, y=333)
        '''There is some unnecessary code.
            It is there just in case I wanna do something different, or use in a another program'''

    def files(self):
        try:
            self.lista.delete(0, 'end')

            file_path = filedialog.askopenfilenames()

            for file in file_path:
                PDFS.append(file)
            for c in PDFS:

                self.lista.insert(t.END, c)
        except Exception as fileex:
            messagebox.showinfo("Error", "Something went wrong! Try Again")
            self.lista.delete(0, 'end')
            # print(fileex)
            return 0

    def mergeri(self):
        '''try:
        self.lista.delete(0, 'end')
        _a_1 = PDFS
        _b_2 = PDFS2
        _c_3 = listaksi
        listakokoelma = _a_1, _b_2, _c_3
        if all(listakokoelma):
        del PDFS[:]
        del PDFS2[:]
        del listaksi[:]
        self.lista.insert(t.END, "----------List empty----------" +
        "\n" + "PDFMergeri1.0")
        return 0'''

        try:
            if PDFS:
                merger = PdfFileMerger()

                for pdf in PDFS:
                    merger.append(open(pdf, 'rb'))

                result = t.filedialog.asksaveasfile(
                    defaultextension=".pdf", filetypes=[("pdf files", '*.pdf')])
                with open(result.name, 'wb') as fout:
                    merger.write(fout)

                messagebox.showinfo(
                    "Info", "All PDF files have been merged successfully, and have been saved!'")
                self.lista.delete(0, 'end')
                del PDFS[:]
                # del listaksi[:]
                del PDFS2[:]
                return 0
            else:
                messagebox.showinfo("Error", "You must choose files first")
                return 0
        except Exception as mergeri_error:
            messagebox.showinfo("Error", "Wrong filetype or no files chosen!")
            self.lista.delete(0, 'end')
            del PDFS[:]
            del PDFS2[:]
            print(mergeri_error)
            # del listaksi[:]
            return 0

            '''if listaksi:

                        excels = [pd.ExcelFile(name)
                                                    for name in listaksi]
                        frames1 = [x.parse(x.sheet_names[0], header=None, index_col=None)
                                        for x in excels]
                        frames1[1:] = [ekseli[1:]
                            for ekseli in frames1[1:]]
                        combined = pd.concat(frames1)

                        combined.to_excel(
                            'result' + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".xlsx", header=False, index=False)
                        self.lista.delete(0, 'end')

                        EXCELIT2 = ()

                        messagebox.showinfo(
                            "Info", "All spreadsheet files have been merged successfully, and have been saved!'")
                        del listaksi[:]
                        del PDFS[:]
                        del PDFS2[:]
                        self.lista.delete(0, 'end')
                        return 0'''

        '''except Exception:
                messagebox.showinfo("Error", "Wrong filetype or no files chosen!")
                self.lista.delete(0, 'end')
                del PDFS[:]
                del PDFS2[:]
                del listaksi[:]
                return 0'''

    def mergeri_2(self):

        try:
            if PDFS2:

                merger = PdfFileMerger()

                for pdf in PDFS2:
                    merger.append(open(pdf, 'rb'))

                result = t.filedialog.asksaveasfile(
                    defaultextension=".pdf", filetypes=[("pdf files", '*.pdf')])

                with open(result.name, 'wb') as fout:
                    merger.write(fout)

                messagebox.showinfo(
                    "Info", "All image files converted to PDF files! A merged result has been made")

                self.lista.delete(0, 'end')
                del PDFS2[:]
                del PDFS[:]
                # del listaksi[:]
                return 0
            else:
                messagebox.showinfo("Error", "No files chosen")
                return 0

        except Exception as mergeri_2_error:
            messagebox.showinfo("Error", "Wrong filetype or no files chosen!")
            self.lista.delete(0, 'end')
            print(mergeri_2_error)
            del PDFS[:]
            del PDFS2[:]
            # del listaksi[:]

    def avaakuvat(self):
        try:
            global kuvat2
            if not PDFS or PDFS2:
                kuvat = filedialog.askopenfilenames()

                kuvat2 = self.tk.splitlist(kuvat)
                for c in kuvat2:

                    self.lista.insert(t.END, c)
            else:
                messagebox.showinfo(
                    "You have unmerged PDF files on your list. Clear the list first.")
                del PDFS[:]
                del PDFS2[:]
                # self.lista.delete(0, 'end')

                return 0

            laskuri2 = 0

            for formaatti in kuvat2:

                laskuri1 = 0

                muoto = what(formaatti)

                if muoto == "jpeg":

                    laskuri2 = laskuri2 + 1
                    im = Image.open(formaatti)

                    if im.mode == "RBGA":
                        im = im.convert("RBG")
                    newfile = ((formaatti) + str(laskuri2) + ".pdf")
                    cwd = os.getcwd()
                    im.save(newfile, "PDF", resolution=100.0)

                    tiedostonimi = (str(cwd) + "\\" + str(laskuri2) + ".pdf")

                    if os.path.exists(tiedostonimi):
                        uusi = os.replace(newfile, str(laskuri2) + "_" +
                                          datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf")
                        max_mtime = 0

                        # This loop was copied from stackoverflow
                        for dirname, subdirs, files in os.walk("."):
                            for fname in files:
                                full_path = os.path.join(dirname, fname)
                                mtime = os.stat(full_path).st_mtime
                                if mtime > max_mtime:
                                    max_mtime = mtime
                                    max_dir = dirname
                                    max_file = fname

                        tiedosto = max_file

                        PDFS2.append(tiedosto)

                    else:
                        uusi = os.rename(newfile, str(laskuri2) + "_" +
                                         datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".pdf")
                        max_mtime = 0

                        # This loop was copied from stackoverflow
                        for dirname, subdirs, files in os.walk("."):
                            for fname in files:
                                full_path = os.path.join(dirname, fname)
                                mtime = os.stat(full_path).st_mtime
                                if mtime > max_mtime:
                                    max_mtime = mtime
                                    max_dir = dirname
                                    max_file = fname

                        tiedosto_2 = max_file
                        PDFS2.append(tiedosto_2)

                    laskuri1 = laskuri1 + 1

                elif muoto == "png":

                    cwd = os.getcwd()
                    tiedostonimi = (str(cwd) + "\\" + str(laskuri2) + ".pdf")

                    laskuri2 = laskuri2 + 1
                    im = Image.open(formaatti)

                    ima3 = im.convert('RGB').save('kuva' + ".jpg")
                    PNG_FILE = ((formaatti) + str(laskuri2) + ".pdf")
                    ima4 = Image.open('kuva.jpg')

                    ima4.save(PNG_FILE, "PDF", resolution=100.0)

                    if os.path.exists(tiedostonimi):

                        PNG_FILE = ((formaatti) + str(laskuri2) + ".pdf")
                        uusi_2 = os.replace(PNG_FILE, str(
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

                        tiedosto_3 = max_file
                        PDFS2.append(tiedosto_3)

                    else:

                        PNG_FILE = ((formaatti) + str(laskuri2) + ".pdf")
                        uusi_2 = os.replace(PNG_FILE, str(
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

                        tiedosto_3 = max_file
                        PDFS2.append(tiedosto_3)

                else:
                    messagebox.showinfo("Warning", "No image format recognized")
                    kuvat2 = ("")
                    self.lista.delete(0, 'end')
                    return 0

            self.mergeri_2()

            self.lista.delete(0, 'end')
        except Exception:
            messagebox.showinfo("Error", "Something went wrong! Try Again")
            return 0

    def clear_list(self):

        try:
            self.lista.delete(0, 'end')
            a_1 = PDFS
            b_2 = PDFS2
            # c_3 = listaksi
            listakokoelma = a_1, b_2
            if any(listakokoelma):
                del PDFS[:]
                del PDFS2[:]

                self.lista.insert(t.END, "----------List empty----------" +
                                  "\n" + "PDFMergeri1.0")

                '''If all and any fails somehow continue -->'''

            elif PDFS:
                del PDFS[:]
                self.lista.insert(t.END, "----------List empty----------" +
                                  "\n" + "PDFMergeri1.0")
            elif PDFS2:
                del PDFS2[:]
                self.lista.insert(t.END, "----------List empty----------" +
                                  "\n" + "PDFMergeri1.0")
            else:
                messagebox.showinfo("Info", "List already empty!")
                self.lista.insert(t.END, "----------List empty----------" +
                                  "\n" + "PDFMergeri1.0")
                '''elif listaksi:
                        # del listaksi[:]
                        self.lista.insert(t.END, "----------List empty----------" +
                                          "\n" + "PDFMergeri1.0")

                    else:
                        messagebox.showinfo("Info", "List already empty!")
                        self.lista.insert(t.END, "----------List empty----------" +
                                          "\n" + "PDFMergeri1.0")'''
        except Exception:
            messagebox.showinfo("Error", "Something went wrong")
            return 0

    def help(self):
        messagebox.showinfo("Help", "Adding PDF files for merging: You can select multiple files by pressing down CTRL and selecting the files, or if you forget to select a file you can still press Add PDFs to add them. After adding the files press the Merge button. \n When converting images to PDF, you can choose multiple files by pressing down CTRL and choose the images. After choosing and opening the images mergeri will automatically convert them to separate PDF files. Merged files will be saved in the folder of your choice")
        return 0

    def quit(self):

        self.master.destroy()

    # messagebox.showinfo("Something went wrong! Exiting..")
    # Mergergui.master.destroy()

    '''def excelmuuntaja(self):

            try:
                if not PDFS or PDFS2:

                    # self.lista.delete(0, 'end')

                    EXCELIT2 = t.filedialog.askopenfilenames()
                    # EXCELIT2 = self.tk.splitlist(file_path)

                    listaksi = list(EXCELIT2)

                    for file in listaksi:
                        ekseli = pd.read_excel(file)
                        # all_data = EXCELIT.append(ekseli)

                    for c in EXCELIT2:
                        self.lista.insert(t.END, c)

                else:
                    del PDFS[:]
                    del PDFS2[:]
                    # self.lista.delete(0, 'end')
                    mergeri(listaksi)

                    return listaksi

            except Exception:
                messagebox.showinfo("Error", "Something went wrong! Try Again")
                return 0'''

    '''def excelmuuntaja2(self):
                try:
                    excels = [pd.ExcelFile(name) for name in listaksi]
                    frames1 = [x.parse(x.sheet_names[0], header=None, index_col=None)
                                    for x in excels]
                    frames1[1:] = [ekseli[1:] for ekseli in frames1[1:]]
                    combined = pd.concat(frames1)

                    combined.to_excel(
                        'result' + "_" + datetime.datetime.now().strftime("%I%M%S%p%B%d%Y") + ".xlsx", header=False, index=False)
                    self.lista.delete(0, 'end')

                    EXCELIT2 = ()
                except Exception:
                    messagebox.showinfo("Error", "Something went wrong")
                    return 0'''


if __name__ == "__main__":
    main()
