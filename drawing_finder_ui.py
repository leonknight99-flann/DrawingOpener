import tkinter as tk
from tkinter import messagebox as mb
import os
import sys
import subprocess

drawingPath = '\\\\Filesrv\\Drawings\\PROD\\'
enquiryPath = '\\\\Filesrv\\CustomerEnquiries\\'

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class MainApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs) 
        self.title("Drawing Opener")
        self.geometry("450x180")
        self.resizable(False, False)
        self.iconbitmap(resource_path('FlannMicrowave.ico'))

        self.drawingNumber = tk.StringVar()
        self.drawingNumberHistory = ['empty']*5
        self.enquiryNumber = tk.StringVar()

        self.transparencyBool = tk.IntVar()
        
        self.menuBar = tk.Menu(self)
        self.config(menu=self.menuBar)
        
        self.fileMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label="File", menu=self.fileMenu)
        self.fileMenu.add_command(label='Exit', command=self.exit_program)

        self.historyMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label="History", menu=self.historyMenu)
        for name in self.drawingNumberHistory:
            self.historyMenu.add_command(label=name, command=None)

        self.rfMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label='RF Tools', menu=self.rfMenu)
        self.rfMenu.add_checkbutton(label='Transparency', variable=self.transparencyBool, onvalue=1, offvalue=0, command=lambda: self.transparency_menu())
        self.rfMenu.add_command(label='Waveguide calculator', command=lambda: self.coming_soon())
        self.rfMenu.add_command(label='VSWR calculator', command=lambda: self.coming_soon())

        self.helpMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label="Help", menu=self.helpMenu)
        self.helpMenu.add_command(label='Help', command=lambda: self.help_menu())
        self.helpMenu.add_command(label='About', command=lambda: self.about_menu())

        self.introTextDrawing = tk.Label(text="Enter drawing number:")
        self.introTextDrawing.grid(row=0, column=0, columnspan=1, padx=10, pady=5)

        self.expressionFieldDrawing = tk.Entry(textvariable=self.drawingNumber)
        self.expressionFieldDrawing.grid(row=0, column=1, columnspan=6, padx=0, pady=5)

        self.openButtonDrawing = tk.Button(text='Open', command=lambda: [self.open_drawing(self.drawingNumber.get()), 
                                                                  self.clear_text_entry()])
        self.openButtonDrawing.grid(row=0,column=7, padx=0, pady=5)

        self.closeButtonDrawing = tk.Button(text='Close All', command=lambda: [subprocess.call('taskkill /f /im InventorView.exe', creationflags=subprocess.CREATE_NO_WINDOW), 
                                                                        subprocess.call('taskkill /f /im dwgviewr.exe', creationflags=subprocess.CREATE_NO_WINDOW), 
                                                                        self.clear_message_box()])
        self.closeButtonDrawing.grid(row=0,column=8, padx=0, pady=5)

        self.introTextEnquiries = tk.Label(text="Enter enquiry number:")
        self.introTextEnquiries.grid(row=1, column=0, columnspan=1, padx=0, pady=0)

        self.expressionFieldEnquiries = tk.Entry(textvariable=self.enquiryNumber)
        self.expressionFieldEnquiries.grid(row=1, column=1, columnspan=6, padx=0, pady=0)

        self.openButtonEnquries = tk.Button(text='Open', command=lambda: [self.open_enquiry(self.enquiryNumber.get()), 
                                                                  self.clear_text_entry()])
        self.openButtonEnquries.grid(row=1,column=7, padx=0, pady=0)

        self.openFolderBool = tk.IntVar()
        self.openFolderCheckbox = tk.Checkbutton(text='Open Folder', variable=self.openFolderBool, onvalue=1, offvalue=0)
        self.openFolderCheckbox.grid(row=2,column=8, columnspan=2, padx=0, pady=5)

        self.openDrawingABool = tk.IntVar(value=True)
        self.openDrawingCBool = tk.IntVar()
        self.openDrawingRBool = tk.IntVar()

        self.openDrawingTypeText = tk.Label(text="Open Drawing Type:")
        self.openDrawingTypeText.grid(row=2,column=0, padx=10)

        self.openDrawingACheckbox = tk.Checkbutton(text="A", variable=self.openDrawingABool, onvalue=1, offvalue=0)
        self.openDrawingACheckbox.grid(row=2,column=1, columnspan=1, padx=0, pady=5)
        self.openDrawingCCheckbox = tk.Checkbutton(text="C", variable=self.openDrawingCBool, onvalue=1, offvalue=0)
        self.openDrawingCCheckbox.grid(row=2,column=2, columnspan=1, padx=0, pady=5)
        self.openDrawingRCheckbox = tk.Checkbutton(text="R", variable=self.openDrawingRBool, onvalue=1, offvalue=0)
        self.openDrawingRCheckbox.grid(row=2,column=3, columnspan=1, padx=0, pady=5)

        self.messageBox = tk.Text(height=2, width=50)
        self.messageBoxScrollBar = tk.Scrollbar(command=self.messageBox.yview, orient="vertical")
        self.messageBox.configure(yscrollcommand=self.messageBoxScrollBar.set)
        self.messageBox.grid(row=3, column=0, columnspan=10, padx=10, pady=0)
        self.messageBoxScrollBar.grid(row=3, column=11, sticky='ns')

        self.bind('<Return>', lambda event=None: self.openButtonDrawing.invoke())

    def exit_program(self):
        sys.exit()

    def clear_text_entry(self):
        self.expressionFieldDrawing.delete(0, tk.END)
        self.expressionFieldEnquiries.delete(0, tk.END)

    def clear_message_box(self):
        self.messageBox.delete('1.0', tk.END)

    def transparency_menu(self):
        if self.transparencyBool.get():
            self.attributes("-alpha", 0.6)

        if not self.transparencyBool.get():
            self.attributes("-alpha", 1)

    def help_menu(self):
        mb.showinfo('Help', "Insert the number of the drawing without any prefix 0s.\nFor example 00047421 enter as 47421\n\n"
                    "This will then try to open the .idw file, failing that a .dwg file, and failing that the file explorer folder\n\n"
                    "The 'Close All' button closes all .idw and .dwg files instantly by the windows task kill process - to help clear up the desktop if lots of drawings have been opened\n\n"
                    "The 'Open Folder' checkbox will open the folder that the drawing is located in as well as the drawing\n\n"
                    "Please report any bugs and suggest any ideas to Leon")

    def about_menu(self):
        mb.showinfo('About', "Leon's drawing opener\nVersion: 1.2.1")

    def coming_soon(self):
        mb.showinfo('Message','Feature coming soon')

    # def open_waveguide_calculator(self):
    #     waveguide_calculator = WaveguideCalculatorPage()
    #     waveguide_calculator.grab_set()

    def filter_file_names(self, fileList, letterType):
        filteredList = []
        filteredList = filteredList + list(filter(lambda d: '.idw' in d, fileList))
        filteredList = filteredList + list(filter(lambda d: '.dwg' in d, fileList))
        filteredList = filteredList + list(filter(lambda d: '.tif' in d, fileList))
        filteredList = filteredList + list(filter(lambda d: '.png' in d, fileList))
        if not filteredList:
            mb.showerror('Error', f'No drawing of letter {letterType} found.')
            return
        return filteredList[0]


    def open_drawing(self, drawingNumber):
        while True:
            if not drawingNumber.isdigit():
                mb.showerror('Error', 'Invalid number entered')
                break
            
            if not len(drawingNumber) == 8:
                drawingNumber = (8-len(drawingNumber))*'0' + drawingNumber 

            self.historyMenu.add_command(label=drawingNumber, command=lambda: self.open_drawing(drawingNumber))
            self.historyMenu.delete(0)

            A = 'P' + drawingNumber[:4] + '\\'
            B = 'P' + drawingNumber[4:6] + '\\'

            dirList = os.listdir(drawingPath + A + B)

            drawingFiles = list(filter(lambda d: drawingNumber in d, dirList))
            self.messageBox.insert(tk.END, f'{drawingFiles}')
            self.messageBox.see(tk.END)
            drawingFileList = []
            if self.openDrawingABool.get():
                drawingFileA = list(filter(lambda d: '_A' in d, drawingFiles))
                drawingFileList.append(self.filter_file_names(drawingFileA, 'A'))
            if self.openDrawingCBool.get():
                drawingFileC = list(filter(lambda d: '_C' in d, drawingFiles))
                drawingFileList.append(self.filter_file_names(drawingFileC, 'C'))
            if self.openDrawingRBool.get():
                drawingFileR = list(filter(lambda d: '_R' in d, drawingFiles))
                drawingFileList.append(self.filter_file_names(drawingFileR, 'R'))

            if drawingFileList == [None]:
                self.messageBox.insert(tk.END, f'\nNo A C R drawings: {drawingFiles[0]}')
                os.startfile(drawingPath + A + B + drawingFiles[0])
                break

            for i in drawingFileList:    
                try:
                    os.startfile(drawingPath + A + B + i)
                    self.messageBox.insert(tk.END, f'\n{i}')
                except Exception:
                    self.messageBox.insert(tk.END, f'\nError in opening file {i} complete manually or re-run unselecting type')
                    self.messageBox.see(tk.END)
                    break
            if self.openFolderBool.get():
                os.startfile(drawingPath + A + B)
            self.messageBox.insert(tk.END, '\n\n')
            break

    def open_enquiry(self, enquiryNumber):
        while True:
            if not enquiryNumber.isdigit() or len(enquiryNumber) != 9:
                mb.showerror('Error', 'Invalid number entered')
                break
            self.messageBox.insert(tk.END, f'Opening enquiry: {enquiryNumber}')
            try:
                os.startfile(enquiryPath + enquiryNumber)
            except Exception:
                self.messageBox.insert(tk.END, '\nError in opening enquiry - check number')
                self.messageBox.see(tk.END)
                break
            break
            

# class WaveguideCalculatorPage(tk.Toplevel):
#     def __init__(self, *args, **kwargs):
#         super().__init__(*args, **kwargs) 
#         self.title("Waveguide Calculator")
#         self.geometry("450x120")
#         self.resizable(False, False)
#         self.iconbitmap(resource_path('images\\FlannMicrowave.ico'))
        
#         self.label = tk.Label(text="Coming Soon!")
#         self.label.grid(row=0, column=0)
      

if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
