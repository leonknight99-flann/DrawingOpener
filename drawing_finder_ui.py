import tkinter as tk
from tkinter import messagebox as mb
import os
import sys
import subprocess

drawingPath = '\\\\Filesrv\\Drawings\\PROD\\'

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class WaveguideCalculatorApplication(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs) 
        self.title("Waveguide Calculator")
        self.geometry("450x120")
        self.resizable(False, False)
        self.iconbitmap(resource_path('images\\FlannMicrowave.ico'))
        
        self.label = tk.Label(text="Coming Soon!")
        self.label.pack()
        

class MainApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs) 
        self.title("Drawing Opener")
        self.geometry("450x180")
        self.resizable(False, False)
        self.iconbitmap(resource_path('images\\FlannMicrowave.ico'))

        self.drawingNumber = tk.StringVar()
        self.drawingNumberHistory = ['empty']*5
        
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
        self.rfMenu.add_command(label='Waveguide calculator', command=lambda: self.open_waveguide_calculator())
        self.rfMenu.add_command(label='VSWR calculator', command=lambda: self.coming_soon())
        
        self.transparencyBool = tk.IntVar()

        self.helpMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label="Help", menu=self.helpMenu)
        self.helpMenu.add_checkbutton(label='Transparency', variable=self.transparencyBool, onvalue=1, offvalue=0, command=lambda: self.transparency_menu())
        self.helpMenu.add_command(label='Help', command=lambda: self.help_menu())
        self.helpMenu.add_command(label='About', command=lambda: self.about_menu())

        self.introText = tk.Label(text="Enter drawing number:")
        self.introText.grid(row=0, column=0, columnspan=5, padx=10, pady=10)

        self.expressionField = tk.Entry(textvariable=self.drawingNumber)
        self.expressionField.grid(row=0, column=5, padx=0, pady=10)

        self.openButton = tk.Button(text='Open', command=lambda: [self.open_drawing(self.drawingNumber.get()), 
                                                                  self.clear_text_entry()])
        self.openButton.grid(row=0,column=6, padx=0, pady=10)

        self.openFolderBool = tk.IntVar()
        self.openFolderCheckbox = tk.Checkbutton(text='Open Folder', variable=self.openFolderBool, onvalue=1, offvalue=0)
        self.openFolderCheckbox.grid(row=1,column=6, columnspan=3, padx=0, pady=0)

        self.closeButton = tk.Button(text='Close All', command=lambda: [subprocess.call('taskkill /f /im InventorView.exe', creationflags=subprocess.CREATE_NO_WINDOW), 
                                                                        subprocess.call('taskkill /f /im dwgviewr.exe', creationflags=subprocess.CREATE_NO_WINDOW), 
                                                                        self.clear_message_box()])
        self.closeButton.grid(row=0,column=8, padx=0, pady=10)

        self.messageBoxText = tk.Label(text="Output Messages:")
        self.messageBoxText.grid(row=1, column=1, padx=10)
        self.messageBox = tk.Text(height=4, width=50)
        self.messageBoxScrollBar = tk.Scrollbar(command=self.messageBox.yview, orient="vertical")
        self.messageBox.configure(yscrollcommand=self.messageBoxScrollBar.set)
        self.messageBox.grid(row=2, column=1, columnspan=10, padx=10, pady=0)
        self.messageBoxScrollBar.grid(row=2, column=11, sticky='ns')

        self.bind('<Return>', lambda event=None: self.openButton.invoke())

    def exit_program(self):
        sys.exit()

    def clear_text_entry(self):
        self.expressionField.delete(0, tk.END)

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
        mb.showinfo('About', "Leon's drawing opener\nVersion: 1.2.0")

    def coming_soon(self):
        mb.showinfo('Message','Feature coming soon!')

    def open_waveguide_calculator(self):
        self.waveguide_calculator = WaveguideCalculatorApplication()
        self.waveguide_calculator.mainloop()

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
            drawingFileA = list(filter(lambda d: '_A' in d, drawingFiles))
            drawingFileIDW = list(filter(lambda d: '.idw' in d, drawingFileA))
            drawingFileA = list(filter(lambda d: not '.bak' in d, drawingFileA))
            try:
                os.startfile(drawingPath + A + B + drawingFileIDW[0])
            except Exception:
                try:
                    os.startfile(drawingPath + A + B + drawingFileA[0])  # Opens drawing 
                    self.messageBox.insert(tk.END, f'\n{drawingFileA[0]}')
                    self.messageBox.see(tk.END)
                except Exception:
                    self.messageBox.insert(tk.END, '\nError in opening files - complete manually')
                    self.messageBox.see(tk.END)
                    os.startfile(drawingPath + A + B)  # Opens file explore folder of the location of the drawing
                    break
            if self.openFolderBool.get():
                os.startfile(drawingPath + A + B)
            self.messageBox.insert(tk.END, '\n\n')
            break

        
if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
