import tkinter as tk
from tkinter import messagebox as mb
from tkinter import ttk
import os
import sys
import subprocess
import math
import pyodbc

drawingPath = '\\\\Filesrv\\Drawings\\PROD\\'
enquiryPath = '\\\\Filesrv\\CustomerEnquiries\\'

mydb = pyodbc.connect("DRIVER={SQL Server};SERVER=SQLSRV22;DATABASE=EngAdmin;UID=FLUser;PWD=MelonBall", readonly=True)
mydb_cursor = mydb.cursor()

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class MainApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs) 
        self.title("Drawing Opener")
        self.geometry("450x165")
        self.resizable(False, False)
        self.iconbitmap(resource_path('FlannMicrowave.ico'))

        self.drawingNumber = tk.StringVar()
        self.drawingNumberHistory = ['empty']*5
        self.enquiryNumber = tk.StringVar()
        
        self.sub_app_waveguide = None
        self.sub_app_vswr = None
        self.transparencyBool = tk.IntVar()
        self.openFolderBool = tk.IntVar()
        self.dcnCheckBool = tk.IntVar(value=True)

        self.openDrawingABool = tk.IntVar(value=True)
        self.openDrawingCBool = tk.IntVar()
        self.openDrawingRBool = tk.IntVar()
        
        self.menuBar = tk.Menu(self)
        self.config(menu=self.menuBar)
        
        self.fileMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label='File', menu=self.fileMenu)

        self.drawingTypeMenu = tk.Menu(self.fileMenu, tearoff=False)
        self.fileMenu.add_cascade(label='Drawing Type', menu=self.drawingTypeMenu)
        self.drawingTypeMenu.add_checkbutton(label="A", variable=self.openDrawingABool, onvalue=1, offvalue=0)
        self.drawingTypeMenu.add_checkbutton(label="C", variable=self.openDrawingCBool, onvalue=1, offvalue=0)
        self.drawingTypeMenu.add_checkbutton(label="R", variable=self.openDrawingRBool, onvalue=1, offvalue=0)

        self.optionMenu = tk.Menu(self.fileMenu, tearoff=False)
        self.fileMenu.add_cascade(label='Options', menu=self.optionMenu)
        self.optionMenu.add_checkbutton(label='Transparency', variable=self.transparencyBool, onvalue=1, offvalue=0, command=lambda: self.transparency_menu())
        self.optionMenu.add_checkbutton(label='Open Folder', variable=self.openFolderBool, onvalue=1, offvalue=0)
        self.optionMenu.add_checkbutton(label='DCN Check', variable=self.dcnCheckBool, onvalue=1, offvalue=0)

        self.fileMenu.add_command(label='Exit', command=self.exit_program)
        
        self.historyMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label="History", menu=self.historyMenu)
        for name in self.drawingNumberHistory:
            self.historyMenu.add_command(label=name, command=None)

        self.rfMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label='RF Tools', menu=self.rfMenu)
        self.rfMenu.add_command(label='Waveguide calculator', command=lambda: self.open_waveguide_calculator())
        self.rfMenu.add_command(label='VSWR calculator', command=lambda: self.open_vswr_calculator())

        self.helpMenu = tk.Menu(self.menuBar, tearoff=False)
        self.menuBar.add_cascade(label="Help", menu=self.helpMenu)
        self.helpMenu.add_command(label='Help', command=lambda: self.help_menu())
        self.helpMenu.add_command(label='About', command=lambda: self.about_menu())

        self.introTextDrawing = tk.Label(self, text="Enter drawing number:")
        self.introTextDrawing.grid(row=0, column=0, columnspan=1, padx=10, pady=(5,0))

        self.expressionFieldDrawing = tk.Entry(self, textvariable=self.drawingNumber)
        self.expressionFieldDrawing.grid(row=0, column=1, columnspan=6, padx=0, pady=(5,0))

        self.openButtonDrawing = tk.Button(self, text='Open', command=lambda: [self.open_drawing(self.drawingNumber.get()), 
                                                                  self.clear_text_entry()])
        self.openButtonDrawing.grid(row=0,column=8, padx=0, pady=(5,0))

        self.closeButtonDrawing = tk.Button(self, text='Close All', command=lambda: [subprocess.call('taskkill /f /im InventorView.exe', creationflags=subprocess.CREATE_NO_WINDOW), 
                                                                        subprocess.call('taskkill /f /im dwgviewr.exe', creationflags=subprocess.CREATE_NO_WINDOW), 
                                                                        self.clear_message_box()])
        self.closeButtonDrawing.grid(row=0,column=9,columnspan=2, padx=0, pady=(5,0))

        self.introInspection = tk.Label(self, text="Enter inspection partID:")
        self.introInspection.grid(row=1,column=0, padx=10)

        self.expressionFieldInspection = tk.Entry(self, state='disabled')
        self.expressionFieldInspection.grid(row=1, column=1, columnspan=6, padx=0, pady=0)

        self.openButtonInspection = tk.Button(self, text='Open', command=lambda: [self.coming_soon()])
        self.openButtonInspection.grid(row=1,column=8, padx=0, pady=0)

        self.introTextEnquiries = tk.Label(self, text="Enter enquiry number:")
        self.introTextEnquiries.grid(row=2, column=0, columnspan=1, padx=0, pady=0)

        self.expressionFieldEnquiries = tk.Entry(self, textvariable=self.enquiryNumber)
        self.expressionFieldEnquiries.grid(row=2, column=1, columnspan=6, padx=0, pady=0)

        self.openButtonEnquries = tk.Button(self, text='Open', command=lambda: [self.open_enquiry(self.enquiryNumber.get()), 
                                                                  self.clear_text_entry()])
        self.openButtonEnquries.grid(row=2,column=8, padx=0, pady=0)

        self.messageBox = tk.Text(self, height=2, width=50)
        self.messageBoxScrollBar = tk.Scrollbar(self, command=self.messageBox.yview, orient="vertical")
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
        mb.showinfo('Help', " - Insert the number of the drawing without any prefix 0s.\nFor example 00047421 enter as 47421\n\n"
                    " - This will then try to open the .idw file, failing that a .dwg file, then .tif and finally .png\n\n"
                    " - If a specific drawing is required such as C or R then use the checkboxes\n\n"
                    " - The 'Close All' button closes all .idw and .dwg files instantly by the windows task kill process - to help clear up the desktop if lots of drawings have been opened\n\n"
                    " - The 'Open Folder' checkbox will open the folder that the drawing is located\n\n"
                    " - Customer enquiry folder opens via the open button after the exact 9 digit number has been entered\n\n"
                    "Please report any bugs and suggest any ideas to Leon")

    def about_menu(self):
        mb.showinfo('About', "Leon's drawing opener\nVersion: 1.3.2")

    def coming_soon(self):
        mb.showinfo('Message','Feature coming soon')

    def open_waveguide_calculator(self):
        if self.sub_app_waveguide is None or not self.sub_app_waveguide.winfo_exists():
            self.sub_app_waveguide = WaveguideCalculatorPage(self)
        else:
            self.sub_app_waveguide.focus()

    def open_vswr_calculator(self):
        if self.sub_app_vswr is None or not self.sub_app_vswr.winfo_exists():
            self.sub_app_vswr = UnitCalculatorPage(self)
        else:
            self.sub_app_vswr.focus()

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
            
            print(self.dcn_check(drawingNumber))
            print(self.dcnCheckBool.get())
            if self.dcn_check(drawingNumber) and self.dcnCheckBool.get():
                mb.showerror('DCN', 'DCN(s) found')
            
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

    def dcn_check(self, DC_Code):
        for row in mydb_cursor.execute("select * from DCNTab where DC_CODE like ?", '%'+DC_Code+'%'):
            try:
                if row.CHANGE_NOTE_NO:
                    return True
            except Exception:
                self.messageBox.insert(tk.END, f'\nError in checking for DCN')
            return False


    def open_enquiry(self, enquiryNumber):
        while True:
            if not enquiryNumber.isdigit() or len(enquiryNumber) != 9:
                mb.showerror('Error', 'Invalid number entered')
                break
            self.messageBox.insert(tk.END, f'\nOpening enquiry: {enquiryNumber}')
            try:
                os.startfile(enquiryPath + enquiryNumber)
            except Exception:
                self.messageBox.insert(tk.END, '\nError in opening enquiry - check number')
                self.messageBox.see(tk.END)
                break
            break
            

class WaveguideCalculatorPage(tk.Toplevel):
    def __init__(self, main_app):
        self.main_app = main_app
        super().__init__(main_app)
        self.title("Waveguide Calculator")
        self.geometry("450x120")
        self.resizable(False, False)

        self.attributes("-topmost",True)

        self.iconbitmap(resource_path('FlannMicrowave.ico'))
        
        self.label = tk.Label(self, text="Coming Soon!")
        self.label.grid(row=0, column=0)
      

class UnitCalculatorPage(tk.Toplevel):
    """
    ConversionTool is a tkinter-based GUI for performing unit conversions.
    
    Supported conversions:
    - Length: km, m, cm, mm, mi, yd, ft, in
    - Loss: VSWR, RL, S parameter
    """
    def __init__(self, main_app):
        self.main_app = main_app
        super().__init__(main_app)
        self.title("Unit Calculator")
        self.geometry("350x180")
        self.resizable(False, False)

        self.attributes("-topmost",True)

        self.iconbitmap(resource_path('FlannMicrowave.ico'))

        self.conversion_fns = {
            'Length': self.length_conversion,
            'Loss': self.loss_conversion,
        }

        """
        Setup the user interface components.
        """
        options = {
            'Length': ['in', 'mm', 'cm', 'm', 'km', 'mi', 'yd', 'ft'],
            'Loss': ['VSWR', 'dB', 'S11']
        }

        self.entries = []

        for i, (name, option_values) in enumerate(options.items()):
            tk.Label(self, text=f'{name} Conversion').grid(column=i, row=0, sticky='W', padx=10, pady=5)

            from_ = ttk.Combobox(self, values=option_values)
            from_.grid(column=i, row=1, sticky='W', padx=10, pady=0)
            from_.current(0)

            to_ = ttk.Combobox(self, values=option_values)
            to_.grid(column=i, row=2, sticky='W', padx=10, pady=0)
            to_.current(1)

            entry = tk.Entry(self)
            entry.grid(column=i, row=3, sticky='W', padx=10, pady=5)

            result = tk.Label(self, text='')
            result.grid(column=i, row=4, sticky='W', padx=10, pady=5)

            self.entries.append((name, entry, result, from_, to_))

        tk.Button(self, text='Calculate', command=self.calculate).grid(column=0, row=5, sticky='W', padx=10, pady=5)


    def length_conversion(self, from_unit, to_unit, value):
        """
        Convert a length value from one unit to another.
        
        Parameters:
        from_unit (str): The original unit of the length.
        to_unit (str): The target unit for the conversion.
        value (float): The length value to convert.
        
        Returns:
        float: The converted length value.
        """
        length_dict = {'km': 1000.0, 'm': 1.0, 'cm': 0.01, 'mm': 0.001, 
                       'mi': 1609.34, 'yd': 0.9144, 'ft': 0.3048, 'in': 0.0254}
        return (value * length_dict[from_unit]) / length_dict[to_unit]
    
    def loss_conversion(self, from_unit, to_unit, value):
        if from_unit == 'VSWR':
            s11 = (value - 1) / (value + 1)
            if to_unit == 'S11':
                return s11
            elif to_unit == 'dB':
                return 20 * math.log10(s11)
        elif from_unit == 'dB':
            s11 = 10 ** (value / 20)
            if to_unit == 'S11':
                return s11
            elif to_unit == 'VSWR':
                return (1 + s11) / (1 - s11)
        elif from_unit == 'S11':
            if to_unit == 'dB':
                return 20 * math.log10(value)
            elif to_unit == 'VSWR':
                return (1 + value) / (1 - value)
        return value
    
    def calculate(self):
        """
        Perform the unit conversions based on the current entries and update the result labels.
        """
        for name, entry, result, from_, to_ in self.entries:
            value = entry.get()
            if not value:  # Skip if the entry is empty
                continue
            try:
                value = float(value)
                conversion_fn = self.conversion_fns[name]
                result['text'] = conversion_fn(from_.get(), to_.get(), value)
            except ValueError:
                tk.messagebox.showerror('Error', f'Invalid input for {name} conversion')
            

if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
