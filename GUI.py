# -*- coding: utf-8 -*-
"""
Created on Sat Jul 18 23:35:04 2020

@author: wwink
"""

import tkinter as tk
from tkinter.filedialog import askopenfilename
import processor as prc


class MainApplication(tk.Frame):

    
    def open_file(self):
        """Open a file for editing."""
        self.filepath = askopenfilename(
            filetypes=[("Excel file", "*.xlsx"), ("All Files", "*.*")]
        )
        if not self.filepath:
            return
        
        self.window.title(f"Localize Excel - {self.filepath}")
    #Check if entered value is parsable to float
    def ent_num_validation(self, in_str, act_typ):
        if act_typ == '1': #insert
            try:
                float(in_str)
                return True
            except ValueError:
                return False
        return True
        
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)

        self.window = parent
        
        self.window.title("Localization Keys")
        self.window.rowconfigure(1, minsize=100, weight=1)
        self.window.columnconfigure(0, minsize=100, weight=1)
        self.filepath = None
        
        excel_processor = prc.ExcelProcessor()
        
        #Left frame
        fr_top = tk.Frame(self.window, bd=2)
        #%%Open button
        btn_open = tk.Button(fr_top, text="Open", command = self.open_file)
        btn_open.grid(row=0, column=0, sticky="ew")
        
        #%%bottom left frame for text and label
        fr_ver_entry = tk.Frame(fr_top)
        fr_ver_entry.grid(row = 1, column = 0)
       
        #%%Text entry box
        ent_version = tk.Entry(fr_ver_entry, validate="key", validatecommand = (self.register(self.ent_num_validation),'%P','%d'))        
        ent_version.grid(row = 0, column=1, sticky="ew")
        
        #%%Label text entry box
        lab_version = tk.Label(fr_ver_entry,text="Version:")
        lab_version.grid(row = 0, column = 0, sticky="ew")
        #%%Check box
        var_be = tk.BooleanVar()
        var_fe = tk.BooleanVar()
        
        check_be = tk.Checkbutton(fr_top, text="BE", variable=var_be)
        check_fe = tk.Checkbutton(fr_top, text="FE", variable=var_fe)
        
        check_be.grid(row=0, column=1, sticky="ew")
        check_fe.grid(row=1, column=1, sticky="ew")
        #%%Top Frame
        fr_top.grid(row = 0)
        
        #%% Execute buttin
        btn_exec = tk.Button(self.window, text = "Execute")
        btn_exec.grid(row = 1, sticky = "n")
        btn_exec.bind("<Button-1>", excel_processor.set_attr(self.filepath, ent_version.get()))
        
        