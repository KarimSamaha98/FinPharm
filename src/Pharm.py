# from pickle import PickleBuffer
import tkinter
import tkinter.messagebox
from tkinter import *
import customtkinter
import tkinter.ttk as ttk
from tkinter import filedialog
from time import sleep
from PIL import ImageTk, Image
import time
from datetime import datetime
from openpyxl import Workbook
import os
from Writer import *
from tkinter import messagebox




customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):


    def __init__(self, WIDTH, HEIGHT):
        super().__init__()

        self.title("Pharmacy Accounting and Management Toolkit")
        self.WIDTH = WIDTH
        self.HEIGHT = HEIGHT
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # call .on_closing() when app gets closed
        self.init_menu()
        self.init_param()
        self.db = Writer(self.PATH.get(), self.FILENAME)
        
    
    def init_param(self):
        self.FILENAME = "{}_{}.xlsx".format(datetime.now().month, datetime.now().year)
        self.TYPE = tkinter.StringVar(master=self, value='None')
        self.CLIENT_NAME = tkinter.StringVar(master=self, value='None')
        self.INSURANCE_REF = tkinter.StringVar(master=self, value='0')
        self.INVOICE_REF = tkinter.StringVar(master=self, value='0')
        self.GROUP_NAME = tkinter.StringVar(master=self, value='None')
        self.TOTAL_AMOUNT = tkinter.StringVar(master=self, value='0')
        self.PATIENT_SHARE = tkinter.StringVar(master=self, value='0')
        self.PAYMENT_STATUS = tkinter.StringVar(master=self, value='None')
        self.PATH = tkinter.StringVar(master=self, value="C:\\Users\\Karim\\OneDrive\\Documents\\PHARM_DATA")
        self.SSNBR = tkinter.StringVar(master=self, value='None')
        self.APPROVED_AMOUNT = tkinter.StringVar(master=self, value='0')
        self.OPERATOR = tkinter.StringVar(master=self, value='None')
        self.NOTE = tkinter.StringVar(master=self, value='None')
        self.ADMIN = tkinter.StringVar(master=self, value='None')
        self.PASSWORD = tkinter.StringVar(master=self, value='None')
        return
           
    def init_menu(self):
        # configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=9)
        self.grid_rowconfigure(0, weight=1)
        self.frame_left = customtkinter.CTkFrame(master=self,corner_radius=0)
        self.frame_left.grid(row=0, column=0, sticky="nswe")
        self.frame_right = None

        # configure widgets
        self.label_1 = customtkinter.CTkLabel(master=self.frame_left,text="Main Menu",text_font=("Roboto Medium", -16))  # font name and size in px
        self.label_1.grid(row=1, column=0, columnspan=2, sticky="nswe", pady=10)
        self.button_1 = customtkinter.CTkButton(master=self.frame_left,text="Add Transaction",command=self.show_add_transaction)
        self.button_1.grid(row=2, column=0, columnspan=2, sticky="nswe",  pady=10, padx=20)
        self.button_3 = customtkinter.CTkButton(master=self.frame_left,text="View Transaction",command=self.show_view_transaction)
        self.button_3.grid(row=3, column=0, columnspan=2, sticky="nswe",  pady=10, padx=20)
        self.button_2 = customtkinter.CTkButton(master=self.frame_left,text="Edit Transaction",command=self.show_edit_transaction)
        self.button_2.grid(row=4, column=0, columnspan=2,  sticky="nswe", pady=10, padx=20)
        self.button_4 = customtkinter.CTkButton(master=self.frame_left,text="View Client",command=self.show_view_client)
        self.button_4.grid(row=5, column=0, columnspan=2, sticky="nswe",  pady=10, padx=20)
        self.button_5 = customtkinter.CTkButton(master=self.frame_left,text="Settings",command=self.show_settings)
        self.button_5.grid(row=6, column=0, columnspan=2, sticky="nswe",  pady=10, padx=20)
    
    
    def init_frame_right(self):
        # instantiate right frame
        self.frame_right = customtkinter.CTkFrame(master=self, width=2000, height=self.HEIGHT)
        self.frame_right.grid(row=0, column=1, padx=20, pady=20)
        self.frame_right.columnconfigure([0,1,2,3,4,5,6,7,8,9],weight=1)
        self.frame_right.rowconfigure([0,1,2,3,4,5,6,7,8,9],weight=1)

    def create_add_transaction(self):
        self.title = customtkinter.CTkLabel(master=self.frame_right, text="Add a Transaction",text_font=("Roboto Medium", -32))  # font name and size in px 
        self.title.grid(row=0, column=0, columnspan=9, pady=20, padx=20, sticky='nw')
        self.button_1 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Transaction Type", anchor='w')
        self.button_1.grid(row=1, column=0, pady=10, padx=20, sticky='W') 
        self.type = customtkinter.CTkOptionMenu(master=self.frame_right,
                                       values=["Personal", "Bankers", "Globemed"], variable = self.TYPE, command=self.update_entries)
        self.type.grid(row=1, column=1, columnspan=9,padx=20, pady=10, sticky='W')       

    def create_settings(self):
        self.title = customtkinter.CTkLabel(master=self.frame_right, text="Settings",text_font=("Roboto Medium", -32))  # font name and size in px 
        self.title.grid(row=0, column=0, columnspan=9, pady=20, padx=20, sticky='nw')

        self.button_1 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Admin", anchor='w')
        self.button_1.grid(row=1, column=0, pady=10, padx=20, sticky='W') 

        self.button_1 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Password", anchor='w')
        self.button_1.grid(row=2, column=0, pady=10, padx=20, sticky='W')

        self.entry_43 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.ADMIN, width=100)
        self.entry_43.grid(row=1, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        self.entry_43 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.PASSWORD, width=100)
        self.entry_43.grid(row=2, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.button_8 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Validate",
                                                command=self.validate_entry, width=200)
        
        self.button_8.grid(row=10, column=9, padx=20, pady=30)  


    def validate_entry(self):
        if self.ADMIN.get() == 'GSAMAHA' and self.PASSWORD.get() == '1957':
            self.show_dir_param()
        else:
            messagebox.showinfo("Admin Access", "Wrong credentials. Try again or contact support.", icon="warning", parent=None)

    def show_dir_param(self):
        print("Show parameters")
        self.entry_8 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="", textvariable=self.PATH, width=250, state=tkinter.DISABLED)
        self.entry_8.grid(row=3, column=1, columnspan=2, sticky='E', padx=20, pady=30)
        
        
        self.button_12 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Select Directory",
                                                command=self.askDirectory)
        
        self.button_12.grid(row=4, column=1, padx=20, pady=30)
        return 
        
    def show_bankers_entry(self):
        ## Labels
        self.button_4 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Additional Notes", anchor='w'
                                                )
        self.button_4.grid(row=1, column=5, pady=10, padx=20, sticky='W')

        self.button_4 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Client Name", anchor='w'
                                                )
        self.button_4.grid(row=2, column=0, pady=10, padx=20, sticky='W')
        
        self.button_5 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Insurance Reference", anchor='w'
                                                )
        
        self.button_5.grid(row=3, column=0, pady=10, padx=20, sticky='W')
        
        self.button_6 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Invoice Reference", anchor='w'
                                                )
        
        self.button_6.grid(row=4, column=0, pady=10, padx=20, sticky='W')
        
        
        self.button_8 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Group Name", anchor='w'
                                                )
        
        self.button_8.grid(row=5, column=0, pady=10, padx=20, sticky='W')
        self.button_9 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Total Amount", anchor='w'
                                                )
        
        self.button_9.grid(row=6, column=0, pady=10, padx=20, sticky='W')

        self.button_10 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Patient Share", anchor='w'
                                                )
        
        self.button_10.grid(row=7, column=0, pady=10, padx=20, sticky='W')

        self.button_11 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Operator", anchor='w'
                                                )
        
        self.button_11.grid(row=8, column=0, pady=10, padx=20, sticky='W')
    
        
        ## Entries Left
        self.entry_43 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.NOTE, width=400)
        self.entry_43.grid(row=1, column=6,  columnspan=3, padx=20, pady=10, sticky='W')

        self.entry_1 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.CLIENT_NAME, width=400)
        self.entry_1.grid(row=2, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        self.entry_2 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.INSURANCE_REF, width=100)
        self.entry_2.grid(row=3, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        
        self.entry_3 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.INVOICE_REF, width=100)
        self.entry_3.grid(row=4, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        
        self.entry_4 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.GROUP_NAME, width=400)
        self.entry_4.grid(row=5, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        self.entry_5 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.TOTAL_AMOUNT, width=100)
        self.entry_5.grid(row=6, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.entry_6 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.PATIENT_SHARE, width=100)
        self.entry_6.grid(row=7, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.type = customtkinter.CTkOptionMenu(master=self.frame_right,
                                       values=["Dana", "Patricia", "Roula", "Rasha", "Nour"], variable = self.OPERATOR)   
        self.type.grid(row=8, column=1,  columnspan=4, padx=20, pady=10, sticky='W')  

        self.button_8 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Submit Transaction",
                                                command=self.write_data, width=200)
        
        self.button_8.grid(row=10, column=9, padx=20, pady=30)    
    
    def show_personal_entry(self):
        ## Labels
        self.button_4 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Additional Notes", anchor='w'
                                                )
        self.button_4.grid(row=1, column=5, pady=10, padx=20, sticky='W')

        self.button_4 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Client Name", anchor='w'
                                                )
        self.button_4.grid(row=2, column=0, pady=10, padx=20, sticky='W')    

        self.button_5 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Payment Status", anchor='w'
                                                )
        
        self.button_5.grid(row=3, column=0, pady=10, padx=20, sticky='W')
        
        self.button_6 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Invoice Reference", anchor='w'
                                                )
        
        self.button_6.grid(row=4, column=0, pady=10, padx=20, sticky='W')
        
        
        self.button_9 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Total Amount", anchor='w'
                                                )
        
        self.button_9.grid(row=5, column=0, pady=10, padx=20, sticky='W')

        self.button_11 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Operator", anchor='w'
                                                )
        
        self.button_11.grid(row=6, column=0, pady=10, padx=20, sticky='W')        
    
        
        ## Entries Left
        self.entry_43 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.NOTE, width=400)
        self.entry_43.grid(row=1, column=6,  columnspan=3, padx=20, pady=10, sticky='W')

        self.entry_1 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.CLIENT_NAME, width=400)
        self.entry_1.grid(row=2, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        self.type = customtkinter.CTkOptionMenu(master=self.frame_right,
                                       values=["P", "NP"], variable = self.PAYMENT_STATUS)
        self.type.grid(row=3, column=1, columnspan=9,padx=20, pady=10, sticky='W')   
        
        self.entry_3 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.INVOICE_REF, width=100)
        self.entry_3.grid(row=4, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        
        self.entry_5 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.TOTAL_AMOUNT, width=100)
        self.entry_5.grid(row=5, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.type = customtkinter.CTkOptionMenu(master=self.frame_right,
                                       values=["Dana", "Patricia", "Roula", "Rasha", "Nour"], variable = self.OPERATOR)   
        self.type.grid(row=6, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.button_8 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Submit Transaction",
                                                command=self.write_data, width=200)
        
        self.button_8.grid(row=10, column=9, padx=20, pady=30)            
   
    def show_globemed_entry(self):
        ## Labels
        self.button_4 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Additional Notes", anchor='w'
                                                )
        self.button_4.grid(row=1, column=5, pady=10, padx=20, sticky='W')

        self.button_4 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Client Name", anchor='w'
                                                )
        self.button_4.grid(row=2, column=0, pady=10, padx=20, sticky='W')
        
        self.button_5 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Insurance Reference", anchor='w'
                                                )
        
        self.button_5.grid(row=3, column=0, pady=10, padx=20, sticky='W')
        
        self.button_6 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Invoice Reference", anchor='w'
                                                )
        
        self.button_6.grid(row=4, column=0, pady=10, padx=20, sticky='W')
        
        
        self.button_8 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="SSNBR", anchor='w'
                                                )
        
        self.button_8.grid(row=5, column=0, pady=10, padx=20, sticky='W')
        self.button_9 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Total Amount", anchor='w'
                                                )
        
        self.button_9.grid(row=6, column=0, pady=10, padx=20, sticky='W')

        self.button_10 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Patient Share", anchor='w'
                                                )
        
        self.button_10.grid(row=7, column=0, pady=10, padx=20, sticky='W')

        self.button_11 = customtkinter.CTkLabel(master=self.frame_right,
                                                text="Operator", anchor='w'
                                                )
        
        self.button_11.grid(row=8, column=0, pady=10, padx=20, sticky='W') 
        
        ## Entries Left
        self.entry_43 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.NOTE, width=400)
        self.entry_43.grid(row=1, column=6,  columnspan=3, padx=20, pady=10, sticky='W')

        self.entry_1 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.CLIENT_NAME, width=400)
        self.entry_1.grid(row=2, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        self.entry_2 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.INSURANCE_REF, width=100)
        self.entry_2.grid(row=3, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        
        self.entry_3 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.INVOICE_REF, width=100)
        self.entry_3.grid(row=4, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        
        self.entry_4 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="None", textvariable=self.SSNBR, width=400)
        self.entry_4.grid(row=5, column=1,  columnspan=4, padx=20, pady=10, sticky='W')
        
        self.entry_5 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.TOTAL_AMOUNT, width=100)
        self.entry_5.grid(row=6, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.entry_6 = customtkinter.CTkEntry(master=self.frame_right, placeholder_text="0", textvariable=self.PATIENT_SHARE, width=100)
        self.entry_6.grid(row=7, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.type = customtkinter.CTkOptionMenu(master=self.frame_right,
                                       values=["Dana", "Patricia", "Roula", "Rasha", "Nour"], variable = self.OPERATOR)   
        self.type.grid(row=8, column=1,  columnspan=4, padx=20, pady=10, sticky='W')

        self.button_8 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Submit Transaction",
                                                command=self.write_data, width=200)
        
        self.button_8.grid(row=10, column=9, padx=20, pady=30)    
       
   
    def write_data(self):
        try:
            if self.TYPE.get() == 'Bankers':
                DATA = [self.INVOICE_REF.get(), self.INSURANCE_REF.get(), self.GROUP_NAME.get(), self.CLIENT_NAME.get(), int(self.TOTAL_AMOUNT.get()), 'None', self.PATIENT_SHARE.get(), 'None', 'None', self.PAYMENT_STATUS.get(), 'None', str(datetime.now()), self.NOTE.get(), self.OPERATOR.get()]
                self.db.write_data('Bankers', DATA)
            if self.TYPE.get() == 'Personal':
                if self.PAYMENT_STATUS.get() == 'P':
                    purchase_amount = int(self.TOTAL_AMOUNT.get())
                    credit_amount = 0
                    payment_date =  str(datetime.now())
                elif self.PAYMENT_STATUS.get() == 'NP':
                    purchase_amount = 0
                    credit_amount = int(self.TOTAL_AMOUNT.get())
                    payment_date = 'None'
                DATA =  [self.INVOICE_REF.get(), self.CLIENT_NAME.get(), int(self.TOTAL_AMOUNT.get()), purchase_amount, credit_amount, self.PAYMENT_STATUS.get(), payment_date,  str(datetime.now()), self.NOTE.get(), self.OPERATOR.get()]
                self.db.write_data('Personal', DATA)
            if self.TYPE.get() == 'Globemed':
                DATA = [self.INVOICE_REF.get(), self.CLIENT_NAME.get(), self.INSURANCE_REF.get(), self.SSNBR.get(), int(self.TOTAL_AMOUNT.get()), 'None', self.PATIENT_SHARE.get(), int(self.TOTAL_AMOUNT.get())-int(self.PATIENT_SHARE.get()), 'None', 'None', 'None', str(datetime.now()), self.NOTE.get(), self.OPERATOR.get()]
                self.db.write_data('Globemed', DATA)
            self.reset_param() #Reset the values
        except:
            messagebox.showinfo("Writing Data", "An error occured while writing the data, make sure that all fields have been correctly filled and try again", icon="warning", parent=None)
        return
    
    def reset_param(self):
        self.TYPE.set('None')
        self.CLIENT_NAME.set('None')
        self.INSURANCE_REF.set('0')
        self.INVOICE_REF.set('0')
        self.GROUP_NAME.set('None')
        self.TOTAL_AMOUNT.set('0')
        self.PATIENT_SHARE.set('0')

    def show_add_transaction(self):
        if self.frame_right != None:
            self.frame_right.grid_forget()
        self.init_frame_right()
        self.create_add_transaction()
        print("Add Transaction Button Pressed") 

    def show_edit_transaction(self):
        return 

    def show_view_transaction(self):
        return 

    def show_view_client(self):
        return 

    def show_settings(self):
        if self.frame_right != None:
            self.frame_right.grid_forget()
        self.init_frame_right()
        self.create_settings()
        print("Settings Button Pressed") 
        return 

    def on_closing(self, event=0):
        self.destroy()
             
        
    def askDirectory(self):
        self.withdraw()
        self.PATH.set(filedialog.askdirectory())
        self.db = Writer(self.PATH.get(), self.FILENAME)
        self.deiconify()

    def update_entries(self, choice):
        if choice == 'Bankers':
            self.frame_right.grid_forget()
            self.init_frame_right()
            self.show_add_transaction()
            self.show_bankers_entry()
        elif choice == 'Personal':
            self.frame_right.grid_forget()
            self.init_frame_right()
            self.show_add_transaction()
            self.show_personal_entry()
        elif choice == 'Globemed':
            self.frame_right.grid_forget()
            self.init_frame_right()
            self.show_add_transaction()
            self.show_globemed_entry()            
    


if __name__ == "__main__":
    app = App(1920,1080)
    while True:
        app.update()
