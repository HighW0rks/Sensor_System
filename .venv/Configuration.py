import Chart
import customtkinter as ctk
import time
import main

class ConfigurationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.main_app = main.MainApp()
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.config_file = "config.txt"
        self.geometry("800x600")
        self.X_as_middle = 800/2
        self.resizable(width=False, height=False)
        self.title("Configuration")
        self.Titletext = None
        self.config_option_name = None
        self.dark_white_mode_var = None
        self.loading()

    def loading(self):
        self.Frame_Input = ctk.CTkFrame(self)
        self.Frame_Input.grid(row=1, column=0,padx=(self.X_as_middle-self.Frame_Input.winfo_reqwidth())/2,pady=40, stick="nsew")
        self.Frame_Bool = ctk.CTkFrame(self)
        self.Frame_Bool.grid(row=1,column=1,padx=(self.X_as_middle-self.Frame_Bool.winfo_reqwidth())/2,pady=40, sticky="nsew")
        self.Close_Save_Button = ctk.CTkButton(self, text="Save & Close", command=self.Close_Save)
        self.Close_Save_Button.place(x=self.X_as_middle-self.Close_Save_Button.winfo_reqwidth()/2, y=600-50)
        self.Title()
        self.Section_Input()
        self.Section_Bool()

    def Title(self):
        self.Titletext = ctk.CTkLabel(master=self, text="Configuration", font=ctk.CTkFont(size=20, weight="bold"))
        self.Titletext.grid(row=0, column=0,columnspan=3, sticky="N")

    def Section_Input(self):
        self.text_option_seconden = ctk.CTkLabel(master=self.Frame_Input, text="Seconden")
        self.text_option_seconden.grid(row=1,column=0,padx=(20,0), pady=20,sticky="w")
        self.text_option_seconden_step = ctk.CTkLabel(master=self.Frame_Input, text="Stappen tussenin")
        self.text_option_seconden_step.grid(row=2,column=0,sticky="w")
        self.config_option_seconden = ctk.CTkEntry(master=self.Frame_Input)
        self.config_option_seconden.insert(0,self.readfile_value(1))
        self.config_option_seconden.grid(row=1,column=1, padx=(10,20))
        self.config_option_seconden.bind("<Return>", lambda event: self.text_config(1, self.config_option_seconden))
        self.config_option_seconden_step = ctk.CTkEntry(master=self.Frame_Input)
        self.config_option_seconden_step.insert(0,self.readfile_value(2))
        self.config_option_seconden_step.grid(row=2,column=1)
        self.config_option_seconden_step.bind("<Return>", lambda event: self.text_config(2, self.config_option_seconden_step))

    def Close_Save(self):
        self.text_config(2,self.config_option_seconden_step)
        self.text_config(1, self.config_option_seconden)
        self.destroy()

    def Section_Bool(self):
        dark_mode_state = self.read_dark_mode_state_from_file()
        self.dark_white_mode_var = ctk.BooleanVar(value=dark_mode_state)
        self.dark_white_mode = ctk.CTkSwitch(master=self.Frame_Bool, text="Dark/Light mode", variable=self.dark_white_mode_var, command=lambda: self.bool_config(3))
        self.dark_white_mode.grid(row=3, column=1, padx=20, pady=20)

    def read_dark_mode_state_from_file(self):
        with open(self.config_file) as read:
            for line in read:
                if "Modes" in line:
                    return line.split(": ")[1].strip() == "True"
        return False

    def readfile_value(self, row):
        with open(self.config_file,"r") as read:
            self.config_text_value = read.readlines()
        self.parameter_text_configvalue = str(self.config_text_value[row - 1]).split(": ")[1].strip()
        return self.parameter_text_configvalue

    def text_config(self, row, name):
        with open(self.config_file, "r") as read:
            self.config_text_value = read.readlines()
        self.parameter_text_configname = str(self.config_text_value[row - 1]).split(": ")[0].strip()
        self.config_new_value = name.get()
        self.updated_text_row = f"{self.parameter_text_configname}: {self.config_new_value}\n"
        self.config_text_value[row - 1] = self.updated_text_row
        with open(self.config_file, "w") as write:
            write.writelines(self.config_text_value)

    def bool_config(self, row):
        with open(self.config_file, "r") as read:
            self.config_bool_value = read.readlines()
        self.parameter_bool_configname = str(self.config_bool_value[row - 1]).split(": ")[0].strip()
        self.parameter_bool_configvalue = str(self.config_bool_value[row - 1]).split(": ")[1].strip()
        self.new_bool_value = str(not self.dark_white_mode_var.get())
        updated_row = f"{self.parameter_bool_configname}: {self.dark_white_mode_var.get()}\n"
        self.config_bool_value[row - 1] = updated_row
        with open(self.config_file, "w") as write:
            write.writelines(self.config_bool_value)
        self.after(0, self.restart)

    def restart(self):
        self.main_app.background_color()

    def destroy(self):
        app = Chart.App()
        app.deiconify()
        super().destroy()




