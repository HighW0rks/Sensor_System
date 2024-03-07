import customtkinter as ctk
import propar
import serial
import main

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.main_app = main.MainApp()
        self.title("PPM Meter")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.ppm_meter_text = None
        self.error_message = None
        self.value = None
        self.instrument = None
        self.connection()
        self.loading()


    def loading(self):
        self.ppm_meter_text = ctk.CTkLabel(self, text=self.value)
        self.ppm_meter_text.grid(row=0,column=0)

    def connection(self):
        try:
            self.instrument = propar.instrument('COM1')
            self.value = self.instrument.readParameter(dde_nr=1, channel=None)

        except Exception:
            self.error_message = ctk.CTkLabel(self, text="Something has gone wrong \nplease contact the developer")
            self.error_message.grid(row=0,column=0)


    def destroy(self):
        self.main_app.deiconify()
        super().destroy()
        main.MainApp()


