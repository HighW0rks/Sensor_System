import customtkinter as ctk
import tkchart
import random
import Configuration
import Chart
import time
import PPM_Meter
import sys

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.geometry("500x500")
        self.resizable(width=False, height=False)
        self.title("Test App")
        self.background_color()
        self.loading()
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def background_color(self):
        with open("config.txt") as read:
            for line in read:
                if "Modes" in line:
                    value = line.split(": ")[1].strip()
        if value == "True":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")

    def loading(self):
        self.chart_button = ctk.CTkButton(self, text="Chart", command=self.open_chart)
        self.chart_button.grid(row=0,column=0)
        self.PPM_Meter_button = ctk.CTkButton(self, text="PPM Meter", command=self.open_PPM_Meter)
        self.PPM_Meter_button.grid(row=2,column=0)


    def open_chart(self):
        self.withdraw()
        app = Chart.App()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def open_PPM_Meter(self):
        self.withdraw()
        app = PPM_Meter.App()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def close_app(self):
        self.destroy()
        sys.exit()

if __name__ == "__main__":
    mainapp = MainApp()
    mainapp.mainloop()
