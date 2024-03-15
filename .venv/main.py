import customtkinter as ctk
import tkinter as tk
import Configuration
import Chart
import Sensor
import sys

import main
import propar
import serial
import threading
import time

class First_check(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("250x100")
        self.resizable(width=False, height=False)
        self.title("Error")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.Error_text = ctk.CTkLabel(self, text="")
        self.update()
        self.Location_button = ctk.CTkButton(self, text="Select a folder",
                                             command=Configuration.ConfigurationApp().Folder_Location_Config)
        self.Location_check()

    def Location_check(self):
        Location = Configuration.ConfigurationApp().readfile_value(4)
        if len(Location) <= 0:
            self.Error_text.configure(text="Error!\nNo folder location found")
            self.update()
            self.Location_button.place(x=self.winfo_reqwidth()/2-self.Location_button.winfo_reqwidth()/4, y=50)
            self.after(200,self.Location_check())
        else:
            self.destroy()
            app = MainApp()
            app.mainloop()

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.geometry("500x500")
        self.resizable(width=False, height=False)
        self.title("Test App")
        self.status_sensor_var = tk.IntVar(value=0)
        self.status_flow_var = tk.IntVar(value=0)
        self.main_run = True
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
        self.frame_button = ctk.CTkFrame(self)
        self.frame_button.grid(row=0, column=0,padx=(20,0), pady=(20,0))
        self.frame_status = ctk.CTkFrame(self)
        self.frame_status.grid(row=0, column=1,padx=(66,0),pady=(20,0))
        self.status_info()
        self.buttons()
        self.thread()
    def thread(self):
        self.status_thread = threading.Thread(target=self.status_thread_function)
        self.status_thread.daemon = True
        self.status_thread.start()

    def buttons(self):
        self.chart_button = ctk.CTkButton(master=self.frame_button, text="Chart", command=self.open_chart)
        self.chart_button.grid(row=0,column=0,padx=30,pady=(20,0))
        self.PPM_Meter_button = ctk.CTkButton(master=self.frame_button, text="PPM Meter", command=self.open_PPM_Meter)
        self.PPM_Meter_button.grid(row=1,column=0,padx=30,pady=(0,20))
        self.start_button = ctk.CTkButton(self, text="Start", width=300, height=100, command=main.start_program)
        self.start_button.grid(row=1, column=0, columnspan=2, padx=(20,0),pady=(200,0), sticky="nsew")

    def status_info(self):
        self.status_sensor = ctk.CTkCheckBox(master=self.frame_status, text="Sensor", state="disabled",variable=self.status_sensor_var)
        self.status_sensor.grid(row=0, column=0,padx=40,pady=(20,5),sticky="w")
        self.status_flow = ctk.CTkCheckBox(master=self.frame_status, text="Flow controller", state="disabled", variable=self.status_flow_var)
        self.status_flow.grid(row=1, column=0,padx=40,pady=(4,20),sticky="w")

    def status_thread_function(self):
        while self.main_run == True:
            try:
                self.serial = serial.Serial("COM5", 19200, timeout=1)
                self.serial.write(b"\x02\x30\x35\x30\x30\x30\x37\x03")
                self.response = self.serial.read(1024)
                if self.response:
                    self.sensor_status = 1
                else:
                    self.sensor_status = 0
                self.serial.close()
            except Exception as e:
                self.sensor_status = 0
                # print(e)
            try:
                self.flow = propar.instrument("COM3", channel=1)
                value = self.flow.readParameter(8)
                if value >= 0:
                    self.flow_status = 1
                else:
                    self.flow_status = 0
                # print("Good")
            except Exception as e:
                self.flow_status = 0
                # print(e)
            try:
                self.after(100, self.update_status_ui)
            except Exception as e:
                pass
                # print(e)
            time.sleep(0.1)
    def update_status_ui(self):
        self.status_sensor_var.set(self.sensor_status)
        self.status_flow_var.set(self.flow_status)

    def open_chart(self):
        self.main_run = False
        try:
            self.serial.close()
        except Exception:
            pass
        self.destroy()
        app = Chart.App()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()
        super().destroy()

    def open_PPM_Meter(self):
        self.main_run = False
        super().destroy()
        app = Sensor.App()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def close_app(self):
        super().destroy()
        exit()

if __name__ == "__main__":
    mainapp = First_check()
    mainapp.mainloop()

class start_program():
    def __init__(self):
        self.check_connection()


    def check_connection(self):
        try:
            self.flow1 = propar.instrument("COM3", channel=1)
            print("flow 1 connected")
        except Exception:
            print("Flow 1 didn't connect")
        else:
            value_1 = self.flow1.readParameter(8)
            print(value_1)
            if value_1 >= 0:
                self.flow_value_1 = value_1
            else:
                self.flow_value_1 = 0

        try:
            self.flow2 = propar.instrument("COM3", channel=2)
            print("Flow 2 connected")
            value_2 = self.flow2.readParameter(8)
            if value_2 is not None and value_2 >= 0:
                self.flow_value_2 = value_2
            else:
                self.flow_value_2 = 0
        except Exception:
            print("Flow 2 didn't connect")

        try:
            self.flow3 = propar.instrument("COM3", channel=3)
            print("flow 3 connected")
        except Exception:
            print("Flow 3 didn't connect")
        else:
            value_3 = self.flow3.readParameter(8)
            if value_3 is not None and value_3 >= 0:
                self.flow_value_3 = value_3
            else:
                self.flow_value_3 = 0


