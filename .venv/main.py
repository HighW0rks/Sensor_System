import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk
import Chart
import sys
import os
import serial.tools.list_ports
import serial
import main
import propar
import subprocess
import threading
import time
from connection import Connection
import psutil

def terminate_existing_main_processes():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'main.exe':
            proc.terminate()

class File_check(ctk.CTk):
    def __init__(self, con):
        super().__init__()
        self.geometry("250x100")
        self.resizable(width=False, height=False)
        self.title("Error")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.Error_text = ctk.CTkLabel(self, text="")
        self.update()
        self.con = con
        self.Location_button = ctk.CTkButton(self, text="Select a folder",command=Chart.ConfigurationApp().Folder_Location_Config)
        self.Location_check()

    def Location_check(self):
        Location = Chart.ConfigurationApp().readfile_value(4)
        if len(Location) <= 0:
            self.Error_text.configure(text="Error!\nNo folder location found")
            self.update()
            self.Location_button.place(x=self.winfo_reqwidth()/2-self.Location_button.winfo_reqwidth()/4, y=50)
            self.after(200, self.Location_check)
        else:
            self.destroy()
            app = MainApp(self.con)
            app.mainloop()

class MainApp(ctk.CTk):
    def __init__(self,con):
        super().__init__()
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.geometry("500x550")
        self.resizable(width=False, height=False)
        self.title("Test App")
        self.status_sensor_var = tk.IntVar(value=0)
        self.status_flow_var = tk.IntVar(value=0)
        self.main_run = True
        self.serienummer = "Unknown"
        self.connection = con
        self.sensor_status = 0
        self.flow_status = 0
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
        self.frame()
        self.status_info()
        self.sensor_info_con()
        self.sensor_info()
        self.buttons()
        self.thread()
        self.ppm_meter()

    def frame(self):
        self.frame_button = ctk.CTkFrame(self)
        self.frame_button.grid(row=0, column=0, padx=(20, 0), pady=(20, 20), sticky="n")
        self.frame_status = ctk.CTkFrame(self)
        self.frame_status.grid(row=0, column=1, padx=(66, 0),pady=(20, 0), sticky="n")
        self.frame_ppm = ctk.CTkFrame(self)
        self.frame_ppm.grid(row=1, column=1, padx=(66,0), sticky="n")
        self.information_title = ctk.CTkFrame(self)
        self.information_title.grid(row=2, column=0, columnspan=2,sticky="n",padx=(20,0),pady=(20,10))
        self.information_tab = ctk.CTkFrame(self)
        self.information_tab.grid(row=3, column=0,columnspan=2,padx=(20,0),sticky="n")

    def thread(self):
        sensor_thread = threading.Thread(target=self.sensor_thread)
        sensor_thread.daemon = True
        sensor_thread.start()
        flow_thread = threading.Thread(target=self.flow_thread)
        flow_thread.daemon = True
        flow_thread.start()

    def buttons(self):
        self.empty_button_frame = ctk.CTkLabel(master=self.frame_button, text="", height=0)
        self.empty_button_frame.grid(row=0,column=0,padx=100)
        self.chart_button = ctk.CTkButton(master=self.frame_button, text="Chart", command=self.open_chart)
        self.chart_button.grid(row=1,column=0,pady=(0,15))
        self.config_button = ctk.CTkButton(master=self.frame_button, text="Config", command=self.open_config)
        self.config_button.grid(row=2,column=0, pady=(0,15))
        self.start_button = ctk.CTkButton(self, text="Start", width=300, height=100, command=self.start_program)
        self.start_button.grid(row=5, column=0, columnspan=2, padx=(20,0),pady=(10,20), sticky="nsew")
        self.channel_option = ctk.CTkComboBox(self, values=["Ch1", "Ch2", "Ch3", "Ch4", "Ch6"])
        self.channel_option.grid(row=4, column=0, columnspan=2, padx=(20,0), pady=(20,0))
        self.channel_option.set("Select a channel")

    def status_info(self):
        self.empty_status = ctk.CTkLabel(master=self.frame_status, text="",height=0).grid(row=0,column=0,padx=100)
        self.status_sensor = ctk.CTkCheckBox(master=self.frame_status, text="Sensor", state="disabled",variable=self.status_sensor_var)
        self.status_sensor.grid(row=1, column=0,padx=40,pady=(0,5),sticky="w")
        self.status_flow = ctk.CTkCheckBox(master=self.frame_status, text="Flow controller", state="disabled", variable=self.status_flow_var)
        self.status_flow.grid(row=2, column=0,padx=40,pady=(4,20),sticky="w")

    def sensor_info_con(self):
        try:
            self.connection.sensor.write(b"\x02\x30\x39\x30\x30\x34\x34\x30\x30\x30\x62\x03")
            self.serienummer = str(self.connection.sensor.read(1024).decode()[5:13])
            if not len(self.serienummer) > 0:
                self.serienummer = "Unknown"
        except Exception as e:
            print(e)
            pass
    def sensor_info(self):
        title = ctk.CTkLabel(master=self.information_title, text="Sensor Information", font=ctk.CTkFont(size=15,weight="bold"))
        title.grid(row=0, column=0, padx=80, pady=5)
        self.version = ctk.CTkLabel(master=self.information_tab, text=f"Serie Nummer\n{self.serienummer}")
        self.version.grid(row=0, column=0,padx=20,pady=10)

    def sensor_thread(self):
        self.sensor_info_con()
        while self.main_run:
            if self.serienummer == "Unknown":
                self.sensor_info_con()
            try:
                self.serial = self.connection.sensor
                self.serial.write(b"\x02\x30\x35\x30\x30\x30\x37\x03")
                response = self.serial.read(1024)
                if response:
                    self.sensor_status = 1
                    self.ppm_value = response.decode().split()[5]
                    with open("transfer.txt", "w") as write:
                        write.writelines(self.ppm_value)
                    self.version.configure(text=f"Serie Nummer\n{self.serienummer}")
                    self.ppm_meter_label.configure(text=f"PPM value\n\n{self.ppm_value}")
                    x = False
                else:
                    self.serienummer = "Unknown"
                    self.sensor_status = 0
                    self.version.configure(text="Serie Nummer\nUnknown")
                    self.ppm_meter_label.configure(text="Sensor not found")
                    x = True
            except Exception as e:
                self.serienummer = "Unknown"
                self.sensor_status = 0
                print(f"Error in sensor communication: {e}")
                self.version.configure(text="Serie Nummer\nUnknown")
                self.ppm_meter_label.configure(text="Sensor not found")
                self.connection.initialize_sensor()
                x = True
                pass

    def flow_thread(self):
        while self.main_run:
            try:
                flow = self.connection.flow_1
                value = flow.readParameter(8)
                if value != None:
                    self.flow_status = True
                else:
                    self.flow_status = False
            except Exception as e:
                print(f"Error in flow communication: {e}")
                self.flow_status = 0
            self.after(1000, self.status_update)
            time.sleep(1)

    def status_update(self):
        self.status_sensor_var.set(self.sensor_status)
        self.status_flow_var.set(self.flow_status)
        if self.sensor_status and self.flow_status and self.channel_option.get() != "Select a channel":
            self.start_button.configure(state="enabled")
        else:
            self.start_button.configure(state="disabled")

    def ppm_meter(self):
        self.empty_ppm_frame = ctk.CTkLabel(master=self.frame_ppm, text="")
        self.empty_ppm_frame.grid(row=0, column=0, padx=100)
        self.ppm_meter_label = ctk.CTkLabel(master=self.frame_ppm, text="Connecting...")
        self.ppm_meter_label.grid(row=0,column=0, pady=20)

    def open_chart(self):
        # self.main_run = False
        app = Chart.App(self.connection,self.serienummer)
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        print("test")
        app.mainloop()

    def open_config(self):
        app = main.Configuration(self.connection)
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def start_program(self):
        main.StartProgram(self.connection,self.channel_option.get())

    def close_app(self):
        terminate_existing_main_processes()


class Configuration(ctk.CTk):
    def __init__(self,con):
        super().__init__()
        self.con = con
        self.title("Configuration")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.option_add("*TCombobox*Listbox*Font", value=0)
        self.comport_sensor = None
        self.comport_flow = None
        self.combobox()
        self.dark_light_mode()
        self.text()

    def listen_comport(self):
        return serial.tools.list_ports.comports()

    def text(self):
        label_sensor = ctk.CTkLabel(self, text="Sensor").grid(row=0, column=1)
        label_flow = ctk.CTkLabel(self, text="Flow Controller").grid(row=0,column=0)
        button_exit = ctk.CTkButton(self, text="Save & Restart App", command=self.restart_app).grid(row=3, column=0, columnspan=2, pady=(30,0))

    def combobox(self):
        self.flow_com = ttk.Combobox(self, width=25, values=self.listen_comport(),font=(15))
        self.flow_com.grid(row=1, column=0)
        self.flow_com.set("Select a comport")
        self.flow_com.configure(font=(20))
        self.sensor_com = ttk.Combobox(self, width=25, values=self.listen_comport(), font=(15))
        self.flow_com.bind('<<ComboboxSelected>>', self.get_flow)
        self.sensor_com.grid(row=1,column=1)
        self.sensor_com.set("Select a comport")
        self.sensor_com.bind('<<ComboboxSelected>>', self.get_sensor)

    def check(self):
        if self.comport_flow != None and self.comport_sensor != None:
            print("All good")

    def get_flow(self, event):
        self.comport_flow = self.flow_com.get()
        value = self.comport_flow.split(" ")[0]
        Chart.ConfigurationApp().text_config(6, value)
        self.check()

    def get_sensor(self, event):
        self.comport_sensor = self.sensor_com.get()
        value = self.comport_sensor.split(" ")[0]
        Chart.ConfigurationApp().text_config(5, value)
        self.check()

    def dark_light_mode(self):
        with open("config.txt", "r") as r:
            self.read = r.readlines()
            self.modes_bool = str(self.read[2].split(": ")[1].split("\n")[0])
            dark_mode_state = ctk.BooleanVar(value=(self.modes_bool == "True"))
            print(dark_mode_state)
        self.dark_white_mode = ctk.CTkSwitch(self, text="Dark/Light mode", variable=dark_mode_state, command= self.set_dark_light_mode)
        self.dark_white_mode.grid(row=2, column=0, columnspan=2, padx=20, pady=20)

    def set_dark_light_mode(self):
        with open("config.txt","r") as r:
            value = r.readlines()[2].split(": ")[1].split("\n")[0]
        if value == "True":
            value = "False"
        else:
            value = "True"
        with open("config.txt","w") as w:
            self.read[2] = f"Modes: {value}\n"
            w.writelines(self.read)
        self.restart()

    def restart(self):
        main.MainApp.background_color(ctk.CTk)

    def restart_app(self):
        os.execv(sys.executable, [sys.executable] + sys.argv)

    def destroy(self):
        super().destroy()

class StartProgram():
        def __init__(self,con,channel):
            self.row_value = 0
            self.connection = con
            self.channel = channel
            self.read_script_thread()

        def read_script_thread(self):
            thread = threading.Thread(target=self.read_script)
            thread.daemon = True
            thread.start()

        def read_script(self):
            try:
                with open(fr"Sensor\Script {self.channel}.txt", "r") as read:
                    read_value = read.readlines()
                    if not self.row_value < len(read_value):
                        print("End of file reached")
                        return
                    self.seconde_value = read_value[self.row_value].split(". ")[1].split(", ")[0].strip()
                    self.percentage_value = read_value[self.row_value].split(", ")[1].strip()
                    self.channel_value = read_value[self.row_value].split(", ")[2].strip()
                print(f"Sec: {self.seconde_value}  | Flow: ","{:.0f}".format(int(self.percentage_value) * 32000 / 100),f" | Channel: {self.channel_value}")
                self.set_value()
            except Exception as e:
                print(e)
                pass
        def set_value(self):
            try:
                self.connection.flow_1.writeParameter(9, "{:.0f}".format(int(self.percentage_value) * 32000 / 100))
                self.row_value += 1
                time.sleep(int(self.seconde_value))
                self.read_script()
            except Exception as e:
                print(e)
                pass

if __name__ == "__main__":
    log_file_path = "debug_log.txt"

    log_file = open(log_file_path, "w")
    os.dup2(log_file.fileno(), 1)  # Redirect stdout
    os.dup2(log_file.fileno(), 2)  # Redirect stderr
    con = Connection()
    mainapp = File_check(con)
    mainapp.mainloop()