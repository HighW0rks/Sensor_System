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
import xlsxwriter
import datetime

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
        self.running = True
        self.serienummer = "Unknown"
        self.connection = con
        self.serial = self.connection.sensor
        self.sensor_status = 0
        self.flow_status = 0
        self.time = 0
        self.row_value = 0
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
        self.zero_point()

    def frame(self):
        self.frame_button = ctk.CTkFrame(self)
        self.frame_button.grid(row=0, column=0, padx=(20, 0), pady=(20, 20), sticky="n")
        self.frame_status = ctk.CTkFrame(self)
        self.frame_status.grid(row=0, column=1, padx=(66, 0),pady=(20, 0), sticky="n")
        self.frame_ppm = ctk.CTkFrame(self)
        self.frame_ppm.grid(row=1, column=1, padx=(66,0), sticky="n")
        self.frame_zero_point = ctk.CTkFrame(self)
        self.frame_zero_point.grid(row=1, column=0, padx=(20, 0))
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
        if self.main_run:
            try:
                self.connection.sensor.write(b"\x02\x30\x39\x30\x30\x34\x34\x30\x30\x30\x62\x03")
                read = self.connection.sensor.read(1024)
                print("Read SN: ", read)
                print(len(read))
                self.serienummer = str(read.decode()[5:13])
                if len(read) == 0 or len(read) > 24 :
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
                self.serial.write(b"\x02\x30\x35\x30\x30\x30\x37\x03")
                response = self.serial.read(1024).decode()
                print("Response: ",response)
                if response:
                    self.sensor_status = 1
                    self.channel_1 = response.split()[1]
                    self.channel_2 = response.split()[2]
                    self.channel_3 = response.split()[3]
                    self.channel_4 = response.split()[4]
                    self.ppm_value = response.split()[5]
                    self.channel_6 = response.split()[6]
                    self.temperature = response.split()[7]
                    self.air_pressure = response.split()[8]

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
                # self.connection.initialize_sensor()
                x = True
                pass

    def flow_thread(self):
        while self.main_run:
            try:
                flow1 = self.connection.flow_1
                flow2 = self.connection.flow_2
                flow3 = self.connection.flow_3
                self.flow1_value = flow1.readParameter(8)
                self.flow2_value = flow2.readParameter(8)
                self.flow3_value = flow3.readParameter(8)
                if self.flow1_value != None:
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
            self.start_button.configure(state="normal")
        else:
            self.start_button.configure(state="disabled")

        if self.sensor_status:
            self.zero_point_button.configure(state="normal")
        else:
            self.zero_point_button.configure(state="disabled")

    def ppm_meter(self):
        self.empty_ppm_frame = ctk.CTkLabel(master=self.frame_ppm, text="")
        self.empty_ppm_frame.grid(row=0, column=0, padx=100)
        self.ppm_meter_label = ctk.CTkLabel(master=self.frame_ppm, text="Connecting...")
        self.ppm_meter_label.grid(row=0,column=0, pady=17)

    def zero_point(self):
        self.empty_zero_point_frame = ctk.CTkLabel(master=self.frame_zero_point, text="",height=0)
        self.empty_zero_point_frame.grid(row=0, column=0, padx=100)
        self.zero_point_button = ctk.CTkButton(master=self.frame_zero_point, text="Zero Point",command=self.send_zero_point_command)
        self.zero_point_button.grid(row=1, column=0, pady=(0,17))

    def send_zero_point_command(self):
        self.main_run = False
        self.serial.write(b"\x02\x34\x30\x30\x30\x30\x36\x03")
        time.sleep(1)
        self.main_run = True

    def start_program(self):
        self.channel = self.channel_option.get()
        self.start_excel()
        self.read_script_thread()

    def start_excel(self):
        Location = Chart.ConfigurationApp().readfile_value(4)
        self.get_time()
        self.workbook = xlsxwriter.Workbook(fr"{Location}\{self.serienummer}-{self.current_time}.xlsx")
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.write(0, 0, 'Time')
        self.worksheet.set_column('A:A', 20)
        self.worksheet.write(0, 1, 'Channel 1')
        self.worksheet.write(0, 2, 'Channel 2')
        self.worksheet.write(0, 3, 'Channel 3')
        self.worksheet.write(0, 4, 'Channel 4')
        self.worksheet.write(0, 5, 'PPM Value')
        self.worksheet.write(0, 6, 'Channel 6')
        self.worksheet.write(0, 7, 'Temperature')
        self.worksheet.write(0, 8, 'Air Pressure')
        self.write_excel_thread()

    def write_excel_thread(self):
        thread = threading.Thread(target=self.write_excel)
        thread.daemon = True
        thread.start()

    def read_script_thread(self):
        thread = threading.Thread(target=self.read_script)
        thread.daemon = True
        thread.start()

    def write_excel(self):
        while self.running:
            self.time += 1
            sensor = float(self.ppm_value)

            self.get_time()
            self.worksheet.write(self.time, 0, self.current_time)
            self.worksheet.write(self.time, 1, self.channel_1)
            self.worksheet.write(self.time, 2, self.channel_2)
            self.worksheet.write(self.time, 3, self.channel_3)
            self.worksheet.write(self.time, 4, self.channel_4)
            self.worksheet.write(self.time, 5, sensor)
            self.worksheet.write(self.time, 6, self.channel_6)
            self.worksheet.write(self.time, 7, self.temperature)
            self.worksheet.write(self.time, 8, self.air_pressure)
            time.sleep(1)

    def get_time(self):
        self.current_time = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        self.day_time = datetime.datetime.now().strftime("%Y-%m-%d")

    def read_script(self):
        try:
            with open(fr"Sensor\Script {self.channel}.txt", "r") as read:
                read_value = read.readlines()
                if not self.row_value < len(read_value):
                    print("End of file reached")
                    self.running = False
                    self.workbook.close()
                    return
                self.seconde_value = read_value[self.row_value].split(". ")[1].split("; ")[0].strip()
                self.percentage_value = read_value[self.row_value].split("; ")[1].strip()
                self.channel_value = read_value[self.row_value].split("; ")[2].strip()
            print(f"Sec: {self.seconde_value}  | Flow: ", "{:.0f}".format(float(self.percentage_value) * 32000 / 100),
                  f" | Channel: {self.channel_value}")
            self.set_value()
        except Exception as e:
            print(e)
            pass

    def set_value(self):
        if self.channel_value == "1":
            self.connection.flow_1.writeParameter(9, "{:.0f}".format(float(self.percentage_value) * 32000 / 100))
            self.row_value += 1
            time.sleep(float(self.seconde_value))
            self.read_script()
        if self.channel_value == "2":
            self.connection.flow_2.writeParameter(9, "{:.0f}".format(float(self.percentage_value) * 32000 / 100))
            self.row_value += 1
            time.sleep(float(self.seconde_value))
            self.read_script()
        else:
            self.connection.flow_3.writeParameter(9, "{:.0f}".format(float(self.percentage_value) * 32000 / 100))
            self.row_value += 1
            time.sleep(float(self.seconde_value))
            self.read_script()

    def open_chart(self):
        # self.main_run = False
        app = Chart.App(self.connection,self.serienummer)
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        print("test")
        app.mainloop()

    def open_config(self):
        self.main_run = False
        app = main.Configuration(self.connection,self.serial)
        app.protocol("WM_DELETE_WINDOW", lambda: (app.destroy(), self.switch_main_run(), self.thread()))
        app.mainloop()
    def switch_main_run(self):
        self.main_run = not self.main_run

    def close_app(self):
        terminate_existing_main_processes()


class Configuration(ctk.CTk):
    def __init__(self,con,sensor):
        super().__init__()
        self.con = con
        self.sensor = sensor
        self.title("Configuration")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.option_add("*TCombobox*Listbox*Font", value=0)
        self.comport_sensor = None
        self.comport_flow = None
        self.channel_switch_var = False
        self.combobox()
        self.button_switch()
        self.text()

    def listen_comport(self):
        return serial.tools.list_ports.comports()

    def text(self):
        label_sensor = ctk.CTkLabel(self, text="Sensor").grid(row=0, column=1)
        label_flow = ctk.CTkLabel(self, text="Flow Controller").grid(row=0,column=0)
        button_exit = ctk.CTkButton(self, text="Save & Restart App", command=self.restart_app).grid(row=4, column=0, columnspan=2, pady=(30,0))

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

    def button_switch(self):
        with open("config.txt", "r") as r:
            self.read = r.readlines()
            self.modes_bool = str(self.read[2].split(": ")[1].split("\n")[0])
            dark_mode_state = ctk.BooleanVar(value=(self.modes_bool == "True"))
            print(dark_mode_state)
        self.dark_white_mode = ctk.CTkSwitch(self, text="Dark/Light mode", variable=dark_mode_state, command= self.set_dark_light_mode)
        self.dark_white_mode.grid(row=2, column=0, columnspan=2, padx=20, pady=20)
        with open("config.txt", "r") as r:
            self.read = r.readlines()
            self.modes_bool = str(self.read[6].split(": ")[1].split("\n")[0])
            self.channel_mode_state = ctk.BooleanVar(value=(self.modes_bool == "True"))
        channel_text = ctk.CTkLabel(self, text="K1/K2/K3    K4/K6").place(x=188, y=110)
        self.channel_switch = ctk.CTkSwitch(self, text="", variable=self.channel_mode_state, command=self.channel)
        self.channel_switch.grid(row=3, column=0, columnspan=2, padx=(80,20), pady=20)

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

    def channel(self):
        with open("config.txt", "r") as file:
            value = file.readlines()[6].split(": ")[1].split("\n")[0]
        if value == "True":
            print("K1/K2/K3")
            self.sensor.write(b"\x02\x38\x31\x30\x30\x30\x33\x62\x03")
            value = "False"
        else:
            print("K4/K6")
            self.sensor.write(b"\x02\x38\x31\x30\x30\x31\x33\x61\x03")
            value = "True"
        with open("config.txt", "w") as w:
            self.read[6] = f"Channel: {value}\n"
            w.writelines(self.read)

    def restart(self):
        main.MainApp.background_color(ctk.CTk)

    def restart_app(self):
        os.execv(sys.executable, [sys.executable] + sys.argv)

    def destroy(self):
        super().destroy()

if __name__ == "__main__":
    # log_file_path = r"C:\Users\rijken.b\Documents\Test Systeem\debug_log.txt"
    #
    # log_file = open(log_file_path, "a")
    # os.dup2(log_file.fileno(), 1)  # Redirect stdout
    # os.dup2(log_file.fileno(), 2)  # Redirect stderr
    con = Connection()
    mainapp = File_check(con)
    mainapp.mainloop()