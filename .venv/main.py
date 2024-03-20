import customtkinter as ctk
import tkinter as tk
import Chart
import Sensor
import sys
import main
import propar
import threading
import time
from connection import Connection

class First_check(ctk.CTk):
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
        self.geometry("500x500")
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
        self.frame_button.grid(row=0, column=0, padx=(20, 0), pady=(20, 0))
        self.frame_status = ctk.CTkFrame(self)
        self.frame_status.grid(row=0, column=1, padx=(66, 0))
        self.frame_ppm = ctk.CTkFrame(self)
        self.frame_ppm.grid(row=1, column=0, padx=(20,0), sticky="n")
        self.information_title = ctk.CTkFrame(self)
        self.information_title.grid(row=2, column=0, columnspan=2,sticky="n",pady=(20,0))
        self.information_tab = ctk.CTkFrame(self)
        self.information_tab.grid(row=3, column=0,columnspan=2, sticky="n")

    def thread(self):
        sensor_thread = threading.Thread(target=self.sensor_thread)
        sensor_thread.daemon = True
        sensor_thread.start()
        flow_thread = threading.Thread(target=self.flow_thread)
        flow_thread.daemon = True
        flow_thread.start()

    def buttons(self):
        self.chart_button = ctk.CTkButton(master=self.frame_button, text="Chart", command=self.open_chart)
        self.chart_button.grid(row=0,column=0,padx=30,pady=(20,20))
        # self.PPM_Meter_button = ctk.CTkButton(master=self.frame_button, text="PPM Meter", command=self.open_PPM_Meter)
        # self.PPM_Meter_button.grid(row=1,column=0,padx=30,pady=(0,20))
        self.start_button = ctk.CTkButton(self, text="Start", width=300, height=100, command=self.start_program)
        self.start_button.grid(row=4, column=0, columnspan=2, padx=(20,0),pady=(200,0), sticky="nsew")

    def status_info(self):
        self.status_sensor = ctk.CTkCheckBox(master=self.frame_status, text="Sensor", state="disabled",variable=self.status_sensor_var)
        self.status_sensor.grid(row=0, column=0,padx=40,pady=(20,5),sticky="w")
        self.status_flow = ctk.CTkCheckBox(master=self.frame_status, text="Flow controller", state="disabled", variable=self.status_flow_var)
        self.status_flow.grid(row=1, column=0,padx=40,pady=(4,20),sticky="w")

    def sensor_info_con(self):
        try:
            self.connection.sensor.write(b"\x02\x30\x39\x30\x30\x34\x34\x30\x30\x30\x62\x03")
            self.serienummer = self.connection.sensor.read(1024).decode().split("0900")[1].split(" ")[0]
            if not len(self.serienummer) > 0:
                self.serienummer = "Unknown"
        except Exception:
            pass
    def sensor_info(self):
        title = ctk.CTkLabel(master=self.information_title, text="Sensor Information", font=ctk.CTkFont(size=15,weight="bold"))
        title.grid(row=0, column=0, padx=80, pady=5)
        self.version = ctk.CTkLabel(master=self.information_tab, text=f"Version\n{self.serienummer}")
        self.version.grid(row=0, column=0)

    def sensor_thread(self):
        time.sleep(3)
        while self.main_run:
            try:
                self.serial = self.connection.sensor
                self.serial.write(b"\x02\x30\x35\x30\x30\x30\x37\x03")
                response = self.serial.read(1024)
                if response:
                    self.sensor_info_con()
                    self.sensor_status = 1
                    self.ppm_value = response.decode().split()[5]
                    with open("transfer.txt", "w") as write:
                        write.writelines(self.ppm_value)
                    self.version.configure(text=f"Version\n{self.serienummer}")
                    self.ppm_meter_label.configure(text=f"PPM value\n\n{self.ppm_value}")
                else:
                    self.sensor_status = 0
                    self.version.configure(text="Version\nUnknown")
                    self.ppm_meter_label.configure(text="Sensor not found")
            except Exception as e:
                self.sensor_status = 0
                print(f"Error in sensor communication: {e}")
                self.version.configure(text="Version\nUnknown")
                self.ppm_meter_label.configure(text="Sensor not found")
                self.connection.initialize_sensor()
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
            self.update_status_ui(self.sensor_status, self.flow_status)
            time.sleep(1)
    def ppm_meter(self):
        self.ppm_meter_label = ctk.CTkLabel(master=self.frame_ppm, text="Connecting...")
        self.ppm_meter_label.grid(row=0,column=0,padx=20, pady=20)

    def update_status_ui(self, sensor_status, flow_status):
        self.status_sensor_var.set(sensor_status)
        self.status_flow_var.set(flow_status)

    def open_chart(self):
        # self.main_run = False
        app = Chart.App(self.connection)
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    # def open_PPM_Meter(self):
    #     self.main_run = False
    #     app = Sensor.App(self.connection)
    #     app.protocol("WM_DELETE_WINDOW", app.destroy)
    #     app.mainloop()

    def start_program(self):
        main.StartProgram(self.connection)


    def close_app(self):
        self.main_run = False
        super().destroy()
        exit()

if __name__ == "__main__":
    con = Connection()
    mainapp = First_check(con)
    mainapp.mainloop()

class StartProgram():
        def __init__(self,con):
            self.row_value = 0
            self.connection = con
            self.check_connection()
            self.read_script_thread()

        def check_connection(self):
            self.flow1 = self.connection.flow_1
            value_1 = self.flow1.readParameter(8)
            if value_1 is not None and value_1 >= 0:
                self.flow_value_1 = value_1
            else:
                self.flow_value_1 = 0

            self.flow2 = self.connection.flow_2
            value_2 = self.flow2.readParameter(8)
            if value_2 is not None and value_2 >= 0:
                self.flow_value_2 = value_2
            else:
                self.flow_value_2 = 0

            self.flow3 = self.connection.flow_3
            value_3 = self.flow3.readParameter(8)
            if value_3 is not None and value_3 >= 0:
                self.flow_value_3 = value_3
            else:
                self.flow_value_3 = 0

        def read_script_thread(self):
            thread = threading.Thread(target=self.read_script)
            thread.daemon = True
            thread.start()

        def read_script(self):
            with open("script.txt", "r") as read:
                read_value = read.readlines()
                if not self.row_value < len(read_value):
                    print("End of file reached")
                    return
                self.seconde_value = read_value[self.row_value].split(". ")[1].split(", ")[0].strip()
                self.percentage_value = read_value[self.row_value].split(", ")[1].strip()
                self.channel_value = read_value[self.row_value].split(", ")[2].strip()
            print(f"Sec: {self.seconde_value}  | Flow: ","{:.0f}".format(int(self.percentage_value) * 32000 / 100),f" | Channel: {self.channel_value}")
            self.set_value()

        def set_value(self):
            self.flow1.writeParameter(9, "{:.0f}".format(int(self.percentage_value) * 32000 / 100))
            self.row_value += 1
            time.sleep(int(self.seconde_value))
            self.read_script()
