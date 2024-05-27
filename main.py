# Standard library
import datetime
import os
import sys
import threading
import time
import subprocess
# Third-Party library
import openpyxl
import xlwings
import psutil
import serial
import serial.tools.list_ports
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk
import tkchart
import requests
import logging
# Uncommon library
from connection import Connection
import Chart
from Config import readfile_value, text_config

execute_path = os.path.abspath(sys.argv[0])
icon = os.path.dirname(execute_path) + r"\skalar_analytical_bv_logo_Zoy_icon.ico"
version = None
def set_priority():
    import win32api, win32process, win32con

    pid = win32api.GetCurrentProcessId()
    handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, True, pid)
    win32process.SetPriorityClass(handle, win32process.HIGH_PRIORITY_CLASS)

def log():
    open('log.txt', 'w').truncate()
    while True:
        logging.basicConfig(filename='log.txt', level=logging.NOTSET)

        # Redirect stdout and stderr to the log file
        sys.stdout = sys.stderr = open('log.txt', 'a')

def update():
    global version
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json'
    }

    response = requests.get("https://api.github.com/repos/HighW0rks/Sensor_System/releases/latest", headers=headers)
    if response.status_code == 200:
        latest_release = response.json()
    else:
        print("Failed to retrieve the latest release.")
        return

    if latest_release:
        response = requests.get("https://api.github.com/repos/HighW0rks/Sensor_System/tags", headers=headers)
        if response.status_code == 200:
            latest_tag = response.json()[0]['name']
            version = latest_tag
        else:
            file_check()

        if latest_tag != readfile_value(12):
            UpdateApp(latest_tag).mainloop()
        else:
            file_check()


class UpdateApp(ctk.CTk):
    def __init__(self, latest_tag):
        super().__init__()
        self.title("Update")
        self.tag = latest_tag
        self.resizable(height=False, width=False)
        self.iconbitmap(icon)
        self.background_color()
        self.app()

    def background_color(self):
        # Set application appearance mode based on configuration
        if readfile_value(7) == "False":
            ctk.set_appearance_mode("Light")  # Set light mode
        else:
            ctk.set_appearance_mode("Dark")  # Set dark mode

    def app(self):
        ctk.CTkLabel(self, text="New update available!").grid(row=0, column=0, sticky="s")
        ctk.CTkLabel(self, text=f"{readfile_value(12)}                {self.tag}").grid(row=1, column=0, sticky="s")
        ctk.CTkLabel(self, text="➞", font=ctk.CTkFont(size=25)).place(x=133, y=28)
        ctk.CTkButton(self, text="Update", command=self.start_update, width=250, height=40,font=ctk.CTkFont(size=20)).grid(row=2, column=0, padx=20, pady=20, sticky="n")
        ctk.CTkButton(self, text="skip update", command=self.skip).grid(row=2, column=0, pady=(80, 20), sticky="s")

    def start_update(self):
        try:
            # Replace 'Update.exe' with the full path if it's not in the current directory
            os.system('start /B Update.exe')
        except Exception as e:
            print(f"Error: {e}")
            for i in range(5):
                print("App will destroy in 5 sec:")
                print(i + 1)
            terminate_existing_main_processes()

    def skip(self):
        self.destroy()
        file_check()


def file_check():
    # location = readfile_value(8)
    # file_locations = [
    #     f"{location}",
    #     f"{location}/2SN100224",
    #     f"{location}/2SN1001073",
    #     f"{location}/2SN1001098",
    #     f"{location}/2SN100224/425 --- test Data QC 2SN100224.xlsx",
    #     f"{location}/2SN100224/Scripts",
    #     f"{location}/2SN100224/Scripts/Script Ch1.txt",
    #     f"{location}/2SN100224/Scripts/Script Ch2.txt",
    #     f"{location}/2SN100224/Scripts/Script Ch3.txt",
    #     f"{location}/2SN1001073/425 --- test Data QC 2SN1001073.xlsx",
    #     f"{location}/2SN1001073/Scripts",
    #     f"{location}/2SN1001073/Scripts/Script Ch1.txt",
    #     f"{location}/2SN1001073/Scripts/Script Ch2.txt",
    #     f"{location}/2SN1001073/Scripts/Script Ch3.txt",
    #     f"{location}/2SN1001073/Scripts/Script Ch4.txt",
    #     f"{location}/2SN1001073/Scripts/Script Ch6.txt",
    #     f"{location}/2SN1001098/425 --- test Data QC 2SN1001098 Goud.xlsx",
    #     f"{location}/2SN1001098/Scripts",
    #     f"{location}/2SN1001098/Scripts/Script Ch1.txt",
    #     f"{location}/2SN1001098/Scripts/Script Ch2.txt",
    #     f"{location}/2SN1001098/Scripts/Script Ch3.txt",
    #     f"{location}/2SN1001098/Scripts/Script Ch4.txt",
    #     f"{location}/2SN1001098/Scripts/Script Ch6.txt"
    # ]
    # for i in range(len(file_locations)):
    #     file = file_locations[i]
    #     if not os.path.exists(file):
    #         if i == 0:
    #             mainapp = FileApp(file, True)
    #             mainapp.mainloop()
    #             return
    #         else:
    #
    #             mainapp = FileApp(file)
    #             mainapp.mainloop()
    #             return
    con = Connection()
    app = MainApp(con)  # Create an instance of the main application
    app.mainloop()


class FileApp(ctk.CTk):
    def __init__(self, file, main_folder=False):
        super().__init__()
        self.resizable(width=False, height=False)  # Disable window resizing
        self.title("Error")  # Set window title
        self.iconbitmap(icon)  # Set window icon
        self.file = file
        self.main_folder = main_folder
        self.text_frame = None
        self.Location_button = None
        self.Error_text = None
        self.restart_button = None
        self.new_folder = None
        self.background_color()
        self.error()

    def background_color(self):
        # Set application appearance mode based on configuration
        if readfile_value(7) == "False":
            ctk.set_appearance_mode("Light")  # Set light mode
        else:
            ctk.set_appearance_mode("Dark")  # Set dark mode

    def error(self):
        print("Error start")
        self.text_frame = ctk.CTkFrame(self)
        self.text_frame.grid(row=0, column=0, padx=20, pady=(20, 0))
        if len(self.file) <= 2:
            self.file = "config empty"
        if self.main_folder:
            self.Error_text = ctk.CTkLabel(master=self.text_frame, text=f"The main folder is not found:\n{self.file}")
            self.Error_text.grid(row=0, column=0, padx=20, pady=20)
            self.Location_button = ctk.CTkButton(self, text="Select a folder", command=self.folder_location_config)
            self.Location_button.grid(row=1, column=0, padx=20, pady=(20, 0))
        else:
            self.Error_text = ctk.CTkLabel(master=self.text_frame, text=f"File/Folder doesn't exist:\n{self.file}")
            self.Error_text.grid(row=0, column=0, padx=20, pady=20)
        self.restart_button = ctk.CTkButton(self, text="Continue...", command=self.destroy)
        self.restart_button.grid(row=2, column=0, pady=(20, 20))

    def folder_location_config(self):
        self.new_folder = tk.filedialog.askdirectory()
        text_config(8, self.new_folder)

    def destroy(self):
        super().destroy()
        file_check()


class MainApp(ctk.CTk):
    def __init__(self, con):
        global version
        super().__init__()
        self.iconbitmap(icon)  # Set application icon
        self.geometry("500x600")  # Set window geometry
        self.resizable(width=False, height=False)  # Disable window resizing
        self.title("Skalar Saxon Tester")  # Set window title
        self.status_sensor_var = tk.IntVar(value=0)  # Variable to track sensor status
        self.status_flow_var = tk.IntVar(value=0)  # Variable to track flow status
        self.main_run = True  # Flag to indicate if the main loop is running
        self.running = True  # Flag to indicate if the program is running
        self.program_run = True  # Flag to indicate if the program is in running state
        self.connection = con  # Connection object
        self.sensor_status = 0  # Variable to store sensor status
        self.flow_status = 0  # Variable to store flow status
        self.time = 0  # Variable to store time
        self.row_value = 0  # Variable to store row value
        self.version = version
        ctk.CTkLabel(self, text=self.version).place(x=455, y=573)
        self.channel_option = None  # Variable to store selected channel option
        self.sensor_type = None
        self.frame_button = None
        self.frame_status = None
        self.frame_ppm = None
        self.frame_zero_point = None
        self.information_title = None
        self.information_serienummer = None
        self.information_temperature = None
        self.information_channel_status = None
        self.chart_button = None
        self.start_button = None
        self.stop_button = None
        self.status_sensor = None
        self.status_flow = None
        self.sensor_version = None
        self.serienummer_label = None
        self.status_channel_label = None
        self.temperature_label = None
        self.select_type_sensor = None
        self.folder_location = None
        self.excel_file = None
        self.channel_1 = None
        self.channel_2 = None
        self.channel_3 = None
        self.channel_4 = None
        self.ppm_value = None
        self.channel_6 = None
        self.temperature = None
        self.air_pressure = None
        self.flow1_value = None
        self.flow2_value = None
        self.flow3_value = None
        self.ppm_meter_label = None
        self.zero_point_button = None
        self.folder = None
        self.channel = None
        self.excel_row = None
        self.workbook = None
        self.worksheet = None
        self.current_time = None
        self.seconden_value = None
        self.percentage_value = None
        self.channel_value = None
        self.channel_status = "Unknown"  # Variable to store channel status
        self.serienummer = "Unknown"
        self.background_color()  # Set application background color
        self.loading()  # Load application components
        self.protocol("WM_DELETE_WINDOW", terminate_existing_main_processes)  # Define action on window close

    def loading(self):
        # Load application components
        self.frame()
        self.ppm_meter()
        self.status_info()
        self.sensor_info_con()
        self.sensor_info()
        self.buttons()
        self.thread()
        self.set_sensor()
        self.zero_point()

    def frame(self):
        # Create frames for organizing widgets
        self.frame_button = ctk.CTkFrame(self)
        self.frame_button.grid(row=0, column=0, padx=(15, 0), pady=(20, 20), sticky="n")
        self.frame_status = ctk.CTkFrame(self)
        self.frame_status.grid(row=0, column=1, padx=(66, 0), pady=(20, 0), sticky="n")
        self.frame_ppm = ctk.CTkFrame(self)
        self.frame_ppm.grid(row=1, column=1, padx=(66, 0), sticky="n")
        self.frame_zero_point = ctk.CTkFrame(self)
        self.frame_zero_point.grid(row=1, column=0, padx=(15, 0))
        self.information_title = ctk.CTkFrame(self)
        self.information_title.grid(row=2, column=0, columnspan=2, sticky="n", padx=(20, 0), pady=(20, 10))
        self.information_serienummer = ctk.CTkFrame(self)
        self.information_serienummer.grid(row=3, column=0, padx=(20, 0), sticky="n")
        self.information_temperature = ctk.CTkFrame(self)
        self.information_temperature.grid(row=3, column=0, columnspan=2, padx=(20, 0), sticky="n")
        self.information_channel_status = ctk.CTkFrame(self)
        self.information_channel_status.grid(row=3, column=1, padx=(42, 0), sticky="n")

    def thread(self):
        # Start threads for sensor and flow
        threading.Thread(target=self.sensor_thread, daemon=True).start()
        threading.Thread(target=self.flow_thread, daemon=True).start()

    def buttons(self):
        ctk.CTkLabel(master=self.frame_button, text="", height=0).grid(row=0, column=0, padx=100)
        self.chart_button = ctk.CTkButton(master=self.frame_button, text="Chart", command=self.open_chart)
        self.chart_button.grid(row=1, column=0, pady=(0, 15))
        ctk.CTkButton(master=self.frame_button, text="Sensor", command=self.open_sensor).grid(row=2, column=0,
                                                                                              pady=(0, 15))
        ctk.CTkButton(master=self.frame_button, text="Config", command=self.open_config).grid(row=3, column=0,
                                                                                              pady=(0, 15))
        self.start_button = ctk.CTkButton(self, text="Start", width=300, height=100, state="disabled",command=self.start_program)
        self.start_button.grid(row=6, column=0, columnspan=2, padx=(20, 0),pady=(10, 20), sticky="nsew")
        self.stop_button = ctk.CTkButton(self, text="Stop", width=300, height=100, command=lambda: self.start_stop(1))

    def status_info(self):
        ctk.CTkLabel(master=self.frame_status, text="", height=0).grid(row=0, column=0, padx=100)
        self.status_sensor = ctk.CTkCheckBox(master=self.frame_status, text="Sensor", state="disabled",
                                             variable=self.status_sensor_var)
        self.status_sensor.grid(row=1, column=0, padx=40, pady=(0, 5), sticky="w")
        self.status_flow = ctk.CTkCheckBox(master=self.frame_status, text="Flow controller", state="disabled",
                                           variable=self.status_flow_var)
        self.status_flow.grid(row=2, column=0, padx=40, pady=(4, 20), sticky="w")

    def sensor_info_con(self):
        if self.connection.status_sensor:
            if self.main_run:
                try:
                    self.connection.sensor.write(b"\x02\x30\x39\x30\x30\x34\x34\x30\x30\x30\x62\x03")
                    read = self.connection.sensor.read(1024)
                    if len(read) == 0:
                        self.serienummer = "Unknown"
                    elif len(read) > 24:
                        self.serienummer = "Error"
                    else:
                        self.serienummer = str(read.decode()[5:13])
                    self.connection.sensor.write(b"\x02\x36\x38\x30\x30\x30\x63\x03")
                    read = self.connection.sensor.read(1024).decode()
                    self.sensor_version = str(read[61:81])
                    self.sensor_type = str(read.split()[4])
                except Exception as e:
                    print("test1", e)
                    pass

    def status_channel(self):
        try:
            if len(self.ppm_value) > 0:
                if self.channel_4 == self.ppm_value or self.channel_6 == self.ppm_value:
                    self.status_channel_label.configure(text="Channel\nK1/K2/K3")
                else:
                    self.status_channel_label.configure(text="Channel\nK4/K6")
        except Exception:
            pass

    def sensor_info(self):
        ctk.CTkLabel(master=self.information_title, text="Sensor Information",
                     font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=80, pady=5)
        self.serienummer_label = ctk.CTkLabel(master=self.information_serienummer,
                                              text=f"Serie Nummer\n{self.serienummer}")
        self.serienummer_label.grid(row=0, column=0, padx=20, pady=10)
        self.status_channel_label = ctk.CTkLabel(master=self.information_channel_status, text="Channel\nUnknown")
        self.status_channel_label.grid(row=0, column=0, padx=20, pady=10)
        self.temperature_label = ctk.CTkLabel(master=self.information_temperature, text="Temperature\nUnknown")
        self.temperature_label.grid(row=0, column=0, padx=20, pady=10)

    def set_sensor(self):
        self.select_type_sensor = ctk.CTkComboBox(self, values=["V153", "V176", "V200"], justify="center",command=self.set_channel)
        self.select_type_sensor.grid(row=4, column=0, columnspan=2, padx=(20, 0), pady=(20, 0))
        self.select_type_sensor.set("Select a sensor")

    def set_channel(self, event=None):
        type_sensor = self.select_type_sensor.get()
        if type_sensor == "V153":
            channel_values = ["Ch1", "Ch2", "Ch3"]
            self.folder_location = "2SN100224"
            self.excel_file = "425 --- test Data QC 2SN100224.xlsx"

        elif type_sensor == "V176":
            channel_values = ["Ch1", "Ch2", "Ch3", "Ch4", "Ch6"]
            self.folder_location = "2SN1001073"
            self.excel_file = "425 --- test Data QC 2SN1001073.xlsx"

        else:
            channel_values = ["Ch1", "Ch2", "Ch3", "Ch4", "Ch6"]
            self.folder_location = "2SN1001098"
            self.excel_file = "425 --- test Data QC 2SN1001098 Goud.xlsx"

        # Destroy the existing channel_option if it exists
        if self.channel_option is not None:
            self.channel_option.destroy()

        # Create a new channel_option ComboBox with updated values
        self.channel_option = ctk.CTkComboBox(self, values=channel_values, justify="center")
        self.channel_option.grid(row=5, column=0, columnspan=2, padx=(20, 0), pady=(20, 0))
        self.channel_option.set("Select a channel")

    def sensor_thread(self):
        # Continuous loop for sensor communication
        while self.main_run:
            # Check if channel label status is unknown and update if necessary
            if self.status_channel_label.cget('text') == "Channel\nUnknown":
                self.status_channel()

            # Check if serial number is unknown and update if necessary
            if self.serienummer_label.get() == "Unknown":
                self.sensor_info_con()
            elif self.serienummer_label.get() == "Error":
                self.sensor_info_con()

            try:
                # Send command to sensor and read response
                self.connection.sensor.write(b"\x02\x30\x35\x30\x30\x30\x37\x03")
                response = self.connection.sensor.read(1024).decode()
                configure = False

                if response:
                    print("Response: ", response)
                    # Parse response and update values
                    self.sensor_status = 1
                    self.channel_1, self.channel_2, self.channel_3, self.channel_4, self.ppm_value, self.channel_6, self.temperature, self.air_pressure = response.split()[
                                                                                                                                                          1:9]

                    # Write PPM value to a file
                    with open("transfer.txt", "w") as write:
                        write.write(self.ppm_value)

                    # Update labels with new values
                    self.serienummer_label.configure(text=f"Serie Nummer\n{self.serienummer}")
                    self.ppm_meter_label.configure(text=f"PPM value\n{self.ppm_value}")
                    self.temperature_label.configure(text=f"Temperature\n{self.temperature}")
                else:
                    # Set sensor status to unknown if no response
                    self.serienummer = "Unknown"
                    print("No serienummer found")
                    self.sensor_status = 0
                    configure = True
            except Exception as e:
                # Handle errors in sensor communication
                self.serienummer = "Unknown"
                print(e)
                self.sensor_status = 0
                configure = True
                time.sleep(1)

            # Reconfigure sensor if necessary
            if configure:
                self.sensor_info_label_configure()
                self.connection.initialize_sensor()

    def sensor_info_label_configure(self):
        # Configurates the labels
        self.serienummer_label.configure(text="Serie Nummer\nUnknown")
        self.ppm_meter_label.configure(text="Sensor not found")
        self.temperature_label.configure(text="Temperature\nUnknown")
        self.status_channel_label.configure(text="Channel\nUnknown")

    def flow_thread(self):
        # Continuous loop for flow communication
        while self.main_run:
            try:
                # Read flow values from flow controllers
                self.flow1_value = self.connection.flow_1.readParameter(8)
                self.flow2_value = self.connection.flow_2.readParameter(8)
                self.flow3_value = self.connection.flow_3.readParameter(8)

                # Update flow status based on received values
                if self.flow1_value is not None:
                    self.flow_status = True
                else:
                    self.flow_status = False
            except Exception:
                self.flow_status = False

            # Schedule status update and sleep for 1 second
            self.after(1000, self.status_update)
            time.sleep(1)

    def status_update(self):
        # Update status variables and UI elements
        self.status_sensor_var.set(self.sensor_status)
        self.status_flow_var.set(self.flow_status)
        try:
            # Enable start button if sensor and flow are connected and channel is selected
            if self.sensor_status and self.flow_status and self.channel_option.get() != "Select a channel":
                self.start_button.configure(state="normal")
            else:
                self.start_button.configure(state="disabled")
        except Exception:
            pass

        # Enable zero point button if sensor is connected
        if self.sensor_status:
            self.zero_point_button.configure(state="normal")
        else:
            self.zero_point_button.configure(state="disabled")

    def ppm_meter(self):
        # Create empty label frame for PPM meter
        ctk.CTkLabel(master=self.frame_ppm, text="").grid(row=0, column=0, padx=100)

        # Initialize PPM meter label
        self.ppm_meter_label = ctk.CTkLabel(master=self.frame_ppm, text="Connecting...")
        self.ppm_meter_label.grid(row=0, column=0, pady=17)

    def zero_point(self):
        # Create empty label frame for zero point
        ctk.CTkLabel(master=self.frame_zero_point, text="", height=0).grid(row=0, column=0, padx=100)

        # Initialize zero point button
        self.zero_point_button = ctk.CTkButton(master=self.frame_zero_point, text="Zero Point", state="disabled",
                                               command=self.send_zero_point_command)
        self.zero_point_button.grid(row=1, column=0, pady=(0, 17))

    def send_zero_point_command(self):
        # Send zero point command to sensor
        self.main_run = False
        self.connection.sensor.write(b"\x02\x34\x30\x30\x30\x30\x36\x03")
        time.sleep(1)
        self.main_run = True

    def start_program(self):
        # Start the program
        self.start_stop(0)
        self.row_value = 0

        # Read folder configuration
        self.folder = readfile_value(8)

        # Get selected channel
        self.channel = self.channel_option.get()

        # Start Excel
        self.start_excel()

        # Start reading script thread
        threading.Thread(target=self.read_script, daemon=True).start()

    def start_excel(self):
        # Initialize Excel settings
        self.excel_row = 2  # Start from row 2
        self.workbook = openpyxl.load_workbook(
            fr"{self.folder}/{self.folder_location}/{self.excel_file}")  # Load workbook
        self.worksheet = self.workbook["Sheet1"]  # Select Sheet1
        self.worksheet['E2'] = self.serienummer  # Set serial number
        self.worksheet['E3'] = self.folder_location  # Set model
        self.worksheet['E4'] = self.sensor_version  # Set sensor version
        self.worksheet['J3'] = datetime.datetime.now().strftime("%Y-%m-%d")  # Set date

        # Select appropriate worksheet based on the selected channel
        if self.channel == "Ch1":
            self.worksheet = self.workbook["Ch.1 2-60%"]
        elif self.channel == "Ch2":
            self.worksheet = self.workbook["Ch.2 500-20000 2%"]
        elif self.channel == "Ch3":
            self.worksheet = self.workbook["Ch.3 0-1000 2%"]
        elif self.channel == "Ch4":
            self.worksheet = self.workbook["Ch.4 500-20000 2%"]
        elif self.channel == "Ch6":
            self.worksheet = self.workbook["Ch.6 0-20%"]

        # Set column headers
        self.worksheet['A1'] = 'Time'
        self.worksheet['B1'] = 'Channel 1'
        self.worksheet['C1'] = 'Channel 2'
        self.worksheet['D1'] = 'Channel 3'
        self.worksheet['E1'] = 'Channel 4'
        self.worksheet['F1'] = 'Auto channel'
        self.worksheet['G1'] = 'Channel 6'
        self.worksheet['H1'] = 'Temperature'
        self.worksheet['I1'] = 'Air Pressure'

        threading.Thread(target=self.write_excel, daemon=True).start()

    def write_excel(self):
        # Continuous loop for writing Excel data
        while self.running:
            # Write data to Excel
            self.get_time()
            self.worksheet[f'A{self.excel_row}'] = self.current_time
            self.worksheet[f'B{self.excel_row}'] = self.channel_1
            self.worksheet[f'C{self.excel_row}'] = self.channel_2
            self.worksheet[f'D{self.excel_row}'] = self.channel_3
            self.worksheet[f'E{self.excel_row}'] = self.channel_4
            self.worksheet[f'F{self.excel_row}'] = float(self.ppm_value)
            self.worksheet[f'G{self.excel_row}'] = self.channel_6
            self.worksheet[f'H{self.excel_row}'] = self.temperature
            self.worksheet[f'I{self.excel_row}'] = self.air_pressure

            # Increment row
            self.excel_row += 1

            # Wait for 1 second before writing next data
            time.sleep(1)

    def get_time(self):
        # Get current time in the specified format
        self.current_time = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

    def read_script(self):
        # Read script file and execute commands
        if self.running:
            # Open script file
            with open(fr"{self.folder}/{self.folder_location}/Scripts/Script {self.channel}.txt", "r") as read:
                read_value = read.readlines()
                # Check if end of file is reached
                if not self.row_value < len(read_value):
                    print("End of file reached")
                    # Save Excel file with timestamp
                    self.get_time()
                    if not os.path.exists(f"{self.folder}/{self.folder_location}/{self.serienummer}"):
                        os.mkdir(f"{self.folder}/{self.folder_location}/{self.serienummer}")
                    self.save_location = f"{self.folder}/{self.folder_location}/{self.serienummer}/{self.channel}_{self.current_time}.xlsx"
                    self.workbook.save(self.save_location)
                    # Stop program
                    self.start_stop(1)
                    self.after(0,self.open_validate)
                else:
                    # Parse script values
                    self.seconden_value = read_value[self.row_value].split(". ")[1].split("; ")[0].strip()
                    self.percentage_value = read_value[self.row_value].split("; ")[1].strip()
                    self.channel_value = read_value[self.row_value].split("; ")[2].strip()
                    # Print values for debugging
                    print(f"Sec: {self.seconden_value}  | Flow: ",
                          "{:.0f}".format(float(self.percentage_value) * 32000 / 100),
                          f" | Channel: {self.channel_value}")
                    # Set flow value
                    self.set_value()
        else:
            print("Excel has been written")

    def set_value(self):
        # Set flow value based on channel
        try:
            if self.running:
                if self.channel_value == "1":
                    self.connection.flow_1.writeParameter(9,
                                                          "{:.0f}".format(float(self.percentage_value) * 32000 / 100))
                    self.row_value += 1
                    time.sleep(float(self.seconden_value))
                    self.read_script()
                elif self.channel_value == "2":
                    self.connection.flow_2.writeParameter(9,
                                                          "{:.0f}".format(float(self.percentage_value) * 32000 / 100))
                    self.row_value += 1
                    time.sleep(float(self.seconden_value))
                    self.read_script()
                else:
                    self.connection.flow_3.writeParameter(9,
                                                          "{:.0f}".format(float(self.percentage_value) * 32000 / 100))
                    self.row_value += 1
                    time.sleep(float(self.seconden_value))
                    self.read_script()
        except Exception:
            pass

    def start_stop(self, x):
        # Start or stop the program based on the input parameter 'x'
        if x == 0:
            # If x is 0, show the stop button and hide the start button
            self.stop_button.grid(row=6, column=0, columnspan=2, padx=(20, 0), pady=(10, 20), sticky="nsew")
            self.start_button.grid_forget()
            self.running = True  # Set running flag to True
        if x == 1:
            # If x is 1, show the start button and hide the stop button
            self.start_button.grid(row=6, column=0, columnspan=2, padx=(20, 0), pady=(10, 20), sticky="nsew")
            self.stop_button.grid_forget()
            self.running = False  # Set running flag to False

    def open_chart(self):
        # Open the chart application
        app = Chart.App(self.connection, self.sensor_type, self.serienummer)
        app.protocol("WM_DELETE_WINDOW", app.destroy)  # Handle window close event
        app.mainloop()

    def open_config(self):
        # Open the configuration application
        self.main_run = False  # Set main run flag to False
        app = Configuration(self.connection, self.connection.sensor)
        # Handle window close event, switch main run flag, and start thread on exit
        app.protocol("WM_DELETE_WINDOW",lambda: (app.destroy(), self.switch_main_run(), self.thread(), self.reset_channel_info()))
        app.mainloop()

    def open_validate(self):
        app = validate(self.save_location, self.channel)
        app.protocol("WM_DELETE_WINDOW", app.destroy)  # Handle window close event
        app.mainloop()

    def background_color(self):
        # Set application appearance mode based on configuration
        if readfile_value(7) == "False":
            ctk.set_appearance_mode("Light")  # Set light mode
        else:
            ctk.set_appearance_mode("Dark")  # Set dark mode

    def open_sensor(self):
        app = SensorApp()
        app.protocol("WM_DELETE_WINDOW", app.destroy)  # Handle window close event
        app.mainloop()

    def reset_channel_info(self):
        # Resets the channel info for the K1/K2/K3 or K4/K6 gets displayed correctly
        self.status_channel_label.configure(text="Channel\nUnknown")
        self.ppm_value = None

    def switch_main_run(self):
        # Switch the value of the main run flag
        self.main_run = not self.main_run


class Configuration(ctk.CTk):
    def __init__(self, con, sensor):
        super().__init__()
        self.con = con
        self.sensor = sensor
        self.title("Configuration")
        self.iconbitmap(icon)
        self.option_add("*TCombobox*Listbox*Font", value=0)
        self.comport_sensor = None
        self.comport_flow = None
        self.frame_comport = None
        self.frame_switch = None
        self.frame_folder = None
        self.flow_com = None
        self.sensor_com = None
        self.dark_white_mode = None
        self.channel_switch = None
        self.Folder_Location = None
        self.channel_switch_var = False
        self.frame()
        self.combobox()
        self.button_switch()
        self.text()
        self.directory()

    def frame(self):
        # Create frames for organizing widgets
        self.frame_comport = ctk.CTkFrame(self)
        self.frame_comport.grid(row=0, column=0, columnspan=2)
        self.frame_switch = ctk.CTkFrame(self)
        self.frame_switch.grid(row=1, column=0, pady=(10, 0))
        self.frame_folder = ctk.CTkFrame(self)
        self.frame_folder.grid(row=1, column=1)

    def text(self):
        # Add labels and buttons
        ctk.CTkLabel(master=self.frame_comport, text="Sensor").grid(row=0, column=1)
        ctk.CTkLabel(master=self.frame_comport, text="Flow Controller").grid(row=0, column=0)
        ctk.CTkButton(self, text="Save & Restart App", command=self.restart_app).grid(row=4, column=0, columnspan=2,
                                                                                      pady=(30, 0))

    def combobox(self):
        # Create combo boxes for selecting serial ports
        self.flow_com = ttk.Combobox(master=self.frame_comport, width=25, values=serial.tools.list_ports.comports())
        self.flow_com.grid(row=1, column=0)
        self.flow_com.set("Select a comport")
        self.sensor_com = ttk.Combobox(master=self.frame_comport, width=25, values=serial.tools.list_ports.comports())
        self.flow_com.bind('<<ComboboxSelected>>', self.get_flow)
        self.sensor_com.grid(row=1, column=1)
        self.sensor_com.set("Select a comport")
        self.sensor_com.bind('<<ComboboxSelected>>', self.get_sensor)

    def get_sensor(self):
        # Get selected sensor comport
        self.comport_sensor = self.sensor_com.get()
        value = self.comport_sensor.split(" ")[0]
        text_config(9, value)

    def get_flow(self):
        # Get selected flow controller comport
        self.comport_flow = self.flow_com.get()
        value = self.comport_flow.split(" ")[0]
        text_config(10, value)

    def button_switch(self):
        # Reads the config and makes an appearance switch
        dark_mode_state = ctk.BooleanVar(value=(readfile_value(7) == "False"))
        self.dark_white_mode = ctk.CTkSwitch(master=self.frame_switch, text="Dark/Light mode", variable=dark_mode_state,
                                             command=self.background_color)
        self.dark_white_mode.grid(row=2, column=0, columnspan=2, padx=20, pady=20)

        # Reads the config and makes a channel switch
        channel_mode_state = ctk.BooleanVar(value=(readfile_value(11) == "True"))
        self.channel_switch = ctk.CTkSwitch(master=self.frame_switch, text="K1/K2/K3 | K4/K6",
                                            variable=channel_mode_state, command=self.channel)
        self.channel_switch.grid(row=3, column=0, columnspan=2, pady=(0, 20))

    def background_color(self):
        # Set application appearance mode based on configuration
        if readfile_value(7) == "True":
            ctk.set_appearance_mode("Light")  # Set light mode
            text_config(7, "False")
        else:
            ctk.set_appearance_mode("Dark")  # Set dark mode
            text_config(7, "True")

    def channel(self):
        # Reads the config file and writes to the sensor to switch
        if readfile_value(11) == "True":
            print("Switched from K4/K6 to K1/K2/K3")
            self.sensor.write(b"\x02\x38\x31\x30\x30\x30\x33\x62\x03")
            value = "False"
        else:
            print("Switched from K1/K2/K3 to K4/K6")
            self.sensor.write(b"\x02\x38\x31\x30\x30\x31\x33\x61\x03")
            value = "True"
        text_config(11, value)

    def directory(self):
        # Add button for selecting main folder
        ctk.CTkButton(master=self.frame_folder, text="Select the main folder", command=self.folder_location).grid(row=0,
                                                                                                                  column=0)

    def folder_location(self):
        # Get folder location
        self.Folder_Location = tk.filedialog.askdirectory()
        # Sets folder location
        text_config(8, self.Folder_Location)

    def restart_app(self):
        script_path = sys.argv[0]

        # Launch a new instance of the script
        subprocess.Popen([sys.executable, script_path])

        # Terminate the current instance
        terminate_existing_main_processes()

    def destroy(self):
        # Destroys the application
        super().destroy()


class validate(ctk.CTk):
    def __init__(self, file, channel):
        super().__init__()
        self.title("Validate")
        self.resizable(width=False, height=False)
        self.iconbitmap(icon)
        self.file_location = file
        self.channel = channel
        self.read_excel()

    def read_excel(self):
        print(self.file_location)
        print(type(self.file_location))
        workbook = xlwings.Book({self.file_location})
        sheet = workbook.sheets["Sheet1"]
        if self.channel == "Ch1":
            corr = sheet.range('D27').value
            slope = sheet.range('D28').value
        elif self.channel == "Ch2":
            corr = sheet.range('D57').value
            slope = sheet.range('D58').value
        elif self.channel == "Ch3":
            corr = sheet.range('D78').value
            slope = sheet.range('D79').value
        elif self.channel == "Ch4":
            corr = sheet.range('D108').value
            slope = sheet.range('D109').value
        else:
            corr = sheet.range('D127').value
            slope = sheet.range('D128').value

        ctk.CTkLabel(self, text=corr).grid(row=0, column=0, padx=10, sticky="nw")
        ctk.CTklabel(self, text=slope).grid(row=1, column=0, padx=10, sticky="nw")
        if corr > float(0.99):
            ctk.CTkLabel(self, text="Corr passed").grid(row=0, column=1, padx=10, sticky="nw")
        else:
            ctk.CTkLabel(self, text="Corr failed").grid(row=0, column=1, padx=10, sticky="nw")
        if slope > float(0.99):
            ctk.CTkLabel(self, text="Slope passed").grid(row=1, colomn=0, padx=10, sticky="nw")
        else:
            ctk.CTkLabel(self, text="Slope failed").grid(row=1, column=1, padx=10, sticky="nw")

    def destroy(self):
        # Destroys the application
        super().destroy()


class SensorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PPM Chart")
        self.iconbitmap(icon)
        self.resizable(width=False, height=False)
        self.Place_Button_Y = 28
        self.Place_Button_X = 140
        try:
            self.ppm_value = [float(''.join(map(str, open("transfer.txt", "r").readlines())))]
        except Exception:
            self.ppm_value = [0]
        self.Menu_One = None
        self.Menu_Two = None
        self.Menu_Restart = None
        self.Menu_Quit = None
        self.Menu_Config = None
        self.value_x_axis = None
        self.value_x_steps = None
        self.value_y_axis = None
        self.value_y_steps = None
        self.Masterchart = None
        self.Drawline = None
        self.Frame_Input = None
        self.Close_Save_Button = None
        self.Menu_One_Show = True
        self.Menu_Two_Show = True
        self.main_run = True
        self.loading()

    def loading(self):
        self.topbar()
        self.chart()
        # threading.Thread(target=self.loop, daemon=True).start()

    def topbar(self):
        self.Menu_One = ctk.CTkButton(self, text="Exit", command=self.topbar_menu_one)
        self.Menu_One.grid(row=0, column=0, sticky="W")
        self.Menu_Two = ctk.CTkButton(self, text="Settings", command=self.topbar_menu_two)
        self.Menu_Two.grid(row=0, column=0, padx=140, sticky="W")

    def topbar_menu_one(self):
        if self.Menu_One_Show:
            self.Menu_Restart = ctk.CTkButton(self, text="Restart", command=self.restart)
            self.Menu_Restart.place(x=0, y=self.Place_Button_Y)
            self.Menu_Quit = ctk.CTkButton(self, text="Quit", command=self.destroy)
            self.Menu_Quit.place(x=0, y=self.Place_Button_Y * 2)
        else:
            self.Menu_Restart.place_forget()
            self.Menu_Quit.place_forget()
        self.Menu_One_Show = not self.Menu_One_Show

    def topbar_menu_two(self):
        if self.Menu_Two_Show:
            self.Menu_Config = ctk.CTkButton(self, text="Config", command=self.open_settings)
            self.Menu_Config.place(x=self.Place_Button_X, y=self.Place_Button_Y)
        else:
            self.Menu_Config.place_forget()
            try:
                self.select_type_sensor.destroy()
                self.channel_option.destroy()
            except Exception:
                pass
        self.Menu_Two_Show = not self.Menu_Two_Show

    def get_x_axis_values(self):
        self.value_x_axis = int(readfile_value(3))
        self.value_x_steps = int(readfile_value(4))
        self.value_y_axis = int(readfile_value(5))
        self.value_y_steps = int(readfile_value(6))

    def chart(self):
        self.get_x_axis_values()
        x_axis_values = tuple(num for num in range(self.value_x_steps, self.value_x_axis + 1, self.value_x_steps))
        self.Masterchart = tkchart.LineChart(
            master=self,
            width=800,
            height=400,
            axis_size=5,
            y_axis_section_count=int(self.value_y_axis / self.value_y_steps),
            x_axis_section_count=int(self.value_x_axis / self.value_x_steps),
            y_axis_label_count=10,
            x_axis_label_count=10,
            y_axis_data="PPM meting",
            x_axis_data="Seconden",
            x_axis_values=x_axis_values,
            y_axis_values=(0, self.value_y_axis),
            y_axis_precision=0,
            y_axis_section_color="#404040",
            x_axis_section_color="#404040",
            x_axis_font_color="#707070",
            y_axis_font_color="#707070",
            x_axis_data_font_color="lightblue",
            y_axis_data_font_color="lightblue",
            bg_color="#202020",
            fg_color="#202020",
            axis_color="#707070",
            data_font_style=("Arial", 15, "bold"),
            axis_font_style=("Arial", 10, "bold"),
            x_space=40,
            y_space=20,
            x_axis_data_position="side",
            y_axis_data_position="side"
        )
        self.Masterchart.grid(row=2, column=0, rowspan=3, padx=50, pady=50)
        self.Drawline = tkchart.Line(master=self.Masterchart,
                                     color="lightblue",
                                     size=2,
                                     fill="enabled")
        # self.area_label = ctk.CTkLabel(self, text="Total Area: N/A")
        # self.area_label.grid(row=5, column=0)

    # def loop(self):
    #     total_area = 0
    #     y = 0
    #     while self.main_run:
    #         y_old = self.ppm_value[0]
    #         self.Masterchart.show_data(data=self.ppm_value, line=self.Drawline)
    #         self.ppm_value = [float(''.join(map(str, open("transfer.txt", "r").readlines())))]
    #         time.sleep(self.value_x_steps)
    #         y_new = self.ppm_value[0]
    #         if y_old <= 0 and y_new <= 0:
    #             if total_area != 0:
    #                 print(f"Total Area recorded: {int(total_area)}")
    #                 total_area = 0
    #                 y = 0
    #         else:
    #             if y_old != y_new:
    #                 if y_old > y_new:
    #                     y = y_old - y_new
    #                     total_area += self.value_x_steps * ((y / 2) + y_old)
    #                 else:
    #                     y = y_new - y_old
    #                     total_area += self.value_x_steps * ((y / 2) + y_new)
    #             else:
    #                 y = 0
    #                 total_area += self.value_x_steps * y_new
    #         if not total_area == 0:
    #             self.area_label.configure(text=f"Total Area: {total_area}")
    #         print(f"Old: {y_old} | New: {y_new} | difference: {y} | Area: {total_area}")

    def open_settings(self):
        app = ConfigurationApp()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def restart(self):
        super().destroy()
        SensorApp().mainloop()

    def destroy(self):
        self.main_run = False
        super().destroy()


class ConfigurationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.iconbitmap(icon)
        self.X_as_middle = 800 / 2
        self.resizable(width=False, height=False)
        self.title("Configuration")
        self.title_label = None
        self.config_option_name = None
        self.Frame_Input = None
        self.Close_Save_Button = None
        self.x_text_sec = None
        self.x_text_grid = None
        self.x_seconden = None
        self.x_grid = None
        self.y_text_sec = None
        self.y_seconden = None
        self.y_grid = None
        self.y_text_grid = None
        self.loading()

    def loading(self):
        self.Frame_Input = ctk.CTkFrame(self)
        self.Frame_Input.grid(row=1, column=0, padx=(self.X_as_middle - self.Frame_Input.winfo_reqwidth()) / 2, pady=40,
                              stick="nsew")
        self.Close_Save_Button = ctk.CTkButton(self, text="Save & Close", command=self.destroy)
        self.Close_Save_Button.grid(row=2, column=0, pady=(0, 20))
        ctk.CTkLabel(master=self, text="Configuration", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0,
                                                                                                       sticky="N")
        self.section_input()

    def section_input(self):
        self.x_text_sec = ctk.CTkLabel(master=self.Frame_Input, text="X seconden")
        self.x_text_sec.grid(row=1, column=0, padx=(20, 0), pady=20, sticky="w")
        self.x_text_grid = ctk.CTkLabel(master=self.Frame_Input, text="X Grid lines")
        self.x_text_grid.grid(row=2, column=0, padx=(20, 0))
        self.x_seconden = ctk.CTkEntry(master=self.Frame_Input)
        self.x_seconden.insert(0, readfile_value(3))
        self.x_seconden.grid(row=1, column=1, padx=20)
        self.x_seconden.bind("<Return>", lambda event: text_config(3, self.x_seconden))
        self.x_grid = ctk.CTkEntry(master=self.Frame_Input)
        self.x_grid.insert(0, readfile_value(4))
        self.x_grid.grid(row=2, column=1, padx=20)
        self.x_grid.bind("<Return>", lambda event: text_config(4, self.x_grid))

        self.y_text_sec = ctk.CTkLabel(master=self.Frame_Input, text="Y seconden")
        self.y_text_sec.grid(row=3, column=0, padx=(20, 0), pady=20, sticky="w")
        self.y_text_grid = ctk.CTkLabel(master=self.Frame_Input, text="Y Grid lines")
        self.y_text_grid.grid(row=4, column=0, padx=(20, 0))
        self.y_seconden = ctk.CTkEntry(master=self.Frame_Input)
        self.y_seconden.insert(0, readfile_value(5))
        self.y_seconden.grid(row=3, column=1, padx=20)
        self.y_seconden.bind("<Return>", lambda event: text_config(5, self.y_seconden))
        self.y_grid = ctk.CTkEntry(master=self.Frame_Input)
        self.y_grid.insert(0, readfile_value(6))
        self.y_grid.grid(row=4, column=1, padx=20)
        self.y_grid.bind("<Return>", lambda event: text_config(6, self.y_grid))

    def close_save(self):
        text_config(3, self.x_seconden)
        text_config(4, self.x_grid)
        text_config(5, self.y_seconden)
        text_config(6, self.y_grid)

    def destroy(self):
        self.close_save()
        super().destroy()


def terminate_existing_main_processes():
    # Function to terminate existing instances of the main application process
    current_pid = os.getpid()  # Get the PID of the current process
    for proc in psutil.process_iter(['pid', 'name']):
        # Iterate through all running processes
        if proc.info['pid'] == current_pid:
            # Check if the process PID matches the PID of the current process
            proc.terminate()  # Terminate the process


if __name__ == "__main__":
    # Initialize and run the application
    set_priority()
    threading.Thread(target=log, daemon=True).start()
    update()
