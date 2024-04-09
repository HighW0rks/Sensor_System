# Standard library
import datetime
import os
import shutil
import sys
import threading
import time
# Third-Party library
import openpyxl
import psutil
import serial
import serial.tools.list_ports
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk
# Uncommon library
import main
from connection import Connection
import Chart
import propar

def terminate_existing_main_processes():
    # Function to terminate existing instances of the main application process
    for proc in psutil.process_iter(['pid', 'name']):
        # Iterate through all running processes
        if proc.info['name'] == 'Skalar Saxon Tester.exe':
            # Check if the process name matches the main application
            proc.terminate()  # Terminate the process

class File_check(ctk.CTk):
    def __init__(self, con):
        super().__init__()
        self.geometry("250x100")  # Set window geometry
        self.resizable(width=False, height=False)  # Disable window resizing
        self.title("Error")  # Set window title
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")  # Set window icon
        self.con = con  # Connection object
        self.Location_check()  # Check folder location

    def Location_check(self):
        # Function to check if the required folder location exists
        self.Error_text = ctk.CTkLabel(self, text="")  # Initialize error label
        self.Location_button = ctk.CTkButton(self, text="Select a folder", command=Chart.ConfigurationApp().Folder_Location_Config)
        # Initialize button to select folder location

        Location = Chart.ConfigurationApp().readfile_value(4)  # Read folder location from configuration
        if not os.path.exists(f"{Location}/2SN100224"):
            # Check if the required folder does not exist
            self.Error_text.configure(text="Error!\nNo folder location found")  # Display error message
            self.Location_button.place(x=self.winfo_reqwidth() / 2 - self.Location_button.winfo_reqwidth() / 4, y=50)
            # Place button in the center horizontally
            self.update()  # Update the window
            self.after(200, self.Location_check)  # Repeat the folder location check after 200 milliseconds
        else:
            self.destroy()  # Destroy the current window
            app = MainApp(self.con)  # Create an instance of the main application
            app.mainloop()  # Start the main application loop


class MainApp(ctk.CTk):
    def __init__(self, con):
        super().__init__()
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")  # Set application icon
        self.geometry("500x550")  # Set window geometry
        self.resizable(width=False, height=False)  # Disable window resizing
        self.title("Test App")  # Set window title
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
        self.channel_option = None  # Variable to store selected channel option
        self.serienummer = "Unknown"  # Variable to store serial number
        self.channel_status = "Unknown"  # Variable to store channel status
        self.background_color()  # Set application background color
        self.loading()  # Load application components
        self.protocol("WM_DELETE_WINDOW", self.close_app)  # Define action on window close

    def background_color(self):
        # Set application appearance mode based on configuration
        with open("config.txt") as read:
            for line in read:
                if "Modes" in line:
                    value = line.split(": ")[1].strip()
        if value == "True":
            ctk.set_appearance_mode("Light")  # Set light mode
        else:
            ctk.set_appearance_mode("Dark")  # Set dark mode

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
        self.information_channel_status = ctk.CTkFrame(self)
        self.information_channel_status.grid(row=3, column=1, padx=(20, 0), sticky="n")

    def thread(self):
        # Start threads for sensor and flow
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
        self.start_button.grid(row=6, column=0, columnspan=2, padx=(20,0),pady=(10,20), sticky="nsew")
        self.stop_button = ctk.CTkButton(self, text="Stop", width=300, height=100, command= lambda: self.start_stop(1))

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
                self.serienummer = str(read.decode()[5:13])
                if len(read) == 0 or len(read) > 24 :
                    self.serienummer = "Unknown"
                self.connection.sensor.write(b"\x02\x36\x38\x30\x30\x30\x63\x03")
                read = self.connection.sensor.read(1024)
                self.sensor_version = str(read.decode()[61:81])
            except Exception as e:
                print(e)
                pass


    def status_channel(self):
        try:
            if len(self.ppm_value) > 0:
                if self.channel_1 == self.ppm_value or self.channel_2 == self.ppm_value or self.channel_3 == self.ppm_value:
                    self.status_channel_label.configure(text="K1/K2/K3")
                else:
                    self.status_channel_label.configure(text="K4/K6")
        except Exception:
            pass


    def sensor_info(self):
        title = ctk.CTkLabel(master=self.information_title, text="Sensor Information", font=ctk.CTkFont(size=15,weight="bold"))
        title.grid(row=0, column=0, padx=80, pady=5)
        self.serienummer_label = ctk.CTkLabel(master=self.information_serienummer, text=f"Serie Nummer\n{self.serienummer}")
        self.serienummer_label.grid(row=0, column=0,padx=20,pady=10)
        self.status_channel_label = ctk.CTkLabel(master=self.information_channel_status, text="Unknown")
        self.status_channel_label.grid(row=0, column=0, padx=20, pady=10)


    def set_sensor(self):
        self.select_type_sensor = ctk.CTkComboBox(self, values=["V153", "V176", "V200"], justify="center", command=self.set_channel)
        self.select_type_sensor.grid(row=4, column=0, columnspan=2, padx=(20,0), pady=(20,0))
        self.select_type_sensor.set("Select a sensor")


    def set_channel(self, event=None):
        self.type_sensor = self.select_type_sensor.get()
        if self.type_sensor == "V153":
            channel_values = ["Ch1", "Ch2", "Ch3"]
            self.folder_location = "2SN100224"
            self.excel_file = "425 --- test Data QC 2SN100224.xlsx"

        elif self.type_sensor == "V176":
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
        # Initial sensor information retrieval
        self.sensor_info_con()

        # Continuous loop for sensor communication
        while self.main_run:
            # Check if channel label status is unknown and update if necessary
            if self.status_channel_label.cget('text') == "Unknown":
                self.status_channel()

            # Check if serial number is unknown and update if necessary
            if self.serienummer == "Unknown":
                self.sensor_info_con()

            try:
                # Send command to sensor and read response
                self.connection.sensor.write(b"\x02\x30\x35\x30\x30\x30\x37\x03")
                response = self.connection.sensor.read(1024).decode()
                print("Response: ", response)
                configure = False

                if response:
                    # Parse response and update values
                    self.sensor_status = 1
                    self.channel_1, self.channel_2, self.channel_3, self.channel_4, self.ppm_value, self.channel_6, self.temperature, self.air_pressure = response.split()[1:9]

                    # Write PPM value to a file
                    with open("transfer.txt", "w") as write:
                        write.write(self.ppm_value)

                    # Update labels with new values
                    self.serienummer_label.configure(text=f"Serie Nummer\n{self.serienummer}")
                    self.ppm_meter_label.configure(text=f"PPM value\n\n{self.ppm_value}")
                    configure = False
                else:
                    # Set sensor status to unknown if no response
                    self.serienummer = "Unknown"
                    self.sensor_status = 0
                    self.serienummer_label.configure(text="Serie Nummer\nUnknown")
                    self.ppm_meter_label.configure(text="Sensor not found")
                    self.status_channel_label.configure(text="Unknown")
                    configure = True
            except Exception as e:
                # Handle errors in sensor communication
                self.serienummer = "Unknown"
                self.sensor_status = 0
                print(f"Error in sensor communication: {e}")
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
        self.status_channel_label.configure(text="Unknown")

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
            except Exception as e:
                # Handle errors in flow communication
                print(f"Error in flow communication: {e}")
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
        self.empty_ppm_frame = ctk.CTkLabel(master=self.frame_ppm, text="")
        self.empty_ppm_frame.grid(row=0, column=0, padx=100)

        # Initialize PPM meter label
        self.ppm_meter_label = ctk.CTkLabel(master=self.frame_ppm, text="Connecting...")
        self.ppm_meter_label.grid(row=0, column=0, pady=17)

    def zero_point(self):
        # Create empty label frame for zero point
        self.empty_zero_point_frame = ctk.CTkLabel(master=self.frame_zero_point, text="", height=0)
        self.empty_zero_point_frame.grid(row=0, column=0, padx=100)

        # Initialize zero point button
        self.zero_point_button = ctk.CTkButton(master=self.frame_zero_point, text="Zero Point",
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
        self.folder = Chart.ConfigurationApp().readfile_value(4)

        # Get selected channel
        self.channel = self.channel_option.get()

        # Start Excel
        self.start_excel()

        # Start reading script thread
        self.read_script_thread()

    def start_excel(self):
        # Initialize Excel settings
        self.excel_row = 2  # Start from row 2
        self.workbook = openpyxl.load_workbook(
            fr"{self.folder}/{self.folder_location}/{self.excel_file}")  # Load workbook
        self.worksheet = self.workbook["Sheet1"]  # Select Sheet1
        self.worksheet['E2'] = self.serienummer  # Set serial number
        self.worksheet['E3'] = self.folder_location # Set model
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

        # Write Excel thread
        self.write_excel_thread()

    def write_excel_thread(self):
        # Start thread for writing Excel data
        thread = threading.Thread(target=self.write_excel)
        thread.daemon = True
        thread.start()

    def read_script_thread(self):
        # Start thread for reading script
        thread = threading.Thread(target=self.read_script)
        thread.daemon = True
        thread.start()

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
        if self.running == True:
            # Open script file
            with open(fr"{self.folder}/{self.folder_location}/Scripts/Script {self.channel}.txt", "r") as read:
                read_value = read.readlines()
                # Check if end of file is reached
                if not self.row_value < len(read_value):
                    print("End of file reached")
                    self.running = False
                    # Save Excel file with timestamp
                    self.get_time()
                    if not os.path.exists(f"{self.folder}/{self.folder_location}/{self.serienummer}"):
                        os.mkdir(f"{self.folder}/{self.folder_location}/{self.serienummer}")
                    self.workbook.save(
                        f"{self.folder}/{self.folder_location}/{self.serienummer}/{self.channel}_{self.current_time}.xlsx")
                    # Stop program
                    self.start_stop(1)
                    return
                else:
                    # Parse script values
                    self.seconde_value = read_value[self.row_value].split(". ")[1].split("; ")[0].strip()
                    self.percentage_value = read_value[self.row_value].split("; ")[1].strip()
                    self.channel_value = read_value[self.row_value].split("; ")[2].strip()
                    # Print values for debugging
                    print(f"Sec: {self.seconde_value}  | Flow: ",
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
                    self.connection.flow_1.writeParameter(9,"{:.0f}".format(float(self.percentage_value) * 32000 / 100))
                    self.row_value += 1
                    time.sleep(float(self.seconde_value))
                    self.read_script()
                elif self.channel_value == "2":
                    self.connection.flow_2.writeParameter(9,"{:.0f}".format(float(self.percentage_value) * 32000 / 100))
                    self.row_value += 1
                    time.sleep(float(self.seconde_value))
                    self.read_script()
                else:
                    self.connection.flow_3.writeParameter(9,"{:.0f}".format(float(self.percentage_value) * 32000 / 100))
                    self.row_value += 1
                    time.sleep(float(self.seconde_value))
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
        app = Chart.App(self.connection, self.serienummer)
        app.protocol("WM_DELETE_WINDOW", app.destroy)  # Handle window close event
        app.mainloop()

    def open_config(self):
        # Open the configuration application
        self.main_run = False  # Set main run flag to False
        app = Configuration(self.connection, self.connection.sensor)
        # Handle window close event, switch main run flag, and start thread on exit
        app.protocol("WM_DELETE_WINDOW", lambda: (app.destroy(), self.switch_main_run(), self.thread(), self.reset_channel_info()))
        app.mainloop()

    def reset_channel_info(self):
        # Resets the channel info for the K1/K2/K3 or K4/K6 gets displayed correctly
        self.status_channel_label.configure(text="Unknown")
        self.ppm_value = None

    def switch_main_run(self):
        # Switch the value of the main run flag
        self.main_run = not self.main_run

    def close_app(self):
        # Close the application
        terminate_existing_main_processes()  # Terminate existing main processes


class Configuration(ctk.CTk):
    def __init__(self, con, sensor):
        super().__init__()
        self.con = con
        self.sensor = sensor
        self.title("Configuration")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.option_add("*TCombobox*Listbox*Font", value=0)
        self.comport_sensor = None
        self.comport_flow = None
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
        self.frame_switch.grid(row=1, column=0, pady=(10,0))
        self.frame_folder = ctk.CTkFrame(self)
        self.frame_folder.grid(row=1, column=1)

    def listen_comport(self):
        # Get available serial ports
        return serial.tools.list_ports.comports()

    def text(self):
        # Add labels and buttons
        label_sensor = ctk.CTkLabel(master=self.frame_comport, text="Sensor").grid(row=0, column=1)
        label_flow = ctk.CTkLabel(master=self.frame_comport, text="Flow Controller").grid(row=0,column=0)
        button_exit = ctk.CTkButton(self, text="Save & Restart App", command=self.restart_app).grid(row=4, column=0, columnspan=2, pady=(30,0))

    def combobox(self):
        # Create comboboxes for selecting serial ports
        self.flow_com = ttk.Combobox(master=self.frame_comport, width=25, values=self.listen_comport(),font=(15))
        self.flow_com.grid(row=1, column=0)
        self.flow_com.set("Select a comport")
        self.flow_com.configure(font=(20))
        self.sensor_com = ttk.Combobox(master=self.frame_comport, width=25, values=self.listen_comport(), font=(15))
        self.flow_com.bind('<<ComboboxSelected>>', self.get_flow)
        self.sensor_com.grid(row=1,column=1)
        self.sensor_com.set("Select a comport")
        self.sensor_com.bind('<<ComboboxSelected>>', self.get_sensor)

    def check(self):
        # Check if both comports are selected
        if self.comport_flow != None and self.comport_sensor != None:
            print("All good")

    def get_flow(self, event):
        # Get selected flow controller comport
        self.comport_flow = self.flow_com.get()
        value = self.comport_flow.split(" ")[0]
        Chart.ConfigurationApp().text_config(6, value)
        self.check()

    def get_sensor(self, event):
        # Get selected sensor comport
        self.comport_sensor = self.sensor_com.get()
        value = self.comport_sensor.split(" ")[0]
        Chart.ConfigurationApp().text_config(5, value)
        self.check()

    def button_switch(self):
        # Reads the config and makes a appearance switch
        with open("config.txt", "r") as r:
            self.read = r.readlines()
            self.modes_bool = str(self.read[2].split(": ")[1].split("\n")[0])
            dark_mode_state = ctk.BooleanVar(value=(self.modes_bool == "True"))
        self.dark_white_mode = ctk.CTkSwitch(master=self.frame_switch, text="Dark/Light mode", variable=dark_mode_state, command= self.set_dark_light_mode)
        self.dark_white_mode.grid(row=2, column=0, columnspan=2, padx=20, pady=20)

        # Reads the config and makes a channel switch
        with open("config.txt", "r") as r:
            self.read = r.readlines()
            self.modes_bool = str(self.read[6].split(": ")[1].split("\n")[0])
            channel_mode_state = ctk.BooleanVar(value=(self.modes_bool == "True"))
        self.channel_switch = ctk.CTkSwitch(master=self.frame_switch, text="K1/K2/K3 | K4/K6", variable=channel_mode_state, command=self.channel)
        self.channel_switch.grid(row=3, column=0, columnspan=2, pady=(0,20))

    def set_dark_light_mode(self):
        # Reads the config file and sets the value to set the background color
        with open("config.txt","r") as r:
            value = r.readlines()[2].split(": ")[1].split("\n")[0]
        if value == "True":
            value = "False"
        else:
            value = "True"
        with open("config.txt","w") as w:
            self.read[2] = f"Modes: {value}\n"
            w.writelines(self.read)
        self.set_background_color()

    def channel(self):
        # Reads the config file and writes to the sensor to switch
        with open("config.txt", "r") as file:
            value = file.readlines()[6].split(": ")[1].split("\n")[0]
        if value == "True":
            print("Switched from K4/K6 to K1/K2/K3")
            self.sensor.write(b"\x02\x38\x31\x30\x30\x30\x33\x62\x03")
            value = "False"
        else:
            print("Switched from K1/K2/K3 to K4/K6")
            self.sensor.write(b"\x02\x38\x31\x30\x30\x31\x33\x61\x03")
            value = "True"
        with open("config.txt", "w") as w:
            self.read[6] = f"Channel: {value}\n"
            w.writelines(self.read)

    def directory(self):
        # Add button for selecting main folder
        ctk.CTkButton(master=self.frame_folder, text="Select the main folder", command=self.folder_location).grid(row=0, column=0)

    def folder_location(self):
        # Get folder location
        self.Folder_Location = tk.filedialog.askdirectory()
        # Sets folder location
        Chart.ConfigurationApp().text_config(4,self.Folder_Location)

    def set_background_color(self):
        # Sets the background color
        MainApp.background_color(ctk.CTk)

    def restart_app(self):
        # Save configuration and restart application
        with open("config.txt", "w") as w:
            w.writelines(self.read)
        os.execv(sys.executable, [sys.executable] + sys.argv)

    def destroy(self):
        # Destroys the application
        super().destroy()

if __name__ == "__main__":
    # Initialize and run the application
    con = Connection()
    mainapp = File_check(con)
    mainapp.mainloop()