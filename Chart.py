import customtkinter as ctk
import tkinter as tk
import threading
import serial
import tkchart
import ctypes
import time
import xlsxwriter
import datetime
import script
import propar
import warnings
import os
import sys
from Config import text_config, readfile_value

execute_path = os.path.abspath(sys.argv[0])
icon = os.path.dirname(execute_path) + r"\skalar_analytical_bv_logo_Zoy_icon.ico"


class App(ctk.CTk):
    def __init__(self,con,sensor_type, sn):
        super().__init__()
        self.c = con
        self.sensor_type = sensor_type
        self.serienummer = sn
        try:
            self.flow1 = self.c.flow_1
            self.flow2 = self.c.flow_2
            self.flow3 = self.c.flow_3
        except Exception as e:
            print(e)
            pass
        self.status_flow = self.c.status_flow
        self.iconbitmap(icon)
        self.resizable(width=False, height=False)
        self.title("Chart System")
        self.command = b"\x02\x30\x35\x30\x30\x30\x37\x03"
        self.Menu_One_Show = True
        self.Menu_Two_Show = True
        self.Menu_Three_Show = True
        self.Menu_Two_Slider = True
        self.Place_Button_Y = 28
        self.Place_Button_X = 140
        self.sliders_var = []
        self.input_var = []
        self.arrow_up = []
        self.arrow_down = []
        self.channel_option = None
        self.time = 0
        self.start_time = 0
        self.duration = 0
        self.total_delay = 0
        self.sensor_data = 0
        self.Seconds_Between = 10
        self.bronkhorst_data_1 = 0
        self.bronkhorst_data_2 = 0
        self.bronkhorst_data_3 = 0
        self.format_value_one = 0
        self.format_value_two = 0
        self.format_value_three = 0
        self.status_start_stop = False
        self.status_excel = True
        self.loading()

    def loading(self):
        self.Topbar()
        ctk.CTkLabel(text="Test system", master=self, font=ctk.CTkFont(size=15, weight="bold"), justify="center").grid(row=1, column=0, padx=50, pady=10)
        self.Masterchart()
        self.Chartline()
        self.Togglebutton()
        self.excel_button()
        self.loop_per_sec()
        self.loop()


    def Topbar(self):
        self.Menu_One = ctk.CTkButton(self, text="Exit",command=self.Topbar_Menu_One)
        self.Menu_One.grid(row=0,column=0, sticky="W")
        self.Menu_Two = ctk.CTkButton(self, text="Settings", command=self.Topbar_Menu_Two)
        self.Menu_Two.grid(row=0,column=0, padx=140, sticky="W")

    def Topbar_Menu_One(self):
        if self.Menu_One_Show:
            self.Menu_Restart = ctk.CTkButton(self, text="Restart", command=self.restart)
            self.Menu_Restart.place(x=0, y=self.Place_Button_Y)
            self.Menu_Quit = ctk.CTkButton(self, text="Quit", command=self.destroy)
            self.Menu_Quit.place(x=0, y=self.Place_Button_Y*2)
        else:
            self.Menu_Restart.place_forget()
            self.Menu_Quit.place_forget()
        self.Menu_One_Show = not self.Menu_One_Show

    def Topbar_Menu_Two(self):
        if self.Menu_Two_Show:
            self.Menu_Sliders = ctk.CTkButton(self, text="Sliders", command=self.Topbar_Sliders)
            self.Menu_Sliders.place(x=self.Place_Button_X, y=self.Place_Button_Y)
            self.Menu_Config = ctk.CTkButton(self, text="Config", command=self.open_settings)
            self.Menu_Config.place(x=self.Place_Button_X, y=self.Place_Button_Y * 2)
            self.Menu_Script = ctk.CTkButton(self, text="Script", command=self.set_sensor)
            self.Menu_Script.place(x=self.Place_Button_X, y=self.Place_Button_Y*3)
        else:
            self.Menu_Config.place_forget()
            self.Menu_Sliders.place_forget()
            self.Menu_Script.place_forget()
            try:
                self.select_type_sensor.destroy()
                self.channel_option.destroy()
            except Exception:
                pass
        self.Menu_Two_Show = not self.Menu_Two_Show

    def set_sensor(self):
        self.Menu_Script.place_forget()
        self.select_type_sensor = ctk.CTkComboBox(self, values=["V153", "V176", "V200"], justify="center", command=self.set_channel)
        self.select_type_sensor.place(x=self.Place_Button_X, y=self.Place_Button_Y * 3)
        self.select_type_sensor.set("Select a sensor")

    def set_channel(self, event=None):
        self.type_sensor = self.select_type_sensor.get()
        self.select_type_sensor.destroy()
        if self.type_sensor == "V153":
            channel_values = ["Ch1", "Ch2", "Ch3"]
            self.sensor_artikel = "2SN100224"

        elif self.type_sensor == "V176":
            channel_values = ["Ch1", "Ch2", "Ch3", "Ch4", "Ch6"]
            self.sensor_artikel = "2SN1001073"

        else:
            channel_values = ["Ch1", "Ch2", "Ch3", "Ch4", "Ch6"]
            self.sensor_artikel = "2SN1001098"


        if self.channel_option is not None:
            self.channel_option.destroy()

        self.channel_option = ctk.CTkComboBox(self, values=channel_values, justify="center", command= lambda event: self.open_script())
        self.channel_option.place(x=self.Place_Button_X, y=self.Place_Button_Y * 3)
        self.channel_option.set("Select a channel")

    def Masterchart(self):
        self.value_x_axis = int(readfile_value(1))
        self.value_steps = int(readfile_value(2))
        if self.value_x_axis < self.value_steps:
            text_config(1, 100)
            text_config(2, 10)
            self.value_x_axis = 100
            self.value_steps = 10

        x_axis_values = tuple(num for num in range(self.value_steps, self.value_x_axis+1, self.value_steps))
        self.Masterchart = tkchart.LineChart(
            master=self,
            width=800,
            height=400,
            axis_size=5,
            y_axis_section_count=10,
            x_axis_section_count=int(self.value_x_axis/self.value_steps),
            y_axis_label_count=10,
            x_axis_label_count=10,
            y_axis_data="Sensor meting",
            x_axis_data="Seconden",
            x_axis_values=x_axis_values,
            y_axis_values=(0, 100),
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

    def Chartline(self):
        self.Drawline_1 = tkchart.Line(master=self.Masterchart,
                                       color="lightblue",
                                       size=2,
                                       fill="enabled")
        self.Drawline_2 = tkchart.Line(master=self.Masterchart,
                                       color="red",
                                       size=2)
        self.Drawline_3 = tkchart.Line(master=self.Masterchart,
                                       color="green",
                                       size=2)

    def excel_button(self):
        self.Stop_Excel_button = ctk.CTkButton(self, text="Stop saving data", command=self.toggle_excel)
        self.Start_Excel_button = ctk.CTkButton(self, text="Start saving data", command=self.toggle_excel)
        self.Start_Excel_button.grid(row=5, column=0, padx=30, pady=(10, 0))

    def toggle_excel(self):
        self.status_excel = not self.status_excel
        if self.status_excel:
            self.Start_Excel_button.grid(row=5, column=0, padx=30, pady=(10,0))
            self.Stop_Excel_button.grid_forget()
            self.workbook.close()
        else:
            self.Stop_Excel_button.grid(row=5, column=0, padx=30, pady=(10,0))
            self.Start_Excel_button.grid_forget()
            self.start_excel()
        self.update()

    def start_excel(self):
        if self.sensor_type == "153":
            sensor = "2SN100224"
        elif self.sensor_type == "176":
            sensor = "2SN1001073"
        else:
            sensor = "2SN1001098"
        self.get_time()
        Location = readfile_value(8)
        if not os.path.exists(f"{Location}/{sensor}/{self.serienummer}"):
            os.mkdir(f"{Location}/{sensor}/{self.serienummer}")
        self.workbook = xlsxwriter.Workbook(fr"{Location}/{sensor}/{self.serienummer}/CustomReading-{self.current_time}.xlsx")
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.write(0, 0, 'Time')
        self.worksheet.set_column('A:A', 20)
        self.worksheet.write(0, 1, 'Value 1')
        self.worksheet.write(0, 2, 'Value 2')
        self.worksheet.write(0, 3, 'Value 3')
        self.worksheet.write(0, 4, 'PPM Value')
        threading.Thread(target=self.write_excel, daemon=True).start()

    def write_excel(self):
        while self.status_excel == True:
            self.time += 1
            data_1 = float(self.bronkhorst_data_1)
            data_2 = float(self.bronkhorst_data_2)
            data_3 = float(self.bronkhorst_data_3)
            try:
                sensor = float(open("transfer.txt", "r").read())
            except Exception:
                sensor = "Error"
            self.get_time()
            self.worksheet.write(self.time, 0, self.current_time)
            self.worksheet.write(self.time, 1, data_1)
            self.worksheet.write(self.time, 2, data_2)
            self.worksheet.write(self.time, 3, data_3)
            self.worksheet.write(self.time, 4, sensor)
            time.sleep(1)


    def Togglebutton(self):
        self.pausebutton = ctk.CTkButton(self, text="Stop", command=self.toggle_status)
        self.startbutton = ctk.CTkButton(self, text="Start", command=self.toggle_status)
        self.pausebutton.grid(row=6, column=0, padx=30)
        self.startbutton.grid(row=6, column=0, padx=30)

    def toggle_status(self):
        self.status_start_stop = not self.status_start_stop
        if self.status_start_stop:
            self.startbutton.grid_forget()
            self.pausebutton.grid(row=6, column=0, padx=30)
        else:
            self.pausebutton.grid_forget()
            self.startbutton.grid(row=6, column=0, padx=30)
        self.update()

    def Topbar_Sliders(self):
        if self.Menu_Two_Slider:
            self.Slider()
        else:
            for i in range(3):
                try:
                    self.sliders_var[i].grid_forget()
                    self.input_var[i].grid_forget()
                    self.arrow_up[i].grid_forget()
                    self.arrow_down[i].grid_forget()
                except Exception:
                    pass
            self.sliders_var = []
            self.input_var = []
            self.arrow_up = []
            self.arrow_down = []
        self.Menu_Two_Slider = not self.Menu_Two_Slider

    def Slider(self):
        if self.status_flow:
            for i in range(3):
                flow = [self.flow1.readParameter(8) / 320, self.flow2.readParameter(8) / 320, self.flow3.readParameter(8) / 320]
                arrow_up = ctk.CTkButton(self, text="🔼", width=20, height=20, font=ctk.CTkFont(size=20),command=lambda index=i: self.arrow_up_command(index))
                arrow_up.grid(row=3, column=4 + i)
                slider = ctk.CTkSlider(self, width=10, height=200, from_=0, to=100, progress_color="blue",orientation="vertical",command=lambda event, index=i: self.set_entry_value(index))
                slider.grid(row=4, column=4 + i)
                slider.set(flow[i])
                arrow_down = ctk.CTkButton(self, text="🔽", width=20, height=20, font=ctk.CTkFont(size=20),command=lambda index=i: self.arrow_down_command(index))
                arrow_down.grid(row=5, column=4 + i)
                entry = ctk.CTkEntry(self)
                entry.grid(row=2, column=4 + i, sticky="s")
                entry.insert(0, flow[i])
                entry.bind("<KeyRelease>", lambda event, index=i: self.set_slider_value(index))

                self.sliders_var.append(slider)
                self.input_var.append(entry)
                self.arrow_up.append(arrow_up)
                self.arrow_down.append(arrow_down)

    def set_slider_value(self, i, value = None, event=None):
        if value == None:
            try:
                value = float(self.input_var[i].get())
            except Exception:
                value = 0
        self.sliders_var[i].set(value)
        if i == 0:
            self.flow1.writeParameter(9, "{:.0f}".format(value / 100 * 32000))
        elif i == 1:
            self.flow2.writeParameter(9, "{:.0f}".format(value / 100 * 32000))
        else:
            self.flow3.writeParameter(9, "{:.0f}".format(value / 100 * 32000))

    def set_entry_value(self, i, value=None, event=None):
        self.input_var[i].delete(0, ctk.END)
        if value == None:
            value = self.sliders_var[i].get()
        self.input_var[i].insert(0, "{:.1f}".format(value))
        if i == 0:
            self.flow1.writeParameter(9, "{:.0f}".format(value / 100 * 32000))
        elif i == 1:
            self.flow2.writeParameter(9, "{:.0f}".format(value / 100 * 32000))
        else:
            self.flow3.writeParameter(9, "{:.0f}".format(value / 100 * 32000))

    def arrow_up_command(self, i):
        try:
            value = float(self.input_var[i].get())
        except Exception:
            value = 0
        value += 0.1
        self.set_entry_value(i, value)
        self.set_slider_value(i, value)


    def arrow_down_command(self, i):
        try:
            value = float(self.input_var[i].get())
        except Exception:
            value = 0
        value -= 0.1
        self.set_entry_value(i, value)
        self.set_slider_value(i, value)


    def get_time(self):
        self.current_time = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        self.day_time = datetime.datetime.now().strftime("%Y-%m-%d")

    def loop(self):
        if self.status_start_stop:
            if self.status_flow:
                self.Masterchart.show_data(data=[float(self.bronkhorst_data_1)], line=self.Drawline_1)
                self.Masterchart.show_data(data=[float(self.bronkhorst_data_2)], line=self.Drawline_2)
                self.Masterchart.show_data(data=[float(self.bronkhorst_data_3)], line=self.Drawline_3)
        self.after(self.value_steps*1000, self.loop)

    def loop_per_sec(self):
        if self.status_start_stop:
            if self.status_flow:
                parameter_value_1 = self.flow1.readParameter(8)
                if parameter_value_1 is not None:
                    self.bronkhorst_data_1 = parameter_value_1 / 32000 * 100
                else:
                    self.bronkhorst_data_1 = 0

                parameter_value_2 = self.flow2.readParameter(8)
                if parameter_value_2 is not None:
                    self.bronkhorst_data_2 = parameter_value_2 / 32000 * 100
                else:
                    self.bronkhorst_data_2 = 0

                parameter_value_3 = self.flow3.readParameter(8)
                if parameter_value_3 is not None:
                    self.bronkhorst_data_3 = parameter_value_3 / 32000 * 100
                else:
                    self.bronkhorst_data_3 = 0
        self.after(1000, self.loop_per_sec)

    def open_settings(self):
        app = ConfigurationApp()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def open_script(self):
        app = script.app(self.sensor_artikel, self.channel_option.get())
        self.select_type_sensor.destroy()
        self.channel_option.destroy()
        self.Menu_Script.place(x=self.Place_Button_X, y=self.Place_Button_Y * 3)
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def restart(self):
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=UserWarning)
                self.workbook.close()
        except Exception:
            pass
        super().destroy()
        App(self.c,self.sensor_type,self.serienummer).mainloop()

    def destroy(self):
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=UserWarning)
                self.workbook.close()
        except Exception:
            pass
        super().destroy()



class ConfigurationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.iconbitmap(icon)
        self.config_file = "config.txt"
        self.X_as_middle = 800/2
        self.resizable(width=False, height=False)
        self.title("Configuration")
        self.loading()

    def loading(self):
        self.Frame_Input = ctk.CTkFrame(self)
        self.Frame_Input.grid(row=1, column=0,padx=(self.X_as_middle-self.Frame_Input.winfo_reqwidth())/2,pady=40, stick="nsew")
        self.Close_Save_Button = ctk.CTkButton(self, text="Save & Close", command=self.destroy)
        self.Close_Save_Button.grid(row=2, column=0, pady=(0,20))
        ctk.CTkLabel(master=self, text="Configuration", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="N")
        self.Section_Input()

    def Section_Input(self):
        self.text_option_seconden = ctk.CTkLabel(master=self.Frame_Input, text="Seconden")
        self.text_option_seconden.grid(row=1,column=0,padx=(20,0), pady=20,sticky="w")
        self.text_option_seconden_step = ctk.CTkLabel(master=self.Frame_Input, text="Stappen tussenin")
        self.text_option_seconden_step.grid(row=2,column=0,sticky="w")
        self.config_option_seconden = ctk.CTkEntry(master=self.Frame_Input)
        self.config_option_seconden.insert(0,readfile_value(1))
        self.config_option_seconden.grid(row=1,column=1, padx=(10,20))
        self.config_option_seconden.bind("<Return>", lambda event: text_config(1, self.config_option_seconden))
        self.config_option_seconden_step = ctk.CTkEntry(master=self.Frame_Input)
        self.config_option_seconden_step.insert(0,readfile_value(2))
        self.config_option_seconden_step.grid(row=2,column=1)
        self.config_option_seconden_step.bind("<Return>", lambda event: text_config(2, self.config_option_seconden_step))


    def Close_Save(self):
        text_config(2,self.config_option_seconden_step)
        text_config(1, self.config_option_seconden)

    def destroy(self):
        self.Close_Save()
        super().destroy()