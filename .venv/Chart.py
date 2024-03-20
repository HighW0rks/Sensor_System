import customtkinter as ctk
import tkinter as tk
import threading
import serial
import tkchart
import ctypes
import time
import main
import xlsxwriter
import datetime
import script
import propar

class App(ctk.CTk):
    def __init__(self,con):
        super().__init__()
        self.c = con
        print(self.c)
        self.flow1 = self.c.flow_1
        self.flow2 = self.c.flow_2
        self.flow3 = self.c.flow_3
        self.status_flow1 = self.c.status_flow_1
        self.status_flow2 = self.c.status_flow_2
        self.status_flow3 = self.c.status_flow_3
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.resizable(width=False, height=False)
        self.title("Chart System")
        self.command = b"\x02\x30\x35\x30\x30\x30\x37\x03"
        self.Menu_One_Show = True
        self.Menu_Two_Show = True
        self.Menu_Three_Show = True
        self.Menu_Two_Slider = True
        self.Place_Button_Y = 28
        self.Place_Button_X = 140
        self.sliderchannel_One = None
        self.sliderchannel_Two = None
        self.sliderchannel_Three = None
        self.inputchannel_One = None
        self.inputchannel_Two = None
        self.inputchannel_Three = None
        self.error_message_slider_entry = None
        self.controllers = []
        self.time = 0
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
        self.threat()
        self.connection()
        self.Topbar()
        self.load_text()
        self.Masterchart()
        self.Chartline()
        self.Togglebutton()
        self.excel_button()
        self.loop_per_sec()
        self.loop()

    def threat(self):
        self.connect_thread = threading.Thread(target=self.connection_saxon)
        self.connect_thread.start()

    def connection_saxon(self):
        running = True
        try:
            self.serial = self.c.sensor
            while running == True:
                self.data_saxon()
        except serial.serialutil.SerialException as e:
            print("Error port not found", e)
            running = False
        except IndexError as e:
            print("Error sensor not found", e)
            running = False

    def data_saxon(self):
        try:
            with open("transfer.txt", "r") as read:
                self.sensor_data = read.readlines(1)
            self.update()
            time.sleep(0.5)
        except serial.serialutil.SerialException as e:
            print("Error port not found")
            running = False

    def connection(self):
        if self.status_flow1:
            self.controllers.append("1")
        else:
            print("Error flow 1 unable to connect")

        # if self.status_flow2:
        #     self.controllers.append("2")
        # else:
        #     print("Error flow 2 unable to connect")
        #
        # if self.status_flow3:
        #     self.controllers.append("3")
        # else:
        #     print("Error flow 3 unable to connect")

        print(self.controllers)

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
            self.Menu_Script = ctk.CTkButton(self, text="Script", command=self.open_script)
            self.Menu_Script.place(x=self.Place_Button_X, y=self.Place_Button_Y*3)
        else:
            self.Menu_Config.place_forget()
            self.Menu_Sliders.place_forget()
            self.Menu_Script.place_forget()
        self.Menu_Two_Show = not self.Menu_Two_Show

    def Topbar_Sliders(self):
        if self.Menu_Two_Slider:
            self.Sliders()
        else:
            self.sliderchannel_One.grid_forget()
            self.sliderchannel_Two.grid_forget()
            self.sliderchannel_Three.grid_forget()
            self.inputchannel_One.grid_forget()
            self.inputchannel_Two.grid_forget()
            self.inputchannel_Three.grid_forget()
            if self.error_message_slider_entry:
                self.error_message_slider_entry.grid_forget()
        self.Menu_Two_Slider = not self.Menu_Two_Slider

    def load_text(self):
        self.StartupLabel = ctk.CTkLabel(text="Test system", master=self, font=ctk.CTkFont(size=15, weight="bold"), justify="center")
        self.StartupLabel.grid(row=1, column=0, padx=50, pady=10)

    def Get_x_axis_values(self):
        self.value_x_axis = int(ConfigurationApp().readfile_value(1))
        self.value_steps = int(ConfigurationApp().readfile_value(2))

    def Masterchart(self):
        self.Get_x_axis_values()
        if self.value_x_axis < self.value_steps:
            with open("config.txt", "w") as config_file:
                config_file.write("Seconden: 100\n")
                config_file.write("Stappen tussenin: 10\n")
                config_file.write("Modes: False")
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
                                       size=2)
        self.Drawline_2 = tkchart.Line(master=self.Masterchart,
                                       color="red",
                                       size=2)
        self.Drawline_3 = tkchart.Line(master=self.Masterchart,
                                       color="green",
                                       size=2)

    def excel_button(self):
        self.Stop_Excel_button = ctk.CTkButton(self, text="Stop saving data", command=self.toggle_excel)
        self.Start_Excel_button = ctk.CTkButton(self, text="Start saving data", command=self.toggle_excel)
        self.Stop_Excel_button.grid(row=5, column=0, padx=30, pady=(10, 0))
        self.Start_Excel_button.grid(row=5, column=0, padx=30, pady=(10, 0))
        self.start_excel()
    def start_excel(self):
        current_time = self.current_time()
        Location = ConfigurationApp().readfile_value(4)
        self.workbook = xlsxwriter.Workbook(fr"{Location}\{current_time}.xlsx")
        self.worksheet = self.workbook.add_worksheet()

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

    def write_excel(self):
        if not self.status_excel:
            self.time += 1
            current_time = self.current_time()
            data_1 = float(self.bronkhorst_data_1)
            data_2 = float(self.bronkhorst_data_2)
            data_3 = float(self.bronkhorst_data_3)
            sensor = float(self.sensor_data)

            self.worksheet.write(0, 0, 'Time')
            self.worksheet.set_column('A:A', 20)
            self.worksheet.write(0, 1, 'Value 1')
            self.worksheet.write(0, 2, 'Value 2')
            self.worksheet.write(0, 3, 'Value 3')
            self.worksheet.write(0, 4, 'PPM Value')

            self.worksheet.write(self.time, 0, current_time)
            self.worksheet.write(self.time, 1, data_1)
            self.worksheet.write(self.time, 2, data_2)
            self.worksheet.write(self.time, 3, data_3)
            self.worksheet.write(self.time, 4, sensor)



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

    def Sliders(self):
        self.inputchannel_One = ctk.CTkEntry(self)
        self.sliderchannel_One = ctk.CTkSlider(master=self, width=10, height=200, from_=0, to=100,
                                               progress_color="blue", orientation="vertical")
        if "1" in self.controllers:
            self.inputchannel_One.grid(row=2, column=5, sticky="S")
            self.inputchannel_One.bind("<Return>", lambda event: self.set_slider_value(1))
            self.sliderchannel_One.grid(row=4, column=5)

        self.inputchannel_Two = ctk.CTkEntry(self)
        self.sliderchannel_Two = ctk.CTkSlider(master=self, width=10, height=200, from_=0, to=100,
                                               progress_color="blue", orientation="vertical")
        if "2" in self.controllers:
            self.inputchannel_Two.grid(row=2, column=6, sticky="S")
            self.inputchannel_Two.bind("<Return>", lambda event: self.set_slider_value(2))
            self.sliderchannel_Two.grid(row=4, column=6)

        self.inputchannel_Three = ctk.CTkEntry(self)
        self.sliderchannel_Three = ctk.CTkSlider(master=self, width=10, height=200, from_=0, to=100,
                                                 progress_color="blue", orientation="vertical")
        if "3" in self.controllers:
            self.inputchannel_Three.grid(row=2, column=7, sticky="S")
            self.inputchannel_Three.bind("<Return>", lambda event: self.set_slider_value(3))
            self.sliderchannel_Three.grid(row=4, column=7)

        self.update_entry_from_slider()
        self.bind("<Motion>", self.update_entry_from_slider)

    def set_slider_value(self, column):
        if column == 1:
            try:
                value_one = int(self.inputchannel_One.get())
                self.sliderchannel_One.set(value_one)
                self.update_entry_from_slider()
                self.error_message_slider_entry.configure(text="")
            except ValueError:
                self.error_slider_value(1)

        elif column == 2:
            try:
                value_two = int(self.inputchannel_Two.get())
                self.sliderchannel_Two.set(value_two)
                self.update_entry_from_slider()
                self.error_message_slider_entry.configure(text="")

            except ValueError:
                self.error_slider_value(2)

        else:
            try:
                value_three = int(self.inputchannel_Three.get())
                self.sliderchannel_Three.set(value_three)
                self.update_entry_from_slider()
                self.error_message_slider_entry.configure(text="")

            except ValueError:
                self.error_slider_value(3)

    def error_slider_value(self, x):
        self.error_message_slider_entry = ctk.CTkLabel(self, text="Please input a valid number")
        self.error_message_slider_entry.grid(row=3, column=x, sticky="n")

    def update_entry_from_slider(self, event=None):
        self.value_one = self.sliderchannel_One.get()
        self.value_two = self.sliderchannel_Two.get()
        self.value_three = self.sliderchannel_Three.get()
        self.format_value_one = "{:.1f}".format(self.value_one)
        self.format_value_two = "{:.1f}".format(self.value_two)
        self.format_value_three = "{:.1f}".format(self.value_three)
        self.inputchannel_One.delete(0, ctk.END, )
        self.inputchannel_One.insert(0, str(self.format_value_one) + " %")
        self.inputchannel_Two.delete(0, ctk.END, )
        self.inputchannel_Two.insert(0, str(self.format_value_two) + " %")
        self.inputchannel_Three.delete(0, ctk.END, )
        self.inputchannel_Three.insert(0, str(self.format_value_three) + " %")
        if "1" in self.controllers:
            data1 = "{:.0f}".format(self.value_one / 100 * 32000)
            self.flow1.writeParameter(9, data1)
        if "2" in self.controllers:
            self.flow2.writeParameter(9,self.value_two / 100 * 32000)
        if "3" in self.controllers:
            self.flow3.writeParameter(9,self.value_three / 100 * 32000)

    def current_time(self):
        current_time = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        return current_time

    def loop(self):
        if self.status_start_stop:
            if "1" in self.controllers:
                data_1 = float(self.bronkhorst_data_1)
                self.Masterchart.show_data(data=[data_1], line=self.Drawline_1)

            if "2" in self.controllers:
                data_2 = float(self.bronkhorst_data_2)
                self.Masterchart.show_data(data=[data_2], line=self.Drawline_2)

            if "3" in self.controllers:
                data_3 = float(self.bronkhorst_data_3)
                self.Masterchart.show_data(data=[data_3], line=self.Drawline_3)
        self.after(self.value_steps*1000, self.loop)

    def loop_per_sec(self):
        if self.status_start_stop:
            if "1" in self.controllers:
                parameter_value_1 = self.flow1.readParameter(8)
                self.test = self.c.sensor

                print(parameter_value_1, " | ",self.sensor_data)
                if parameter_value_1 is not None:
                    self.bronkhorst_data_1 = parameter_value_1 / 32000 * 100
                else:
                    self.bronkhorst_data_1 = 0
            else:
                self.bronkhorst_data_1 = 0

            if "2" in self.controllers:
                parameter_value_2 = self.flow2.readParameter(8)
                if parameter_value_2 is not None:
                    self.bronkhorst_data_2 = parameter_value_2 / 32000 * 100
                else:
                    self.bronkhorst_data_2 = 0
            else:
                self.bronkhorst_data_2 = 0

            if "3" in self.controllers:
                parameter_value_3 = self.flow3.readParameter(8)
                if parameter_value_3 is not None:
                    self.bronkhorst_data_3 = parameter_value_3 / 32000 * 100
                else:
                    self.bronkhorst_data_3 = 0
            else:
                self.bronkhorst_data_3 = 0
            self.write_excel()
        self.after(1000,self.loop_per_sec)

    def open_settings(self):
        app = ConfigurationApp()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def open_script(self):
        app = script.app()
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
        restart = App(self.c)
        restart.mainloop()

    def destroy(self):
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=UserWarning)
                self.workbook.close()
        except Exception:
            pass
        print("Destroy")
        super().destroy()



class ConfigurationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
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
        self.Close_Save_Button = ctk.CTkButton(self, text="Save & Close", command=self.destroy)
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
        self.Input_Folder = ctk.CTkButton(master=self.Frame_Input, text="Select a folder",
                                          command=self.Folder_Location_Config)
        self.Input_Folder.grid(row=3, column=1, padx=20, pady=20)

    def Close_Save(self):
        self.text_config(2,self.config_option_seconden_step)
        self.text_config(1, self.config_option_seconden)

    def Section_Bool(self):
        dark_mode_state = self.read_dark_mode_state_from_file()
        self.dark_white_mode_var = ctk.BooleanVar(value=dark_mode_state)
        self.dark_white_mode = ctk.CTkSwitch(master=self.Frame_Bool, text="Dark/Light mode", variable=self.dark_white_mode_var, command=lambda: self.bool_config("Modes"))
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
        try:
            self.parameter_text_configvalue = str(self.config_text_value[row - 1]).split(": ")[1].strip()
        except IndexError:
            self.parameter_text_configvalue = ""
        return self.parameter_text_configvalue

    def text_config(self, row, name):
        with open(self.config_file, "r") as read:
            self.config_text_value = read.readlines()
        self.parameter_text_configname = str(self.config_text_value[row - 1]).split(": ")[0].strip()
        try:
            self.config_new_value = name.get()
        except Exception:
            self.config_new_value = name
        self.updated_text_row = f"{self.parameter_text_configname}: {self.config_new_value}\n"
        self.config_text_value[row - 1] = self.updated_text_row
        with open(self.config_file, "w") as write:
            write.writelines(self.config_text_value)

    def bool_config(self, name):
        with open(self.config_file, "r") as file:
            lines = file.readlines()

        for index, line in enumerate(lines):
            if name in line:
                parameter_name, parameter_value = map(str.strip, line.split(":"))
                if parameter_value == "True":
                    new_value = "False"
                else:
                    new_value = "True"
                lines[index] = f"{parameter_name}: {new_value}\n"
                break

        with open(self.config_file, "w") as file:
            file.writelines(lines)

        self.restart()

    def Folder_Location_Config(self):
        self.Folder_Location = tk.filedialog.askdirectory()
        self.text_config(4,self.Folder_Location)

    def restart(self):
        main.MainApp.background_color(ctk.CTk)

    def destroy(self):
        self.Close_Save()
        super().destroy()
        app = App()
        app.mainloop()


if __name__ == '__main__':
    App().mainloop()