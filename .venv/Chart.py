import Configuration
import customtkinter as ctk
import tkchart
import ctypes
import main
import xlsxwriter
import datetime
import script

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.main_app = main.MainApp()
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.resizable(width=False, height=False)
        self.title("Chart System")
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
        self.time = 0
        self.Seconds_Between = 10
        self.format_value_one = 50
        self.format_value_two = 50
        self.format_value_three = 50
        self.status_start_stop = False
        self.status_excel = True
        self.loading()

    def loading(self):
        self.Topbar()
        self.load_text()
        self.Masterchart()
        self.Chartline()
        self.Togglebutton()
        self.excel_button()
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
            self.Menu_Config = ctk.CTkButton(self, text="Config", command=self.open_settings)
            self.Menu_Config.place(x=self.Place_Button_X, y=self.Place_Button_Y)
            self.Menu_Sliders = ctk.CTkButton(self, text="Sliders", command=self.Topbar_Sliders)
            self.Menu_Sliders.place(x=self.Place_Button_X, y=self.Place_Button_Y*2)
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
        self.value_x_axis = int(Configuration.ConfigurationApp().readfile_value(1))
        self.value_steps = int(Configuration.ConfigurationApp().readfile_value(2))

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
        Location = Configuration.ConfigurationApp().readfile_value(4)
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
            data_1 = float(self.format_value_one)
            data_2 = float(self.format_value_two)
            data_3 = float(self.format_value_three)

            self.worksheet.write(0, 0, 'Time')
            self.worksheet.set_column('A:A', 20)
            self.worksheet.write(0, 1, 'Value 1')
            self.worksheet.write(0, 2, 'Value 2')
            self.worksheet.write(0, 3, 'Value 3')

            self.worksheet.write(self.time, 0, current_time)
            self.worksheet.write(self.time, 1, data_1)
            self.worksheet.write(self.time, 2, data_2)
            self.worksheet.write(self.time, 3, data_3)



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
        self.inputchannel_One.grid(row=2, column=5, sticky="S")
        self.inputchannel_One.bind("<Return>", lambda event: self.set_slider_value(1))
        self.sliderchannel_One = ctk.CTkSlider(master=self, width=10, height=200, from_=0, to=100,
                                               progress_color="blue", orientation="vertical")
        self.sliderchannel_One.grid(row=4, column=5)

        self.inputchannel_Two = ctk.CTkEntry(self)
        self.inputchannel_Two.grid(row=2, column=6, sticky="S")
        self.inputchannel_Two.bind("<Return>", lambda event: self.set_slider_value(2))
        self.sliderchannel_Two = ctk.CTkSlider(master=self, width=10, height=200, from_=0, to=100,
                                               progress_color="blue", orientation="vertical")
        self.sliderchannel_Two.grid(row=4, column=6)

        self.inputchannel_Three = ctk.CTkEntry(self)
        self.inputchannel_Three.grid(row=2, column=7, sticky="S")
        self.inputchannel_Three.bind("<Return>", lambda event: self.set_slider_value(3))
        self.sliderchannel_Three = ctk.CTkSlider(master=self, width=10, height=200, from_=0, to=100,
                                               progress_color="blue", orientation="vertical")
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

    def current_time(self):
        current_time = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        return current_time

    def loop(self):
        if self.status_start_stop:
            data_1 = float(self.format_value_one)
            self.Masterchart.show_data(data=[data_1], line=self.Drawline_1)

            data_2 = float(self.format_value_two)
            self.Masterchart.show_data(data=[data_2], line=self.Drawline_2)

            data_3 = float(self.format_value_three)
            self.Masterchart.show_data(data=[data_3], line=self.Drawline_3)
            self.write_excel()
        self.after(self.value_steps*1000, self.loop)

    def open_settings(self):
        self.withdraw()
        app = Configuration.ConfigurationApp()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def open_script(self):
        self.withdraw()
        app = script.app()
        app.protocol("WM_DELETE_WINDOW", app.destroy)
        app.mainloop()

    def restart(self):
        super().destroy()
        restart = App()
        restart.mainloop()

    def destroy(self):
        self.workbook.close()
        super().destroy()
        app = main.MainApp()
        app.mainloop()